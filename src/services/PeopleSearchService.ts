import { SPFI } from "@pnp/sp";
import "@pnp/sp/search";
import { IUserInfo } from "../models/IUserInfo";

// Same out-of-the-box "People Results" source already used in
// DepartmentService.ts — restricts results to actual people (user profiles),
// not documents/sites.
const PEOPLE_RESULTS_SOURCE_ID = "b09a7990-05ea-4af9-81ef-edfab16c4e31";

// Minimum length a stem must keep after stripping a gender/number suffix.
// Prevents short words like "Lehre" collapsing into an overly broad "Lehr*".
const MIN_STEM_LENGTH = 4;

// Common German gender/number endings, longest first so "innen" is tried
// before "in" and doesn't get cut short.
const GENDER_SUFFIXES = ["innen", "en", "er", "in", "e"];

/**
 * Strips a common German gender/number ending off a single word, e.g.
 * "Lernende" -> "Lernend", "Lernender" -> "Lernend", "Mitarbeiterin" ->
 * "Mitarbeiter". Used so one typed word matches multiple grammatical forms
 * of the same job title instead of requiring an exact match.
 */
export const stripGenderSuffix = (word: string): string => {
  const wWord = (word ?? "").trim();
  for (const suffix of GENDER_SUFFIXES) {
    if (
      wWord.toLowerCase().endsWith(suffix) &&
      wWord.length - suffix.length >= MIN_STEM_LENGTH
    ) {
      return wWord.slice(0, wWord.length - suffix.length);
    }
  }
  return wWord;
};

/**
 * Builds a JobTitle wildcard phrase for SharePoint KQL from a typed filter
 * term: only the last word gets gender-suffix-stripped and a trailing `*`,
 * so "Junior Lernende" becomes `Junior Lernend*` and matches "Junior
 * Lernender", "Junior Lernenden", etc. Also means the term no longer has to
 * be the complete job title — a leading fragment is enough.
 */
const buildJobTitleWildcardPhrase = (term: string): string => {
  const words = (term ?? "").trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return "";
  const stem = stripGenderSuffix(words[words.length - 1]);
  return [...words.slice(0, -1), stem].join(" ");
};

/**
 * Loose match used everywhere a JobTitle filter term is compared against an
 * actual JobTitle: case-insensitive, substring-based (so a partial title is
 * enough) and gender-suffix-insensitive (so "Lernende" also matches
 * "Lernender"/"Lernenden"). This is the client-side counterpart to the
 * wildcard search query above, used e.g. to re-validate a profile's real
 * JobTitle or to power the exclude filter.
 */
export const jobTitleMatchesTerm = (
  actualJobTitle: string | undefined,
  term: string
): boolean => {
  const wActual = (actualJobTitle ?? "").trim().toLowerCase();
  const wTerm = (term ?? "").trim().toLowerCase();
  if (!wActual || !wTerm) return false;

  if (wActual.indexOf(wTerm) !== -1) return true;

  const stem = buildJobTitleWildcardPhrase(wTerm).toLowerCase();
  return !!stem && wActual.indexOf(stem) !== -1;
};

/**
 * True if a person should be KEPT — i.e. their JobTitle does NOT match any
 * of the given exclude terms (same loose matching as jobTitleMatchesTerm).
 * Empty/undefined excludeTerms never filters anyone out.
 */
export const matchesJobTitleExclude = (
  actualJobTitle: string | undefined,
  excludeTerms: string[]
): boolean => {
  const wExcludeTerms = (excludeTerms ?? [])
    .map((t) => t.trim())
    .filter(Boolean);
  if (wExcludeTerms.length === 0) return true;
  return !wExcludeTerms.some((term) =>
    jobTitleMatchesTerm(actualJobTitle, term)
  );
};

/**
 * Flat, tree-free lookup of everyone whose JobTitle matches one of the given
 * values — e.g. ["Lernende", "Lernender"].
 *
 * Unlike useGetUserProperties (which walks DirectReports / ExtendedManagers
 * starting from one root user), this doesn't care who anyone's manager is.
 * It's a straight SharePoint Search query against the "JobTitle" managed
 * property, so it also finds people who have no manager set at all — which
 * is exactly the case for apprentices who rotate departments.
 */
export const getUsersByJobTitle = async (
  sp: SPFI,
  jobTitles: string[]
): Promise<IUserInfo[]> => {
  const wJobTitles = (jobTitles ?? []).map((t) => t.trim()).filter(Boolean);
  if (wJobTitles.length === 0) return [];

  // Loose, gender-insensitive wildcard match per title, OR'd together, e.g.
  // typing "Lernende" builds JobTitle:"Lernend*" — matches "Lernende",
  // "Lernender", "Lernenden", etc., and doesn't require the full title.
  const queryText = wJobTitles
    .map((title) => {
      const phrase = buildJobTitleWildcardPhrase(title).replace(
        /"/g,
        '\\"'
      );
      return `JobTitle:"${phrase}*"`;
    })
    .join(" OR ");

  try {
    const results = await sp.search({
      Querytext: queryText,
      SourceId: PEOPLE_RESULTS_SOURCE_ID,
      RowLimit: 500,
      TrimDuplicates: true,
      SelectProperties: [
        "PreferredName",
        "WorkEmail",
        "JobTitle",
        "Department",
        "PictureURL",
        "AccountName",
      ],
    });

    // Standard SharePoint Search REST response shape: an array of Cells,
    // each a {Key, Value} pair, per row — same family of object as the
    // RefinementResults used in DepartmentService.ts, just the results
    // table instead of the refiners.
    const rows =
      results?.RawSearchResults?.PrimaryQueryResult?.RelevantResults?.Table
        ?.Rows ?? [];

    const people: IUserInfo[] = rows.map((row) => {
      const cellMap = new Map<string, string>(
        (row.Cells ?? []).map((cell) => [cell.Key, cell.Value])
      );
      return {
        displayName: cellMap.get("PreferredName") ?? "",
        email: cellMap.get("WorkEmail") ?? "",
        title: cellMap.get("JobTitle") ?? "",
        department: cellMap.get("Department") ?? "",
        pictureUrl: cellMap.get("PictureURL") || undefined,
        id: cellMap.get("AccountName") || undefined,
        // Flat directory entries never have direct reports in this view —
        // even if the person technically does, PersonCard's "click to
        // drill in" behavior is harmless here since there's no tree to
        // navigate into from this component.
        hasDirectReports: false,
        userType: "Employee",
      } as IUserInfo;
    });

    return people
      .filter((p) => !!p.displayName)
      .sort(
        (a, b) =>
          (a.department ?? "").localeCompare(b.department ?? "") ||
          a.displayName.localeCompare(b.displayName)
      );
  } catch (error) {
    console.log(error);
    return [];
  }
};

/**
 * Loads the distinct list of "JobTitle" values currently in use across the
 * tenant — same pattern as getDepartments in DepartmentService.ts. Useful to
 * populate a dropdown in the web part's property pane so editors pick an
 * existing job title instead of retyping it (and risking a typo that
 * silently returns zero results).
 */
export const getJobTitles = async (sp: SPFI): Promise<string[]> => {
  try {
    const results = await sp.search({
      Querytext: "*",
      SourceId: PEOPLE_RESULTS_SOURCE_ID,
      RowLimit: 1,
      TrimDuplicates: false,
      Refiners: "JobTitle",
      SelectProperties: ["JobTitle"],
    });

    const refiners =
      results?.RawSearchResults?.PrimaryQueryResult?.RefinementResults
        ?.Refiners;
    const jobTitleRefiner = refiners?.filter(
      (refiner) => refiner.Name === "JobTitle"
    )[0];

    if (!jobTitleRefiner || !jobTitleRefiner.Entries) {
      return [];
    }

    return jobTitleRefiner.Entries.map((entry) => entry.RefinementValue)
      .filter((value) => !!value)
      .sort((a, b) => a.localeCompare(b));
  } catch (error) {
    console.log(error);
    return [];
  }
};

/**
 * Returns people with the specified JobTitle who are somewhere
 * below the selected manager.
 *
 * By default all hierarchy levels are checked. Pass directReportsOnly=true
 * to only match people whose immediate manager is exactly the given
 * manager (i.e. one level down, not the whole subtree).
 */
export const getUsersUnderManagerByJobTitle = async (
  sp: SPFI,
  managerLoginName: string,
  jobTitle: string,
  directReportsOnly: boolean = false
): Promise<IUserInfo[]> => {
  const normalizedJobTitle =
    (jobTitle ?? "").trim();

  const normalizedManager =
    (managerLoginName ?? "").trim();

  // JobTitle is required.
  if (!normalizedJobTitle) {
    return [];
  }

  // No manager selected:
  // return everyone with this JobTitle.
  if (!normalizedManager) {
    return getUsersByJobTitle(
      sp,
      [normalizedJobTitle]
    );
  }

  const normalizedManagerLower =
    normalizedManager.toLowerCase();

  // First find everybody with the requested JobTitle.
  const candidates =
    await getUsersByJobTitle(
      sp,
      [normalizedJobTitle]
    );

  if (candidates.length === 0) {
    return [];
  }

  const result: IUserInfo[] = [];

  const [batchedSP, execute] = sp.batched();

  for (const candidate of candidates) {
    const loginName =
      candidate.id ?? candidate.email;

    if (!loginName) {
      continue;
    }

    // Never show the selected manager himself.
    if (
      loginName.trim().toLowerCase() ===
      normalizedManagerLower
    ) {
      continue;
    }

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    batchedSP.profiles
      .getPropertiesFor(loginName)
      .then((profile) => {
        const managers: string[] =
          ((profile?.ExtendedManagers ?? []) as unknown[])
            .filter(
              (manager): manager is string =>
                typeof manager === "string"
            )
            .map((manager: string) =>
              manager.trim().toLowerCase()
            );

        // ExtendedManagers is ordered top-of-hierarchy -> immediate
        // manager, so the LAST entry is always the person's direct
        // manager. "directReportsOnly" checks exactly that last entry;
        // otherwise any entry in the chain counts (whole subtree).
        const isBelowManager = directReportsOnly
          ? managers[managers.length - 1] === normalizedManagerLower
          : managers.indexOf(normalizedManagerLower) !== -1;

        // Loose match (substring + gender-suffix-insensitive) instead of an
        // exact comparison, consistent with the search query above.
        const hasMatchingJobTitle = jobTitleMatchesTerm(
          profile?.Title,
          normalizedJobTitle
        );

        if (
          isBelowManager &&
          hasMatchingJobTitle
        ) {
          result.push(candidate);
        }
      })
      .catch((error) => {
        console.log(
          `Could not load profile for ${loginName}`,
          error
        );
      });
  }

  await execute();

  return result
    .filter(
      (person, index, array) =>
        array.findIndex(
          (p) =>
            (p.id ?? p.email) ===
            (person.id ?? person.email)
        ) === index
    )
    .sort((a, b) =>
      (a.displayName ?? "").localeCompare(
        b.displayName ?? ""
      )
    );
};