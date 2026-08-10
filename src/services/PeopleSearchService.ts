import { SPFI } from "@pnp/sp";
import "@pnp/sp/search";
import { IUserInfo } from "../models/IUserInfo";

// Same out-of-the-box "People Results" source already used in
// DepartmentService.ts — restricts results to actual people (user profiles),
// not documents/sites.
const PEOPLE_RESULTS_SOURCE_ID = "b09a7990-05ea-4af9-81ef-edfab16c4e31";

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

  // Exact-match each title, OR'd together:
  // JobTitle:"Lernende" OR JobTitle:"Lernender"
  const queryText = wJobTitles
    .map((title) => `JobTitle:"${title.replace(/"/g, '\\"')}"`)
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
