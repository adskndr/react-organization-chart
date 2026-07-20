import { SPFI } from "@pnp/sp";
import "@pnp/sp/search";

// Well-known, out-of-the-box "People Results" result source ID in SharePoint Online.
// Restricting the search to this source keeps the refiner values limited to people
// (i.e. actual "Department" values from user profiles) instead of document metadata.
const PEOPLE_RESULTS_SOURCE_ID = "b09a7990-05ea-4af9-81ef-edfab16c4e31";

/**
 * Loads the distinct list of "Department" values currently in use across the tenant,
 * using the SharePoint Search refiner for the "Department" managed property.
 */
export const getDepartments = async (sp: SPFI): Promise<string[]> => {
  try {
    const results = await sp.search({
      Querytext: "*",
      SourceId: PEOPLE_RESULTS_SOURCE_ID,
      RowLimit: 1,
      TrimDuplicates: false,
      Refiners: "Department",
      SelectProperties: ["Department"],
    });

    const refiners =
      results?.RawSearchResults?.PrimaryQueryResult?.RefinementResults
        ?.Refiners;
    const departmentRefiner = refiners?.filter(
      (refiner) => refiner.Name === "Department"
    )[0];

    if (!departmentRefiner || !departmentRefiner.Entries) {
      return [];
    }

    return departmentRefiner.Entries.map((entry) => entry.RefinementValue)
      .filter((value) => !!value)
      .sort((a, b) => a.localeCompare(b));
  } catch (error) {
    console.log(error);
    return [];
  }
};
