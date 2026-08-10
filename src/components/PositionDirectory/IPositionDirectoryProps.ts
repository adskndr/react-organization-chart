import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { IGraphPhotoClient } from "../../services/PhotoService";

export interface IPositionDirectoryProps {
  context: WebPartContext;
  // e.g. ["Lernende", "Lernender"] — everyone with a matching JobTitle shows
  // up here, no matter who their manager is (or whether they have one).
  jobTitles: string[];
  // Clusters the grid under a heading per Department, instead of one flat grid.
  groupByDepartment: boolean;
  showActionsBar: boolean;
  graphClient?: IGraphPhotoClient;
  sp: SPFI;
}
