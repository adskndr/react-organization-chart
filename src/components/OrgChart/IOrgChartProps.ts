import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { IGraphPhotoClient } from "../../services/PhotoService";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

export interface IOrgChartProps {
  defaultUser: string;
  context: WebPartContext;
  startFromUser: IPropertyFieldGroupOrPerson[];
  coLeadUser?: IPropertyFieldGroupOrPerson[];
  managerLevels?: number;
  showActionsBar: boolean;
  showPeers?: boolean;
  departmentFilterSelected?: string[];
  departmentFilterText?: string;

  /**
   * Optional JobTitle / Position filter.
   *
   * Example:
   * "Lernende"
   */
  jobTitleFilterText?: string;

  /**
   * Optional JobTitle / Position exclude filter — comma/semicolon separated.
   * People whose JobTitle matches one of these values (same loose,
   * gender-insensitive, partial match as jobTitleFilterText) are hidden
   * from the chart, regardless of which other filters are active.
   *
   * Example:
   * "Lernende, Praktikant"
   */
  jobTitleExcludeFilterText?: string;

  graphClient?: IGraphPhotoClient;
  sp: SPFI;
}