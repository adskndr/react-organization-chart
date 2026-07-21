import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { IGraphPhotoClient } from "../../services/PhotoService";
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

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
  graphClient?: IGraphPhotoClient;
  sp: SPFI;
}