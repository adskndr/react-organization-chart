import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

export interface IOrgChartProps {
  title: string;
  defaultUser: string;
  context: WebPartContext;
  startFromUser: IPropertyFieldGroupOrPerson[];
  coLeadUser?: IPropertyFieldGroupOrPerson[];
  showAllManagers: boolean;
  showGuestUsers: boolean;
  showActionsBar: boolean;
  showPeers?: boolean;
  departmentFilterSelected?: string[];
  departmentFilterText?: string;
  showTitle?: boolean;
  titleHeadingLevel?: string;
  titleFontSize?: number;
  sp: SPFI;
}