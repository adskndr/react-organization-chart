import { IUserInfo } from "../../models";
import { SPFI } from "@pnp/sp";
import { ServiceScope } from "@microsoft/sp-core-library";
import { IGraphPhotoClient } from "../../services/PhotoService";

export interface IPersonCardProps {
  userInfo: IUserInfo;
  onUserSelected: (user: IUserInfo) => void;
  selectedUser?: IUserInfo;
  showActionsBar?: boolean;
  graphClient?: IGraphPhotoClient;
  serviceScope: ServiceScope;
  sp: SPFI;
}