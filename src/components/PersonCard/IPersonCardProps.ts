import { IUserInfo } from "../../models";
import { SPFI } from "@pnp/sp";
import { IGraphPhotoClient } from "../../services/PhotoService";

export interface IPersonCardProps {
  userInfo: IUserInfo;
  onUserSelected: (user: IUserInfo) => void;
  selectedUser?: IUserInfo;
  showActionsBar?: boolean;
  graphClient?: IGraphPhotoClient;
  sp: SPFI;
}