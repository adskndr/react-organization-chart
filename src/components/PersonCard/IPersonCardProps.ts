import { IUserInfo } from "../../models";
import { SPFI } from "@pnp/sp";
import { IGraphPhotoClient } from "../../services/PhotoService";

export interface IPersonCardProps {
  userInfo: IUserInfo;
  onUserSelected: (user: IUserInfo) => void;
  selectedUser?: IUserInfo;
  showActionsBar?: boolean;
  graphClient?: IGraphPhotoClient;
  // Deliberately untyped (not `ServiceScope` from @microsoft/sp-core-library):
  // in a real SPFx node_modules tree there can be more than one physical
  // copy of that package (one top-level, one nested under
  // @microsoft/sp-loader, sometimes another under @pnp/spfx-controls-react's
  // own dependency chain). TypeScript treats those as separate nominal types
  // because of a private field on the class, so importing our own copy here
  // just to type this prop causes a build error the moment it's passed into
  // LivePersona's own (differently-sourced) ServiceScope type. It's cast
  // back at the one place it's actually used — see PersonCard.tsx.
  serviceScope: unknown;
  sp: SPFI;
}