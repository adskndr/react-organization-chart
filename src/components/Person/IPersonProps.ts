import { PersonaSize } from "@fluentui/react/lib/Persona";
import { IGraphPhotoClient } from "../../services/PhotoService";
export interface IPersonProps {
  userEmail: string;
  text: string;
  secondaryText: string;
  tertiaryText?: string;
  pictureUrl?:string;
  graphClient?: IGraphPhotoClient;
  size?: PersonaSize;
}