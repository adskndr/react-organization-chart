import { ResponseType } from "@microsoft/microsoft-graph-client";

// A minimal structural type for what we actually use from MSGraphClientV3.
// Importing the real MSGraphClientV3 type here runs into duplicate nested
// copies of @microsoft/sp-http-msgraph in node_modules (one under
// @microsoft/sp-component-base, one at the top level), which TypeScript
// treats as incompatible nominal types even though they're identical in
// shape. Using a small structural interface sidesteps that entirely.
export interface IGraphPhotoClient {
  api(path: string): {
    responseType(type: ResponseType): { get(): Promise<Blob> };
  };
}

// Caches resolved object URLs (and known-missing photos) per user so we never
// fetch or request the same person's photo twice in a session.
const photoCache = new Map<string, string | undefined>();

/**
 * Loads a user's profile photo directly from Entra ID via Microsoft Graph
 * (`/users/{id}/photo/$value`) and returns a local object URL for it.
 * Returns undefined if the user has no photo, or the call fails (e.g. missing
 * "User.Read.All" Graph permission — see package-solution.json).
 */
export const getUserPhotoUrl = async (
  graphClient: IGraphPhotoClient | undefined,
  userPrincipalNameOrEmail: string | undefined
): Promise<string | undefined> => {
  if (!graphClient || !userPrincipalNameOrEmail) return undefined;

  if (photoCache.has(userPrincipalNameOrEmail)) {
    return photoCache.get(userPrincipalNameOrEmail);
  }

  try {
    const photoBlob: Blob = await graphClient
      .api(`/users/${encodeURIComponent(userPrincipalNameOrEmail)}/photo/$value`)
      .responseType(ResponseType.BLOB)
      .get();

    const objectUrl = URL.createObjectURL(photoBlob);
    photoCache.set(userPrincipalNameOrEmail, objectUrl);
    return objectUrl;
  } catch (error) {
    // No photo set, no permission, or user not found — cache as "no photo"
    // so we don't keep retrying (and keep showing initials instead).
    photoCache.set(userPrincipalNameOrEmail, undefined);
    return undefined;
  }
};