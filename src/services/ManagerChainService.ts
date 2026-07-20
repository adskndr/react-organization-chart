import { SPFI } from "@pnp/sp";
import "@pnp/sp/profiles";

/**
 * Returns how many levels of managers actually exist above the given person
 * (i.e. the length of their ExtendedManagers chain), so the "manager levels"
 * property pane field can be capped to a value that makes sense — no point
 * letting an admin dial in "7" if there are only 3 levels above this person.
 */
export const getManagerChainLength = async (
  sp: SPFI,
  loginName: string
): Promise<number> => {
  if (!loginName) return 0;
  try {
    const profile = await sp.profiles.getPropertiesFor(loginName);
    return profile?.ExtendedManagers?.length ?? 0;
  } catch (error) {
    console.log(error);
    return 0;
  }
};
