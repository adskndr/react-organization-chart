/* eslint-disable @typescript-eslint/no-non-null-assertion */
import { SPFI } from "@pnp/sp";
import "@pnp/sp/profiles";
import "@pnp/sp/batching";
import { IUserInfo } from "../models/IUserInfo";
import * as React from "react";
import { get, set } from "idb-keyval";
import { sortBy, filter } from "lodash";
import { IPersonProperties } from "../models/IPersonProperties";

/*************************************************************************************/
// Hook to get user profile information
// *************************************************************************************/

type getUserProfileFunc = (
  sp: SPFI,
  currentUser: string,
  managerLevels?: number
) => Promise<ProfileDataResponse>;

type ProfileDataResponse = Maybe<{
  managersList: IUserInfo[];
  reportsLists: IUserInfo[];
  peersList: IUserInfo[];
  currentUserProfile: IPersonProperties;
}>;

export const useGetUserProperties = (): {
  getUserProfile: getUserProfileFunc;
} => {
  const getUserProfile = React.useCallback(
    async (
      sp: SPFI,
      currentUser: string,
      managerLevels: number = 0
    ): Promise<ProfileDataResponse> => {
      if (!currentUser) return;
      const loginName = currentUser;
      const cacheCurrentUser: Maybe<IPersonProperties> = await get(
        `${loginName}__orgchart__`
      );
      let currentUserProfile: IPersonProperties;
      if (!cacheCurrentUser) {
        currentUserProfile = await sp.profiles.getPropertiesFor(loginName);
        // console.log(currentUserProfile);
        await set(`${loginName}__orgchart__`, currentUserProfile);
      } else {
        currentUserProfile = cacheCurrentUser;
      }
      // get Managers and Direct Reports
      let reportsLists: IUserInfo[] = [];
      let managersList: IUserInfo[] = [];
      let peersList: IUserInfo[] = [];

      const wDirectReports: Maybe<string[]> =
        currentUserProfile && currentUserProfile.DirectReports;
      const wExtendedManagers: Maybe<string[]> =
        currentUserProfile && currentUserProfile.ExtendedManagers;
      const wPeers: Maybe<string[]> =
        currentUserProfile && currentUserProfile.Peers;

      // Get Direct Reports if exists
      if (wDirectReports && wDirectReports.length > 0) {
        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        reportsLists = await getDirectReports(sp, wDirectReports);
      }
      // Get Peers if exists (colleagues reporting to the same manager)
      if (wPeers && wPeers.length > 0) {
        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        peersList = await getPeers(sp, wPeers);
      }
      // Get Managers if exists — ExtendedManagers is ordered starting with the
      // top of the hierarchy down to the immediate manager, so the last
      // `managerLevels` entries are exactly "N levels up" from this person.
      if (managerLevels > 0 && wExtendedManagers && wExtendedManagers.length > 0) {
        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        managersList = await getExtendedManagers(
          sp,
          wExtendedManagers,
          managerLevels
        );
      }

      return { managersList, reportsLists, peersList, currentUserProfile };
    },
    []
  );

  return { getUserProfile };
};

const getDirectReports = async (
  sp: SPFI,
  directReports: string[]
): Promise<IUserInfo[]> => {
  const _reportsList: IUserInfo[] = [];
  const [batchedSP, execute] = sp.batched();
  
  for (const userReport of directReports) {
    const cacheDirectReport: Maybe<IPersonProperties> = await get(
      `${userReport}__orgchart__`
    );
    if (!cacheDirectReport) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      batchedSP.profiles
        .getPropertiesFor(userReport)
        .then(async (directReport: IPersonProperties) => {
          // eslint-disable-next-line @typescript-eslint/no-use-before-define
          const userInfo = await manpingUserProperties(directReport);
          _reportsList.push(userInfo);
          await set(`${userReport}__orgchart__`, directReport);
        });
    } else {
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
      const userInfo = await manpingUserProperties(cacheDirectReport);
      _reportsList.push(userInfo);
    }
  }
  await execute();
  return sortBy(_reportsList, ["displayName"]);
};

const getPeers = async (
  sp: SPFI,
  peers: string[]
): Promise<IUserInfo[]> => {
  const _peersList: IUserInfo[] = [];
  const [batchedSP, execute] = sp.batched();

  for (const peer of peers) {
    const cachePeer: Maybe<IPersonProperties> = await get(
      `${peer}__orgchart__`
    );
    if (!cachePeer) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      batchedSP.profiles
        .getPropertiesFor(peer)
        .then(async (peerProfile: IPersonProperties) => {
          // eslint-disable-next-line @typescript-eslint/no-use-before-define
          const userInfo = await manpingUserProperties(peerProfile);
          _peersList.push(userInfo);
          await set(`${peer}__orgchart__`, peerProfile);
        });
    } else {
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
      const userInfo = await manpingUserProperties(cachePeer);
      _peersList.push(userInfo);
    }
  }
  await execute();
  return sortBy(_peersList, ["displayName"]);
};

const getExtendedManagers = async (
  sp: SPFI,
  extendedManagers: string[],
  managerLevels: number
): Promise<IUserInfo[]> => {
  const wManagers: IUserInfo[] = [];
  // ExtendedManagers is ordered starting with the TOP of the hierarchy down
  // to the immediate manager (last entry = nearest manager). So "N levels
  // above" means the last N entries, kept in that same (top-to-near) order —
  // which is exactly the order we want for top-to-bottom rendering.
  const levelsToFetch =
    managerLevels > 0 ? extendedManagers.slice(-managerLevels) : [];
  const [batchedSP, execute] = sp.batched();

  for (const manager of levelsToFetch) {
    const cacheManager: Maybe<IPersonProperties> = await get(
      `${manager}__orgchart__`
    );
    if (!cacheManager) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      batchedSP.profiles
        .getPropertiesFor(manager)
        .then(async (_profile: IPersonProperties) => {
          // eslint-disable-next-line @typescript-eslint/no-use-before-define
          const userInfo = await manpingUserProperties(_profile);
          wManagers.push(userInfo);
          await set(`${manager}__orgchart__`, _profile);
        });
    } else {
      // eslint-disable-next-line @typescript-eslint/no-use-before-define
      const userInfo = await manpingUserProperties(cacheManager);
      wManagers.push(userInfo);
    }
  }
  await execute();
  // Batched/cached lookups can resolve out of order, so re-sort by the
  // original chain order before returning ("top level first").
  const chainOrder = new Map(levelsToFetch.map((login, idx) => [login, idx]));
  return wManagers.sort(
    (a, b) => (chainOrder.get(a.id ?? "") ?? 0) - (chainOrder.get(b.id ?? "") ?? 0)
  );
};

// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
function userTypeMapper(userType: string | undefined) {
  switch (userType) {
    case "0":
      return "Employee";
    case "1":
      return "Guest";
    default:
      return "Unknown";
  }
}

export const manpingUserProperties = async (
  userProperties: IPersonProperties
): Promise<IUserInfo> => {
  console.log(userProperties);

  return {
    displayName: userProperties.DisplayName as string,
    email: userProperties.Email as string,
    title: userProperties.Title as string,
    pictureUrl: userProperties.PictureUrl,
    id: userProperties.AccountName,
    userUrl: userProperties.UserUrl,
    numberDirectReports: userProperties.DirectReports.length,
    hasDirectReports: userProperties.DirectReports.length > 0 ? true : false,
    hasPeers: userProperties.Peers.length > 0 ? true : false,
    numberPeers: userProperties.Peers.length,
    department:
      filter(userProperties?.UserProfileProperties, { Key: "Department" })[0]
        ?.Value ?? "",
    workPhone:
      filter(userProperties?.UserProfileProperties, { Key: "WorkPhone" })[0]
        ?.Value ?? "",
    cellPhone:
      filter(userProperties?.UserProfileProperties, { Key: "CellPhone" })[0]
        ?.Value ?? "",
    location:
      filter(userProperties?.UserProfileProperties, { Key: "SPS-Location" })[0]
        ?.Value ?? "",
    office:
      filter(userProperties?.UserProfileProperties, { Key: "Office" })[0]
        ?.Value ?? "",
    manager:
      filter(userProperties?.UserProfileProperties, { Key: "Manager" })[0]
        ?.Value ?? "",
    userType: userTypeMapper(
      filter(userProperties?.UserProfileProperties, { Key: "SPS-UserType" })[0]
        ?.Value
    ),
    loginName: userProperties.loginName,
  };
};