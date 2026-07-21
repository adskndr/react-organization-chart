import * as React from "react";
import { IOrgChartProps } from "./IOrgChartProps";
import { IOrgChartState } from "./IOrgChartState";
import { OrgChartReducer } from "./OrgChartReducer";
import {
  useGetUserProperties,
  manpingUserProperties,
} from "../../hooks/useGetUserProperties";
import { Stack } from "@fluentui/react/lib/Stack";
import { PersonCard } from "../PersonCard/PersonCard";
import { IUserInfo } from "../../models/IUserInfo";
import { EOrgChartTypes } from "./EOrgChartTypes";
import {
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Text,
} from "@fluentui/react";

import { useOrgChartStyles } from "./useOrgChartStyles";
import { useWrappedContentWidth } from "../../hooks/useWrappedContentWidth";

import "./OrgChart.module.scss";
import { Placeholder } from "../Placeholder/PlaceholderComponent";


const initialState: IOrgChartState = {
  isLoading: true,
  renderDirectReports: [],
  renderManagers: [],
  renderPeers: [],
  error: undefined,
  currentUser: undefined,
  coLeadUser: undefined,
};

export const OrgChart: React.FunctionComponent<IOrgChartProps> = (
  props: IOrgChartProps
) => {
  const { getUserProfile } = useGetUserProperties();
  const [state, dispatch] = React.useReducer(OrgChartReducer, initialState);
  const { orgChartClasses } = useOrgChartStyles();

  const {
    renderManagers,
    renderDirectReports,
    renderPeers,
    currentUser,
    coLeadUser,
    isLoading,
    error,
  }: IOrgChartState = state;

  // CSS alone can't shrink-wrap a wrapping flex box to "the width of the
  // widest row it actually wrapped into" (see useWrappedContentWidth) — so
  // we measure it and apply it as an explicit width below, for both boxes
  // that can hold more than one card side by side.
  const [teamBoxRef, teamBoxWidth] = useWrappedContentWidth<HTMLDivElement>([
    renderDirectReports,
  ]);
  const [leadershipBoxRef, leadershipBoxWidth] = useWrappedContentWidth<HTMLDivElement>(
    [renderPeers, currentUser, coLeadUser]
  );

  const {
    context,
    managerLevels,
    startFromUser,
    coLeadUser: coLeadUserPicker,
    showActionsBar,
    showPeers,
    departmentFilterSelected,
    departmentFilterText,
    graphClient,
    sp,
  }: IOrgChartProps = props;


  const startFromUserId: Maybe<string> = React.useMemo(
    () => startFromUser && startFromUser[0] && startFromUser[0].id,
    [startFromUser]
  );
  const coLeadUserId: Maybe<string> = React.useMemo(() => {
    const id =
      coLeadUserPicker && coLeadUserPicker[0] && coLeadUserPicker[0].id;
    // If someone accidentally picks the same person as both "start from
    // user" and "co-lead", ignore the co-lead rather than showing the same
    // person twice side by side.
    return id && id !== startFromUserId ? id : undefined;
  }, [coLeadUserPicker, startFromUserId]);
  const onUserSelected = React.useCallback((selectedUser: IUserInfo) => {
    dispatch({
      type: EOrgChartTypes.SET_CURRENT_USER,
      payload: selectedUser,
    });
  }, []);

  const matchesDepartmentFilter = React.useCallback(
    (department?: string): boolean => {
      const wSelected: string[] = (departmentFilterSelected ?? []).map((d) =>
        d.trim().toLowerCase()
      );
      const wText: string = (departmentFilterText ?? "").trim().toLowerCase();

      if (wSelected.length === 0 && !wText) return true;

      const wDepartment: string = (department ?? "").trim().toLowerCase();
      if (wSelected.length > 0 && wSelected.indexOf(wDepartment) > -1) {
        return true;
      }
      if (wText && wDepartment.indexOf(wText) > -1) {
        return true;
      }
      return false;
    },
    [departmentFilterSelected, departmentFilterText]
  );

  const isDeputy = React.useCallback((user: IUserInfo): boolean => {
    return /\bstv\.?/i.test(user.title ?? "");
  }, []);

  const isCoLead = React.useCallback((user: IUserInfo): boolean => {
    return /\bco[-\s]?\w*(leiter|leitung|lead)\b/i.test(user.title ?? "");
  }, []);

  const sortReportsPriority = React.useCallback(
    (users: IUserInfo[]): IUserInfo[] => {
      const rank = (user: IUserInfo): number => {
        if (isDeputy(user)) return 0;
        if (isCoLead(user)) return 1;
        return 2;
      };
      return [...users].sort((a, b) => {
        const rankDiff = rank(a) - rank(b);
        if (rankDiff !== 0) return rankDiff;
        return (a.displayName ?? "").localeCompare(b.displayName ?? "");
      });
    },
    [isDeputy, isCoLead]
  );

  const isDepartmentFilterActive: boolean =
    (departmentFilterSelected ?? []).length > 0 ||
    !!(departmentFilterText ?? "").trim();

  const loadOrgChart = React.useCallback(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    async (selectedUser: string): Promise<any> => {
      const wRenderDirectReports: JSX.Element[] = [];
      const wRenderPeers: JSX.Element[] = [];

      try {
        const profileResponse = await getUserProfile(
          sp,
          selectedUser,
          0,
        );
        if (profileResponse) {
          // Merge direct reports of the main lead and, if configured and we're
          // at the top of the department, the co-lead's direct reports too.
          const combinedReportsMap = new Map<string, IUserInfo>();
          profileResponse.reportsLists.forEach((report) => {
            combinedReportsMap.set(report.id ?? report.email, report);
          });

          if (coLeadUserId && selectedUser === startFromUserId) {
            const coLeadProfileResponse = await getUserProfile(
              sp,
              coLeadUserId
            );
            coLeadProfileResponse?.reportsLists.forEach((report) => {
              combinedReportsMap.set(report.id ?? report.email, report);
            });
          }

          const filteredDirectReports = sortReportsPriority(
            Array.from(combinedReportsMap.values()).filter((report) =>
              matchesDepartmentFilter(report.department)
            )
          );
          const filteredPeers = profileResponse.peersList.filter((peer) =>
            matchesDepartmentFilter(peer.department)
          );

          if (showPeers === true) {
            for (const peerInfo of filteredPeers) {
              wRenderPeers.push(
                <PersonCard
                  key={`peer-${peerInfo.id}`}
                  userInfo={peerInfo}
                  onUserSelected={onUserSelected}
                  selectedUser={currentUser}
                  showActionsBar={showActionsBar}
                  graphClient={graphClient}
                  serviceScope={context.serviceScope}
                sp={sp}
                 />
              );
            }
          }

          let reportIndex = 0;
          while (reportIndex < filteredDirectReports.length) {
            const directReport = filteredDirectReports[reportIndex];

            if (isCoLead(directReport)) {
              // Collect all consecutive co-leads (already sorted next to each
              // other by sortReportsPriority) into one visual group.
              const coLeadGroup: IUserInfo[] = [directReport];
              let lookahead = reportIndex + 1;
              while (
                lookahead < filteredDirectReports.length &&
                isCoLead(filteredDirectReports[lookahead])
              ) {
                coLeadGroup.push(filteredDirectReports[lookahead]);
                lookahead++;
              }

              if (coLeadGroup.length > 1) {
                wRenderDirectReports.push(
                  <Stack
                    key={`co-lead-group-${coLeadGroup[0].id}`}
                    horizontal
                    verticalAlign="center"
                    tokens={{ childrenGap: 8 }}
                    className={orgChartClasses.coLeadGroup}
                  >
                    {coLeadGroup.map((coLead) => (
                      <PersonCard
                        key={`report-${coLead.id}`}
                        userInfo={coLead}
                        onUserSelected={onUserSelected}
                        selectedUser={currentUser}
                        showActionsBar={showActionsBar}
                        graphClient={graphClient}
                        serviceScope={context.serviceScope}
                sp={sp}
                       />
                    ))}
                  </Stack>
                );
                reportIndex = lookahead;
                continue;
              }
            }

            wRenderDirectReports.push(
              <PersonCard
                key={`report-${directReport.id}`}
                userInfo={directReport}
                onUserSelected={onUserSelected}
                selectedUser={currentUser}
                showActionsBar={showActionsBar}
                graphClient={graphClient}
                serviceScope={context.serviceScope}
                sp={sp}
               />
            );
            reportIndex++;
          }
        }

        dispatch({
          type: EOrgChartTypes.SET_HAS_ERROR,
          payload: { hasError: false, errorMessage: "" },
        });
      } catch (error) {
        console.log(error);
        dispatch({
          type: EOrgChartTypes.SET_IS_LOADING,
          payload: false,
        });
        dispatch({
          type: EOrgChartTypes.SET_HAS_ERROR,
          payload: {
            hasError: true,
            errorMessage: "error",
          },
        });
      }

      return { wRenderDirectReports, wRenderPeers };
    },

    [
      sp,
      getUserProfile,
      startFromUserId,
      coLeadUserId,
      onUserSelected,
      currentUser,
      showActionsBar,
      showPeers,
      graphClient,
      matchesDepartmentFilter,
      sortReportsPriority,
      isCoLead,
      orgChartClasses.coLeadGroup,
    ]
  );

  React.useEffect(() => {
    let cancelled = false;

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    (async () => {
      try {
        if (startFromUserId === undefined) {
          // The person was removed from "Start from user" — clear out
          // whatever was shown before instead of leaving it stale.
          dispatch({ type: EOrgChartTypes.SET_CURRENT_USER, payload: undefined });
          dispatch({ type: EOrgChartTypes.SET_RENDER_MANAGERS, payload: [] });
          dispatch({ type: EOrgChartTypes.SET_RENDER_DIRECT_REPORTS, payload: [] });
          dispatch({ type: EOrgChartTypes.SET_RENDER_PEERS, payload: [] });
          dispatch({ type: EOrgChartTypes.SET_IS_LOADING, payload: false });
          return;
        }
        if (startFromUserId === ''){
          dispatch({
            type: EOrgChartTypes.SET_IS_LOADING,
            payload: false,
          });
          dispatch({
            type: EOrgChartTypes.SET_HAS_ERROR,
            payload: {
              hasError: true,
              errorMessage: "User don't have email defined",
            },
          });
          return;
        }
        const profileResponse = await getUserProfile(
          sp,
          startFromUserId,
          managerLevels,
          false // reports/peers aren't needed here — loadOrgChart fetches those
        );
        if (cancelled) return;

        const wCurrentUser: IUserInfo = await manpingUserProperties(
          // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
          profileResponse!.currentUserProfile
        );
        if (cancelled) return;

        dispatch({
          type: EOrgChartTypes.SET_CURRENT_USER,
          payload: wCurrentUser,
        });

        // The managers row is always anchored to the configured "start from
        // user" — level 1 is their direct manager, level 2 the next one up,
        // and so on — regardless of which card is currently being browsed.
        // getExtendedManagers already returns them top-to-nearest, which is
        // exactly the order we want top-to-bottom on screen. Each manager
        // gets their own box (same size, same style) connected by a short
        // vertical line — this stays true whether there's 1 level or several.
        const managersList = profileResponse!.managersList; // eslint-disable-line @typescript-eslint/no-non-null-assertion
        const wRenderManagers: JSX.Element[] = [];
        managersList.forEach((managerInfo, index) => {
          wRenderManagers.push(
            <div
              key={`manager-box-${managerInfo.id}`}
              className={orgChartClasses.managerBox}
            >
              <PersonCard
                userInfo={managerInfo}
                onUserSelected={onUserSelected}
                selectedUser={currentUser}
                showActionsBar={showActionsBar}
                graphClient={graphClient}
                serviceScope={context.serviceScope}
                sp={sp}
               />
            </div>
          );
          if (index < managersList.length - 1) {
            wRenderManagers.push(
              <div
                key={`manager-connector-${managerInfo.id}`}
                className={orgChartClasses.boxConnector}
               />
            );
          }
        });
        dispatch({
          type: EOrgChartTypes.SET_RENDER_MANAGERS,
          payload: wRenderManagers,
        });

        dispatch({
          type: EOrgChartTypes.SET_HAS_ERROR,
          payload: { hasError: false, errorMessage: "" },
        });
      } catch (error) {
        if (cancelled) return;
        console.log(error);
        dispatch({
          type: EOrgChartTypes.SET_IS_LOADING,
          payload: false,
        });
        dispatch({
          type: EOrgChartTypes.SET_HAS_ERROR,
          payload: {
            hasError: true,
            errorMessage: "error",
          },
        });
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [
    getUserProfile,
    sp,
    startFromUserId,
    managerLevels,
    showActionsBar,
    graphClient,
  ]);


  React.useEffect(() => {
    let cancelled = false;

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    (async () => {
      if (!coLeadUserId) {
        dispatch({ type: EOrgChartTypes.SET_CO_LEAD_USER, payload: undefined });
        return;
      }
      try {
        const profileResponse = await getUserProfile(sp, coLeadUserId);
        if (cancelled) return;

        const wCoLeadUser: IUserInfo = await manpingUserProperties(
          // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
          profileResponse!.currentUserProfile
        );
        if (cancelled) return;

        dispatch({
          type: EOrgChartTypes.SET_CO_LEAD_USER,
          payload: wCoLeadUser,
        });
      } catch (error) {
        if (cancelled) return;
        console.log(error);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [getUserProfile, sp, coLeadUserId]);

  React.useEffect(() => {
    let cancelled = false;

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    (async () => {
      if (!currentUser || !currentUser.id) return;
      dispatch({
        type: EOrgChartTypes.SET_IS_LOADING,
        payload: true,
      });

      const { wRenderDirectReports, wRenderPeers } = await loadOrgChart(
        currentUser.id
      );
      if (cancelled) return;

      dispatch({
        type: EOrgChartTypes.SET_RENDER_DIRECT_REPORTS,
        payload: wRenderDirectReports,
      });
      dispatch({
        type: EOrgChartTypes.SET_RENDER_PEERS,
        payload: wRenderPeers,
      });
      dispatch({
        type: EOrgChartTypes.SET_IS_LOADING,
        payload: false,
      });
    })();

    return () => {
      cancelled = true;
    };
  }, [currentUser, loadOrgChart]);

  if (!startFromUserId) {
    return (
      <Placeholder
      iconName="Edit"
      iconText="Configure your Organization Chart Web Part"
      description={"Please configure web part"}
      buttonLabel="Configure"
      onConfigure={context.propertyPane.open}
    />
     
    );
  }


  if (isLoading) {
    return (
      <Stack style={{ minHeight: 200 }} verticalAlign="center" horizontalAlign="center">
        <Spinner
          size={SpinnerSize.large}
          label={"loading Organization Chart..."}
          labelPosition={"bottom"}
         />
      </Stack>
    );
  }

  if (error && error.hasError) {
    return (
      <Stack
        horizontal
        horizontalAlign="center"
        styles={{root:{padding: 20}}}
        tokens={{ childrenGap: 10 }}
      >
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {error.errorMessage}
        </MessageBar>
      </Stack>
    );
  }

  return (
    <>
      <Stack  styles={{root:{padding: 20}}} >
        <Stack horizontalAlign="center" verticalAlign="center">
          {renderManagers.length > 0 && (
            <>
              <Stack horizontalAlign="center" verticalAlign="center">
                {renderManagers}
              </Stack>
              <div className={orgChartClasses.boxConnector} />
            </>
          )}
          <div
            ref={leadershipBoxRef}
            className={orgChartClasses.leadershipBox}
            style={leadershipBoxWidth !== undefined ? { width: leadershipBoxWidth } : undefined}
          >
            {renderPeers}
            <div className={orgChartClasses.leadershipGroup}>
              <PersonCard
                key={`current-${currentUser?.id}`}
                // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
                userInfo={currentUser!}
                onUserSelected={onUserSelected}
                selectedUser={currentUser}
                showActionsBar={showActionsBar}
                graphClient={graphClient}
                serviceScope={context.serviceScope}
                sp={sp}
               />
              {coLeadUser && (
                <PersonCard
                  key={`co-lead-${coLeadUser.id}`}
                  userInfo={coLeadUser}
                  onUserSelected={onUserSelected}
                  selectedUser={currentUser}
                  showActionsBar={showActionsBar}
                  graphClient={graphClient}
                  serviceScope={context.serviceScope}
                  sp={sp}
                 />
              )}
            </div>
          </div>
        </Stack>
        {renderDirectReports.length > 0 && (
          <Stack horizontalAlign="center">
            <div className={orgChartClasses.boxConnector} />
          </Stack>
        )}
        {isDepartmentFilterActive && renderDirectReports.length === 0 && (
          <Stack horizontal horizontalAlign="center" styles={{ root: { padding: 10 } }}>
            <Text variant="medium">
              No direct reports found for the selected department filter.
            </Text>
          </Stack>
        )}
        {renderDirectReports.length > 0 && (
          <div
            ref={teamBoxRef}
            className={orgChartClasses.teamBox}
            style={teamBoxWidth !== undefined ? { width: teamBoxWidth } : undefined}
          >
            {renderDirectReports}
          </div>
        )}
      </Stack>
    </>
  );
};