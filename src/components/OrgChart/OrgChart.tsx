import * as React from "react";
import { IOrgChartProps } from "./IOrgChartProps";
import { IOrgChartState } from "./IOrgChartState";
import { OrgChartReducer } from "./OrgChartReducer";
import {
  useGetUserProperties,
  manpingUserProperties,
} from "../../hooks/useGetUserProperties";
import { IStackStyles, Stack } from "@fluentui/react/lib/Stack";
import { PersonCard } from "../PersonCard/PersonCard";
import { IUserInfo } from "../../models/IUserInfo";
import { EOrgChartTypes } from "./EOrgChartTypes";
import {
  MessageBar,
  MessageBarType,
  Overlay,
  Spinner,
  SpinnerSize,
  Text,
} from "@fluentui/react";

import { useOrgChartStyles } from "./useOrgChartStyles";

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

const titleStyle: IStackStyles = {
  root: {
    paddingBottom: 40,
  },
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

  const {
    context,
    showAllManagers,
    showGuestUsers,
    startFromUser,
    coLeadUser: coLeadUserPicker,
    showActionsBar,
    showPeers,
    departmentFilterSelected,
    departmentFilterText,
    showTitle,
    titleHeadingLevel,
    titleFontSize,
    title,
    sp,
  }: IOrgChartProps = props;


  const startFromUserId: Maybe<string> = React.useMemo(
    () => startFromUser && startFromUser[0].id,
    [startFromUser]
  );
  const coLeadUserId: Maybe<string> = React.useMemo(
    () => coLeadUserPicker && coLeadUserPicker[0] && coLeadUserPicker[0].id,
    [coLeadUserPicker]
  );
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
      const wRenderManagers: JSX.Element[] = [];
      const wRenderDirectReports: JSX.Element[] = [];
      const wRenderPeers: JSX.Element[] = [];

      try {
        const profileResponse = await getUserProfile(
          sp,
          selectedUser,
          startFromUserId,
          showAllManagers,
          showGuestUsers,
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

          for (const managerInfo of profileResponse.managersList) {
            wRenderManagers.push(
              <React.Fragment key={`manager-${managerInfo.id}`}>
                <PersonCard
                  userInfo={managerInfo}
                  onUserSelected={onUserSelected}
                  selectedUser={currentUser}
                  showActionsBar={showActionsBar}
                  sp={sp}
                 />
                <div
                  className={orgChartClasses.separatorVertical}
                 />
              </React.Fragment>
            );
          }

          if (showPeers !== false) {
            for (const peerInfo of filteredPeers) {
              wRenderPeers.push(
                <PersonCard
                  key={`peer-${peerInfo.id}`}
                  userInfo={peerInfo}
                  onUserSelected={onUserSelected}
                  selectedUser={currentUser}
                  showActionsBar={showActionsBar}
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

      return { wRenderDirectReports, wRenderManagers, wRenderPeers };
    },

    [
      sp,
      getUserProfile,
      startFromUserId,
      coLeadUserId,
      showAllManagers,
      showGuestUsers,
      onUserSelected,
      currentUser,
      showActionsBar,
      showPeers,
      matchesDepartmentFilter,
      sortReportsPriority,
      isCoLead,
      orgChartClasses.separatorVertical,
      orgChartClasses.coLeadGroup,
    ]
  );

  React.useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    (async () => {
      try {
        if (startFromUserId === undefined)  return;
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
        const profileResponse = await getUserProfile(sp, startFromUserId);
        const wCurrentUser: IUserInfo = await manpingUserProperties(
          // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
          profileResponse!.currentUserProfile
        );
        dispatch({
          type: EOrgChartTypes.SET_CURRENT_USER,
          payload: wCurrentUser,
        });
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
    })();
  }, [getUserProfile, sp, startFromUserId]);

  React.useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    (async () => {
      if (!coLeadUserId) {
        dispatch({ type: EOrgChartTypes.SET_CO_LEAD_USER, payload: undefined });
        return;
      }
      try {
        const profileResponse = await getUserProfile(sp, coLeadUserId);
        const wCoLeadUser: IUserInfo = await manpingUserProperties(
          // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
          profileResponse!.currentUserProfile
        );
        dispatch({
          type: EOrgChartTypes.SET_CO_LEAD_USER,
          payload: wCoLeadUser,
        });
      } catch (error) {
        console.log(error);
      }
    })();
  }, [getUserProfile, sp, coLeadUserId]);

  React.useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    (async () => {
      if (!currentUser || !currentUser.id) return;
      dispatch({
        type: EOrgChartTypes.SET_IS_LOADING,
        payload: true,
      });

      const { wRenderDirectReports, wRenderManagers, wRenderPeers } = await loadOrgChart(
        currentUser.id
      );
      dispatch({
        type: EOrgChartTypes.SET_RENDER_MANAGERS,
        payload: wRenderManagers,
      });
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
  }, [currentUser, loadOrgChart]);

  if (!startFromUser) {
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
      <Overlay style={{ height: "100%", position: "fixed" }}>
        <Stack style={{ height: "100%" }} verticalAlign="center">
          <Spinner
            styles={{ root: { zIndex: 9999 } }}
            size={SpinnerSize.large}
            label={"loading Organization Chart..."}
            labelPosition={"bottom"}
           />
        </Stack>
      </Overlay>
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
        {showTitle !== false && (
          <Stack horizontal horizontalAlign="center" styles={titleStyle}>
            {React.createElement(
              titleHeadingLevel || "h2",
              { style: { fontSize: titleFontSize || 28, margin: 0, fontWeight: 600 } },
              title
            )}
          </Stack>
        )}
        <Stack horizontalAlign="center" verticalAlign="center">
          {renderManagers}
          <Stack
            horizontal
            horizontalAlign="center"
            verticalAlign="center"
            tokens={{ childrenGap: 15 }}
            wrap
          >
            {renderPeers}
            <Stack
              horizontal
              horizontalAlign="center"
              verticalAlign="center"
              tokens={{ childrenGap: 15 }}
              className={orgChartClasses.leadershipBox}
            >
              <PersonCard
                key={`current-${currentUser?.id}`}
                // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
                userInfo={currentUser!}
                onUserSelected={onUserSelected}
                selectedUser={currentUser}
                showActionsBar={showActionsBar}
                sp={sp}
               />
              {coLeadUser && (
                <PersonCard
                  key={`co-lead-${coLeadUser.id}`}
                  userInfo={coLeadUser}
                  onUserSelected={onUserSelected}
                  selectedUser={currentUser}
                  showActionsBar={showActionsBar}
                  sp={sp}
                 />
              )}
            </Stack>
          </Stack>
          {renderDirectReports.length && (
            <>
              <div className={orgChartClasses.separatorVertical} />
              <div className={orgChartClasses.separatorHorizontal} />
            </>
          )}
        </Stack>
        {isDepartmentFilterActive && renderDirectReports.length === 0 && (
          <Stack horizontal horizontalAlign="center" styles={{ root: { padding: 10 } }}>
            <Text variant="medium">
              No direct reports found for the selected department filter.
            </Text>
          </Stack>
        )}
        <Stack
          horizontal
          horizontalAlign="center"
          styles={{root:{padding: 10}}}
          tokens={{ childrenGap: 15 }}
          wrap
          className={
            renderDirectReports.length ? orgChartClasses.teamBox : undefined
          }
        >
          {renderDirectReports}
        </Stack>
      </Stack>
    </>
  );
};