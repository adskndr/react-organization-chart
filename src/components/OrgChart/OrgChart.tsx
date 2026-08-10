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
  Spinner,
  SpinnerSize,
} from "@fluentui/react";

import { useOrgChartStyles } from "./useOrgChartStyles";
import { useWrappedContentWidth } from "../../hooks/useWrappedContentWidth";

import "./OrgChart.module.scss";
import { Placeholder } from "../Placeholder/PlaceholderComponent";
import {
  getUsersByJobTitle,
  getUsersUnderManagerByJobTitle,
} from "../../services/PeopleSearchService";

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
  const [state, dispatch] = React.useReducer(
    OrgChartReducer,
    initialState
  );

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

  const [teamBoxRef, teamBoxWidth] = useWrappedContentWidth([
    renderDirectReports,
  ]);

  const [leadershipBoxRef, leadershipBoxWidth] =
    useWrappedContentWidth([
      renderPeers,
      currentUser,
      coLeadUser,
    ]);

  const {
    context,
    managerLevels,
    startFromUser,
    coLeadUser: coLeadUserPicker,
    showActionsBar,
    showPeers,
    departmentFilterSelected,
    departmentFilterText,
    jobTitleFilterText,
    graphClient,
    sp,
  }: IOrgChartProps = props;

  const startFromUserId: Maybe = React.useMemo(
    () =>
      startFromUser &&
      startFromUser[0] &&
      startFromUser[0].id,
    [startFromUser]
  );

  const coLeadUserId: Maybe = React.useMemo(() => {
    const id =
      coLeadUserPicker &&
      coLeadUserPicker[0] &&
      coLeadUserPicker[0].id;

    // Wenn dieselbe Person als Start-User und Co-Lead
    // ausgewählt wurde, Co-Lead ignorieren.
    return id && id !== startFromUserId
      ? id
      : undefined;
  }, [coLeadUserPicker, startFromUserId]);

  const onUserSelected = React.useCallback(
    (selectedUser: IUserInfo) => {
      dispatch({
        type: EOrgChartTypes.SET_CURRENT_USER,
        payload: selectedUser,
      });
    },
    []
  );

  const matchesDepartmentFilter = React.useCallback(
    (department?: string): boolean => {
      const selectedDepartments: string[] = (
        departmentFilterSelected ?? []
      ).map((department) =>
        department.trim().toLowerCase()
      );

      const departmentText: string = (
        departmentFilterText ?? ""
      )
        .trim()
        .toLowerCase();

      if (
        selectedDepartments.length === 0 &&
        !departmentText
      ) {
        return true;
      }

      const userDepartment: string = (
        department ?? ""
      )
        .trim()
        .toLowerCase();

      if (
        selectedDepartments.length > 0 &&
        selectedDepartments.indexOf(userDepartment) > -1
      ) {
        return true;
      }

      if (
        departmentText &&
        userDepartment.indexOf(departmentText) > -1
      ) {
        return true;
      }

      return false;
    },
    [departmentFilterSelected, departmentFilterText]
  );

  const isDeputy = React.useCallback(
    (user: IUserInfo): boolean => {
      return /\bstv.?/i.test(user.title ?? "");
    },
    []
  );

  const isCoLead = React.useCallback(
    (user: IUserInfo): boolean => {
      return /\bco[-\s]?\w*(leiter|leitung|lead)\b/i.test(
        user.title ?? ""
      );
    },
    []
  );

  const sortReportsPriority = React.useCallback(
    (users: IUserInfo[]): IUserInfo[] => {
      const rank = (user: IUserInfo): number => {
        if (isDeputy(user)) return 0;
        if (isCoLead(user)) return 1;
        return 2;
      };

      return [...users].sort((a, b) => {
        const rankDiff = rank(a) - rank(b);

        if (rankDiff !== 0) {
          return rankDiff;
        }

        return (a.displayName ?? "").localeCompare(
          b.displayName ?? ""
        );
      });
    },
    [isDeputy, isCoLead]
  );

  const isDepartmentFilterActive: boolean =
    (departmentFilterSelected ?? []).length > 0 ||
    !!(departmentFilterText ?? "").trim();

  const isJobTitleFilterActive: boolean =
    !!(jobTitleFilterText ?? "").trim();

  /*
   * Lädt die Personen für den OrgChart.
   *
   * JobTitle:
   *   - ohne Start-User:
   *       alle Personen mit passendem JobTitle
   *
   *   - mit Start-User:
   *       nur Personen unterhalb des Start-Users
   *
   * Normal:
   *   - Direct Reports und Peers des ausgewählten Users
   */
  const loadOrgChart = React.useCallback(
    async (
      selectedUser: string
    ): Promise<{
      wRenderDirectReports: JSX.Element[];
      wRenderPeers: JSX.Element[];
    }> => {
      const wRenderDirectReports: JSX.Element[] = [];
      const wRenderPeers: JSX.Element[] = [];

      try {
        /*
         * JOBTITLE FILTER
         */
        if (isJobTitleFilterActive) {
          let people: IUserInfo[];

          const jobTitles = (jobTitleFilterText ?? "")
            .split(/[;,]/)
            .map((title) => title.trim())
            .filter(Boolean);

          /*
           * JobTitle + Start-User:
           * Nur Personen unterhalb des gewählten Managers.
           */
          if (startFromUserId) {
            const peoplePerJobTitle =
              await Promise.all(
                jobTitles.map((jobTitle) =>
                  getUsersUnderManagerByJobTitle(
                    sp,
                    startFromUserId,
                    jobTitle
                  )
                )
              );

            people = ([] as IUserInfo[]).concat(
              ...peoplePerJobTitle
            );
          } else {
            /*
             * NUR JobTitle:
             * Alle Personen in der Organisation.
             */
            people = await getUsersByJobTitle(
              sp,
              jobTitles
            );
          }

          /*
           * Duplikate entfernen.
           */
          people = people.filter(
            (person, index, array) =>
              array.findIndex(
                (p) =>
                  (p.id ?? p.email) ===
                  (person.id ?? person.email)
              ) === index
          );

          const filteredPeople = sortReportsPriority(
            people.filter((person) =>
              matchesDepartmentFilter(
                person.department
              )
            )
          );

          /*
           * JobTitle-Ergebnisse als normale Karten darstellen.
           */
          for (const person of filteredPeople) {
            wRenderDirectReports.push(
              <PersonCard
                key={`jobtitle-${person.id ?? person.email}`}
                userInfo={person}
                onUserSelected={onUserSelected}
                selectedUser={currentUser}
                showActionsBar={showActionsBar}
                graphClient={graphClient}
                serviceScope={context.serviceScope}
                sp={sp}
              />
            );
          }

          dispatch({
            type: EOrgChartTypes.SET_HAS_ERROR,
            payload: {
              hasError: false,
              errorMessage: "",
            },
          });

          return {
            wRenderDirectReports,
            wRenderPeers,
          };
        }

        /*
         * NORMALER ORGCHART
         */
        const profileResponse = await getUserProfile(
          sp,
          selectedUser,
          0
        );

        if (profileResponse) {
          /*
           * Direct Reports des Hauptleiters.
           */
          const combinedReportsMap = new Map<
            string,
            IUserInfo
          >();

          profileResponse.reportsLists.forEach(
            (report) => {
              combinedReportsMap.set(
                report.id ?? report.email,
                report
              );
            }
          );

          /*
           * Falls ein Co-Lead konfiguriert ist,
           * dessen Reports ebenfalls hinzufügen.
           */
          if (
            coLeadUserId &&
            selectedUser === startFromUserId
          ) {
            const coLeadProfileResponse =
              await getUserProfile(
                sp,
                coLeadUserId
              );

            coLeadProfileResponse?.reportsLists.forEach(
              (report) => {
                combinedReportsMap.set(
                  report.id ?? report.email,
                  report
                );
              }
            );
          }

          const filteredDirectReports =
            sortReportsPriority(
              Array.from(
                combinedReportsMap.values()
              ).filter((report) =>
                matchesDepartmentFilter(
                  report.department
                )
              )
            );

          const filteredPeers =
            profileResponse.peersList.filter((peer) =>
              matchesDepartmentFilter(
                peer.department
              )
            );

          /*
           * Peers darstellen.
           */
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

          /*
           * Direct Reports darstellen.
           */
          let reportIndex = 0;

          while (
            reportIndex <
            filteredDirectReports.length
          ) {
            const directReport =
              filteredDirectReports[reportIndex];

            /*
             * Co-Leads gruppieren.
             */
            if (isCoLead(directReport)) {
              const coLeadGroup: IUserInfo[] = [
                directReport,
              ];

              let lookahead = reportIndex + 1;

              while (
                lookahead <
                  filteredDirectReports.length &&
                isCoLead(
                  filteredDirectReports[lookahead]
                )
              ) {
                coLeadGroup.push(
                  filteredDirectReports[lookahead]
                );

                lookahead++;
              }

              if (coLeadGroup.length > 1) {
                wRenderDirectReports.push(
                  <Stack
                    key={`co-lead-group-${coLeadGroup[0].id}`}
                    horizontal
                    verticalAlign="center"
                    tokens={{
                      childrenGap: 8,
                    }}
                    className={
                      orgChartClasses.coLeadGroup
                    }
                  >
                    {coLeadGroup.map((coLead) => (
                      <PersonCard
                        key={`report-${coLead.id}`}
                        userInfo={coLead}
                        onUserSelected={
                          onUserSelected
                        }
                        selectedUser={currentUser}
                        showActionsBar={
                          showActionsBar
                        }
                        graphClient={graphClient}
                        serviceScope={
                          context.serviceScope
                        }
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
          payload: {
            hasError: false,
            errorMessage: "",
          },
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

      return {
        wRenderDirectReports,
        wRenderPeers,
      };
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
      jobTitleFilterText,
      isJobTitleFilterActive,
      context.serviceScope,
    ]
  );

  /*
   * ============================================================
   * START USER / MANAGER
   * ============================================================
   *
   * WICHTIG:
   *
   * Wenn NUR JobTitle aktiv ist und kein Start-User vorhanden ist,
   * darf hier KEIN Manager geladen werden.
   *
   * Dadurch kann kein alter Manager-Zustand stehen bleiben.
   */
  React.useEffect(() => {
    let cancelled = false;

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    (async () => {
      try {
        /*
         * NUR JOBTITLE:
         *
         * Der normale OrgChart wird komplett geleert.
         *
         * Wichtig:
         * Hier wird ausdrücklich NICHT getUserProfile()
         * aufgerufen.
         */
        if (
          isJobTitleFilterActive &&
          !startFromUserId
        ) {
          dispatch({
            type: EOrgChartTypes.SET_CURRENT_USER,
            payload: undefined,
          });

          dispatch({
            type: EOrgChartTypes.SET_RENDER_MANAGERS,
            payload: [],
          });

          dispatch({
            type: EOrgChartTypes.SET_RENDER_PEERS,
            payload: [],
          });

          dispatch({
            type: EOrgChartTypes.SET_CO_LEAD_USER,
            payload: undefined,
          });

          dispatch({
            type: EOrgChartTypes.SET_IS_LOADING,
            payload: false,
          });

          return;
        }

        /*
         * Kein Start-User und kein JobTitle:
         * Alles zurücksetzen.
         */
        if (
          startFromUserId === undefined &&
          !isJobTitleFilterActive
        ) {
          dispatch({
            type: EOrgChartTypes.SET_CURRENT_USER,
            payload: undefined,
          });

          dispatch({
            type: EOrgChartTypes.SET_RENDER_MANAGERS,
            payload: [],
          });

          dispatch({
            type: EOrgChartTypes.SET_RENDER_DIRECT_REPORTS,
            payload: [],
          });

          dispatch({
            type: EOrgChartTypes.SET_RENDER_PEERS,
            payload: [],
          });

          dispatch({
            type: EOrgChartTypes.SET_CO_LEAD_USER,
            payload: undefined,
          });

          dispatch({
            type: EOrgChartTypes.SET_IS_LOADING,
            payload: false,
          });

          return;
        }

        /*
         * Ungültiger User.
         */
        if (startFromUserId === "") {
          dispatch({
            type: EOrgChartTypes.SET_IS_LOADING,
            payload: false,
          });

          dispatch({
            type: EOrgChartTypes.SET_HAS_ERROR,
            payload: {
              hasError: true,
              errorMessage:
                "User don't have email defined",
            },
          });

          return;
        }

        if (cancelled || !startFromUserId) {
          return;
        }

        /*
         * User-Profil laden.
         *
         * Dieser Teil wird bei NUR JobTitle nicht erreicht.
         */
        const profileResponse =
          await getUserProfile(
            sp,
            startFromUserId,
            managerLevels,
            false
          );

        if (cancelled) return;

        const wCurrentUser: IUserInfo =
          await manpingUserProperties(
            // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
            profileResponse!.currentUserProfile
          );

        if (cancelled) return;

        dispatch({
          type: EOrgChartTypes.SET_CURRENT_USER,
          payload: wCurrentUser,
        });

        /*
         * Manager rendern.
         *
         * Das passiert nur, wenn tatsächlich ein Start-User
         * vorhanden ist.
         */
        const managersList =
          profileResponse!.managersList;

        const wRenderManagers: JSX.Element[] = [];

        managersList.forEach(
          (managerInfo, index) => {
            wRenderManagers.push(
              <div
                key={`manager-box-${managerInfo.id}`}
                className={
                  orgChartClasses.managerBox
                }
              >
                <PersonCard
                  userInfo={managerInfo}
                  onUserSelected={
                    onUserSelected
                  }
                  selectedUser={currentUser}
                  showActionsBar={
                    showActionsBar
                  }
                  graphClient={graphClient}
                  serviceScope={
                    context.serviceScope
                  }
                  sp={sp}
                />
              </div>
            );

            if (
              index <
              managersList.length - 1
            ) {
              wRenderManagers.push(
                <div
                  key={`manager-connector-${managerInfo.id}`}
                  className={
                    orgChartClasses.boxConnector
                  }
                />
              );
            }
          }
        );

        dispatch({
          type: EOrgChartTypes.SET_RENDER_MANAGERS,
          payload: wRenderManagers,
        });

        dispatch({
          type: EOrgChartTypes.SET_HAS_ERROR,
          payload: {
            hasError: false,
            errorMessage: "",
          },
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
    isJobTitleFilterActive,
    onUserSelected,
    currentUser,
    context.serviceScope,
    orgChartClasses.managerBox,
    orgChartClasses.boxConnector,
  ]);

  /*
   * ============================================================
   * CO-LEAD
   * ============================================================
   */
  React.useEffect(() => {
    let cancelled = false;

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    (async () => {
      /*
       * Bei reinem JobTitle-Filter gibt es keinen Co-Lead.
       */
      if (
        isJobTitleFilterActive &&
        !startFromUserId
      ) {
        dispatch({
          type: EOrgChartTypes.SET_CO_LEAD_USER,
          payload: undefined,
        });

        return;
      }

      if (!coLeadUserId) {
        dispatch({
          type: EOrgChartTypes.SET_CO_LEAD_USER,
          payload: undefined,
        });

        return;
      }

      try {
        const profileResponse =
          await getUserProfile(
            sp,
            coLeadUserId
          );

        if (cancelled) return;

        const wCoLeadUser: IUserInfo =
          await manpingUserProperties(
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
  }, [
    getUserProfile,
    sp,
    coLeadUserId,
    isJobTitleFilterActive,
    startFromUserId,
  ]);

  /*
   * ============================================================
   * DIRECT REPORTS / JOBTITLE / NORMAL ORGCHART
   * ============================================================
   */
  React.useEffect(() => {
    let cancelled = false;

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    (async () => {
      /*
       * ========================================================
       * NUR JOBTITLE
       * ========================================================
       *
       * Kein Start-User:
       * - keine Manager
       * - kein Current User
       * - keine Peers
       * - nur JobTitle-Ergebnisse
       */
      if (
        isJobTitleFilterActive &&
        !startFromUserId
      ) {
        dispatch({
          type: EOrgChartTypes.SET_RENDER_MANAGERS,
          payload: [],
        });

        dispatch({
          type: EOrgChartTypes.SET_CURRENT_USER,
          payload: undefined,
        });

        dispatch({
          type: EOrgChartTypes.SET_CO_LEAD_USER,
          payload: undefined,
        });

        dispatch({
          type: EOrgChartTypes.SET_RENDER_PEERS,
          payload: [],
        });

        dispatch({
          type: EOrgChartTypes.SET_IS_LOADING,
          payload: true,
        });

        try {
          const {
            wRenderDirectReports,
            wRenderPeers,
          } = await loadOrgChart("");

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

        return;
      }

      /*
       * ========================================================
       * JOBTITLE + MANAGER
       * ========================================================
       */
      if (
        isJobTitleFilterActive &&
        startFromUserId
      ) {
        dispatch({
          type: EOrgChartTypes.SET_IS_LOADING,
          payload: true,
        });

        try {
          const {
            wRenderDirectReports,
            wRenderPeers,
          } = await loadOrgChart(
            startFromUserId
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

        return;
      }

      /*
       * ========================================================
       * NORMALER ORGCHART
       * ========================================================
       */
      if (
        !currentUser ||
        !currentUser.id
      ) {
        return;
      }

      dispatch({
        type: EOrgChartTypes.SET_IS_LOADING,
        payload: true,
      });

      try {
        const {
          wRenderDirectReports,
          wRenderPeers,
        } = await loadOrgChart(
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
    currentUser,
    loadOrgChart,
    isJobTitleFilterActive,
    startFromUserId,
  ]);

  /*
   * ============================================================
   * PLACEHOLDER
   * ============================================================
   */
  if (
    !startFromUserId &&
    !isJobTitleFilterActive
  ) {
    return (
      <Placeholder
        iconName="Edit"
        iconText="Configure your Organization Chart Web Part"
        description="Please configure web part"
        buttonLabel="Configure"
        onConfigure={
          context.propertyPane.open
        }
      />
    );
  }

  /*
   * ============================================================
   * LOADING
   * ============================================================
   */
  if (isLoading) {
    return (
      <Stack
        style={{ minHeight: 200 }}
        verticalAlign="center"
        horizontalAlign="center"
      >
        <Spinner
          size={SpinnerSize.large}
          label="loading Organization Chart..."
          labelPosition="bottom"
        />
      </Stack>
    );
  }

  /*
   * ============================================================
   * ERROR
   * ============================================================
   */
  if (error && error.hasError) {
    return (
      <Stack
        horizontal
        horizontalAlign="center"
        styles={{
          root: {
            padding: 20,
          },
        }}
        tokens={{
          childrenGap: 10,
        }}
      >
        {error.errorMessage}
      </Stack>
    );
  }

  /*
   * ============================================================
   * RENDER
   * ============================================================
   */
  return (
    <>
      <Stack
        styles={{
          root: {
            padding: 20,
          },
        }}
      >
        {/*
         * ======================================================
         * MANAGER
         * ======================================================
         *
         * WICHTIG:
         *
         * Sobald der JobTitle-Filter aktiv ist, wird
         * renderManagers NICHT mehr angezeigt.
         *
         * Das ist die entscheidende Änderung.
         *
         * Vorher:
         *
         *   (!isJobTitleFilterActive || startFromUserId)
         *
         * Jetzt:
         *
         *   !isJobTitleFilterActive
         *
         * Dadurch kann bei einem reinen JobTitle-Filter
         * niemals die Manager-Box erscheinen.
         */}
        {!isJobTitleFilterActive &&
          renderManagers.length > 0 && (
            <>
              {renderManagers}
            </>
          )}

        {/*
         * ======================================================
         * LEADERSHIP BOX
         * ======================================================
         *
         * Bei reinem JobTitle-Filter ebenfalls nicht anzeigen,
         * weil dort kein Current User benötigt wird.
         */}
        {(!isJobTitleFilterActive ||
          startFromUserId) && (
          <div
            ref={leadershipBoxRef}
            className={
              orgChartClasses.leadershipBox
            }
            style={
              leadershipBoxWidth !==
              undefined
                ? {
                    width:
                      leadershipBoxWidth,
                  }
                : undefined
            }
          >
            {renderPeers}

            {currentUser && (
              <PersonCard
                key={`current-${currentUser.id}`}
                userInfo={currentUser}
                onUserSelected={
                  onUserSelected
                }
                selectedUser={currentUser}
                showActionsBar={
                  showActionsBar
                }
                graphClient={graphClient}
                serviceScope={
                  context.serviceScope
                }
                sp={sp}
              />
            )}

            {coLeadUser && (
              <PersonCard
                key={`co-lead-${coLeadUser.id}`}
                userInfo={coLeadUser}
                onUserSelected={
                  onUserSelected
                }
                selectedUser={currentUser}
                showActionsBar={
                  showActionsBar
                }
                graphClient={graphClient}
                serviceScope={
                  context.serviceScope
                }
                sp={sp}
              />
            )}
          </div>
        )}

        {/*
         * ======================================================
         * NO RESULTS
         * ======================================================
         */}
        {(isDepartmentFilterActive ||
          isJobTitleFilterActive) &&
          renderDirectReports.length === 0 && (
            <Stack
              horizontal
              horizontalAlign="center"
              styles={{
                root: {
                  padding: 10,
                },
              }}
            >
              {isJobTitleFilterActive
                ? `No people found with JobTitle "${jobTitleFilterText}".`
                : "No direct reports found for the selected department filter."}
            </Stack>
          )}

        {/*
         * ======================================================
         * DIRECT REPORTS / JOBTITLE RESULTS
         * ======================================================
         */}
        {renderDirectReports.length > 0 && (
          <div
            ref={teamBoxRef}
            className={
              orgChartClasses.teamBox
            }
            style={
              teamBoxWidth !== undefined
                ? {
                    width: teamBoxWidth,
                  }
                : undefined
            }
          >
            {renderDirectReports}
          </div>
        )}
      </Stack>
    </>
  );
};

