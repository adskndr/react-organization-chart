import { mergeStyles, mergeStyleSets } from "@fluentui/react";
import { PERSON_CARD_HEIGHT } from "../../common/cardDimensions";

const currentTheme = window.__themeState__.theme;

const orgChartClasses = mergeStyleSets({
  tilesContainer: mergeStyles({
    marginBottom: 10,
    marginTop: 0,
    gridGap: "10px",
    padding: 10,
    justifyContent: "center",
  }),

  separatorVertical: mergeStyles({
    height: 25,
    borderStyle: "solid",
    borderWidth: 1,
    borderColor: currentTheme.neutralQuaternary,
  }),

  coLeadGroup: mergeStyles({
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",
    gap: "15px",
    width: "fit-content",
    maxWidth: "100%",
    flexGrow: 0,
    flexShrink: 1,
  }),

  managerBox: mergeStyles({
    border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
    borderRadius: 4,
    padding: "4px 6px",
    boxSizing: "border-box",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    width: "fit-content",
    maxWidth: "100%",
    alignSelf: "center",
    flexGrow: 0,
    flexShrink: 0,
  }),

  managersGroup: mergeStyles({
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    width: "fit-content",
    maxWidth: "100%",
    alignSelf: "center",
    flexGrow: 0,
    flexShrink: 0,
  }),

  leadershipBox: mergeStyles({
    border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
    borderRadius: 4,
    padding: "4px 6px",
    boxSizing: "border-box",
    width: "fit-content",
    maxWidth: "100%",
    minHeight: PERSON_CARD_HEIGHT + 8,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    rowGap: "15px",
    alignSelf: "center",
    flexGrow: 0,
    flexShrink: 1,
  }),

  leadershipGroup: mergeStyles({
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",
    gap: "15px",
    width: "fit-content",
    maxWidth: "100%",
    flexGrow: 0,
    flexShrink: 1,
  }),

  peersGroup: mergeStyles({
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",
    flexWrap: "wrap",
    gap: "15px",
    width: "fit-content",
    maxWidth: "100%",
    flexGrow: 0,
    flexShrink: 1,
  }),

  teamBox: mergeStyles({
    border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
    borderRadius: 4,
    padding: "4px 6px",
    boxSizing: "border-box",
    width: "fit-content",
    maxWidth: "100%",
    display: "flex",
    flexWrap: "wrap",
    alignItems: "center",
    justifyContent: "center",
    columnGap: "15px",
    rowGap: "15px",
    alignSelf: "center",
    flexGrow: 0,
    flexShrink: 1,
  }),

  boxConnector: mergeStyles({
    width: 0,
    height: 12,
    borderLeftStyle: "solid",
    borderLeftWidth: 1,
    borderLeftColor: currentTheme.neutralQuaternary,
    alignSelf: "center",
    flexGrow: 0,
    flexShrink: 0,
  }),
});

export const useOrgChartStyles = () => {
  return { orgChartClasses };
};