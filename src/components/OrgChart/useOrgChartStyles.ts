import { mergeStyles, mergeStyleSets } from "@fluentui/react";
import { PERSON_CARD_HEIGHT } from "../../common/cardDimensions";

const currentTheme = window.__themeState__.theme;

// Computed once at module load (the theme doesn't change during the page's
// lifetime) instead of on every render of every component using this hook.
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
    // No border by design — the tighter gap (see tokens in OrgChart.tsx)
    // is what visually groups the co-leads, without boxing them in.
  }),

  managerBox: mergeStyles({
    border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
    borderRadius: 4,
    padding: "4px 6px",
    boxSizing: "border-box",

    display: "inline-flex",
    justifyContent: "center",
    alignItems: "center",

    width: "max-content",
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

    width: "max-content",
    maxWidth: "100%",

    minHeight: PERSON_CARD_HEIGHT + 8,

    display: "flex",
    flexDirection: "column",
    alignItems: "center",

    rowGap: "15px",
    alignSelf: "center",
    flexGrow: 0,
    flexShrink: 0,
}),

  leadershipGroup: mergeStyles({
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",

    width: "max-content",
    maxWidth: "100%",

    columnGap: "15px",

    flexGrow: 0,
    flexShrink: 0,
}),

  teamBox: mergeStyles({
    border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
    borderRadius: 4,
    paddingTop: 4,
    paddingBottom: 4,
    paddingLeft: 6,
    paddingRight: 6,
    boxSizing: "border-box",
    maxWidth: "100%",
    display: "inline-flex",
    flexWrap: "wrap",
    justifyContent: "center",
    alignItems: "center",
    columnGap: "15px",
    rowGap: "15px",
    alignSelf: "center",
  }),

   peersGroup: mergeStyles({
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",

    width: "max-content",
    maxWidth: "100%",

    flexWrap: "wrap",
    gap: "15px",

    flexGrow: 0,
    flexShrink: 0,
  }),

  boxConnector: mergeStyles({
    width: 0,
    height: 12,
    borderLeftStyle: "solid",
    borderLeftWidth: 1,
    borderLeftColor: "#ffffff",
  }), 

  managersGroup: mergeStyles({
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",

    width: "max-content",
    maxWidth: "100%",

    flexGrow: 0,
    flexShrink: 0,
  }),
});

// eslint-disable-next-line @typescript-eslint/explicit-module-boundary-types, @typescript-eslint/explicit-function-return-type
export const useOrgChartStyles = () => {
  return { orgChartClasses };
};