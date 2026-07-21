import { mergeStyles, mergeStyleSets } from "@fluentui/react";

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
    padding: 16,
    // Fixed footprint (matches the person card's fixed 260px width + this
    // box's own padding/border) so every manager box is guaranteed the
    // exact same size, regardless of that manager's content.
    width: 294,
    boxSizing: "border-box",
    display: "flex",
    justifyContent: "center",
  }),

  leadershipBox: mergeStyles({
    border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
    borderRadius: 4,
    padding: 16,
    boxSizing: "border-box",
    display: "flex",
    flexWrap: "wrap",
    justifyContent: "center",
    alignItems: "center",
    columnGap: "15px",
    rowGap: "15px",
  }),

  leadershipGroup: mergeStyles({
    display: "flex",
    alignItems: "center",
    columnGap: "15px",
    flexShrink: 0,
  }),

  teamBox: mergeStyles({
    border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
    borderRadius: 4,
    padding: 16,
  }),

  boxConnector: mergeStyles({
    width: 0,
    height: 20,
    borderLeftStyle: "solid",
    borderLeftWidth: 1,
    borderLeftColor: "#ffffff",
  }),
});

// eslint-disable-next-line @typescript-eslint/explicit-module-boundary-types, @typescript-eslint/explicit-function-return-type
export const useOrgChartStyles = () => {
  return { orgChartClasses };
};