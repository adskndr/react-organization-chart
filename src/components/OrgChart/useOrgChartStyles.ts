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

  // Person cards are now a fixed size, so this group just needs a bit more
  // breathing room between co-leads — no border needed to read as a pair.
  coLeadGroup: mergeStyles({
    display: "flex",
    alignItems: "center",
    columnGap: "16px",
  }),

  // No fixed width/height here anymore — with the person card itself now a
  // fixed size, the box naturally hugs its content (padding + border only),
  // so it always matches however many managers are actually shown.
  managerBox: mergeStyles({
    border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
    borderRadius: 4,
    padding: 16,
    display: "inline-flex",
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
    columnGap: "16px",
    rowGap: "16px",
  }),

  leadershipGroup: mergeStyles({
    display: "flex",
    alignItems: "center",
    columnGap: "16px",
    flexShrink: 0,
  }),

  teamBox: mergeStyles({
    border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
    borderRadius: 4,
    padding: 16,
    boxSizing: "border-box",
    display: "flex",
    flexWrap: "wrap",
    justifyContent: "center",
    alignItems: "flex-start",
    columnGap: "16px",
    rowGap: "16px",
  }),

  boxConnector: mergeStyles({
    width: 0,
    height: 20,
    borderLeftStyle: "solid",
    borderLeftWidth: 1,
    borderLeftColor: currentTheme.white,
  }),
});

// eslint-disable-next-line @typescript-eslint/explicit-module-boundary-types, @typescript-eslint/explicit-function-return-type
export const useOrgChartStyles = () => {
  return { orgChartClasses };
};
