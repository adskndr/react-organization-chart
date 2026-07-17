import { mergeStyles, mergeStyleSets } from "@fluentui/react";

const currentTheme = window.__themeState__.theme;

// eslint-disable-next-line @typescript-eslint/explicit-module-boundary-types, @typescript-eslint/explicit-function-return-type
export const useOrgChartStyles = () => {

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

    separatorHorizontal: mergeStyles({
      width: "100%",
      borderStyle: "solid",
      borderWidth: 1,
      borderColor: currentTheme.neutralQuaternary,
    }),

    coLeadGroup: mergeStyles({
      // No border by design — the tighter gap (see tokens in OrgChart.tsx)
      // is what visually groups the co-leads, without boxing them in.
    }),

    leadershipBox: mergeStyles({
      border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
      borderRadius: 4,
      padding: 16,
    }),

    teamBox: mergeStyles({
      border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
      borderRadius: 4,
      padding: 16,
    }),
  });

  return { orgChartClasses };
};