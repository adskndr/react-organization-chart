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
    padding: 8,
    // Sized off its own content instead of a hardcoded pixel value: it
    // always holds exactly one card, but this way the box never drifts out
    // of sync if the card's own width ever changes.
    width: "fit-content",
    boxSizing: "border-box",
    display: "flex",
    justifyContent: "center",
  }),

  leadershipBox: mergeStyles({
    border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
    borderRadius: 4,
    padding: 16,
    boxSizing: "border-box",
    // fit-content is what actually makes this dynamic: the box hugs
    // whatever it contains — one card, or lead + co-lead side by side —
    // and grows/shrinks automatically as cards are added or removed,
    // instead of relying on a width that was only ever correct for one
    // card. maxWidth keeps it from overflowing the page if many peer
    // cards are shown alongside the lead/co-lead group.
    width: "fit-content",
    maxWidth: "100%",
    minHeight: PERSON_CARD_HEIGHT + 32, // card height + top/bottom padding
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
    padding: 8,
    boxSizing: "border-box",
    width: "fit-content",
    maxWidth: "100%",
    // The box sizes to its own content (fit-content) rather than stretching
    // to the full row width, so it's no longer stretched/centered by its
    // parent's default flex behavior — without this it lands flush left
    // (most noticeable with just 1-2 direct reports). alignSelf forces it
    // to center itself horizontally regardless of parent alignment.
    alignSelf: "center",
  }),

  boxConnector: mergeStyles({
    width: 0,
    height: 12,
    borderLeftStyle: "solid",
    borderLeftWidth: 1,
    borderLeftColor: "#ffffff",
  }),
});

// eslint-disable-next-line @typescript-eslint/explicit-module-boundary-types, @typescript-eslint/explicit-function-return-type
export const useOrgChartStyles = () => {
  return { orgChartClasses };
};