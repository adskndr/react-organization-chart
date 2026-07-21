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
    paddingTop: 4,
    paddingBottom: 4,
    paddingLeft: 6,
    paddingRight: 6,
    boxSizing: "border-box",
    // inline-flex shrink-wraps to its content on its own — no need for an
    // explicit width to force that (see teamBox below for why that matters
    // once there's more than one card to wrap).
    display: "inline-flex",
    justifyContent: "center",
    alignItems: "center",
  }),

  leadershipBox: mergeStyles({
    border: `1px solid ${currentTheme.neutralQuaternaryAlt}`,
    borderRadius: 4,
    // Kept in sync with managerBox/teamBox padding so every grey box looks
    // identical regardless of what it wraps (manager, start-from-user +
    // co-lead, or the team row).
    paddingTop: 4,
    paddingBottom: 4,
    paddingLeft: 6,
    paddingRight: 6,
    boxSizing: "border-box",
    // maxWidth keeps it from overflowing the page once there are enough
    // peer/co-lead cards to need more than one row.
    maxWidth: "100%",
    minHeight: PERSON_CARD_HEIGHT + 8, // card height + top/bottom padding
    display: "inline-flex",
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
    // Kept in sync with managerBox's padding so a single employee's box
    // looks identical to the manager's.
    paddingTop: 4,
    paddingBottom: 4,
    paddingLeft: 6,
    paddingRight: 6,
    boxSizing: "border-box",
    // inline-flex shrink-wraps to its content by itself, including once
    // flex-wrap kicks in and the cards spill onto a second row (e.g. 3+
    // direct reports) — the box's own width becomes the widest row's
    // width, automatically. The earlier width:"fit-content" approach
    // computed that less reliably once wrapping was involved, which is
    // exactly what caused the box to look "off" specifically at 3 people.
    maxWidth: "100%",
    display: "inline-flex",
    flexWrap: "wrap",
    justifyContent: "center",
    columnGap: "15px",
    rowGap: "15px",
    // teamBox's parent (the outer page Stack) doesn't set an explicit
    // horizontalAlign, so it defaults to stretching children to full
    // width — inline-flex only controls how this box sizes its own
    // content, not whether its parent stretches it. Without this, the
    // box would stretch full-width and only wrap once it ran out of an
    // entire page row, instead of hugging its cards.
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