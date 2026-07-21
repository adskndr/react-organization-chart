import { IButtonStyles, IDocumentCardActionsStyles, IStackStyles, mergeStyles, mergeStyleSets } from "@fluentui/react";
import type { Theme } from "spfx-uifabric-themes";
import { PERSON_CARD_WIDTH, PERSON_CARD_HEIGHT } from "../../common/cardDimensions";
const currentTheme: Theme = window.__themeState__.theme;

// All computed once at module load instead of on every render of every
// PersonCard — this hook runs once per person shown in the chart, so this
// avoids re-running mergeStyleSets/mergeStyles dozens of times per render.
//
// The fixed card height lives here, on the persona content row, NOT on the
// outer .tile — the actions bar (Chat/Mail/Call/Org) is a separate element
// that DocumentCard stacks below this row. Putting the fixed height on
// .tile with overflow:hidden would clip that bar off entirely whenever
// showActionsBar is on.
const stackPersonaStyles: Partial<IStackStyles> = {
  root: {
    padding: 15,
    height: PERSON_CARD_HEIGHT,
    boxSizing: "border-box",
    overflow: "hidden",
  },
};

const buttonStyles: IButtonStyles = {
  icon: {
    fontSize: 12,
  },
  iconHovered: {
    fontWeight: 600,
  },
};

const documentCardActionStyles: Partial<IDocumentCardActionsStyles> = {
  root: {
    height: 34,
    padding: 0,
    backgroundColor: currentTheme.neutralLighterAlt,
    borderTopWidth: 1,
    borderTopStyle: "solid",
    borderTopColor: currentTheme.neutralLight,
    width: "100%",
  },
};

const personaCardStyles = mergeStyleSets({
  separatorHorizontal: mergeStyles({
    width: "100%",
    borderWidth: 0.5,
    borderStyle: "solid",
    borderColor: currentTheme.neutralLight,
  }),
  iconStyles: mergeStyles({
    fontSize: 16, color: currentTheme.themePrimary
  }),
  tile: mergeStyles({
    // Fixed WIDTH so every card lines up identically (and any box drawn
    // around them sizes predictably) — but height is intentionally left to
    // grow with content: the persona row above is pinned to
    // PERSON_CARD_HEIGHT, and the optional actions bar adds its own height
    // on top of that. A fixed height here would clip the actions bar.
    width: PERSON_CARD_WIDTH,
    // Fluent's DocumentCard ships with its own default margin (meant for
    // loose card grids). A Stack normally resets that on its children as
    // part of implementing childrenGap, which is why cards inside teamBox
    // sit flush against its border — but managerBox is a plain div with no
    // such reset, so that leftover margin showed up as extra space around
    // the manager's card. Zeroing it here makes both identical regardless
    // of what kind of container the card sits in.
    margin: 0,
    boxSizing: "border-box",
    borderStyle: "solid",
    borderWidth: 1,
    borderRadius: 0,
    borderColor: "transparent",
    backgroundColor: currentTheme.white,
    boxShadow: "0 5px 15px rgba(50, 50, 90, .1)",
    selectors: {
      ":hover": {
        borderStyle: "solid",
        borderWidth: 1,
        borderRadius: 0,
        borderColor: currentTheme.themePrimary,
      },
      ":focus": {
        borderStyle: "solid",
        borderWidth: 1,
        borderRadius: 0,
        borderColor: currentTheme.themePrimary,
      },
      "@media(max-width : 480px)": {
        width: "100%",
      },
    }
  }),
});

// eslint-disable-next-line @typescript-eslint/explicit-module-boundary-types, @typescript-eslint/explicit-function-return-type
export const usePersonaCardStyles = () => {
  return {documentCardActionStyles, personaCardStyles, buttonStyles, stackPersonaStyles };
};