import { IButtonStyles, IDocumentCardActionsStyles, IStackStyles, mergeStyles, mergeStyleSets } from "@fluentui/react";
import type { Theme } from "spfx-uifabric-themes";
const currentTheme: Theme = window.__themeState__.theme;

// All computed once at module load instead of on every render of every
// PersonCard — this hook runs once per person shown in the chart, so this
// avoids re-running mergeStyleSets/mergeStyles dozens of times per render.
const stackPersonaStyles: Partial<IStackStyles> = {
  root: { padding: "0 15px", height: 70, boxSizing: "border-box" },
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
    width: 336,
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
      "@media(max-width : 360px)": {
        width: "100%",
      },
    }
  }),
});

// eslint-disable-next-line @typescript-eslint/explicit-module-boundary-types, @typescript-eslint/explicit-function-return-type
export const usePersonaCardStyles = () => {
  return {documentCardActionStyles, personaCardStyles, buttonStyles, stackPersonaStyles };
};
