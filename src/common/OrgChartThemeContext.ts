import * as React from "react";

/**
 * Admin-configurable colours for the chart. Every value is optional; when a
 * value is undefined/empty the component falls back to the SharePoint theme
 * colour it used before this feature existed.
 */
export interface IOrgChartColors {
  /** Background behind the whole chart. */
  backgroundColor?: string;
  /** Background of each person card (and its actions bar). */
  cardBackgroundColor?: string;
  /** Text colour on the person cards (name and job title). */
  textColor?: string;
}

export const OrgChartThemeContext = React.createContext<IOrgChartColors>({});

export const useOrgChartColors = (): IOrgChartColors =>
  React.useContext(OrgChartThemeContext);
