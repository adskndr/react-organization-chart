declare interface IOrganizationChartWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  startFromUserLabel: string;
  coLeadUserLabel:string;
  showactionsLabel:string;
  showAllManagers:string;
  showGuestUsers:string;
  showPeersLabel:string;
  departmentFilterLabel:string;
  departmentFilterTextLabel:string;
  departmentFilterTextDescription:string;
  TitleGroupName:string;
  showTitleLabel:string;
  titleHeadingLevelLabel:string;
  titleFontSizeLabel:string;
}

declare module 'OrganizationChartWebPartStrings' {
  const strings: IOrganizationChartWebPartStrings;
  export = strings;
}