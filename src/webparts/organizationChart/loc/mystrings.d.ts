declare interface IOrganizationChartWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  startFromUserLabel: string;
  coLeadUserLabel:string;
  managerLevelsLabel:string;
  showactionsLabel:string;
  showPeersLabel:string;
  departmentFilterLabel:string;
  departmentFilterTextLabel:string;
  departmentFilterTextDescription:string;
}

declare module 'OrganizationChartWebPartStrings' {
  const strings: IOrganizationChartWebPartStrings;
  export = strings;
}