import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";
import {
  PropertyFieldPeoplePicker,
  PrincipalType,
  IPropertyFieldGroupOrPerson,
} from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import * as strings from "OrganizationChartWebPartStrings";
import { OrgChart } from "../../components/OrgChart/OrgChart";

import { IOrgChartProps } from "../../components/OrgChart/IOrgChartProps";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/profiles";
import { getDepartments } from "../../services/DepartmentService";

let _sp: SPFI;
export interface IOrganizationChartWebPartProps {
  title: string;
  currentUser: string;
  selectedUser: IPropertyFieldGroupOrPerson[];
  coLeadUser: IPropertyFieldGroupOrPerson[];
  showAllManagers: boolean;
  showGuestUsers: boolean;
  showPeers: boolean;
  showActionsBar: boolean;
  departmentFilterSelected: string[];
  departmentFilterText: string;
  showTitle: boolean;
  titleHeadingLevel: string;
  titleFontSize: number;
}

export default class OrganizationChartWebPart extends BaseClientSideWebPart<IOrganizationChartWebPartProps> {
  private _departmentOptions: IPropertyPaneDropdownOption[] = [];
  private _departmentsLoaded: boolean = false;

  public async onInit(): Promise<void> {
    _sp = spfi().using(SPFx(this.context));
    return super.onInit();
  }
  
  public get sp(): SPFI {
    return _sp;
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this._departmentsLoaded) return;

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    getDepartments(this.sp)
      .then((departments: string[]) => {
        this._departmentOptions = departments.map((department) => ({
          key: department,
          text: department,
        }));
        this._departmentsLoaded = true;
        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        console.log(error);
        this._departmentsLoaded = true;
      });
  }

  public render(): void {
    const element: React.ReactElement<IOrgChartProps> = React.createElement(
      OrgChart,
      {
        title: this.properties.title,
        defaultUser: this.properties.currentUser,
        startFromUser: this.properties.selectedUser,
        coLeadUser: this.properties.coLeadUser,
        showAllManagers: this.properties.showAllManagers,
        showGuestUsers: this.properties.showGuestUsers,
        context: this.context,
        showActionsBar: this.properties.showActionsBar,
        showPeers: this.properties.showPeers,
        departmentFilterSelected: this.properties.departmentFilterSelected,
        departmentFilterText: this.properties.departmentFilterText,
        showTitle: this.properties.showTitle,
        titleHeadingLevel: this.properties.titleHeadingLevel,
        titleFontSize: this.properties.titleFontSize,
        sp: this.sp,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: strings.TitleFieldLabel,
                }),
                PropertyFieldPeoplePicker("selectedUser", {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  label: strings.startFromUserLabel,
                  initialData: this.properties.selectedUser,
                  key: "peopleFieldId",
                  multiSelect: false,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                }),
                PropertyFieldPeoplePicker("coLeadUser", {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  label: strings.coLeadUserLabel,
                  initialData: this.properties.coLeadUser,
                  key: "coLeadPeopleFieldId",
                  multiSelect: false,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                }),
                PropertyPaneToggle("showAllManagers", {
                  label: strings.showAllManagers,
                }),
                PropertyPaneToggle("showGuestUsers", {
                  label: strings.showGuestUsers,
                }),
                PropertyPaneToggle("showPeers", {
                  label: strings.showPeersLabel,
                }),
                PropertyPaneToggle("showActionsBar", {
                  label: strings.showactionsLabel,
                }),
                PropertyFieldMultiSelect("departmentFilterSelected", {
                  key: "departmentFilterSelected",
                  label: strings.departmentFilterLabel,
                  options: this._departmentOptions,
                  selectedKeys: this.properties.departmentFilterSelected,
                }),
                PropertyPaneTextField("departmentFilterText", {
                  label: strings.departmentFilterTextLabel,
                  description: strings.departmentFilterTextDescription,
                }),
              ],
            },
            {
              groupName: strings.TitleGroupName,
              groupFields: [
                PropertyPaneToggle("showTitle", {
                  label: strings.showTitleLabel,
                }),
                PropertyPaneDropdown("titleHeadingLevel", {
                  label: strings.titleHeadingLevelLabel,
                  options: [
                    { key: "h1", text: "H1" },
                    { key: "h2", text: "H2" },
                    { key: "h3", text: "H3" },
                    { key: "h4", text: "H4" },
                    { key: "h5", text: "H5" },
                    { key: "h6", text: "H6" },
                  ],
                }),
                PropertyPaneSlider("titleFontSize", {
                  label: strings.titleFontSizeLabel,
                  min: 12,
                  max: 60,
                  step: 1,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}