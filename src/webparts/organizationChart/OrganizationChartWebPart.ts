import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
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
import { getManagerChainLength } from "../../services/ManagerChainService";
import { IGraphPhotoClient } from "../../services/PhotoService";

let _sp: SPFI;
export interface IOrganizationChartWebPartProps {
  currentUser: string;
  selectedUser: IPropertyFieldGroupOrPerson[];
  coLeadUser: IPropertyFieldGroupOrPerson[];
  managerLevels: number;
  showPeers: boolean;
  showActionsBar: boolean;
  departmentFilterSelected: string[];
  departmentFilterText: string;
}

export default class OrganizationChartWebPart extends BaseClientSideWebPart<IOrganizationChartWebPartProps> {
  private _departmentOptions: IPropertyPaneDropdownOption[] = [];
  private _departmentsLoaded: boolean = false;
  private _graphClient?: IGraphPhotoClient;
  private _maxManagerLevels: number = 10;
  private _maxManagerLevelsLoadedFor?: string;

  public async onInit(): Promise<void> {
    _sp = spfi().using(SPFx(this.context));
    try {
      // Used to load profile photos directly from Entra ID (Microsoft Graph)
      // instead of the classic SharePoint userphoto.aspx endpoint.
      this._graphClient = await this.context.msGraphClientFactory.getClient(
        "3"
      );
    } catch (error) {
      console.log(error);
    }
    return super.onInit();
  }
  
  public get sp(): SPFI {
    return _sp;
  }

  // PropertyPaneTextField stores raw user input, so this normalizes whatever
  // ends up in the property bag (string or number) into a safe non-negative
  // integer, defaulting to 0 for anything invalid or empty.
  private get _managerLevels(): number {
    const raw = this.properties.managerLevels;
    const parsed =
      typeof raw === "number" ? raw : parseInt(String(raw ?? ""), 10);
    return Number.isFinite(parsed) && parsed >= 0 ? parsed : 0;
  }

  private async _loadMaxManagerLevels(): Promise<void> {
    const loginName = this.properties.selectedUser?.[0]?.id;
    if (!loginName || loginName === this._maxManagerLevelsLoadedFor) return;

    this._maxManagerLevelsLoadedFor = loginName;
    try {
      this._maxManagerLevels = await getManagerChainLength(this.sp, loginName);
      // Don't leave an already-configured value stranded above the newly
      // known ceiling (e.g. was set to 5, but this person only has 3).
      if (this._managerLevels > this._maxManagerLevels) {
        this.properties.managerLevels = this._maxManagerLevels;
      }
    } catch (error) {
      console.log(error);
    } finally {
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  private _onSelectedUserChanged = (
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void => {
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === "selectedUser") {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this._loadMaxManagerLevels();
    }
  };

  protected onPropertyPaneConfigurationStart(): void {
    if (!this._departmentsLoaded) {
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

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._loadMaxManagerLevels();
  }

  public render(): void {
    const element: React.ReactElement<IOrgChartProps> = React.createElement(
      OrgChart,
      {
        defaultUser: this.properties.currentUser,
        startFromUser: this.properties.selectedUser,
        coLeadUser: this.properties.coLeadUser,
        managerLevels: this._managerLevels,
        context: this.context,
        showActionsBar: this.properties.showActionsBar,
        showPeers: this.properties.showPeers,
        departmentFilterSelected: this.properties.departmentFilterSelected,
        departmentFilterText: this.properties.departmentFilterText,
        graphClient: this._graphClient,
        sp: this.sp,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // PropertyFieldPeoplePicker (from @pnp/spfx-property-controls) keeps its
  // own internal selection state once mounted, keyed off React's `key`
  // prop. With a fixed key, clearing the field can leave that internal
  // state out of sync with `initialData` — the control never remounts, so
  // it can end up neither showing the person as removed nor accepting a
  // new one. Deriving the key from the current value (present vs. empty,
  // and which person) forces a clean remount exactly when the selection
  // actually changes — not while the user is merely typing in the search
  // box — which keeps it in sync.
  private _peoplePickerKey(
    prefix: string,
    value: IPropertyFieldGroupOrPerson[] | undefined
  ): string {
    const id = value && value[0] && value[0].id;
    return `${prefix}-${id ?? "empty"}`;
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
                PropertyFieldPeoplePicker("selectedUser", {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  label: strings.startFromUserLabel,
                  initialData: this.properties.selectedUser,
                  key: this._peoplePickerKey(
                    "peopleFieldId",
                    this.properties.selectedUser
                  ),
                  multiSelect: false,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this._onSelectedUserChanged,
                  properties: this.properties,
                }),
                PropertyFieldPeoplePicker("coLeadUser", {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  label: strings.coLeadUserLabel,
                  initialData: this.properties.coLeadUser,
                  key: this._peoplePickerKey(
                    "coLeadPeopleFieldId",
                    this.properties.coLeadUser
                  ),
                  multiSelect: false,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                }),
                PropertyPaneTextField("managerLevels", {
                  label: strings.managerLevelsLabel,
                  deferredValidationTime: 300,
                  onGetErrorMessage: (value: string): string => {
                    if (value === undefined || value === "") return "";
                    const parsed = parseInt(value, 10);
                    if (isNaN(parsed) || String(parsed) !== value.trim()) {
                      return "Please enter a whole number";
                    }
                    if (parsed < 0) {
                      return "Must be 0 or higher";
                    }
                    if (parsed > this._maxManagerLevels) {
                      return `Only ${this._maxManagerLevels} level(s) available above this person`;
                    }
                    return "";
                  },
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
          ],
        },
      ],
    };
  }
}