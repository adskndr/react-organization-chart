import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import * as strings from "PositionDirectoryWebPartStrings";
import { PositionDirectory } from "../../components/PositionDirectory/PositionDirectory";
import { IPositionDirectoryProps } from "../../components/PositionDirectory/IPositionDirectoryProps";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/search";
import { IGraphPhotoClient } from "../../services/PhotoService";

let _sp: SPFI;

export interface IPositionDirectoryWebPartProps {
  // Comma-separated so editors can cover spelling variants in one field,
  // e.g. "Lernende, Lernender".
  jobTitles: string;
  groupByDepartment: boolean;
  showActionsBar: boolean;
}

export default class PositionDirectoryWebPart extends BaseClientSideWebPart<IPositionDirectoryWebPartProps> {
  private _graphClient?: IGraphPhotoClient;

  public async onInit(): Promise<void> {
    _sp = spfi().using(SPFx(this.context));
    try {
      this._graphClient = await this.context.msGraphClientFactory.getClient("3");
    } catch (error) {
      console.log(error);
    }
    return super.onInit();
  }

  public get sp(): SPFI {
    return _sp;
  }

  private get _jobTitles(): string[] {
    return (this.properties.jobTitles ?? "")
      .split(",")
      .map((t) => t.trim())
      .filter(Boolean);
  }

  public render(): void {
    const element: React.ReactElement<IPositionDirectoryProps> = React.createElement(
      PositionDirectory,
      {
        context: this.context,
        jobTitles: this._jobTitles,
        groupByDepartment: this.properties.groupByDepartment,
        showActionsBar: this.properties.showActionsBar,
        graphClient: this._graphClient,
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
                PropertyPaneTextField("jobTitles", {
                  label: strings.jobTitlesLabel,
                  description: strings.jobTitlesDescription,
                }),
                PropertyPaneToggle("groupByDepartment", {
                  label: strings.groupByDepartmentLabel,
                }),
                PropertyPaneToggle("showActionsBar", {
                  label: strings.showActionsBarLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
