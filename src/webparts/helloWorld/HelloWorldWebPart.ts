import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import type { IReadonlyTheme } from "@microsoft/sp-component-base";

import styles from "./HelloWorldWebPart.module.scss";
import * as strings from "HelloWorldWebPartStrings";
import { escape } from "@microsoft/sp-lodash-subset";

export interface IHelloWorldWebPartProps {
  heading: string;
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    this.domElement.innerHTML = `
      <section class="${styles.helloWorld} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ""
    }">
        <div class="${styles.welcome}">
          <img alt="" src="${
            this._isDarkTheme
              ? require("./assets/welcome-dark.png")
              : require("./assets/welcome-light.png")
          }" class="${styles.welcomeImage}" />
          <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
          <div>${this._environmentMessage}</div>
        </div>
       <div class="${styles.container}">
        <div class="${styles.header}">
          <h1>${escape(this.properties.heading)}</h1>
          <p>${escape(this.properties.description)}</p>
          <a href="#" class="${styles.btn}">Get started</a>
        </div>
        <div class="${styles.imagesGrid}">
          <div class="${styles.imageItem} ${styles.large}">
            <img src="${require("./assets/left.jpg")}">
          </div>
          <div class="${styles.imageItem}">
            <img src="${require("./assets/right.jpg")}">
          </div>
        </div>
        <div class="${styles.imageCard}>
          <div class="${styles.numberedImage}">
            <h1>01</h1>
            <img src="${require("./assets/01.jpg")}">
          </div>
          <div class="${styles.numberedImage}">>
            <h1>02</h1>
            <img src="${require("./assets/02.jpg")}">
          </div>
          <div class="${styles.numberedImage}">>
            <h1>03</h1>
            <img src="${require("./assets/03.jpg")}">
          </div>
        </div>
       </div>
      </section>`;
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
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
              groupName: "Contents",
              groupFields: [
                PropertyPaneTextField("description", {
                  label: "Heading",
                }),
                PropertyPaneTextField("test", {
                  label: "Description",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
