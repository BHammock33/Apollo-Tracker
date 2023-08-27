import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneTextFieldProps,
  PropertyPaneLabel,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NasaApolloMissionViewerWebPart.module.scss';
import * as strings from 'NasaApolloMissionViewerWebPartStrings';

import { IMission } from '../../models';
import { MissionService } from '../../services';

export interface INasaApolloMissionViewerWebPartProps {
  description: string;
  selectedMission: string;
}

export default class NasaApolloMissionViewerWebPart extends BaseClientSideWebPart<INasaApolloMissionViewerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  // select a specific mission
  private _selectedMission: IMission;
  // dom element where the mission details will go
  private _missionDetailElement: HTMLElement;

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.nasaApolloMissionViewer} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <h2>Apollo Mission Viewer</h2>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
        <div class="apolloMissionDetails"></div>
      </div>
    </section>`;

    // get reference to the HTML element where we will show mission details
    this._missionDetailElement = this.domElement.getElementsByClassName('apolloMissionDetails')[0] as HTMLElement;

    // show mission if one found, else show error
    if (this._selectedMission) {
      this._renderMissionDetails(this._missionDetailElement, this._selectedMission);
    } else {
      // this._renderMissionLoadingError();
      this._missionDetailElement.innerHTML = '';
    }

  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    this._selectedMission = this._getSelectedMission();

    return super.onInit();
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        //page 1
        {
          header: {
            description: 'About this web Part'
          },
          groups: [
            {
              groupFields: [
                PropertyPaneLabel('',{
                  text: 'great first web part'
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.BasicGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('selectedMission', <IPropertyPaneTextFieldProps> {
                  label: 'Apollo Mission to Show',
                  onGetErrorMessage: this._validateMissionCode.bind(this)
                })
              ]
            },
            //group one
            //group two
            {
              groupName: 'group 2',
              groupFields: [
                PropertyPaneLabel('',{
                text: 'This is the second group'})
              ]
            }
          ]//groups 
        }
        //page 2
      ] //pages
    };
  } //getPropertyPaneConfiguration()

  private _validateMissionCode(value: string): string{
    const validMissionCodeRegEx = /AS-[2,5][0,1][0-9]/g
    return value.match(validMissionCodeRegEx)
    ? ''
    : 'invalid mission code: shoudl be \'AS-###\'.';
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    this._selectedMission = this._getSelectedMission();

    if(this._selectedMission){
      this._renderMissionDetails(this._missionDetailElement, this._selectedMission);
    }else{
      this._missionDetailElement.innerHTML= '';
    }
}  
  /**
     * Get the selected Apollo mission.
     *
     * @private
     * @returns   {IMission}  Mission selected (null if no mission returned).
     * @memberof  ApolloViewerWebPart
     */
  private _getSelectedMission(): IMission {
    // determine the mission ID, defaulting to Apollo 11
    const selectedMissionId: string = (this.properties.selectedMission)
      ? this.properties.selectedMission
      : 'AS-506';

    // get the specified mission
    return MissionService.getMission(selectedMissionId);
  }
 

 

 

  /**
   * Display the specified mission details in the provided DOM element.
   *
   * @private
   * @param {HTMLElement} element   DOM element where the details should be written to.
   * @param {IMission}    mission   Apollo mission to display.
   * @memberof ApolloViewerWebPart
   */
  private _renderMissionDetails(element: HTMLElement, mission: IMission): void {
    element.innerHTML = `
    <p class="ms-font-m">
      <span class="ms-fontWeight-semibold">Mission: </span>
      ${escape(mission.name)}
    </p>
    <p class="ms-font-m">
      <span class="ms-fontWeight-semibold">Duration: </span>
      ${escape(this._getMissionTimeline(mission))}
    </p>
    <a href="${mission.wiki_href}" target="_blank">
      <span>Learn more about ${escape(mission.name)} on Wikipedia &raquo;</span>
    </a>`;
  }

  /**
   * Returns the duration of the mission.
   *
   * @private
   * @param     {IMission}  mission  Apollo mission to use.
   * @returns   {string}             Mission duration range in the format of [start] - [end].
   * @memberof  ApolloViewerWebPart
   */
  private _getMissionTimeline(mission: IMission): string {
    return (mission.end_date !== '')
      ? `${mission.launch_date.toString()} - ${mission.end_date.toString()}`
      : `${mission.launch_date.toString()}`;
  }

}
