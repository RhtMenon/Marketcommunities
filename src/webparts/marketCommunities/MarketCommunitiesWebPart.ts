import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { HttpClient } from '@microsoft/sp-http';

import styles from './MarketCommunitiesWebPart.module.scss';
import * as strings from 'MarketCommunitiesWebPartStrings';

import {app} from '@microsoft/teams-js'; 

export interface IMarketCommunitiesWebPartProps {
  description: string;
  numberOfBlocks: number;
  selectedList: string;
  availableLists: { key: string; text: string }[];
  seeAllButton: string;
}

export default class MarketCommunitiesWebPart extends BaseClientSideWebPart<IMarketCommunitiesWebPartProps> {
  private communityInfo: any[] = [];

  private isTeams = false;
  private isEmbedded = false;
  protected async onInit(): Promise<void> {
    try {
      await app.initialize();
      const context = await app.getContext();
      console.log("Context:", context);
      if(context.app.host.name.includes("teams") || context.app.host.name.includes("Teams")){
        console.log("The extension is running inside Microsoft Teams");
        this.isTeams = true;
      }else{
        console.log("The extension is running outside Microsoft Teams");
      }
    } catch (exp) {
        console.log("The extension is running outside Microsoft Teams");
  }
  this.isEmbedded = document.body.classList.contains('embedded');
  if (this.isEmbedded) {
    console.log('Body has the embedded class');
  } else {
    console.log('Body does not have the embedded class');
  }
    await this.getAvailableLists();
    await this.getCommunityInfo();
  }

  public render(): void {
    const decodedSeeAllButton = decodeURIComponent(this.properties.seeAllButton);
    console.log("Url for See All button: ", decodedSeeAllButton);
  
    this.domElement.innerHTML = `
    <div id="MarketPlaceParentDiv">
        <div class="${styles.topSection}">
          <div class="${styles.MarketCommunitiesHeading}">Marketplace Communities</div>
          <div><a href="${decodedSeeAllButton}" target="_blank">See all</a></div>
        </div>
        <section class="${styles.MarketCommunities}">
          ${this.communityInfo.slice(0, this.properties.numberOfBlocks).map((group: any, index: number) => {
  
            // Check if webUrl is an object and use the Url property if available
            const href = (group.webUrl && typeof group.webUrl === 'object') ? group.webUrl.Url || '' : group.webUrl || '';

            console.log("Group WebUrl:", group.webUrl.Url);
            let fullLink = "";
            if (group.webUrl.Url && group.webUrl.Url.includes("groups/")) {
                let splitUrl = group.webUrl.Url.split("groups/")[1];
                console.log("SplitUrl:", splitUrl);
                let teamsLink1 = "https://teams.microsoft.com/l/entity/db5e5970-212f-477f-a3fc-2227dc7782bf/vivaengage?context=%7B%22subEntityId%22:%22type=custom,data=group:";
                let teamsLink2 = "%22%7D";
                fullLink = teamsLink1 + splitUrl + teamsLink2;
                console.log("Full Link:", fullLink);
            }

            let link = "";

            if(this.isTeams && this.isEmbedded){
              link = fullLink;
            }else if(!this.isTeams && this.isEmbedded){
              link = "https://aka.ms/VivaEngage/Outlook";
            }else{
              link = group.webUrl.Url;
            }
              
            return `
              <div class="${styles.col} colDiv">
                <div class="${styles.ContentBox}">
                  <div class="${styles.ImgPart}">
                    <a href="${link}" target="_blank">
                      <img src="${group.mugshotUrlTemplate}" alt="${group.fullName}">
                    </a>
                  </div>
                  <div class="${styles.Contents}">
                    <h3 onclick="window.open('${link}')">${group.fullName}</h3>
                    <p>${group.description}</p>
                  </div>
                </div>
              </div>
            `;
          }).join('')}
        </section>
      </div>
    `;

    const handleResponsiveCheck = () => {
        const marketPlaceParent = document.getElementById('MarketPlaceParentDiv'); 
        const colDiv = document.querySelectorAll('.colDiv');

        if (marketPlaceParent && colDiv.length > 0) {
            const windowWidth = marketPlaceParent.getBoundingClientRect().width;

            if (windowWidth < 500) {
                colDiv.forEach(element => {
                    element.classList.add(styles.ColSmallerWidthSm);
                    element.classList.remove(styles.ColSmallerWidthMd);
                    element.classList.remove(styles.ColSmallerWidthLg);
                });
              }  else if (windowWidth < 700) {
                colDiv.forEach(element => {
                    element.classList.add(styles.ColSmallerWidthMd);
                    element.classList.remove(styles.ColSmallerWidthSm);
                    element.classList.remove(styles.ColSmallerWidthLg);
                });
            }
            else if (windowWidth < 900) {
              colDiv.forEach(element => {
                  element.classList.add(styles.ColSmallerWidthLg);
                  element.classList.remove(styles.ColSmallerWidthSm);
                  element.classList.remove(styles.ColSmallerWidthMd);
              });

          } else {
                colDiv.forEach(element => {
                    element.classList.remove(styles.ColSmallerWidthSm);
                    element.classList.remove(styles.ColSmallerWidthMd);
                    element.classList.remove(styles.ColSmallerWidthLg);
                });
            }
        }
    };

    handleResponsiveCheck();
    window.addEventListener('resize', handleResponsiveCheck);
}

  

  private async getAvailableLists() {
    try {
      const response = await this.context.httpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
        HttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
          },
        }
      );
  
      const data = await response.json();
      this.properties.availableLists = data.value.map((list: any) => ({
        key: list.Title,
        text: list.Title,
      }));
  
      // Update the property pane dropdown options
      this.context.propertyPane.refresh();
    } catch (error) {
      console.error('Error fetching available lists:', error);
    }
  }
  
  

  private async getCommunityInfo() {
    if (!this.properties.selectedList) {
      // Do not proceed if no list is selected
      return;
    }
  
    try {
      const response = await this.context.httpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.selectedList}')/items?$filter=isMarketPlace eq 1`,
        HttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
          },
        }
      );
  
      const data = await response.json();
  
      if (data && data.value) {
        this.communityInfo = data.value.map((item: any) => {
          console.log('Mugshot URL from SharePoint:', item.MugShotURL);
  
          return {
            fullName: item.Title,
            description: item.CommunityDescription,
            webUrl: item.CommunityURL,
            mugshotUrlTemplate: item.MugShotURL,
          };
        });
      } else {
        console.error('Data not found in the SharePoint list or no items match the filter condition.');
      }
    } catch (error) {
      console.error('Error fetching community information:', error);
    }
  
    this.render(); // Render the web part after fetching data
  }
  
  

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneDropdown('selectedList', {
                  label: 'Select SharePoint List',
                  options: this.properties.availableLists,
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneSlider('numberOfBlocks', {
                  label: 'Number of Blocks',
                  min: 1,
                  max: 10,
                  step: 1,
                }),
                PropertyPaneTextField('seeAllButton',{

                  label: 'Url for See All button'
                })
              ],
            },
          ],
        },
      ],
    };
  }  
}
