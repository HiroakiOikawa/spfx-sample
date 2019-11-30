import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

// add-1 L10-L13
import * as MSTeams from '@microsoft/teams-js';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MSGraphBeta from '@microsoft/microsoft-graph-types-beta';

import styles from './TeamsTabSampleWebPart.module.scss';
import * as strings from 'TeamsTabSampleWebPartStrings';

export interface ITeamsTabSampleWebPartProps {
  description: string;
  // add-2 L21-22
  teamId: string;
  channelId: string;
}

export default class TeamsTabSampleWebPart extends BaseClientSideWebPart<ITeamsTabSampleWebPartProps> {
  // add-3 L27
  private teamsContext: MSTeams.Context;

  // add-4 L30-L41
  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this.teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  public render(): void {
    // add-5 L45-L58
    let teamId: string = '';
    let channelId: string = '';
 
    if (this.properties.teamId == '' && this.teamsContext) {
      teamId = this.teamsContext.teamId;
    } else {
      teamId = this.properties.teamId;
    }

    if (this.properties.channelId == '' && this.teamsContext) {
      channelId = this.teamsContext.channelId;
    } else {
      channelId = this.properties.channelId;
    }

    // add-6 L61-125
    this.context.msGraphClientFactory.getClient()
    .then((client: MSGraphClient) => {
      new Promise<MSGraphBeta.ChatMessage[]>((resolve: (value?: MSGraphBeta.ChatMessage[]) => void, reject:(reason?: any) => void) => {
        client.api(`https://graph.microsoft.com/beta/teams/${teamId}/channels/${channelId}/messages`).get((error, response: any, rawResponse?: any) => {
          if (response == null) reject(error);
          let messages: MSGraphBeta.ChatMessage[] = new Array();
          for (let index = 0; index < response.value.length; index++) {
            const msg: MSGraphBeta.ChatMessage = response.value[index];
            if (msg.deletedDateTime == null) messages.push(msg);
          }
          resolve(messages);
        });
      })
      .then((messages: MSGraphBeta.ChatMessage[]) => {
        return new Promise<string>((resolve: (value?: string) => void, reject:(reson?: any) => void) => {
          let html: string = "";
          function getMessage(index: number) {
            const message:MSGraphBeta.ChatMessage = messages[index];
            let title: string = message.subject;
            if (title == "") title = message.body.content.substr(0, 30);
            let replyCount: number = 0;
            new Promise<string>((res:(value?: string) => void, rej:(error?: any) => void) => {
              client.api(`https://graph.microsoft.com/beta/teams/${teamId}/channels/${channelId}/messages/${message.id}/replies`).get((error, response: any, rawResponse?: any) => {
                if (response != null) {
                  replyCount = response.value.length;
                }
                let rowHtml: string = `
                <tr>
                  <td><a href="${message.webUrl}" target="_blank">${title}</a></td>
                  <td>${message.from.user.displayName}</td>
                  <td>${message.createdDateTime}</td>
                  <td>${replyCount}</td>
                </tr>`;
                res(rowHtml);
              });
            }).then((rowHtml: string) => {
              html += rowHtml;
              if (index >= messages.length - 1) {
                resolve(html);
              } else {
                getMessage(index+1);
              }
            });
          }

          getMessage(0);
        });
      })
      .then((html:string) => {
        this.domElement.innerHTML = `
          <table style='border: 1px solid black;'>
            <thead style='border: 1px solid black;background-color:lightgray;'>
              <tr>
                <th>タイトル</th>
                <th>投稿者</th>
                <th>投稿日時</th>
                <th>返信数</th>
              </tr>
            </thead>
            <tbody>
              ${html}
            </tbody>
          </table>`;
      });
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  // add-7 L145の","-L151
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyPaneTextField('teamId', {
                  label: strings.TeamIdFieldLabel
                }),
                PropertyPaneTextField('channelId', {
                  label: strings.ChannelIdFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
