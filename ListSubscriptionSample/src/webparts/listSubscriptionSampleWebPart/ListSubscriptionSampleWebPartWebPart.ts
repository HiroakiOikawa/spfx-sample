import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListSubscriptionSampleWebPartWebPart.module.scss';
import * as strings from 'ListSubscriptionSampleWebPartWebPartStrings';

//Add 13-14 変更通知を受信するためのライブラリのインポート
import {ListSubscriptionFactory, IListSubscription} from '@microsoft/sp-list-subscription';
import {Guid} from '@microsoft/sp-core-library';

export interface IListSubscriptionSampleWebPartWebPartProps {
  description: string;
  //Add 19 変更通知を購読するリストのリストIDを格納するプロパティ
  listId: string;
}

export default class ListSubscriptionSampleWebPartWebPart extends BaseClientSideWebPart <IListSubscriptionSampleWebPartWebPartProps> {

  //Add 25-40 変更通知を購読するための一連の処理
  private listSubscriptionFactory: ListSubscriptionFactory;
  private listSubscription: Promise<IListSubscription>; 

  private createListSubscription(): void {
    if (this.properties.listId) {
      this.listSubscriptionFactory = new ListSubscriptionFactory(this);
      this.listSubscription = this.listSubscriptionFactory.createSubscription({
        listId: Guid.parse(this.properties.listId),
        callbacks: {
          connect: this.connect.bind(this),
          notification: this.notificate.bind(this),
          disconnect: this.disconnect.bind(this)
        }
      });
    }
  }

  //Add 43 Web パーツ上に表示するメッセージを格納する変数
  private message: string;

  //Add 46-49 変更通知の購読を開始した際のコールバック処理
  private connect(): void {
    this.message = "ライブラリの変更通知購読開始";
    this.render();
  }

  //Add 52-55 変更通知を受信した際のコールバック処理
  private notificate(): void {
    this.message = "ライブラリから変更通知を受信";
    this.render();
  }

  //Add 58-61 変更通知の購読が終了した際のコールバック処理
  private disconnect(): void {
    this.message = "ライブラリの変更通知購読終了";
    this.render();
  }

  //Add 64-68 変更通知の購読処理を呼び出すためにOnInitをオーバーライド
  protected onInit(): Promise<any> {
    this.createListSubscription();
    let retVal: Promise<any> = Promise.resolve();
    return retVal;
  }

  //Modify 71-84 Web パーツのHTMLレンダリング処理
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.listSubscriptionSampleWebPart }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">変更通知購読</span>
              <p class="${ styles.subTitle }">対象リスト：${this.properties.listId}</p>
              <p class="${ styles.subTitle }">通知状態　：${this.message}</p>
            </div>
          </div>
        </div>
      </div>`;
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
                label: strings.DescriptionFieldLabel
              }),
              // Add 104-106 変更通知購読対象となるリストのリストIDを指定するためのプロパティを定義
              PropertyPaneTextField('listId', {
                label: "List ID"
              })
            ]
          }
        ]
      }
    ]
  };
}
}
