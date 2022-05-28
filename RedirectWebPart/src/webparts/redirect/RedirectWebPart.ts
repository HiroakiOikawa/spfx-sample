import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './RedirectWebPart.module.scss';
import * as strings from 'RedirectWebPartStrings';
import { isNumber } from 'lodash';

export interface IRedirectWebPartProps {
  redirectUrl: string;
  waitTime: number;
  escapeString: string;
}

export default class RedirectWebPart extends BaseClientSideWebPart<IRedirectWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  /**
   * HTML描画処理
   * @returns なし
   */
  public render(): void {
   
    // 表示モードがReadの場合だけリダイレクト処理を実行する。

    if (this.displayMode === DisplayMode.Read) {
      if (this.properties.redirectUrl.length === 0) {
        return;
      }
  
      if (this.properties.escapeString.length > 0 && window.location.href.lastIndexOf(this.properties.escapeString) !== -1) {
        return;
      }  
  
      setTimeout(() => {
          window.location.href = this.properties.redirectUrl;
      }, this.properties.waitTime);

    } else {
      this.domElement.innerHTML = `
      <div>Redirect Web Part</div>
      `;
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * プロパティパネルにおけるWaiteTimeプロパティの入力チェック
   * @param value プロパティの入力値
   * @returns エラーメッセージ
   */
  private async validateWaitTime(value: string): Promise<string> {
    if (value != null && value.length > 0 ) {
      if (!isNaN(Number(value))) {
        return "";
      } else {
        return strings.WaitTimeValueErrorMessage;
      }
    } else {
      return "";
    }
  }

  /**
   * プロパティパネルの定義
   * @returns プロパティパネルの定義
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('redirectUrl', {
                  label: strings.RedirectUrlFieldLabel
                }),
                PropertyPaneTextField('waitTime', {
                  label: strings.WaitTimeFieldLabel,
                  description: strings.WaitTimeDescription,
                  onGetErrorMessage: this.validateWaitTime.bind(this)
                }),
                PropertyPaneTextField('escapeString', {
                  label: strings.EscapeStringFieldLabel,
                  description: strings.EscapeStringDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
