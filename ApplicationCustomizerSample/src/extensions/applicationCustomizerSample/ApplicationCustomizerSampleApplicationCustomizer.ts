import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  // Add 6-7 プレースホルダーを使用するためのクラスをインポート
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'ApplicationCustomizerSampleApplicationCustomizerStrings';

// Add 14-15 スタイルシートとescape処理のインポート（ApplicationCustomizerSample.module.scss を追加して一度 gulp build した後に以下のコードを追加）
import style from './ApplicationCustomizerSample.module.scss';
import {escape} from '@microsoft/sp-lodash-subset';

const LOG_SOURCE: string = 'ApplicationCustomizerSampleApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IApplicationCustomizerSampleApplicationCustomizerProperties {
  // This is an example; replace with your own property
  // Del 27
  //testMessage: string;

  // Add 30-31 ヘッダー、フッターに表示するテキストを保持するためのプロパティ
  topText: string;
  bottomText: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ApplicationCustomizerSampleApplicationCustomizer
  extends BaseApplicationCustomizer<IApplicationCustomizerSampleApplicationCustomizerProperties> { 
  // Add 38-39
  private topPlaceHolder: PlaceholderContent | undefined;
  private bottomPlaceHolder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    //Del 46-51
    // let message: string = this.properties.testMessage;
    // if (!message) {
    //   message = '(No properties were provided.)';
    // }

    // Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);

    //Add 54 ページプレースホルダーの描画処理を呼び出すための指定
    this.context.placeholderProvider.changedEvent.add(this, this.renderPlaceHolders);

    return Promise.resolve();
  }

  //Add 60-102 
  private renderPlaceHolders(): void {
    // ページ上部のプレースホルダーの処理
    // プレースホルダーが存在する場合だけレンダリングするようにしている
    if (!this.topPlaceHolder) {
      this.topPlaceHolder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

      if (!this.topPlaceHolder) {
        return;
      }

      if (this.properties) {
        // topText プロパティで指定された値をヘッダーとして表示する
        let topString: string = this.properties.topText;
        if (this.topPlaceHolder.domElement) {
          this.topPlaceHolder.domElement.innerHTML = `
            <div class="${style.app}">
              <div class="${style.top}">
                <span>${escape(topString)}</span>
              </div>
            </div>
          `;
        }
      }
    }

    // ページ下部のプレースホルダーの処理
    // プレースホルダーが存在する場合だけレンダリングするようにしている
    if (!this.bottomPlaceHolder) {
      this.bottomPlaceHolder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);

      if (!this.bottomPlaceHolder) {
        return;
      }
  
      // bottomText プロパティで指定された値をフッターとして表示する
      if (this.properties) {
        let bottomString: string = this.properties.bottomText;
        if (this.bottomPlaceHolder.domElement) {
          this.bottomPlaceHolder.domElement.innerHTML = `
            <div class="${style.app}">
              <div class="${style.bottom}">
                <span>${escape(bottomString)}</span>
              </div>
            </div>
          `;
        }
      }
    }
  }
}
