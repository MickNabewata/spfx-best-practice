import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneDropdown } from '@microsoft/sp-webpart-base';
import * as strings from 'SpfxBestPracticeWebPartStrings';
import UserCard from './components/userCard/UserCard';
import DocumentCard from './components/documentCard/DocumentCard';
import UserTable from './components/userTable/UserTable';
import DocumentTable from './components/documentTable/DocumentTable';
import SiteUsers from '../../datas/siteUsers';
import Documents from '../../datas/documents';

/** マニフェストで定義したプロパティの型定義 */
export interface ISpfxBestPracticeWebPartProps {
  /** タイトル */
  title: string;
  /** データの種類 */
  dataType: 'SiteUsers' | 'Documents';
  /** 描画の形式 */
  variant: 'Card' | 'Table';
}

/** ベストプラクティス集 Webパーツ */
export default class SpfxBestPracticeWebPart extends BaseClientSideWebPart<ISpfxBestPracticeWebPartProps> {

  /** レンダリング */
  public render(): void {

    // データの種類で分岐
    switch(this.properties.dataType) {
      case 'SiteUsers':
        const siteUsers = new SiteUsers(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
        
        // 描画の形式で分岐
        switch(this.properties.variant) {
          case 'Card':
            ReactDom.render(
              React.createElement(
                UserCard,
                {
                  title: this.properties.title,
                  mode: this.displayMode,
                  titleChangeCallback: this.titleChangeHandler,
                  client: siteUsers
                }
              ),
              this.domElement);
            break;
          case 'Table':
            ReactDom.render(
              React.createElement(
                UserTable,
                {
                  title: this.properties.title,
                  mode: this.displayMode,
                  titleChangeCallback: this.titleChangeHandler,
                  client: siteUsers
                }
              ),
              this.domElement);
            break;
        }

        break;
      case 'Documents':
        const documents = new Documents(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl, 'ドキュメント');

        switch(this.properties.variant) {
          case 'Card':
              ReactDom.render(
                React.createElement(
                  DocumentCard,
                  {
                    title: this.properties.title,
                    mode: this.displayMode,
                    titleChangeCallback: this.titleChangeHandler,
                    client: documents
                  }
                ),
                this.domElement);
            break;
          case 'Table':
            ReactDom.render(
              React.createElement(
                DocumentTable,
                {
                  title: this.properties.title,
                  mode: this.displayMode,
                  titleChangeCallback: this.titleChangeHandler,
                  client: documents
                }
              ),
              this.domElement);
            break;
        }

        break;
    }
  }

  /** タイトル変更イベント */
  private titleChangeHandler = (title: string) => {
    this.properties.title = title;
  }

  /** プロパティウィンドウ定義 */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown('dataType', {
                  label: strings.DataTypeLabel,
                  options: [
                    { key: 'SiteUsers', text: strings.DataTypeSiteUser },
                    { key: 'Documents', text: strings.DataTypeDocument }
                  ]
                }),
                PropertyPaneDropdown('variant', {
                  label: strings.VariantLabel,
                  options: [
                    { key: 'Card', text: strings.VariantCard },
                    { key: 'Table', text: strings.VariantTable }
                  ],
                })
              ]
            }
          ]
        }
      ]
    };
  }

  /** 破棄イベント */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /** バージョン取得 */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
