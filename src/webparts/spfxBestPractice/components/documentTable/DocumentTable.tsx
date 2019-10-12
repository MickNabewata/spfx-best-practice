import * as React from 'react';
import styles from './DocumentTable.module.scss';
import SpfxBestPractice from '../spfxBestPractice/SpfxBestPractice';
import { IServerDocument, IDocument } from '../../../../datas/documents';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';

/** カードレイアウト */
export default class DocumentTable extends SpfxBestPractice<IServerDocument, IDocument> {

  /** テーブル列の定義 */
  protected getColumns(): IColumn[] {
    return [
      {
        key: 'Name',
        name: 'ファイル名',
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        onRender: (item: IDocument) => {
          return <a href={item.ServerRelativeUrl} target='_blank'>{item.Name}</a>;
        },
        isPadded: true
      },
      {
        key: 'Editor',
        name: '更新者',
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: false,
        isResizable: true,
        data: 'string',
        onRender: (item: IDocument) => {
          return item.Editor.Title;
        },
        isPadded: true
      },
      {
        key: 'Modified',
        name: '更新日時',
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: false,
        isResizable: true,
        data: 'string',
        onRender: (item: IDocument) => {
          return item.Modified.toLocaleString();
        },
        isPadded: true
      }
    ];
  }

  /** データのレンダリング */
  protected renderDatas(): JSX.Element {
    return (
      <Fabric>
        <DetailsList
          items={this.state.datas}
          compact={false}
          columns={this.getColumns()}
          selectionMode={SelectionMode.none}
          setKey='none'
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
        />
      </Fabric>
    );
  }
}
