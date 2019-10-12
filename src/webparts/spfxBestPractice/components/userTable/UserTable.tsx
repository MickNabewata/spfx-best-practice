import * as React from 'react';
import styles from './UserTable.module.scss';
import SpfxBestPractice from '../spfxBestPractice/SpfxBestPractice';
import { IServerSiteUser, ISiteUser } from '../../../../datas/siteUsers';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';

/** カードレイアウト */
export default class UserTable extends SpfxBestPractice<IServerSiteUser, ISiteUser> {

  /** テーブル列の定義 */
  protected getColumns(): IColumn[] {
    return [
      {
        key: 'Title',
        name: 'ユーザー名',
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        onRender: (item: ISiteUser) => { return item.Title; },
        isPadded: true
      },
      {
        key: 'Email',
        name: 'メールアドレス',
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: false,
        isResizable: true,
        data: 'string',
        onRender: (item: ISiteUser) => { return <a href={`mailto:${item.Email}`}>{item.Email}</a>; },
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
