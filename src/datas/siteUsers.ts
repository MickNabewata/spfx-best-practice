import IDatas from './iDatas';
import { IServerSiteUser, ISiteUser } from './spoBase';
import ServerSiteUsers from './server/siteUsers';
import TestSiteUsers from './test/siteUsers';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';

/** サイトのユーザー サーバー用 */
export interface IServerSiteUser extends IServerSiteUser {}

/** サイトのユーザー */
export interface ISiteUser extends ISiteUser {}

/** SharePoint Online サイトユーザー操作クラス */
export default class SiteUsers implements IDatas<IServerSiteUser, ISiteUser> {

    /** データ操作クラスインスタンス */
    protected _instance: IDatas<IServerSiteUser, ISiteUser>;

    /** SharePoint Online サイトユーザー操作クラス */
    constructor(client: SPHttpClient, webUrl: string) {

        // 実行環境によってテストデータを返却するかサーバーデータを返却するかを判定
        switch(Environment.type) {
            case EnvironmentType.SharePoint:
            case EnvironmentType.ClassicSharePoint:
                this._instance = new ServerSiteUsers(client, webUrl);
                break;
            default:
                this._instance = new TestSiteUsers(client, webUrl);
                break;
        }
    }

    /** サイトユーザー取得 */
    public get(): Promise<ISiteUser[]> {
        return this._instance.get();
    }
}