import IDatas from './iDatas';
import { IServerListItem, IListItem } from './spoBase';
import ServerDocuments from './server/documents';
import TestDocuments from './test/documents';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';

/** ドキュメント サーバー用 */
export interface IServerDocument extends IServerListItem {
    /** ファイル名 */
    Name: string;
    /** ファイル相対パス */
    ServerRelativeUrl: string;
    /** バージョン */
    UIVersionLabel: string;
}

/** ドキュメント */
export interface IDocument extends IListItem {
    /** ファイル名 */
    Name: string;
    /** ファイル相対パス */
    ServerRelativeUrl: string;
    /** バージョン */
    UIVersionLabel: string;
    /** プレビュー画像URL */
    PreviewImageUrl: string;
}

/** SharePoint Online ドキュメント操作クラス */
export default class Documents implements IDatas<IServerDocument, IDocument> {

    /** データ操作クラスインスタンス */
    protected _instance: IDatas<IServerDocument, IDocument>;

    /** リストタイトル */
    protected _listTitle: string;

    /** SharePoint Online ドキュメント操作クラス */
    constructor(client: SPHttpClient, webUrl: string, listTitle: string) {

        // コンストラクタ引数を保存
        this._listTitle = listTitle;

        // 実行環境によってテストデータを返却するかサーバーデータを返却するかを判定
        switch(Environment.type) {
            case EnvironmentType.SharePoint:
            case EnvironmentType.ClassicSharePoint:
                this._instance = new ServerDocuments(client, webUrl, listTitle);
                break;
            default:
                this._instance = new TestDocuments(client, webUrl);
                break;
        }
    }

    /** ドキュメント取得 */
    public get(): Promise<IDocument[]> {
        return this._instance.get();
    }
}