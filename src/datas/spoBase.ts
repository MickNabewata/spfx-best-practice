import IDatas, { IServerData, IData } from './iDatas';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/** SharePoint リストアイテム サーバー用 */
export interface IServerListItem extends IServerData {
    /** アイテムID */
    Id: number;
    /** タイトル */
    Title: string;
    /** 登録者Id */
    AuthorId: number;
    /** 登録日時 */
    Created: string;
    /** 更新者Id */
    EditorId: number;
    /** 更新者 */
    Modified: string;
}

/** SharePoint リストアイテム クライアント用 */
export interface IListItem extends IData {
    /** アイテムID */
    Id: number;
    /** タイトル */
    Title: string;
    /** 登録者 */
    Author: ISiteUser;
    /** 登録日時 */
    Created: Date;
    /** 更新者 */
    Editor: ISiteUser;
    /** 更新者 */
    Modified: Date;
}

/** サイトのユーザー サーバー用 */
export interface IServerSiteUser extends IServerData {
    /** ID */
    Id: number;
    /** 表示名 */
    Title: string;
    /** メールアドレス */
    Email: string;
}

/** サイトのユーザー クライアント用 */
export interface ISiteUser extends IData {
    /** ID */
    Id: number;
    /** 表示名 */
    Title: string;
    /** メールアドレス */
    Email: string;
}

/** データ操作クラス 基底 with SharePoint Client */
export default class SpoBase
    <ServerData extends IServerData, Data extends IData> 
    implements IDatas<ServerData, Data> 
{
    /** SharePoint 問い合わせ用クライアント */
    protected _client : SPHttpClient;

    /** SharePoint WebサイトのURL */
    protected _webUrl : string;

    /** データ操作クラス 基底 with Microsoft Graph */
    constructor(client: SPHttpClient, webUrl: string) {
        // コンストラクタ引数を保存
        this._client = client;
        this._webUrl = webUrl;
    }

    /** データ取得 */
    public get(): Promise<Data[]> {
        return this.getDatas().then(
            (datas) => {
                return this.convertDatas(datas);
            },
            (err) => {
                return Promise.reject(err);
            }
        );
    }

    /** サーバーからのデータ取得 */
    protected getDatas(): Promise<ServerData[]> {
        return Promise.resolve([]);
    }

    /** サーバーから取得したデータをコード内で使用する型に変換 */
    protected convertDatas(datas: ServerData[]): Promise<Data[]> {
        let ret: Data[] = undefined;

        if(datas) {
            ret = [];
            datas.forEach((data) => {
                ret.push(this.convertData(data));
            });
        }

        return Promise.resolve(ret);
    }

    /** サーバーから取得したデータ1件をコード内で使用する型に変換 */
    protected convertData(data: ServerData): Data {
        return undefined;
    }

    /** リストアイテム型の共通項目をコード内で使用する型に変換 */
    protected convertDefaultListItemData(users: ISiteUser[], data: IServerListItem): IListItem {
        let ret: IListItem = undefined;

        if(data) {
            ret = {
                Id: data.Id,
                Title: data.Title,
                Author: this.getUser(users, data.AuthorId),
                Created: new Date(data.Created),
                Editor: this.getUser(users, data.EditorId),
                Modified: new Date(data.Modified)
            };
        }

        return ret;
    }

    /** SharePoint リクエスト実行 */
    protected doSpoRequest<T>(endPoint: string): Promise<T> {
        return new Promise((resolve: (tasks: T) => void, reject: (reason: any) => void) => {
            if(this._client) {
                this._client.get(endPoint, SPHttpClient.configurations.v1).then(
                    (response: SPHttpClientResponse) => {
                        if(response && response.ok) {
                            response.json().then(
                                (json) => {
                                    if(json) {
                                        if(json.value) {
                                            resolve(json.value);
                                        } else {
                                            resolve(json);
                                        }
                                    } else {
                                        reject('json is null.');
                                    }
                                },
                                (err) => { 
                                    reject(err);
                                 }
                            );
                        } else {
                            if(!response) {
                                reject('response is null.');
                            } else {
                                response.json().then(
                                    (errRes) => {
                                        reject((errRes && errRes.error && errRes.error.message)? errRes.error.message : JSON.stringify(errRes));
                                    },
                                    (err) => {
                                        reject(JSON.stringify(err));
                                    }
                                );
                            }
                        }
                    },
                    (err) => { 
                        reject(err);
                     }
                );
            } else {
                reject('spo client is null.');
            }
        });
    }

    /** ユーザーを取得 */
    protected getUser(users: ISiteUser[], id: number): ISiteUser {
        let ret: ISiteUser;
        
        if(users && id) {
            const user = users.filter((v) => { return (v.Id === id); });
            if(user && user.length > 0) {
                ret = user[0];
            }
        }

        return ret;
    }
}