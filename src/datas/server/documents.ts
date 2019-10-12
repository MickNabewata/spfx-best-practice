import SiteUsers from '../siteUsers';
import { IServerDocument, IDocument } from '../documents';
import SpoBase, { ISiteUser } from '../spoBase';
import { SPHttpClient } from '@microsoft/sp-http';
import { Promise } from 'es6-promise';

/** SharePoint Online ドキュメント操作クラス サーバー接続用 */
export default class ServerDocuments extends SpoBase<IServerDocument, IDocument> {

    /** リストタイトル */
    protected _listTitle: string;

    /** サイトのユーザー一覧 */
    protected _users: ISiteUser[];

    /** SharePoint Online ドキュメント操作クラス */
    constructor(client: SPHttpClient, webUrl: string, listTitle: string) {
        super(client, webUrl);

        // コンストラクタ引数を保存
        this._listTitle = listTitle;
    }

    /** ドキュメント取得 */
    protected getDatas(): Promise<IServerDocument[]> {
        const siteUsers = new SiteUsers(this._client, this._webUrl);
        return siteUsers.get().then(
            (users) => {
                this._users = users;
                const itemsEndpoint = `${this._webUrl}/_api/web/lists/getbytitle('${this._listTitle}')/items`;
                return this.doSpoRequest<IServerDocument[]>(itemsEndpoint).then(
                    (documents) => {
                        let getFileRequests: Promise<IServerDocument>[] = [];
                        if(documents) {
                            documents.forEach((document) => {
                                getFileRequests.push(
                                    new Promise<IServerDocument>((resolve: (result: IServerDocument) => void, reject: (err: any) => void) => {
                                        this.doSpoRequest<IServerDocument>(`${itemsEndpoint}(${document.Id})/file`).then(
                                            (file) => {
                                                if(file) {
                                                    document.Name = file.Name;
                                                    document.ServerRelativeUrl = file.ServerRelativeUrl;
                                                    document.UIVersionLabel = file.UIVersionLabel;
                                                    resolve(document);
                                                } else {
                                                    reject('file is null.');
                                                }
                                            },
                                            (err) => {
                                                reject(err);
                                            }
                                        );
                                    })
                                );
                            });
                        }

                        return Promise.all(getFileRequests);
                    },
                    (err) => {
                        return Promise.reject(err);
                    }
                );
            },
            (err) => {
                return Promise.reject(err);
            }
        );
    }

    /** サーバーから取得したデータ1件をコード内で使用する型に変換 */
    protected convertData(data: IServerDocument): IDocument {
        if(data) {
            return {
                Id: data.Id,
                Title: data.Title,
                Name: data.Name,
                ServerRelativeUrl: data.ServerRelativeUrl,
                UIVersionLabel: data.UIVersionLabel,
                PreviewImageUrl: this.getPreviewImageUrl(data),
                Author: this.getUser(this._users, data.AuthorId),
                Created: new Date(data.Created),
                Editor: this.getUser(this._users, data.EditorId),
                Modified: new Date(data.Modified)
            } as IDocument;
        } else {
            return undefined;
        }
    }

    /** プレビュー画像URLの算出 */
    private getPreviewImageUrl(data: IServerDocument): string {
        if(data && data.Name && data.ServerRelativeUrl) {
            const baseUrl = `${this._webUrl}/_layouts/15/getpreview.ashx`;
            const path = `path=${data.ServerRelativeUrl}`;
            const resolution = 'resolution=1';
            const force = 'force=1';
            return `${baseUrl}?${path}&${resolution}&${force}`;
        } else {
            return '';
        }
    }
}