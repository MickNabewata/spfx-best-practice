import { IServerSiteUser, ISiteUser } from '../siteUsers';
import SpoBase from '../spoBase';

/** SharePoint Online サイトユーザー操作クラス サーバー接続用 */
export default class ServerSiteUsers extends SpoBase<IServerSiteUser, ISiteUser> {

    /** サイトユーザー取得 */
    protected getDatas(): Promise<IServerSiteUser[]> {
        return this.doSpoRequest<IServerSiteUser[]>(`${this._webUrl}/_api/web/siteusers`).then(
            (siteUsers) => {
                return Promise.resolve(siteUsers);
            },
            (err) => { 
                return Promise.reject(err);
             }
        );
    }

    /** サーバーから取得したデータ1件をコード内で使用する型に変換 */
    protected convertData(data: IServerSiteUser): ISiteUser {
        if(data) {
            return {
                Id: data.Id,
                Title: data.Title,
                Email: data.Email
            } as ISiteUser;
        } else {
            return undefined;
        }
    }
}