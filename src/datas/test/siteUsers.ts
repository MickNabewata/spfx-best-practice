import { IServerSiteUser, ISiteUser } from '../siteUsers';
import SpoBase from '../spoBase';

/** SharePoint Online サイトユーザー操作クラス テストデータ生成用 */
export default class TestSiteUsers extends SpoBase<IServerSiteUser, ISiteUser> {

    /** サイトユーザー取得 */
    protected getDatas(): Promise<IServerSiteUser[]> {
        let ret: IServerSiteUser[] = [];

        for(let i = 0; i < 10; i++) {
            const pad_i = i.toString().padStart(3, '0');
            ret[i] = {
                Id: i,
                Title: `User-${pad_i}`,
                Email: `User-${pad_i}@contoso.com`
            };
        }

        return Promise.resolve(ret);
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