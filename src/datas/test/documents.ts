import { IServerDocument, IDocument } from '../documents';
import SpoBase from '../spoBase';
import { ISiteUser } from '../siteUsers';

/** SharePoint Online ドキュメント操作クラス テストデータ生成用 */
export default class TestDocuments extends SpoBase<IServerDocument, IDocument> {

    /** ドキュメント取得 */
    protected getDatas(): Promise<IServerDocument[]> {
        let ret: IServerDocument[] = [];

        for(let i = 0; i < 10; i++) {
            const pad_i = i.toString().padStart(3, '0');
            ret[i] = {
                Id: i,
                Title: `Document-${pad_i}`,
                Name: `Document-${pad_i}.docx`,
                ServerRelativeUrl: `/sites/contoso/Shared%20Documents/Document-${pad_i}.docx`,
                UIVersionLabel: `${i}.0`,
                AuthorId: i,
                Created: new Date().toLocaleString(),
                EditorId: i,
                Modified: new Date().toLocaleString()
            };
        }

        return Promise.resolve(ret);
    }

    /** サーバーから取得したデータ1件をコード内で使用する型に変換 */
    protected convertData(data: IServerDocument): IDocument {
        if(data) {
            const pad_id = data.Id.toString().padStart(3, '0');
            return {
                Id: data.Id,
                Title: data.Title,
                Name: data.Name,
                ServerRelativeUrl: data.ServerRelativeUrl,
                UIVersionLabel: data.UIVersionLabel,
                PreviewImageUrl: 'https://cdn-ak.f.st-hatena.com/images/fotolife/m/micknabewata/20180715/20180715155328.jpg',
                Author: { Id: data.Id, Title: `User-${pad_id}`, Email: `User-${pad_id}@contoso.com` } as ISiteUser,
                Created: new Date(data.Created),
                Editor: { Id: data.Id, Title: `User-${pad_id}`, Email: `User-${pad_id}@contoso.com` },
                Modified: new Date(data.Modified)
            } as IDocument;
        } else {
            return undefined;
        }
    }
}