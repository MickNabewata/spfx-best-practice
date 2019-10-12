/** サーバーとやり取りするための型 */
export interface IServerData {
}

/** コード内で利用するための型 */
export interface IData {
}

/** 
 * データ操作クラス 共通インタフェース
 * 　当サンプルではSharePointからしかデータを取得していないが、
 * 　他にGraph APIなど複数のデータ元を扱うことを考慮してインターフェースがあったほうが良い
 */
export default interface IDatas<O extends IServerData, T extends IData> {
    /** データ取得 */
    get(): Promise<T[]>;
}