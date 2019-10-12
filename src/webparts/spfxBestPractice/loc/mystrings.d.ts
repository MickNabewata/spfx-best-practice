/** locフォルダ配下で定義したローカライズ用ファイルの型定義 */
declare interface ISpfxBestPracticeWebPartStrings {
  /** データ種別プロパティラベル */
  DataTypeLabel: string;
  /** 描画形式プロパティラベル */
  VariantLabel: string;
  /** データ種別 選択肢 サイトのユーザー */
  DataTypeSiteUser: string;
  /** データ種別 選択肢 ドキュメント */
  DataTypeDocument: string;
  /** 描画形式 選択肢 カード */
  VariantCard: string;
  /** 描画形式 選択肢 テーブル */
  VariantTable: string;
}

/** locフォルダ配下で定義したローカライズ用ファイルの読取結果 */
declare module 'SpfxBestPracticeWebPartStrings' {
  const strings: ISpfxBestPracticeWebPartStrings;
  export = strings;
}
