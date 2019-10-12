## SharePoint Frammework ベストプラクティス

### マニフェストファイル

#### 名称は英語で作成し後から日本語に変える

yo @microsoft/sharepointを実行した際、Webパーツや拡張コマンドの名称を入力するタイミングがありますが、  
ここでの入力値はクラス名やファイル名にも使われます。  
よって入力値は英語としマニフェストファイル(当サンプルではSpfxBestPracticeWebPart.manifest.json)を  
後から変更することでWebパーツやコマンドの名称を日本語にします。

#### アイコンを変える

マニフェストファイル(当サンプルではSpfxBestPracticeWebPart.manifest.json)でアイコンを変更できます。  
既定のものでは寂しいので機能をイメージしやすいアイコンに変えておきましょう。  
アイコン名には[Office UI Fabric ICON](https://uifabricicons.azurewebsites.net)で用意されているものを利用できます。

#### WebパーツにはTitleプロパティを設ける

SharePoint標準Webパーツのように、編集モードで入力できるタイトルがあるとよりSharePointらしいデザインになります。  
マニフェストファイルでTitleプロパティを定義しましょう。  
Webパーツ内には編集モード(this.displayMode === DisplayMode.Edit)時に入力できるテキストボックスを設けます。  
テキストボックスの入力値をthis.properties.Titleに代入するとプロパティを保存できます。

### serve.json

#### serveConfigurationsを活用する

gulp serveした時に起動するページのURLはserveConfigurationsでの指定がお勧めです。  
Yeaomanで構築したWebパーツでは初期状態でinitialPageプロパティでページURLを指定していますが、  
これを削除して以下のように書くことでも同じようにgulp serveコマンドを実行することができます。

```
"serveConfigurations": {
    "default": {
        "pageUrl": "https://localhost:5432/workbench"
    },
    "spo": {
        "pageUrl": "https://contoso.sharepoint.com/_layouts/workbench.aspx"
    }
}
```

上記の設定でただgulp serveと実行するだけではdefaultの構成が効いてローカルワークベンチが起動しますが、  
gulp serve --config spoとするとspoの構成が効いてSharePoint上での動作確認ができます。

### tsconfig.json

#### esnextを利用できるようにする

libに設定を追加してesnextを利用できるようにしておきましょう。便利な関数が使えて快適です。  
```
"lib": [
    "es5",
    "dom",
    "es2015.collection",
    "dom.iterable",
    "esnext"
]
```

### scss

#### 共通のscssを定義して各scssでインポートする

実際に運用するものを開発する際はコード量が増え、コンポーネントを分割することも多々あります。  
1つのソリューション内に複数のWebパーツやコマンドを定義することもあるでしょう。  
このような場合には共通のscss(当サンプルではTheme.module.scss)を定義しましょう。  
各scssではこれをimportすることで共通のスタイル定義を1ファイルに纏めることができます。

#### メディアクエリを書きやすくするためのmixin定義

SharePointサイトはモバイル端末でも利用しますので、場合によってはメディアクエリが必要になります。  
当サンプルでは共通scss(Theme.module.scss)でこれを定義しています。

#### SharePointサイトのテーマを反映する

SharePointでは外観の変更機能でテーマ(配色)を変更することができます。  
SPFx側で色を固定してしまうとテーマを変更したときに浮いてしまったり、文字が見えなくなったりします。  
テーマをSPFxに反映するには、以下の基本原則に従います。

- scssやstyleで色を固定しない(color: black;などを書かない)
- できるだけOffice UI Fabricを用いる
- 上記2つを行ってもテーマ色が反映されない場合、scss内で"color: [theme: link]"のようにトークンを用いて色を指定する

上記3番目の基本原則について、利用可能なトークンはJavaScriptのwindowオブジェクトから一覧を取得することができます。  
プロパティ名はwindow.__themeState__.theme です。  
TypeScript内では__themeState__が定義されておらず参照できないため、一度anyにするなど工夫してください。  
2019年10月8日時点で手元の環境を使って確認した結果は当プロジェクト内の  
themeTokenSamples.jsonファイルに記載しておきました。

### ローカライズ

SPFxを複数言語向けに提供する場合、ローカライズが必要になります。  
SPFxではlocフォルダ配下のローカライズファイルを使って固定文字列を複数言語向けに提供することができます。  
また、マニフェストファイルで指定する文字列も複数言語を定義しておきます。

```
"preconfiguredEntries": [{
"groupId": "5c03119e-3074-46fd-976b-c60198311f70",
"group": { "en-us": "Other", "default": "その他" },
"title": { "en-us": "SPFx Best Practicies", "default": "SPFx開発ベストプラクティス集" },
"description": { "en-us": "Best practicies for SPFx.", "default": "SPFx開発のベストプラクティス集です。" },
"officeFabricIconFontName": "HintText",
"properties": {
    "title": "",
    "dataType": "SiteUsers",
    "variant": "Card"
}
}]
```