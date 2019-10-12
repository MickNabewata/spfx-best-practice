import * as React from 'react';
import styles from './SpfxBestPractice.module.scss';
import { DisplayMode } from '@microsoft/sp-core-library';
import IDatas, { IServerData, IData } from '../../../../datas/iDatas';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';

/** プロパティ型定義 */
export interface IProps<ServerData extends IServerData, Data extends IData> {
  /** タイトル */
  title: string;
  /** 表示モード */
  mode: DisplayMode;
  /** タイトル変更時コールバック */
  titleChangeCallback: (title: string) => void;
  /** データ操作クライアント */
  client: IDatas<ServerData, Data>;
}

/** ステート型定義 */
export interface IStates<Data extends IData> {
  /** 取得済データ */
  datas: Data[];
  /** エラーメッセージ */
  err: string;
}

/** コンポーネント基底 */
export default class SpfxBestPractice<ServerData extends IServerData, Data extends IData> extends React.Component<IProps<ServerData, Data>, IStates<Data>> {
  
  /** コンポーネント基底 */
  public constructor(props: IProps<ServerData, Data>) {
    super(props);

    // ステート初期化
    this.state= {
      datas: undefined,
      err: undefined
    };
  }

  /** データ取得 */
  protected getDatas() {
    this.props.client.get().then(
      (datas) => {
        this.setState({ datas: datas, err: undefined });
      },
      (err) => {
        const msg = (err && err.message)? err.message : JSON.stringify(err);
        this.setState({ datas: [], err: msg });
      }
    );
  }

  /** タイトル変更イベント */
  protected handleTitleChanged = (ev: React.FormEvent<HTMLInputElement>, newValue?: string) => {
    this.props.titleChangeCallback(newValue);
  }

  /** タイトルのレンダリング */
  protected renderTitle(): JSX.Element {
    let ret = <React.Fragment />;

    switch(this.props.mode) {
      case DisplayMode.Read:
        ret = <p className={styles.title}>{this.props.title}</p>;
        break;
      case DisplayMode.Edit:
        ret = (
          <TextField 
            borderless 
            placeholder='タイトル'
            value={this.props.title}
            className={styles.titleInput}
            onChange={this.handleTitleChanged} />
        );
        break;
    }

    return ret;
  }

  /** コンテンツのレンダリング */
  protected renderBody(): JSX.Element {
    let ret = <React.Fragment />;

    if(this.state.datas) {
      ret = this.renderDatas();
    } else {
      ret = <Spinner size={SpinnerSize.small} />;
    }

    return ret;
  }

  /** データのレンダリング */
  protected renderDatas(): JSX.Element {
    return <React.Fragment />;
  }

  /** レンダリング */
  public render(): React.ReactElement<IProps<ServerData, Data>> {
    return (
      <div className={ styles.spfxBestPractice }>
        <div className={ styles.container }>
          {this.renderTitle()}
          <div className={styles.error}>{this.state.err}</div>
          {this.renderBody()}
        </div>
      </div>
    );
  }
  
  /** 描画完了イベント */
  public componentDidMount() {
    // データ取得
    this.getDatas();
  }
}
