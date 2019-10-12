import * as React from 'react';
import styles from './UserCard.module.scss';
import SpfxBestPractice from '../spfxBestPractice/SpfxBestPractice';
import { IServerSiteUser, ISiteUser } from '../../../../datas/siteUsers';
import { Persona } from 'office-ui-fabric-react/lib/Persona';

/** カードレイアウト */
export default class UserCard extends SpfxBestPractice<IServerSiteUser, ISiteUser> {

  /** データのレンダリング */
  protected renderDatas(): JSX.Element {
    return (
      <div className={ styles.userCard }>
        <div className={ styles.container }>
          {this.state.datas.map((data, i) => {
            return (
              <div className={ styles.item } key={`userCard-${i}`} >
                <Persona
                  imageUrl=''
                  imageInitials={data.Title.substr(0, 2)}
                  text={data.Title}
                  secondaryText={data.Email}
                  hidePersonaDetails={false}
                />
              </div>
            );
          })}
        </div>
      </div>
    );
  }
}
