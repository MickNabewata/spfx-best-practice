import * as React from 'react';
import styles from './DocumentCard.module.scss';
import SpfxBestPractice from '../spfxBestPractice/SpfxBestPractice';
import { IServerDocument, IDocument } from '../../../../datas/documents';
import { DocumentCard as DocCard, DocumentCardActivity, DocumentCardPreview, DocumentCardTitle } from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';

/** カードレイアウト */
export default class DocumentCard extends SpfxBestPractice<IServerDocument, IDocument> {

  /** データのレンダリング */
  protected renderDatas(): JSX.Element {
    return (
      <div className={ styles.documentCard }>
        <div className={ styles.container }>
          {this.state.datas.map((data, i) => {
            return (
              <div className={ styles.item } key={`userCard-${i}`} >
                <DocCard>
                  <DocumentCardPreview 
                    previewImages={[{
                      name: data.Name,
                      linkProps: {
                        href: data.ServerRelativeUrl,
                        target: '_blank'
                      },
                      previewImageSrc: data.PreviewImageUrl,
                      imageFit: ImageFit.cover,
                      height: 100
                    }]} 
                  />
                  <DocumentCardTitle title={data.Name} shouldTruncate={true} />
                  <DocumentCardActivity 
                    activity={data.Modified.toLocaleString()}
                    people={[{ name: data.Editor.Title, profileImageSrc: '' }]}
                  />
                </DocCard>
              </div>
            );
          })}
        </div>
      </div>
    );
  }
}
