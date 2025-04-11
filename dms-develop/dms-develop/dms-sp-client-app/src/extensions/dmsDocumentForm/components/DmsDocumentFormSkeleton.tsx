import * as React from 'react';
import { Pivot, PivotItem, Shimmer, ShimmerElementType } from '@fluentui/react';
import styles from './DmsDocumentForm.module.scss';

const shimmerElementsHeaders = [
  { type: ShimmerElementType.line, height: 16, width: '20%' },
  { type: ShimmerElementType.gap, width: '60%' },
  { type: ShimmerElementType.line, height: 16, width: '20%' }
];
const shimmerElementsFields = [
  { type: ShimmerElementType.line, height: 24, width: '45%' },
  { type: ShimmerElementType.gap, width: '10%' },
  { type: ShimmerElementType.line, height: 24, width: '45%' }
];
export default class DmsDocumentFormSqueleton extends React.Component<{}, {}> {
  /**
   *
   */
  public render(): React.ReactElement<{}> {
    return <>
      <div className={styles.dmsDocumentForm}>
        <Pivot>
          {Array(5).map((index) =>
          (<PivotItem key={index} itemKey={index}>
            <Shimmer />
          </PivotItem>)
          )}
        </Pivot>
        <div className={styles.formWrapper}>
          {Array(5).map((index) => (
            <>
              <Shimmer shimmerElements={shimmerElementsHeaders} />
              <Shimmer shimmerElements={shimmerElementsFields} />
            </>
          ))
          }
        </div>
      </div>
    </>;
  }
}

