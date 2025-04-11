import * as React from 'react';
import {
    Shimmer,
    ShimmerElementType
} from '@fluentui/react';
import styles from '../ProjectAdminPage.module.scss';

export default class ProjectAdminPageFormSkeleton extends React.Component<{}, {}> {

    public render(): React.ReactElement<{}> {
        return (
            <div>
                <div className={styles.column}>
                    <div className={styles.row}>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                        <div className={styles.fieldWrapper}>
                            <Shimmer width='30%' shimmerElements={[{ type: ShimmerElementType.line, height: 28 }]} />
                            <Shimmer width='100%' shimmerElements={[{ type: ShimmerElementType.line, height: 30 }]} />
                        </div>
                    </div>
                    <div className={styles.footer}>
                        <Shimmer shimmerElements={[
                            { type: ShimmerElementType.line, width: 84, height: 32 },
                            { type: ShimmerElementType.gap, width: '1rem' },
                            { type: ShimmerElementType.line, width: 84, height: 32 },
                        ]} />
                    </div>
                </div>
            </div >
        );
    }
}
