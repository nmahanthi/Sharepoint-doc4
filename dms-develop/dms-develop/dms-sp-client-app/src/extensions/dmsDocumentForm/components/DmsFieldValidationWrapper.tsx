import * as React from 'react';
import styles from './DmsDocumentForm.module.scss';

export interface IDmsFieldValidationWrapperProps {
  validationError?: string;
}

export interface IDmsFieldValidationWrapperState {
}

export default class DmsFieldValidationWrapper extends React.Component<IDmsFieldValidationWrapperProps, IDmsFieldValidationWrapperState> {

  constructor(props: IDmsFieldValidationWrapperProps) {
    super(props);
  }

  public render(): React.ReactElement<{}> {
    const { validationError, children } = this.props;
    return <>
      <div className={`${styles.validationWrapper} ${validationError ? styles.invalidFieldWrapper : ''}`}>
        {children}
        {validationError && <span>
          <div role="alert">
            <p className={`ms-TextField-errorMessage ${styles.validationErrorMessage}`}>
              <span data-automation-id="error-message">{`${validationError}`}</span>
            </p>
          </div>
        </span>}
      </div>
    </>;
  }
}

