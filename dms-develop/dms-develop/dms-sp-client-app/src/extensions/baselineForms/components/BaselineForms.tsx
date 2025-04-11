/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

import styles from './BaselineForms.module.scss';
import FormContainer from './CustomForms/FormContainer';

import { ToastContainer } from 'react-toastify';
import "react-toastify/dist/ReactToastify.css";
import ViewFormContainer from './CustomForms/ViewFormContainer';
import { Label } from '@fluentui/react/lib/Label';
import CONSTANTS from '../../../_Constants';
import { ILinkedItems } from '../../../interfaces/IGlobalInterfaces';
import { AadHttpClient } from '@microsoft/sp-http';
import { INotificationService } from '../../../service/NotificationService';
import { IPermissionsService } from '../../../service/PermissionsService';

export interface IBaselineFormsProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
  itemID: number;
  item: ILinkedItems;
  freezeEndpoint: string;
  dmsClient: AadHttpClient;
  notificationService: INotificationService;
  baseLineFields:string;
  technicalFields:string;
  permissionsService:IPermissionsService

}

const LOG_SOURCE: string = 'BaselineForms';

export default class BaselineForms extends React.Component<IBaselineFormsProps, {}> {
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: BaselineForms mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: BaselineForms unmounted');
  }

  public render(): React.ReactElement<{}> {
    const { displayMode } = this.props;
    let titleHeader = "";
    switch (displayMode) {
      case FormDisplayMode.New:
        titleHeader = "New";
        break;
      case FormDisplayMode.Edit:
        titleHeader = "Edit"
        break;
      case FormDisplayMode.Display:
        titleHeader = "View";
        break;
    }
    return <div className={styles.baselineForms} >
      <div className={styles.header}>
        <Label className={styles.titleHeader}>{titleHeader} {CONSTANTS.ContentTypeNames.Baseline}</Label>
      </div>
      <ToastContainer />
      {displayMode === FormDisplayMode.New &&
        <FormContainer {...this.props} />
      }
      {displayMode === FormDisplayMode.Edit &&
        <FormContainer {...this.props} />
      }
      {displayMode === FormDisplayMode.Display &&
        <ViewFormContainer {...this.props} />
      }
    </div>
      ;
  }
}
