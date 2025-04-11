import * as React from 'react';
import {
  Checkbox,
  TextField,
  PrimaryButton,
  Stack,
  IStackStyles,
  IStackTokens,
  ShimmeredDetailsList,
} from '@fluentui/react';
import {
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
} from '@fluentui/react/lib/DetailsList';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { SPFI } from '@pnp/sp/fi';
import { getSP } from '../../../../pnpjs-config';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/content-types';
import { IProjectAttributesSettingsProps } from './IProjectAttributesSettingsProps';
import { Logger } from '@pnp/logging';
import * as strings from 'ProjectAdminPageWebPartStrings';
import styles from '../ProjectAdminPage.module.scss';
import { IProjectRequest } from '../../model/IProjectRequest';
import { AadHttpClient } from '@microsoft/sp-http';

interface IField {
  title: string;
  internalName: string;
  hidden: boolean;
  choices: string[];
  choiceStates: boolean[];
  editedChoices: string;
}

interface IProjectAttributesSettingsState {
  fields: IField[];
  isDialogVisible: boolean;
  isDisabled: boolean; // State variable to manage the enabled/disabled status
  isLoading: boolean; // State variable to show/hide spinner
  isBusy?: boolean;
  errorMessage?: string;
}
const LOG_SOURCE: string = 'ProjectAttributesSettings';

class ProjectAttributesSettings extends React.Component<
  IProjectAttributesSettingsProps,
  IProjectAttributesSettingsState
> {
  private _sp: SPFI;

  constructor(props: IProjectAttributesSettingsProps) {
    super(props);

    this.state = {
      fields: [],
      isDialogVisible: false,
      isDisabled: false, // Initialize as enabled
      isLoading: true, // Initialize spinner state
    };

    // Initialize _sp with context from props
    this._sp = getSP(this.props.context);
  }

  public componentDidMount(): void {
    this._fetchFields().catch((error) => {
      console.error('Error fetching fields:', error);
      Logger.error(new Error(`${LOG_SOURCE}::componentDidMount, Error fetching fields: ${error}`));
      this.setState({ errorMessage: strings.ErrorLoadingProjectConfig, isLoading: false });
    });
  }

  private async _fetchFields(): Promise<void> {
    const fieldsToFetch = [
      'Projectattr1',
      'Projectattr2',
      'Projectattr3',
      'Projectattr4',
      'Projectattr5',
      'Projectattr6',
      'Projectattr7',
      'Projectattr8',
      'Projectattr9',
      'Projectattr10',
    ];
    const fieldData = await Promise.all(
      fieldsToFetch.map(async (fieldName) => {
        const field = await this._sp.web.fields
          .getByInternalNameOrTitle(fieldName)();
        return {
          title: field.Title,
          internalName: field.InternalName,
          hidden: field.Hidden,
          choices: field.Choices || [],
          choiceStates: (field.Choices || []).map(() => true), // Initialize all choices as checked
          editedChoices: (field.Choices || []).join('\n'), // Initialize editedChoices with all choices joined by newline
        };
      })
    );
    this.setState({ fields: fieldData, isLoading: false });
  }

  private _handleFieldVisibilityChange(
    index: number,
    checked: boolean | undefined
  ): void {
    const { fields } = this.state;
    const newFields = [...fields];
    newFields[index].hidden = !checked;
    this.setState({ fields: newFields });
  }

  private _handleTextFieldChange(
    index: number,
    newValue: string | undefined
  ): void {
    const { fields } = this.state;
    const newFields = [...fields];
    newFields[index].title = newValue || '';
    this.setState({ fields: newFields });
  }

  private _handleSave(): void {
    const { fields } = this.state;
    const { onError, dmsClient, projectEndpoint, context } = this.props;
    const request: IProjectRequest = {
      url: context.pageContext.web.serverRelativeUrl,
      fields: fields.map(field => ({
        fieldName: field.internalName,
        hidden: field.hidden,
        displayName: field.title,
        choices: field.editedChoices.split('\n')
      }))
    };
    // Show spinner until elements become disabled
    this.setState({ isBusy: true });

    dmsClient.post(projectEndpoint, AadHttpClient.configurations.v1, {
      body: JSON.stringify(request)
    }).then((response) => {
      if (response.ok)
        this._onCancel();
      else {
        throw new Error(response.statusText);
      }
    }).catch((error) => {
      console.error(error);
      Logger.error(new Error(`${LOG_SOURCE}::handleSave, Error saving fields: ${error}`));
      onError(`${strings.ErrorSavingProjectConfig}. ${error}`);
      this.setState({ isBusy: false });
    });
  }
  public render(): React.ReactElement<IProjectAttributesSettingsProps> {
    const { fields, isBusy, isLoading } =
      this.state;

    const columns: IColumn[] = [
      {
        key: 'field',
        name: 'Field names',
        fieldName: 'internalName',
        minWidth: 100,
        isResizable: true,
        onRender: (item: IField) => item.internalName,
      },
      {
        key: 'visible',
        name: 'Visible',
        fieldName: 'hidden',
        minWidth: 100,
        isResizable: true,
        onRender: (item: IField, index: number) => (
          <Checkbox
            checked={!item.hidden}
            onChange={(_, checked) =>
              this._handleFieldVisibilityChange(index, checked)
            }
            disabled={isBusy}
          />
        ),
      },
      {
        key: 'displayName',
        name: 'Display name',
        fieldName: 'title',
        minWidth: 200,
        isResizable: true,
        onRender: (item: IField, index: number) => (
          <TextField
            value={item.title}
            onChange={(_, newValue) =>
              this._handleTextFieldChange(index, newValue)
            }
            disabled={isBusy}
          />
        ),
      },
      {
        key: 'choices',
        name: 'Choice values',
        fieldName: 'editedChoices', // Use editedChoices instead of choices
        minWidth: 300,
        isResizable: true,
        onRender: (item: IField, fieldIndex: number) => (
          <TextField
            multiline
            autoAdjustHeight
            value={item.editedChoices}
            onChange={(_, newValue) => {
              const newFields = [...fields];
              newFields[fieldIndex].editedChoices = newValue || '';
              this.setState({ fields: newFields });
            }}
            disabled={isBusy}
          />
        ),
      },
    ];

    const stackStyles: IStackStyles = {
      root: {
        margin: '0 auto',
        maxWidth: '1200px',
        position: 'relative', // Ensure relative positioning for containing spinner
        paddingBottom: '50px', // Add padding to the bottom to accommodate buttons
      },
    };

    const stackTokens: IStackTokens = { childrenGap: 20 };

    return (
      <Stack styles={stackStyles} tokens={stackTokens}>
        <ShimmeredDetailsList
          items={fields}
          columns={columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          shimmerLines={10}
          enableShimmer={isLoading}
        />
        <div className={styles.footer}>
          <DefaultButton text={strings.Cancel} allowDisabledFocus onClick={this._onCancel.bind(this)} disabled={isBusy} />
          <PrimaryButton text={strings.Save} allowDisabledFocus onClick={this._handleSave.bind(this)} disabled={isBusy} />
        </div>
      </Stack>
    );
  }
  private _onCancel(): void {
    const { context } = this.props;
    const urlParams = new URLSearchParams(window.location.search);
    const source = urlParams.get('Source');
    if (source) {
      window.location.href = source;
    } else {
      window.location.href = context.pageContext.web.absoluteUrl;
    }
  }
}

export default ProjectAttributesSettings;
