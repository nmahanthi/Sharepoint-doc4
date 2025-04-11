import * as React from 'react';
// import styles from '../ProjectAdminPage.module.scss';
import { TextField } from '@fluentui/react/lib/TextField';
import {
    Checkbox,
} from '@fluentui/react';
import { ModernTaxonomyPicker } from "@pnp/spfx-controls-react/lib/ModernTaxonomyPicker";
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { SPFI } from '@pnp/sp/fi';
import { getSP } from '../../../../pnpjs-config';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Logger } from '@pnp/logging';
import styles from '../ProjectAdminPage.module.scss';
import { IFieldInfo } from '@pnp/sp/fields';
import * as strings from 'ProjectAdminPageWebPartStrings';
import ProjectAdminPageFormSkeleton from './AdminPageFormSkeleton';
import { FieldNames } from '../../../../constants';
import DmsFieldValidationWrapper from '../../../../extensions/dmsDocumentForm/components/DmsFieldValidationWrapper';
import { AadHttpClient } from '@microsoft/sp-http';
import { IProjectRequest } from '../../model/IProjectRequest';
import { IProjectPropertiesService } from '../../../../service/IProjectPropertiesService';
// Props Interface 
export interface IProjectAdminPageFormProps {
    context: WebPartContext;
    onError: (errorMessage: string) => void;
    libraryIds: string[];
    projectEndpoint: string;
    dmsClient: AadHttpClient;
    projectConfigurationPageURL:string;
    projectPropertiesService: IProjectPropertiesService;
}
// State Interface 
export interface IProjectAdminPageFormState {
    isLoading: boolean;
    isBusy?: boolean;
    projectAliasValue?: string;
    projectIsActiveValue?: boolean;
    endCustomerTerms?: { labels: { name: string, isDefault: boolean, languageTag: string }[], id: string }[];
    directCustomerTerms?: { labels: { name: string, isDefault: boolean, languageTag: string }[], id: string }[];
    productTerms?: { labels: { name: string, isDefault: boolean, languageTag: string }[], id: string }[];
    plOwnershipTerms?: { labels: { name: string, isDefault: boolean, languageTag: string }[], id: string }[];
    fieldInfos?: { [key: string]: IFieldInfo };
    validationErrors: { [key: string]: string | undefined };
    contractCode?: string;
    projectType?: string;
    geographyDisplay?: string;
    creationDate?: Date;
    createdByDisplay?: string;
    isMigratedProject?: boolean;
    lastModifiedDate?: Date;
    lastModifiedByDisplay?: string;
    projectName?: string;
    leadingProductLine?: string;
    projectLanguage?: string;
}

const PROJECT_CONFIG_FIELDS = [
    FieldNames.ProjectAlias,
    FieldNames.DirectCustomer,
    FieldNames.EndCustomer,
    FieldNames.Product,
    FieldNames.PLOwnership
];
const LOG_SOURCE: string = 'ProjectAdminPageForm';

export default class ProjectAdminPageForm extends React.Component<IProjectAdminPageFormProps, IProjectAdminPageFormState> {
    private _sp: SPFI;
    constructor(props: IProjectAdminPageFormProps) {
        super(props);
        this.state = {
            isLoading: true,
            validationErrors: {}
        }
        // Initialize _sp with context from props
        this._sp = getSP(this.props.context);
    }
    public componentDidMount(): void {
        const { onError, projectPropertiesService } = this.props;
        const [batch, execute] = this._sp.batched();
        Promise.all(
            [
                Promise.all(PROJECT_CONFIG_FIELDS.map((fieldName) => batch.web.fields.getByInternalNameOrTitle(fieldName)())),
                projectPropertiesService.getProperties()
            ]
        ).then(([fieldInfoArray, propertyBags]) => {
            const productLabels = propertyBags.dms_prod_display ? propertyBags.dms_prod_display.split('|') : [];
            const productIds = propertyBags.dms_prod_guid ? propertyBags.dms_prod_guid.split('|') : [];
            const productTerms = productLabels.map((label, i) => ({ labels: [{ name: label, isDefault: true, languageTag: "en-US" }], id: productIds[i] }));

            const plOwnershipLabels = propertyBags.dms_pl_owner_display ? propertyBags.dms_pl_owner_display.split('|') : [];
            const plOwnershipIds = propertyBags.dms_pl_owner_guid ? propertyBags.dms_pl_owner_guid.split('|') : [];
            const plOwnershipTerms = plOwnershipLabels.map((label, i) => ({ labels: [{ name: label, isDefault: true, languageTag: "en-US" }], id: plOwnershipIds[i] }));

            const endCustomerLabels = propertyBags.dms_end_cust_display ? propertyBags.dms_end_cust_display.split('|') : [];
            const endCustomerIds = propertyBags.dms_end_cust_guid ? propertyBags.dms_end_cust_guid.split('|') : [];
            const endCustomerTerms = endCustomerLabels.map((label, i) => ({ labels: [{ name: label, isDefault: true, languageTag: "en-US" }], id: endCustomerIds[i] }));

            const directCustomerLabels = propertyBags.dms_dir_cust_display ? propertyBags.dms_dir_cust_display.split('|') : [];
            const directCustomerIds = propertyBags.dms_dir_cust_guid ? propertyBags.dms_dir_cust_guid.split('|') : [];
            const directCustomerTerms = directCustomerLabels.map((label, i) => ({ labels: [{ name: label, isDefault: true, languageTag: "en-US" }], id: directCustomerIds[i] }));

            const fieldInfos = fieldInfoArray.reduce((acc, fieldInfo) => {
                acc[fieldInfo.InternalName] = fieldInfo;
                return acc;
            }, {} as { [key: string]: IFieldInfo });
            this.setState({
                fieldInfos,
                projectAliasValue: propertyBags.dms_proj_alias,
                projectIsActiveValue: propertyBags.dms_is_active === "True",
                contractCode: propertyBags.dms_proj_code,
                projectType: propertyBags.dms_proj_type,
                geographyDisplay: propertyBags.dms_geo_display,
                creationDate: propertyBags.dms_create_date ? new Date(propertyBags.dms_create_date) : undefined,
                createdByDisplay: propertyBags.dms_create_display,
                isMigratedProject: propertyBags.dms_is_old_proj === "True",
                lastModifiedDate: propertyBags.dms_modi_date ? new Date(propertyBags.dms_modi_date): undefined,
                lastModifiedByDisplay: propertyBags.dms_modi_display,
                projectName: propertyBags.dms_proj_name,
                leadingProductLine: propertyBags.dms_prod_line_display,
                projectLanguage: propertyBags.dms_proj_lang,
                productTerms,
                plOwnershipTerms,
                directCustomerTerms,
                endCustomerTerms,
                isLoading: false
            });
        }).catch(err => {
            onError(strings.ErrorLoadingProjectConfig);
            this.setState({ isLoading: false });
            console.error(err);
            Logger.error(new Error(`${LOG_SOURCE}: ${err}`));
        });
        execute().catch(err => {
            onError(strings.ErrorLoadingProjectConfig);
            this.setState({ isLoading: false });
            console.error(err);
            Logger.error(new Error(`${LOG_SOURCE}: ${err}`));
        });
    }

    // render  
    public render(): React.ReactElement<IProjectAdminPageFormProps> {
        const {
            isLoading, projectAliasValue, projectIsActiveValue,
            directCustomerTerms, endCustomerTerms, productTerms, plOwnershipTerms,
            fieldInfos, validationErrors, isBusy, contractCode, projectType, geographyDisplay,
            creationDate, createdByDisplay, isMigratedProject, lastModifiedByDisplay,
            lastModifiedDate, leadingProductLine, projectLanguage, projectName } = this.state;
        const { context } = this.props;

        return (
            <div>
                {isLoading && (
                    <ProjectAdminPageFormSkeleton />
                )}
                {!isLoading &&
                    <div className={styles.column}>
                        <div className={styles.row}>
                            <div className={styles.fieldWrapper}>
                                <TextField label={strings.ProjectNameField} disabled
                                    value={projectName} />
                            </div>
                            <div className={styles.fieldWrapper}>
                                <TextField label={strings.ContractCodeField} disabled
                                    value={contractCode} />
                            </div>
                            <div className={styles.fieldWrapper}>
                                <TextField label={strings.ProjectTypeField} disabled
                                    value={projectType} />
                            </div>
                            <div className={styles.fieldWrapper}>
                                <TextField label={strings.GeographyField} disabled
                                    value={geographyDisplay} />
                            </div>
                            <div className={styles.fieldWrapper}>
                                <TextField label={strings.CreatorField} disabled
                                    value={createdByDisplay} />
                            </div>
                            <div className={styles.fieldWrapper}>
                                <TextField label={strings.CreationDateField} disabled
                                    value={creationDate && !isNaN(creationDate.getTime()) ?
                                        `${creationDate.getFullYear()}-${(creationDate.getMonth() + 1).toString().padStart(2, '0')}-${creationDate.getDate().toString().padStart(2, '0')}` :
                                        ''
                                    } />
                            </div>
                            <div className={styles.fieldWrapper}>
                                <TextField label={strings.EditorField} disabled
                                    value={lastModifiedByDisplay} />
                            </div>
                            <div className={styles.fieldWrapper}>
                                <TextField label={strings.ModificationDateField} disabled
                                    value={lastModifiedDate && !isNaN(lastModifiedDate.getTime()) ?
                                        `${lastModifiedDate.getFullYear()}-${(lastModifiedDate.getMonth() + 1).toString().padStart(2, '0')}-${lastModifiedDate.getDate().toString().padStart(2, '0')}` :
                                        ''}
                                />
                            </div>
                            <div className={styles.fieldWrapper}>
                                <TextField label={strings.LanguageField} disabled
                                    value={projectLanguage} />
                            </div>
                            <div className={styles.fieldWrapper}>
                                <TextField label={strings.ProductLineField} disabled
                                    value={leadingProductLine} />
                            </div>
                            <div className={styles.fieldWrapper}>
                                <TextField label={strings.ProjectAliasField} required errorMessage={validationErrors.ProjectAlias}
                                    value={projectAliasValue} onChange={this._onAliasChange.bind(this)} />
                            </div>
                            <div className={styles.fieldWrapper}>
                                {fieldInfos && <DmsFieldValidationWrapper validationError={validationErrors.DirectCustomer}>
                                    <ModernTaxonomyPicker allowMultipleSelections={true}
                                        label={strings.DirectCustomerField}
                                        context={context as never}
                                        required termSetId={(fieldInfos[FieldNames.DirectCustomer] as IFieldInfo & { TermSetId: string }).TermSetId}
                                        panelTitle={strings.DirectCustomerField}
                                        onChange={(terms) => this._onTaxPickerChange('directCustomerTerms', terms as never[], FieldNames.DirectCustomer)}
                                        initialValues={directCustomerTerms}
                                    />
                                </DmsFieldValidationWrapper>}
                            </div>
                            <div className={styles.fieldWrapper}>
                                {fieldInfos && <DmsFieldValidationWrapper validationError={validationErrors.EndCustomer}>
                                    <ModernTaxonomyPicker allowMultipleSelections={true}
                                        label={strings.EndCustomerField}
                                        context={context as never}
                                        required termSetId={(fieldInfos[FieldNames.EndCustomer] as IFieldInfo & { TermSetId: string }).TermSetId}
                                        panelTitle={strings.EndCustomerField}
                                        onChange={(terms) => this._onTaxPickerChange('endCustomerTerms', terms as never[], FieldNames.EndCustomer)}
                                        initialValues={endCustomerTerms}
                                    />
                                </DmsFieldValidationWrapper>}
                            </div>
                            <div className={styles.fieldWrapper}>
                                {fieldInfos &&
                                    <DmsFieldValidationWrapper validationError={validationErrors.PLOwnership}>
                                        <ModernTaxonomyPicker allowMultipleSelections={true}
                                            label={strings.PLOwnerShipField}
                                            context={context as never}
                                            required termSetId={(fieldInfos[FieldNames.PLOwnership] as IFieldInfo & { TermSetId: string }).TermSetId}
                                            panelTitle={strings.PLOwnerShipFieldPanelTitle}
                                            onChange={(terms) => this._onTaxPickerChange('plOwnershipTerms', terms as never[], FieldNames.PLOwnership)}
                                            initialValues={plOwnershipTerms}
                                        />
                                    </DmsFieldValidationWrapper>
                                }
                            </div>
                            <div className={styles.fieldWrapper}>
                                {fieldInfos &&
                                    <DmsFieldValidationWrapper validationError={validationErrors.AlstomProduct}>
                                        <ModernTaxonomyPicker allowMultipleSelections={false}
                                            label={strings.ProductField}
                                            context={context as never}
                                            required termSetId={(fieldInfos[FieldNames.Product] as IFieldInfo & { TermSetId: string }).TermSetId}
                                            panelTitle={strings.ProductPanelTitle}
                                            onChange={(terms) => this._onTaxPickerChange('productTerms', terms as never[], FieldNames.Product)}
                                            initialValues={productTerms}
                                        />
                                    </DmsFieldValidationWrapper>
                                }
                            </div>
                            <div className={`${styles.fieldWrapper} ${styles.noLabel}`}>
                                <Checkbox label={strings.IsActiveProjectField} required
                                    checked={projectIsActiveValue}
                                    onChange={(e, checked) => this.setState({ projectIsActiveValue: !!checked })}
                                />
                            </div>
                            <div className={`${styles.fieldWrapper} ${styles.noLabel}`}>
                                <Checkbox label={strings.MigratedProjectField} disabled
                                    checked={isMigratedProject} />
                            </div>
                        </div>
                        <div className={styles.footer}>
                            <DefaultButton text={strings.Cancel} allowDisabledFocus onClick={this._onCancel.bind(this)} disabled={isBusy} />
                            <PrimaryButton text={strings.Save} allowDisabledFocus onClick={this._onSave.bind(this)} disabled={isBusy || Object.values(validationErrors).some(v => !!v)} />
                        </div>
                    </div>
                }
            </div >
        );
    }

    private _onTaxPickerChange(propertyName: string, newValue: never[], fieldName: string): void {
        const newState = { ...this.state, [propertyName]: newValue };
        this.setState(newState);
        this._validate(fieldName, newValue);
    }
    private _onAliasChange(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void {
        this.setState({ projectAliasValue: newValue });
        this._validate('ProjectAlias', newValue);
    }

    private _validate(propertyName: string, newValue?: unknown[] | unknown): void {
        if (!newValue || (newValue as unknown[]).length === 0) {
            this.setState({ validationErrors: { ...this.state.validationErrors, [propertyName]: strings.ValidationErrorRequiredValueMissing } });
            return;
        } else {
            this.setState({ validationErrors: { ...this.state.validationErrors, [propertyName]: undefined } });
        }
    }
    private _onCancel(): void {
        // const { context } = this.props;
        // const urlParams = new URLSearchParams(window.location.search);
        // const source = urlParams.get('Source');
        // if (source) {
        //     window.location.href = source;
        // } else {
        //     window.location.href = context.pageContext.web.absoluteUrl;
        // }
        window.location.href= `${this.props.context.pageContext.web.absoluteUrl}/${this.props.projectConfigurationPageURL}`;
    }
    private _onSave(): void {
        const { onError, context, dmsClient, projectEndpoint } = this.props;
        this.setState({ isBusy: true });
        const { projectAliasValue, projectIsActiveValue, directCustomerTerms, endCustomerTerms, productTerms, plOwnershipTerms } = this.state;

        const request: IProjectRequest = {
            url: context.pageContext.web.serverRelativeUrl,
            properties: [
                { key: 'dms_proj_alias', value: projectAliasValue || '', spFieldName: FieldNames.ProjectAlias, spValue: projectAliasValue },
                { key: 'dms_is_active', value: projectIsActiveValue ? "True" : "False" },
                { key: 'dms_dir_cust_guid', value: directCustomerTerms?.map(t => t.id).join('|') || '', spFieldName: FieldNames.DirectCustomer, spValue: directCustomerTerms && directCustomerTerms.length > 0 ? `-1;#${directCustomerTerms[0].labels[0].name}|${directCustomerTerms[0].id}` : '' },
                { key: 'dms_dir_cust_display', value: directCustomerTerms?.map(t => t.labels[0].name).join('|') || '' },
                { key: 'dms_end_cust_guid', value: endCustomerTerms?.map(t => t.id).join('|') || '', spFieldName: FieldNames.EndCustomer, spValue: endCustomerTerms && endCustomerTerms.length > 0 ? `-1;#${endCustomerTerms[0].labels[0].name}|${endCustomerTerms[0].id}` : '' },
                { key: 'dms_end_cust_display', value: endCustomerTerms?.map(t => t.labels[0].name).join('|') || '' },
                { key: 'dms_prod_guid', value: productTerms?.map(t => t.id).join('|') || '', spFieldName: FieldNames.Product, spValue: productTerms && productTerms.length > 0 ? `-1;#${productTerms[0].labels[0].name}|${productTerms[0].id}` : '' },
                { key: 'dms_prod_display', value: productTerms?.map(t => t.labels[0].name).join('|') || '' },
                { key: 'dms_pl_owner_guid', value: plOwnershipTerms?.map(t => t.id).join('|') || '', spFieldName: FieldNames.PLOwnership, spValue: plOwnershipTerms && plOwnershipTerms.length > 0 ? `-1;#${plOwnershipTerms[0].labels[0].name}|${plOwnershipTerms[0].id}` : '' },
                { key: 'dms_pl_owner_display', value: plOwnershipTerms?.map(t => t.labels[0].name).join('|') || '' }
            ]
        };
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
}
