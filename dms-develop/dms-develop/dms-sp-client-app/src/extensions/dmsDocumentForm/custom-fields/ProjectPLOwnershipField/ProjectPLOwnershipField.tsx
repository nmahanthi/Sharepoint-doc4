import * as React from "react";
import styles from "./ProjectPLOwnershipField.module.scss";
import DmsDynamicFieldBase, {
  IDmsDynamicFieldProps,
  IDmsDynamicFieldState,
} from "../../components/DmsDynamicFieldBase";
import { IFieldInfo } from "@pnp/sp/fields";
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react";
import {
  HttpClient,
  HttpClientResponse
} from "@microsoft/sp-http";
import { ProjectPropertiesService } from "../../../../service/ProjectPropertiesService";
import { IProjectPropertiesService } from "../../../../service/IProjectPropertiesService";

export interface IPLOwnershipFieldState extends IDmsDynamicFieldState {
  value: unknown;
  hasChanged: boolean;
  plOwnershipTerms?: {
    labels: { name: string; isDefault: boolean; languageTag: string }[];
    id: string;
  }[];
  allPLOwnershipTerms?: {
    labels: { name: string; isDefault: boolean; languageTag: string }[];
    id: string;
  }[];
  disabledTermIds?: string[];
  selectedPlOwnership: IPickerTerms;
  initialValue?: IPickerTerms;
}
export default class PLOwnershipField extends DmsDynamicFieldBase<IPLOwnershipFieldState> {
  constructor(props: IDmsDynamicFieldProps) {
    super(props);
    this.state = {
      selectedPlOwnership: [],
      isLoading: true,
      value: props.value,
      hasChanged: false,
      disabledTermIds: [],
      allPLOwnershipTerms: [],
    };
  }


  public componentDidMount(): void {
    const { context } = this.props;
    context.serviceScope.whenFinished(() => {
      const projectPropertiesService: IProjectPropertiesService = context.serviceScope.consume(ProjectPropertiesService.serviceKey);
      projectPropertiesService.getProperties()
        .then((propertyBags) => {
          const plOwnershipLabels =
            propertyBags.dms_pl_owner_display
              ? propertyBags.dms_pl_owner_display.split("|")
              : [];
          const plOwnershipIds = propertyBags.dms_pl_owner_guid
            ? propertyBags.dms_pl_owner_guid.split("|")
            : [];
          const plOwnershipTerms = plOwnershipLabels.map((label, i) => ({
            labels: [{ name: label, isDefault: true, languageTag: "en-US" }],
            id: plOwnershipIds[i],
          }));
          const plOwnershipTermIds = plOwnershipTerms.map((term) => term.id);
          this.setState(
            {
              plOwnershipTerms,
            },
            () => {
              this.fetchTaxonomyTerms(plOwnershipTermIds);
            }
          );
        }).catch((err) => {
          this.setState({ isLoading: false });
          console.error(err);
        });
    });
    const { value, data } = this.props;
    const typedValue = (value as { TermGuid: string; Label: string });
    if (!typedValue) return;
    let processedValue: { labels: { name: string; isDefault: boolean; languageTag: string }[], id: string }[];
    if (typedValue.TermGuid) {
      if (value && !data) throw new Error('Passing data as stream is mandatory for taxonomy fields');
      processedValue = [{ labels: [{ name: (data as { Label: string }).Label, isDefault: true, languageTag: "en-US" }], id: typedValue.TermGuid }];
    } else {
      if (!/(?:-\d;#[^;]+(?:;#)?)+/.test(value as string)) {
        throw new Error('Invalid taxonomy field value');
      }
      const values = (value as string).split(';#').filter(v => /.*\|(?:[a-z0-9]+-[a-z0-9]+-[a-z0-9]+-[a-z0-9]+-[a-z0-9]+)/.test(v));
      processedValue = values.map(v => {
        const [label, id] = v.split('|');
        return { labels: [{ name: label, isDefault: true, languageTag: "en-US" }], id };
      });
    }
    this.setState({
      value: processedValue,
      initialValue: processedValue ? [{
        key: processedValue[0].id,
        name: processedValue[0].labels[0].name,
        termSet: (this.props.fieldInfo as IFieldInfo & { TermSetId: string }).TermSetId,
        path: ""
      }] : undefined
    });
  }
  private onTaxonomyPickerChange = (terms: IPickerTerms): void => {

    const termId = terms.map((item) => item.key);
    if (termId && termId.length > 0) {
      this.selectedTaxonomyTerm(termId);
    } else {
      this.setState({
        hasChanged: false,
        value: '',
      });

      return
    }

    this.setState({
      selectedPlOwnership: terms,
      hasChanged: true,
    });
  };
  public render(): React.ReactElement<{}> {
    const {
      plOwnershipTerms,
      disabledTermIds,
      initialValue
    } = this.state;
    const { fieldInfo: propFieldInfo, disabled: propsDisabled } = this.props;
    const fieldInfo = propFieldInfo as IFieldInfo & { TermSetId: string };
    const disabled = propsDisabled || fieldInfo.ReadOnlyField;
    return (
      <div className={styles.ProjectPLOwnershipField}>
        {plOwnershipTerms && plOwnershipTerms.length > 0 && (
          <TaxonomyPicker
            allowMultipleSelections={false}
            termsetNameOrID={`${fieldInfo.TermSetId}`}
            required={true}
            panelTitle="PL Ownership"
            label={`${fieldInfo.Title}`}
            context={this.props.context as never}
            isTermSetSelectable={false}
            disabledTermIds={disabledTermIds}
            disabled={disabled}
            onChange={this.onTaxonomyPickerChange}
            initialValues={initialValue}
          />
        )}
      </div>
    );
  }
  private fetchTaxonomyTerms(plOwnershipTermIds: string[]): void {
    const { fieldInfo: propFieldInfo } = this.props;
    const fieldInfo = propFieldInfo as IFieldInfo & { TermSetId: string };
    const termSetId = fieldInfo.TermSetId;
    const spurl = this.props.context.pageContext.web.absoluteUrl;
    const requestUrl = `${spurl}/_api/v2.1/termStore/sets/${termSetId}/terms`;
    this.props.context.httpClient
      .get(requestUrl, HttpClient.configurations.v1)
      .then((response: HttpClientResponse): Promise<{ value: { id: string; labels: { name: string }[] }[] }> => {
        if (response.ok) {
          return response.json();
        } else {
          return Promise.reject(new Error(response.statusText));
        }
      })
      .then((data) => {
        const allTermsValue = data.value.map((terms) => ({
          id: terms.id,
          labels: terms.labels.map((label) => ({
            name: label.name,
            isDefault: true,
            languageTag: "en-US",
          })),
        }));
        this.setState({
          allPLOwnershipTerms: allTermsValue,
        });
        const disabledTermIds = data.value
          .filter((term) => !plOwnershipTermIds.includes(term.id))
          .map((term) => term.id);
        this.setState({
          disabledTermIds,
        });
      })
      .catch((error) => {
        console.error("Error fetching taxonomy terms: ", error);
      });
  }
  private selectedTaxonomyTerm(plOwnershipTermIds: string[]): void {
    const { fieldInfo: propFieldInfo } = this.props;
    const fieldInfo = propFieldInfo as IFieldInfo & { TermSetId: string };
    const termSetId = fieldInfo.TermSetId;
    const spurl = this.props.context.pageContext.web.absoluteUrl;
    const requestUrl = `${spurl}/_api/v2.1/termStore/sets/${termSetId}/terms`;
    this.props.context.httpClient
      .get(requestUrl, HttpClient.configurations.v1)
      .then((response: HttpClientResponse): Promise<{ value: { id: string }[] }> => {
        if (response.ok) {
          return response.json();
        } else {
          return Promise.reject(new Error(response.statusText));
        }
      })
      .then((data) => {
        const selectedTerm = data.value.filter((term) =>
          plOwnershipTermIds.includes(term.id)
        );

        this.setState({
          value: selectedTerm,
        });
      })
      .catch((error) => {
        console.error("Error fetching taxonomy terms: ", error);
      });
  }
}