import * as React from 'react';
import styles from './ProjectAdminPage.module.scss';
import type { IProjectAdminPageProps } from './IProjectAdminPageProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { MessageBar, MessageBarType, Pivot, PivotItem, Text, ThemeProvider } from '@fluentui/react';
import ProjectAdminPageForm from './Form/AdminPageForm';
import ProjectAttributesSettings from './ProjectAttributesSettings/ProjectAttributesSettings';
import { getSP } from '../../../pnpjs-config';
import * as strings from 'ProjectAdminPageWebPartStrings';
import { ListUrls } from '../../../constants';
import { Logger } from '@pnp/logging';
import ProjectTemplatesSttings from './ProjectTemplatesSettings/ProjectTemplatesSttings';


const LOG_SOURCE: string = 'ProjectAdminPage';

export interface IProjectAdminPageState {
  draftLibraryId?: string;
  applicableLibraryId?: string;
  prevLibraryId?: string;
  errorMessages: string[];
}
export default class ProjectAdminPage extends React.Component<IProjectAdminPageProps, IProjectAdminPageState> {
  /**
   *
   */
  constructor(props: IProjectAdminPageProps) {
    super(props);
    this.state = {
      errorMessages: []
    };
  }
  // ,RootFolder/ServerRelativeUrl&$expand=RootFolder
  public componentDidMount(): void {
    const { context } = this.props;
    const { errorMessages } = this.state;
    const sp = getSP();
    const [batch, execute] = sp.batched();
    Promise.all([
      batch.web.lists.select('Id').filter(`RootFolder/ServerRelativeUrl eq '${context.pageContext.web.serverRelativeUrl}/${ListUrls.Draft}'`)(),
      batch.web.lists.select('Id').filter(`RootFolder/ServerRelativeUrl eq '${context.pageContext.web.serverRelativeUrl}/${ListUrls.ApplicableDocuments}'`)(),
      batch.web.lists.select('Id').filter(`RootFolder/ServerRelativeUrl eq '${context.pageContext.web.serverRelativeUrl}/${ListUrls.PreviousVersions}'`)(),
    ]).then(([draftLib, applicableLib, prevLib]) => {
      if (!draftLib || draftLib.length === 0 ||
        !applicableLib || applicableLib.length === 0 ||
        !prevLib || prevLib.length === 0) {
        this.setState({
          errorMessages: [...errorMessages, strings.ErrorLoadingProjectConfig]
        });
        console.error(new Error('Document libraries not found'));
        Logger.error(new Error(`${LOG_SOURCE}:Document libraries not found`));
      } else {
        this.setState({
          draftLibraryId: draftLib[0].Id,
          applicableLibraryId: applicableLib[0].Id,
          prevLibraryId: prevLib[0].Id
        });
      }
    }).catch((error) => {
      this.setState({
        errorMessages: [...errorMessages, strings.ErrorLoadingProjectConfig]
      });
      console.error(error);
      Logger.error(new Error(`${LOG_SOURCE}: ${error}`));
    });
    execute().catch((error) => {
      this.setState({
        errorMessages: [...errorMessages, strings.ErrorLoadingProjectConfig]
      });
      console.error(error);
      Logger.error(new Error(`${LOG_SOURCE}: ${error}`));
    });
  }
  public render(): React.ReactElement<IProjectAdminPageProps> {
    const {
      context,
      theme,
      dmsClient,
      projectEndpoint,
      projectConfigurationPageURL,
      projectPropertiesService,
      projectTemplateEndpoint
      
    } = this.props;
    const { errorMessages, applicableLibraryId, draftLibraryId, prevLibraryId } = this.state;
    return (
      <ThemeProvider theme={theme} >
        <div className={styles.projectAdminPage}>
          <div>
            <Text as='h1' variant='xLarge' className={styles.AdminPageTitle}>{strings.AdminPageTitle}</Text>
          </div>
          {errorMessages.map((errorMessage, index) => (
            <MessageBar
              key={index}
              messageBarType={MessageBarType.error}
              isMultiline={true}
              truncated={false}
              onDismiss={() => this.setState({ errorMessages: errorMessages.filter((_, i) => i !== index) })}
            >
              {errorMessage}
            </MessageBar>))
          }
          <div>
            {applicableLibraryId && draftLibraryId && prevLibraryId && <Pivot>
              <PivotItem headerText={strings.PorjectAttributesHeader} >
                <ProjectAdminPageForm context={context} onError={this._onError}
                  libraryIds={[applicableLibraryId, draftLibraryId, prevLibraryId]} 
                  dmsClient={dmsClient} projectEndpoint={projectEndpoint}
                  projectConfigurationPageURL = {projectConfigurationPageURL}
                  projectPropertiesService={projectPropertiesService}
                  />
              </PivotItem>
              <PivotItem headerText={strings.DocumentProjectSpecificFieldsHeader}>
                <ProjectAttributesSettings context={context} onError={this._onError}
                  libraryIds={[applicableLibraryId, draftLibraryId, prevLibraryId]} 
                  dmsClient={dmsClient} projectEndpoint={projectEndpoint}/>
              </PivotItem>
              <PivotItem headerText={strings.ProjectTemplatesHeader}>
                <ProjectTemplatesSttings
                  context={context}
                  dmsClient={dmsClient} 
                  projectTemplateEndpoint={projectTemplateEndpoint}             
                />
              </PivotItem>
            </Pivot>}
          </div>
        </div>
      </ThemeProvider>
    );
  }
  private _onError = (error: string): void => {
    this.setState({
      errorMessages: [...this.state.errorMessages, error]
    });
  }
}
