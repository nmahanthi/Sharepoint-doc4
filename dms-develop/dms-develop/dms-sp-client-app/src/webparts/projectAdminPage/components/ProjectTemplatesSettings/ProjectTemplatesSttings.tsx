import * as React from "react";
import {
  TextField,
  PrimaryButton,
  IconButton,
  Spinner,
  Icon,
  Link,
  Stack,
  DefaultButton,
  Dialog,
  DialogType,
  DialogFooter,
} from "@fluentui/react";
import styles from "./ProjectTemplatesSttings.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AadHttpClient } from "@microsoft/sp-http";
import {
  ContentTypeService,
  ContentTypeResponse,
  ContentTypeCreateRequest,
} from "../../../../service/ContentTypeService";
import strings from "ProjectAdminPageWebPartStrings";
interface IProjectTemplatesSttingsState {
  selectedContentType: ContentTypeResponse | null;
  isCreatingNew: boolean;
  newContentType: string;
  newFileName: string;
  newFileContent: string;
  isSaving: boolean;
  errorMessage: string;
  showConfirmNavigation: boolean;
  contentTypes: ContentTypeResponse[];
  showDialog: boolean;
  dialogMessage: string;
  dialogTitle: string;
  isError: boolean;
  showDeleteConfirmDialog: boolean;
  contentTypeToDelete: string | null;
}
export interface IProjectTemplatesSttingsProps {
  context: WebPartContext;
  dmsClient: AadHttpClient;
  projectTemplateEndpoint:string
}

class ProjectTemplatesSttings extends React.Component<
  IProjectTemplatesSttingsProps,
  IProjectTemplatesSttingsState
> {
  private fileInputRef: React.RefObject<HTMLInputElement>;
  private contentTypeService: ContentTypeService;
  constructor(props: IProjectTemplatesSttingsProps) {
    super(props);
    this.state = {
      contentTypes: [],
      selectedContentType: null,
      isCreatingNew: false,
      newContentType: "",
      newFileName: "",
      newFileContent: "",
      isSaving: false,
      errorMessage: "",
      showConfirmNavigation: false,
      showDialog: true,
      dialogMessage: "",
      dialogTitle: strings.TemplateDialogTitle,
      isError: false,
      showDeleteConfirmDialog: false,
      contentTypeToDelete: null,
    };
    this.fileInputRef = React.createRef();
    this.contentTypeService = new ContentTypeService(
      this.props.projectTemplateEndpoint,
      this.props.context.pageContext.site.serverRelativeUrl,
      this.props.dmsClient
    );
  }

  public async componentDidMount(): Promise<void> {
    this.setState({ isSaving: true, showDialog: false });
    try {
      await this.fetchContentTypes();
      this.setState({ isSaving: false, showDialog: true });
    } catch (error) {
      this.setState({
        dialogTitle: strings.ErrorDialogTitle,
        dialogMessage: strings.ErrorMsg,
        showDialog: true,
      });
    }
  }
  private async fetchContentTypes(): Promise<void> {
    try {
      const data = await this.contentTypeService.getAllContentTypes();
      console.log("Functiondata", data);
      this.setState({ contentTypes: data });
    } catch (error) {
      this.setState({ errorMessage: strings.ErrorMsg });
      console.error(error);
    }
  }

  private handleNavigation = (item: ContentTypeResponse): void => {
    if (
      this.state.isCreatingNew &&
      (this.state.newContentType || this.state.newFileContent)
    ) {
      this.setState({ showConfirmNavigation: true, selectedContentType: item });
    } else {
      this.setState({
        selectedContentType: item,
        isCreatingNew: false,
        newContentType: item.displayName,
      });
    }
  };
  public confirmNavigation = () => {
    this.setState({
      showConfirmNavigation: false,
      selectedContentType: null,
      isCreatingNew: false,
      newContentType: "",
      newFileName: "",
      newFileContent: "",
    });
  };

  private addContentType = async () => {
    if (this.state.newContentType && this.state.newFileContent) {
      this.setState({ isSaving: true, showDialog: false });
      const newContentTypeRequest: ContentTypeCreateRequest = {
        displayName: this.state.newContentType,
        template: this.state.newFileContent,
      };

      try {
        await this.contentTypeService.addContentType(newContentTypeRequest);
        this.setState({
          newContentType: "",
          newFileName: "",
          newFileContent: "",
          isSaving: false,
          isCreatingNew: false,
          dialogTitle: strings.CreatedDialogTitle,
          dialogMessage: strings.SuccessMessage,
        });
      } catch (error) {
        this.setState({
          isSaving: false,
          dialogTitle: strings.ErrorDialogTitle,
          dialogMessage: strings.ErrorMsg,
        });
      }
    }
  };

  private updateContentType = async () => {
    if (this.state.selectedContentType && this.state.newContentType) {
      this.setState({ isSaving: true, showDialog: false });
      const updatedContentType: ContentTypeCreateRequest = {
        ...this.state.selectedContentType,
        displayName: this.state.newContentType,
        template: this.state.newFileContent,
      };

      try {
        await this.contentTypeService.updateContentType(
          this.state.selectedContentType.id,
          updatedContentType
        );
        this.setState({
          selectedContentType: null,
          newContentType: "",
          newFileName: "",
          newFileContent: "",
          isSaving: false,
          dialogTitle: strings.UpdatedDialogTitle,
          dialogMessage: strings.UpdatedMessage,
        });
      } catch (error) {
        this.setState({
          isSaving: false,
          dialogTitle: strings.ErrorDialogTitle,
          dialogMessage: strings.ErrorMsg,
        });
      }
    }
  };

  private showDeleteConfirmDialog = (id: string): void => {
    this.setState({
      showDialog: false,
      showDeleteConfirmDialog: true,
      contentTypeToDelete: id,
    });
  };

  private closeDeleteConfirmDialog = (): void => {
    this.setState({
      showDeleteConfirmDialog: false,
      contentTypeToDelete: null,
      showDialog: true,
    });
  
  };

  private deleteContentType = async (): Promise<void> => {
    if (this.state.contentTypeToDelete) {
      this.setState({ isSaving: true, showDialog: false,showDeleteConfirmDialog: false });
      try {
        await this.contentTypeService.deleteContentType(
          this.state.contentTypeToDelete
        );
        this.setState({
          dialogTitle: strings.DeletedDialogTitle,
          dialogMessage: strings.DeletedMessage,
          isSaving: false,
          showDeleteConfirmDialog: false,
          contentTypeToDelete: null,
        });
      } catch (error) {
        this.setState({
          dialogTitle: strings.ErrorDialogTitle,
          dialogMessage: strings.ErrorMsg,
          showDeleteConfirmDialog: false,
          isSaving: false,
          contentTypeToDelete: null,
        });
      }
    }
  };

  handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (file.name.endsWith(".docx")) {
      const reader = new FileReader();
      reader.onload = () => {
        this.setState({
          newFileName: file.name,
          newFileContent: reader.result as string,
          errorMessage: "",
        });
      };
      reader.onerror = () => {
        this.setState({
          errorMessage: strings.ErrorMessage,
          newFileName: "",
          newFileContent: "",
        });
      };
      reader.readAsDataURL(file);
    } else {
      this.setState({
        errorMessage: "Please upload a valid Word document.",
        newFileName: "",
        newFileContent: "",
      });
    }
  };

  closeDialog = () => {
    this.setState({
      showDialog: true,
      dialogTitle: strings.TemplateDialogTitle,
      dialogMessage: "",
      isSaving:false
    });
    window.location.reload();
  };

  renderForm() {
    const isEditing = !!this.state.selectedContentType;
    const isSaveDisabled =
      !this.state.newContentType || !this.state.newFileContent;

    return (
      <div>
        <h3>{isEditing ? "Edit Content Type" : "Create New Content Type"}</h3>
        <div className={styles.formRow}>
          <label className={styles.label}>Display Name</label>
          <TextField
            value={this.state.newContentType}
            onChange={(e, v) => this.setState({ newContentType: v || "" })}
            className={styles.input}
          />
        </div>
        <div className={`${styles.formRow}`}>
          <label className={styles.label}>Template</label>
          <div className={styles.fieldContainer}>
            <TextField
              readOnly
              value={this.state.newFileName}
              className={styles.input}
            />
            <IconButton
              className={styles.iconButtonStyle}
              iconProps={{ iconName: "OpenFolderHorizontal" }}
              onClick={() => this.fileInputRef.current?.click()}
            />
          </div>

          <input
            type="file"
            ref={this.fileInputRef}
            accept=".docx"
            onChange={this.handleFileChange}
            className={styles.fileInput}
          />
        </div>
        <PrimaryButton
          text={isEditing ? "Update" : "Save"}
          className={styles.saveButton}
          disabled={isSaveDisabled}
          onClick={isEditing ? this.updateContentType : this.addContentType}
        />
        <DefaultButton
          text="Cancel"
          className={styles.cancelButton}
          onClick={() =>
            this.setState({
              isCreatingNew: false,
              selectedContentType: null,
              newContentType: "",
              newFileName: "",
              newFileContent: "",
            })
          }
        />
      </div>
    );
  }
  render() {
    console.log("State", this.state);
    return (
      <div className={styles.container}>
        <div className={styles.flexContainer}>
          <div className={styles.sidebar}>
            <h3>Content Types</h3>
            <Stack tokens={{ childrenGap: 10 }} className={styles.contentTypes}>
              {this.state.contentTypes.map((ct, index) => (
                <Stack
                  key={ct.id}
                  horizontal
                  verticalAlign="center"
                  tokens={{ childrenGap: 5 }}
                >
                  <Link
                    disabled={index === 0}
                    onClick={() => this.handleNavigation(ct)}
                    className={
                      this.state.selectedContentType?.id === ct.id
                        ? `${styles.link} ${styles.active}`
                        : styles.link
                    }
                  >
                    {ct.displayName}
                  </Link>
                  <Icon iconName="Link" className={styles.icon} />
                  {ct.displayName !== "Native Document" && (
                    <IconButton
                      iconProps={{ iconName: "Delete" }}
                      title="Delete"
                      onClick={() => this.showDeleteConfirmDialog(ct.id)}
                      className={styles.deleteButton}
                      disabled={
                        this.state.isCreatingNew ||
                        !!this.state.selectedContentType
                      }
                    />
                  )}
                </Stack>
              ))}
              <Link
                className={styles.newLink}
                onClick={() =>
                  this.setState({
                    isCreatingNew: true,
                    selectedContentType: null,
                    newContentType: "",
                    newFileName: "",
                    newFileContent: "",
                  })
                }
              >
                + New
              </Link>
            </Stack>
          </div>
          <div className={styles.content}>
            {this.state.isCreatingNew || this.state.selectedContentType ? (
              this.renderForm()
            ) : (
              <p>Select or create a content type.</p>
            )}
          </div>
        </div>

        <Dialog
          hidden={this.state.showDialog}
          // onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: this.state.showDeleteConfirmDialog
              ? "Confirm Delete"
              : this.state.dialogTitle,
            subText: this.state.showDeleteConfirmDialog
              ? "Are you sure you want to delete this content type?"
              : this.state.dialogMessage,
          }}
        >
          {this.state.isSaving && <Spinner label="Processing..." />}
          <DialogFooter>
            {this.state.showDeleteConfirmDialog ? (
              <>
                <PrimaryButton className={styles.deleteBtn} onClick={this.deleteContentType} text="Delete" />
                <DefaultButton
                  onClick={this.closeDeleteConfirmDialog}
                  text="Cancel"
                />
              </>
            ) : (
              <PrimaryButton disabled={this.state.isSaving} onClick={this.closeDialog} text="Close" />
            )}
      
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}

export default ProjectTemplatesSttings;
