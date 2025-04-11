
import * as React from "react";
 
import {
    TextField,
    PrimaryButton,
    IconButton,
    MessageBar,
    MessageBarType,
    Spinner,
    Icon,
    Link,
    Stack,
    DefaultButton,
} from "@fluentui/react";
import styles from "./ProjectTemplatesSttings.module.scss";

 
interface ContentType {
    id: number;
    name: string;
    fileName: string;
    fileContent: string;
    deletable: boolean;
}
 
interface ProjectTemplatesSttingsState {
    contentTypes: ContentType[];
    selectedContentType: ContentType | null;
    isCreatingNew: boolean;
    newContentType: string;
    newFileName: string;
    newFileContent: string;
    isSaving: boolean;
    errorMessage: string;
    showConfirmNavigation: boolean;
}
 
interface ProjectTemplatesSttingsProps { }
 
class ProjectTemplatesSttings extends React.Component<
    ProjectTemplatesSttingsProps,
    ProjectTemplatesSttingsState
> {
    private fileInputRef: React.RefObject<HTMLInputElement>
 
    constructor(props: ProjectTemplatesSttingsProps) {
        super(props);
 
        // Initialize with default content types
        const initialContentTypes = [
            {
                id: 1,
                name: "Native Document",
                fileName: "native.docx",
                fileContent: "",
                deletable: false
            }
        ];
 
        this.state = {
            contentTypes: initialContentTypes,
            selectedContentType: null,
            isCreatingNew: false,
            newContentType: "",
            newFileName: "",
            newFileContent: "",
            isSaving: false,
            errorMessage: "",
            showConfirmNavigation: false,
        };
        this.fileInputRef = React.createRef();
    }
 
     handleNavigation = (item: ContentType): void => {
        if (
            this.state.isCreatingNew &&
            (this.state.newContentType || this.state.newFileContent)
        ) {
            this.setState({ showConfirmNavigation: true, selectedContentType: item });
        } else {
            this.setState({
                selectedContentType: item,
                isCreatingNew: false,
                newContentType: item.name,
                newFileName: item.fileName,
                newFileContent: item.fileContent,
            });
        }
    };
 
    confirmNavigation = () => {
        this.setState({
            showConfirmNavigation: false,
            selectedContentType: null,
            isCreatingNew: false,
            newContentType: "",
            newFileName: "",
            newFileContent: "",
        });
    };
 
    addContentType = () => {
        if (this.state.newContentType && this.state.newFileContent) {
            this.setState({ isSaving: true });
 
            const newContentType: ContentType = {
                id: Date.now(),
                name: this.state.newContentType,
                fileName: this.state.newFileName,
                fileContent: this.state.newFileContent,
                deletable: true,
            };
 
            setTimeout(() => {
                this.setState((prevState) => ({
                    contentTypes: [...prevState.contentTypes, newContentType],
                    newContentType: "",
                    newFileName: "",
                    newFileContent: "",
                    isSaving: false,
                    isCreatingNew: false,
                }));
            }, 2000);
        }
    };
 
    updateContentType = () => {
        if (this.state.selectedContentType && this.state.newContentType) {
            this.setState({ isSaving: true });
 
            const updatedContentType: ContentType = {
                ...this.state.selectedContentType,
                name: this.state.newContentType,
                fileName: this.state.newFileName,
                fileContent: this.state.newFileContent,
            };
 
            setTimeout(() => {
                this.setState((prevState) => ({
                    contentTypes: prevState.contentTypes.map(ct =>
                        ct.id === updatedContentType.id ? updatedContentType : ct
                    ),
                    selectedContentType: null,
                    newContentType: "",
                    newFileName: "",
                    newFileContent: "",
                    isSaving: false,
                }));
            }, 2000);
        }
    };
 
    deleteContentType = (id: number): void => {
        this.setState((prevState) => ({
            contentTypes: prevState.contentTypes.filter((ct) => ct.id !== id),
        }));
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
                    errorMessage: ""
                });
            };
            reader.onerror = () => {
                this.setState({
                    errorMessage: "Error reading file",
                    newFileName: "",
                    newFileContent: ""
                });
            };
            reader.readAsDataURL(file);
        } else {
            this.setState({
                errorMessage: "Please upload a valid Word document.",
                newFileName: "",
                newFileContent: ""
            });
        }
    };
 
    renderForm() {
        const isEditing = !!this.state.selectedContentType;
        const isSaveDisabled = !this.state.newContentType || !this.state.newFileContent;
 
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
                    onClick={() => this.setState({
                        isCreatingNew: false,
                        selectedContentType: null,
                        newContentType: "",
                        newFileName: "",
                        newFileContent: ""
                    })}
                />
            </div>
        );
    }
 
    renderDetails() {
        if (!this.state.selectedContentType) return null;
 
        return (
            <div>
                <h3>{this.state.selectedContentType.name}</h3>
                <p>Template File: {this.state.selectedContentType.fileName}</p>
                <Link
                    href={this.state.selectedContentType.fileContent}
                    download={this.state.selectedContentType.fileName}
                >
                    Download Template
                </Link>
            </div>
        );
    }
 
    render() {
      console.log("State",this.state);
      
        return (
            <div className={styles.container}>
                <div className={styles.flexContainer}>
                    <div className={styles.sidebar}>
                        <h3>Content Types</h3>
                        <Stack tokens={{ childrenGap: 10 }} className={styles.contentTypes}>
                            {this.state.contentTypes.map((ct) => (
                                <Stack
                                    key={ct.id}
                                    horizontal
                                    verticalAlign="center"
                                    tokens={{ childrenGap: 5 }}
                                >
                                    <Link
                                        onClick={() => this.handleNavigation(ct)}
                                        className={
                                            this.state.selectedContentType?.id === ct.id
                                                ? `${styles.link} ${styles.active}`
                                                : styles.link
                                        }
                                    >
                                        {ct.name}
                                    </Link>
                                    <Icon iconName="Link" className={styles.icon} />
                                    {ct.deletable && (
                                        <IconButton
                                            iconProps={{ iconName: "Delete" }}
                                            title="Delete"
                                            onClick={() => this.deleteContentType(ct.id)}
                                            className={styles.deleteButton}
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
                                        newFileContent: ""
                                    })
                                }
                            >
                                + New
                            </Link>
                        </Stack>
                    </div>
                    <div className={styles.content}>
                        {this.state.isCreatingNew || this.state.selectedContentType
                            ? this.renderForm()
                            : <p>Select or create a content type.</p>
                        }
                    </div>
                </div>
 
                {this.state.isSaving && <Spinner label="Saving..." />}
 
                {this.state.errorMessage && (
                    <MessageBar messageBarType={MessageBarType.error}>
                        {this.state.errorMessage}
                    </MessageBar>
                )}
            </div>
        );
    }
}
 
export default ProjectTemplatesSttings;
 