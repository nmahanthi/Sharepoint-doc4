import { BaseDialog, IDialogConfiguration } from "@microsoft/sp-dialog";
import { SPFI } from "@pnp/sp/fi";
import  CreateRevisionDialog from "./createRevisionDialog";
import * as ReactDOM from "react-dom";
import * as React from "react";

export default class CreateRevisionDialogContainer extends BaseDialog {
    constructor(
        private projectReference: string,
        private sp: SPFI,
        private documentSetRelativeUrl:string
      ) {
        super();
      }

      public render(): void {
        ReactDOM.render(<CreateRevisionDialog
          close={this.close} projectReference={this.projectReference} sp={this.sp} documentSetRelativeUrl={this.documentSetRelativeUrl} 
          showDialog={true} />, this.domElement);
      }

      public getConfig(): IDialogConfiguration {
        return { isBlocking: false };
      }
    
      protected onAfterClose(): void {
        super.onAfterClose();
    
        // Clean up the element for the next dialog
        ReactDOM.unmountComponentAtNode(this.domElement);
      }
}