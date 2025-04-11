declare interface IStartWorkflowMenuCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'StartWorkflowMenuCommandSetStrings' {
  const strings: IStartWorkflowMenuCommandSetStrings;
  export = strings;
}
