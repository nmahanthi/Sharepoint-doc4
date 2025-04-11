declare interface ICreateNewVersionMenuCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'CreateNewVersionMenuCommandSetStrings' {
  const strings: ICreateNewVersionMenuCommandSetStrings;
  export = strings;
}
