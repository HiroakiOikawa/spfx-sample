declare interface IRedirectWebPartStrings {
  RedirectUrlFieldLabel: string;
  WaitTimeFieldLabel: string;
  WaitTimeDescription: string;
  WaitTimeValueErrorMessage: string;
  EscapeStringFieldLabel: string;
  EscapeStringDescription: string;
}

declare module 'RedirectWebPartStrings' {
  const strings: IRedirectWebPartStrings;
  export = strings;
}
