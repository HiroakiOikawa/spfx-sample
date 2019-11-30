declare interface ITeamsTabSampleWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  // add-1 L6-L7
  TeamIdFieldLabel: string;
  ChannelIdFieldLabel: string;
}

declare module 'TeamsTabSampleWebPartStrings' {
  const strings: ITeamsTabSampleWebPartStrings;
  export = strings;
}
