declare interface IDetailViewWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  NewsListLabel: string;
  EventsListLabel: string;
  PoliciesListLabel: string;
  BackUrlLabel: string;
}

declare module 'DetailViewWebPartStrings' {
  const strings: IDetailViewWebPartStrings;
  export = strings;
}
