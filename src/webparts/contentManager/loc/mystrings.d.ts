declare interface IContentManagerWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleLabel: string;
  TitleARLabel: string;
  ShowNewsLabel: string;
  ShowEventsLabel: string;
  ShowBenefitsLabel: string;
  NewsListLabel: string;
  EventsListLabel: string;
  BenefitsListLabel: string;
}

declare module 'ContentManagerWebPartStrings' {
  const strings: IContentManagerWebPartStrings;
  export = strings;
}
