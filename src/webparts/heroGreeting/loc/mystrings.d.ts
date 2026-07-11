declare interface IHeroGreetingWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ShowGreetingLabel: string;
  EyebrowLabel: string;
  TitleLine1Label: string;
  TitleEmphasisLabel: string;
  SubtitleLabel: string;
  ButtonTextLabel: string;
  ButtonUrlLabel: string;
}

declare module 'HeroGreetingWebPartStrings' {
  const strings: IHeroGreetingWebPartStrings;
  export = strings;
}
