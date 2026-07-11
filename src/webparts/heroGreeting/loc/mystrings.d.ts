declare interface IHeroGreetingWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ShowGreetingLabel: string;
  EyebrowLabel: string;
  EyebrowARLabel: string;
  TitleLine1Label: string;
  TitleLine1ARLabel: string;
  TitleEmphasisLabel: string;
  TitleEmphasisARLabel: string;
  SubtitleLabel: string;
  SubtitleARLabel: string;
  ButtonTextLabel: string;
  ButtonTextARLabel: string;
  ButtonUrlLabel: string;
}

declare module 'HeroGreetingWebPartStrings' {
  const strings: IHeroGreetingWebPartStrings;
  export = strings;
}
