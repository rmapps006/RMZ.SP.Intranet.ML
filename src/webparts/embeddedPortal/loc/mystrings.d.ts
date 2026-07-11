declare interface IEmbeddedPortalWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  PortalUrlLabel: string;
  FrameHeightLabel: string;
  SectionTitleLabel: string;
  LinkTextLabel: string;
  HideUrlLabel: string;
  AppSignInModeLabel: string;
}

declare module 'EmbeddedPortalWebPartStrings' {
  const strings: IEmbeddedPortalWebPartStrings;
  export = strings;
}
