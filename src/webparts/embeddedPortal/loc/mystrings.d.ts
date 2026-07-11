declare interface IEmbeddedPortalWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  PortalUrlLabel: string;
  FrameHeightLabel: string;
  SectionTitleLabel: string;
  SectionTitleARLabel: string;
  LinkTextLabel: string;
  LinkTextARLabel: string;
  HideUrlLabel: string;
  AppSignInModeLabel: string;
}

declare module 'EmbeddedPortalWebPartStrings' {
  const strings: IEmbeddedPortalWebPartStrings;
  export = strings;
}
