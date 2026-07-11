declare interface IHrBenefitsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleLabel: string;
  TitleARLabel: string;
  AllBenefitsUrlLabel: string;
  BenefitsJsonLabel: string;
}

declare module 'HrBenefitsWebPartStrings' {
  const strings: IHrBenefitsWebPartStrings;
  export = strings;
}
