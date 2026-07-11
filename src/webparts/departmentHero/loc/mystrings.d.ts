declare interface IDepartmentHeroWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  EyebrowLabel: string;
  EyebrowARLabel: string;
  DepartmentNameLabel: string;
  DepartmentNameARLabel: string;
  DescriptionLabel: string;
  DescriptionARLabel: string;
  OwnerNameLabel: string;
  OwnerRoleLabel: string;
}

declare module 'DepartmentHeroWebPartStrings' {
  const strings: IDepartmentHeroWebPartStrings;
  export = strings;
}
