declare interface IEmployeeDirectoryWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleLabel: string;
  TitleARLabel: string;
  LinkTextLabel: string;
  LinkTextARLabel: string;
  FullDirectoryUrlLabel: string;
  PageSizeLabel: string;
}

declare module 'EmployeeDirectoryWebPartStrings' {
  const strings: IEmployeeDirectoryWebPartStrings;
  export = strings;
}
