declare interface IApprovalsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleLabel: string;
  TitleARLabel: string;
  LinkTextLabel: string;
  LinkTextARLabel: string;
  RequestsListLabel: string;
  DepartmentLabel: string;
  ViewAllUrlLabel: string;
  NewRequestUrlLabel: string;
  ReviewAllUrlLabel: string;
}

declare module 'ApprovalsWebPartStrings' {
  const strings: IApprovalsWebPartStrings;
  export = strings;
}
