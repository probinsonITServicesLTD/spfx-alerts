export interface ISiteData {
    logoUrl: string;
    name: string;
    navigation: any[];
    themeKey: string;
    url: string;
    usesMetadataNavigation: boolean;
  }
  
  export interface ISiteDataResponse {
    "@odata.null"?: boolean;
    value?: string;
  }