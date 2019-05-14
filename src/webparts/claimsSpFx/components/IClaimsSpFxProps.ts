import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IClaimsSpFxProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
}
