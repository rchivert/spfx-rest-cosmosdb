import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRest2Props {
  description: string;
  ctx: WebPartContext;
  webapp_uri: string;
  webapp_appid: string;
}
