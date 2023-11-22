import React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IListItem } from "../Entities/IListItem";

export interface IAppContextProps {
  title: string;
  webpartContext: WebPartContext;
  themeVariant: IReadonlyTheme;
  msGraphClient: MSGraphClientV3;
  listItems: IListItem[];
}

export const AppContext = React.createContext<IAppContextProps>(undefined);
