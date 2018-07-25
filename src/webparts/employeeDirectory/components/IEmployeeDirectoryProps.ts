import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";

export interface IEmployeeDirectoryProps {
  title: string;
  columns: number;
  exclude: string;
  sortBy: string;
  updateProperty: (value: string) => void;
  displayMode: DisplayMode;
  context: IWebPartContext;
}
