import { IUserItem } from "./IUserItem";

export interface IEmployeeDirectoryState {
  users: IUserItem[];
  search: string;
  loading: boolean;
}