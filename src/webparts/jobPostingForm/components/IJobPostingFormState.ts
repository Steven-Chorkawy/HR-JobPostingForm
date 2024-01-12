import { IDepartments } from "../../../interfaces/IDepartments";

export interface IJobPostingFormState {
  departmentLibraries: string[];
  divisions?: string[];
  departments: IDepartments[];
}
