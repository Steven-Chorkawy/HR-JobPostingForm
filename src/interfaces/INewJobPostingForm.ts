import { NewJobPostingFormStatus } from "../enums/NewJobFormStatus";
import { IDepartments } from "./IDepartments";

export interface INewJobFormState {
    departments?: string[];
    divisions?: string[];
    divisionDefaultSelectedKey?: string;
    departmentDivisions?: IDepartments[];
    formResponse: NewJobPostingFormStatus;
    documentSetPath?: string;

    // Used to determine if the File picker input should be visible or not. 
    showExtraFilePicker: boolean;

    // Name of the selected department with " - Job Files" removed.  This name should be the name of a document set in the template library. 
    extraTemplateDocSetName?: string;
    templateFiles?: any[];
    templateFilesFound?: boolean;
}

export interface INewJobPostingFormProps {
    description: string;
    departmentLibraries: string[];
    divisions?: string[];
    departments: IDepartments[];
    context: any;
}