import * as React from 'react';
import { INewJobFormState, INewJobPostingFormProps } from '../interfaces/INewJobPostingForm';
import { IDepartments } from '../interfaces/IDepartments';
import { NewJobPostingFormStatus } from '../enums/NewJobFormStatus';
import { CheckForTemplateDocumentSet, GetTemplateDocuments, RemoveStringFromDepartmentName } from '../HelperMethods/MyHelperMethod';


export default class NewJobPostingForm extends React.Component<INewJobPostingFormProps, INewJobFormState> {
    constructor(props: INewJobPostingFormProps) {
        super(props);
        this.state = {
            divisions: this.props.departmentLibraries.length === 1
                ? this._getDivisions(this.props.departments, this.props.departmentLibraries[0])
                : [],
            formResponse: NewJobPostingFormStatus.New,
            showExtraFilePicker: false,
            templateFiles: []
        };

        GetTemplateDocuments().then((value: any) => {
            this.setState({ templateFilesFound: value.length > 0 });
        });

        // Check to see if the user only has access to one library.
        // This needs to be run here because the Department onChange event will never be called if a user only has one department.
        if (this.props.departmentLibraries && this.props.departmentLibraries.length === 1) {
            // Check to see if there are more templates available for this department.
            CheckForTemplateDocumentSet(this.props.departmentLibraries[0]).then(v => {
                this.setState({
                    showExtraFilePicker: v,
                    extraTemplateDocSetName: v ? RemoveStringFromDepartmentName(this.props.departmentLibraries[0]) : "",
                    templateFiles: []
                });
            });
        }
    }

    //#region Private Methods.
    private _getDivisions = (departments: IDepartments[], selectedDepartment: string): string[] => {
        let output: any = [];
        departments.filter(f => f.name === selectedDepartment).map(f => output.push(...f.divisions as any));
        return output;
    }
    //#endregion
}
