import * as React from 'react';
import { INewJobFormState, INewJobPostingFormProps } from '../interfaces/INewJobPostingForm';
import { IDepartments } from '../interfaces/IDepartments';
import { NewJobPostingFormStatus } from '../enums/NewJobFormStatus';
import { CheckForTemplateDocumentSet, GET_INVALID_CHARACTERS, GetTemplateDocuments, RemoveStringFromDepartmentName } from '../HelperMethods/MyHelperMethod';
import { Field, FieldWrapper, FormRenderProps } from '@progress/kendo-react-form';
import { ComboBox, IComboBoxOption, TextField } from '@fluentui/react';


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

    /**
 * Steps defined to reset the extra template to their initial state. 
 * 
 * Default State...
 *  showExtraFilePicker: false.
 *  extraTemplateDocSetName: undefined.     * 
 */
    private _clearExtraTemplates = (formRenderProps: FormRenderProps) => {
        this.setState({
            // showExtraFilePicker: false,          // No need to reset this because CheckForTemplateDocumentSet() will determine if it should be True or False.
            // extraTemplateDocSetName: undefined,  // This will also be set by CheckForTemplateDocumentSet().
            templateFiles: []
        });

        formRenderProps.onChange('TemplateFiles', { value: undefined });
    }

    //#region Validators
    private titleValidator = (value: any) => {
        if (!value)
            return "Please Enter a title.  Titles cannot contain the following characters.  \" * : < > ? / \\ | #";
        return GET_INVALID_CHARACTERS.some(v => { return value.includes(v); }) ? "Title cannot contain the following characters.  \" * : < > ? / \\ | #" : "";
    }

    private departmentValidator = (value: any) => {
        return value ? "" : "Please select a department.";
    }

    private divisionValidator = (value: any) => {
        return value ? "" : "Please select a division.";
    }
    //#endregion

    private LabelTitleInput = (fieldRenderProps: any) => {
        const { validationMessage, visited, label, id, valid, ...others } = fieldRenderProps;
        const showValidationMessage = visited && validationMessage;
        return (
            <FieldWrapper>
                <TextField label={label} id={id} {...others} errorMessage={showValidationMessage && validationMessage} />
            </FieldWrapper>
        );
    }

    private LabelComboboxInput = (fieldRenderProps: any) => {
        const { validationMessage, visited, label, id, valid, ...others } = fieldRenderProps;
        const showValidationMessage = visited && validationMessage;
        const options: IComboBoxOption[] = fieldRenderProps.data && fieldRenderProps.data.map((d: any) => { return { key: d, text: d }; });
        return (
            <FieldWrapper>
                <ComboBox
                    id={id}
                    defaultSelectedKey={fieldRenderProps.defaultSelectedKey && fieldRenderProps.defaultSelectedKey}
                    label={label}
                    options={options}
                    disabled={fieldRenderProps.disabled}
                    errorMessage={showValidationMessage && validationMessage}
                    onChange={(event, option, index) => { if (fieldRenderProps.onChange) { fieldRenderProps.onChange(option?.text); } }}
                />
            </FieldWrapper>
        );
    }

    private FilerPickerInput = (fieldRenderProps: any) => {
        const { label, onSave, ...others } = fieldRenderProps;
        const FOLDER_PATH = `https://claringtonnet.sharepoint.com/sites/Careers/JobPostingTemplates/${this.state.extraTemplateDocSetName}`;

        return (
            <FieldWrapper>
                <FilePicker
                    buttonIcon="FileImage"
                    label={label}
                    buttonLabel={"Select Part Time Template File"}
                    onSave={(filePickerResult: IFilePickerResult[]) => onSave(filePickerResult)}
                    context={this.props.context}
                    defaultFolderAbsolutePath={FOLDER_PATH}
                    hideRecentTab={true}
                    hideWebSearchTab={true}
                    hideStockImages={true}
                    hideOrganisationalAssetTab={true}
                    hideOneDriveTab={true}
                    hideLocalUploadTab={true}
                    hideLocalMultipleUploadTab={true}
                    hideLinkUploadTab={true}
                    {...others}
                />
            </FieldWrapper>
        );
    }

    //#region Private Methods
    private _getDivisions = (departments: IDepartments[], selectedDepartment: string): string[] => {
        let output = [];
        departments.filter(f => f.name === selectedDepartment).map(f => output.push(...f.divisions));
        return output;
    }

    private _areDepartmentsEmpty = () => (this.props.departments && this.props.departments.length === 0) ? true : false;
    private _areDivisionsEmpty = () => (this.props.divisions && this.props.divisions.length === 0) ? true : false;

    private _onSubmit = async (e): Promise<void> => {
        try {
            e.Title = FormatTitle(e.JobTitle, e.Division);
            let checkRes = await CheckForExistingDocumentSet(e.Title, e.Department);

            this.setState({
                formResponse: checkRes ? NewJobPostingFormStatus.DuplicateName : NewJobPostingFormStatus.New,
                documentSetPath: await FormatDocumentSetPath(e.Department, encodeURIComponent(e.Title))
            });

            if (!checkRes) {
                CreateDocumentSet(e).then(value =>
                    this.setState({ formResponse: NewJobPostingFormStatus.Success })
                ).catch(reason => {
                    this.setState({ formResponse: NewJobPostingFormStatus.Failed });
                });
            }
        }
        catch (e) {
            console.error("onSubmit has failed.");
            console.error(e);
            this.setState({ formResponse: NewJobPostingFormStatus.Failed });
        }
    }
    //#endregion

    public render(): React.ReactElement<any> {
        return (
            <div style={{ padding: '2em' }}>
                <div style={{ fontSize: FontSizes.size32 }}>{this.props.description}</div>
                {
                    this.state.templateFilesFound === false &&
                    <MessageBar messageBarType={MessageBarType.warning}>
                        <div>Cannot locate template documents!  All new Job Posting Folder will not contain template documents.</div>
                        <div>Please go to <a href='https://claringtonnet.sharepoint.com/sites/Careers/JobPostingTemplates' data-interception="off" target="_blank">THE TEMPLATE LIBRARY</a> and confirm that the '{MASTER_TEMPLATE_FOLDER_NAME}' folder exists and contains one or more files.</div>
                    </MessageBar>}
                <hr />
                <Form
                    onSubmit={(e: INewJobFormSubmit) => this._onSubmit(e)}
                    initialValues={{
                        "Department": this.props.departmentLibraries && this.props.departmentLibraries.length === 1 ? this.props.departmentLibraries[0] : null,
                        PartTimePosition: false
                    }}
                    render={(formRenderProps) => (
                        <FormElement style={{ maxWidth: 650 }}>
                            <Field
                                id={"Department"}
                                name={"Department"}
                                label={"* Select Department"}
                                component={this.LabelComboboxInput}
                                data={this.props.departmentLibraries}
                                validator={this.departmentValidator}
                                defaultSelectedKey={formRenderProps.valueGetter('Department')}
                                disabled={this._areDepartmentsEmpty()}
                                onChange={value => {
                                    // this._clearExtraTemplates(formRenderProps);
                                    formRenderProps.onChange('Department', { value: value });
                                    if (this.props.departments) {
                                        let newDivisions = this._getDivisions(this.props.departments, value);
                                        if (newDivisions && newDivisions.length > 1) {
                                            this.setState({
                                                divisions: newDivisions,
                                                divisionDefaultSelectedKey: newDivisions[0]
                                            });
                                            formRenderProps.onChange('Division', { value: newDivisions[0] });
                                        }
                                    }

                                    // Check to see if there are more templates available for this department.
                                    CheckForTemplateDocumentSet(value).then(v => {
                                        this.setState({
                                            showExtraFilePicker: v,
                                            extraTemplateDocSetName: v ? RemoveStringFromDepartmentName(value) : "",
                                            templateFiles: []
                                        });
                                        formRenderProps.onChange('TemplateFiles', { value: undefined });
                                    });
                                }}
                            />
                            {
                                this._areDepartmentsEmpty() &&
                                <MessageBar messageBarType={MessageBarType.error}>
                                    Unable to load your departments.  Please try refreshing this page or contact <Link href="mailto:helpdesk@clarington.net?subject=Careers Job Posting Form Cannot Load Departments" target="_blank" underline>helpdesk@clarington.net</Link>
                                </MessageBar>
                            }
                            <Field
                                id={"Division"}
                                name={"Division"}
                                label={"* Select Division"}
                                key={JSON.stringify(this.state.divisions)}
                                component={this.LabelComboboxInput}
                                data={this.state.divisions}
                                validator={this.divisionValidator}
                                defaultSelectedKey={this.state.divisionDefaultSelectedKey && this.state.divisionDefaultSelectedKey}
                                disabled={this._areDivisionsEmpty()}
                                onChange={value => formRenderProps.onChange('Division', { value: value })}
                            />
                            {
                                this._areDivisionsEmpty() &&
                                <MessageBar messageBarType={MessageBarType.error}>
                                    Unable to load your divisions.  Please try refreshing this page or contact <Link href="mailto:helpdesk@clarington.net?subject=Careers Job Posting Form Cannot Load Divisions" target="_blank" underline>helpdesk@clarington.net</Link>
                                </MessageBar>
                            }
                            <Field
                                id={"JobTitle"}
                                name={"JobTitle"}
                                label={"* Job Title"}
                                component={this.LabelTitleInput}
                                validator={this.titleValidator}
                            />
                            <Field
                                id={'PartTimePosition'}
                                name={'PartTimePosition'}
                                label={'Part Time Position'}
                                component={Toggle}
                                defaultChecked={false}
                                onText={"Yes."}
                                offText={'No.'}
                                onChanged={checked => {
                                    this._clearExtraTemplates(formRenderProps);
                                    formRenderProps.onChange('PartTimePosition', { value: checked });
                                }}
                            />
                            {
                                this.state.showExtraFilePicker && formRenderProps.valueGetter('PartTimePosition') === true &&
                                <div>
                                    <Field
                                        id={'TemplateFile'}
                                        name={'TemplateFile'}
                                        label={'Part Time Template File'}
                                        component={this.FilerPickerInput}
                                        onSave={(value: any) => {
                                            formRenderProps.onChange('TemplateFiles', { value: value });
                                            this.setState({ templateFiles: value });
                                        }}
                                    />
                                    Files selected: {this.state.templateFiles.length}
                                    <ul>
                                        {this.state.templateFiles.map(item => (
                                            // <li key={item.fileName}><a href={item.fileAbsoluteUrl} target="_blank" rel="noreferrer">{item.fileName}</a></li>
                                            <li key={item.fileName}>{item.fileName}</li>
                                        ))}
                                    </ul>
                                </div>
                            }
                            <hr />
                            <div className="k-form-buttons" style={{ marginTop: "20px" }}>
                                <Stack horizontal tokens={{ childrenGap: 40 }}>
                                    <PrimaryButton text="Submit" type="submit" disabled={this._areDepartmentsEmpty() || this._areDivisionsEmpty()} />
                                    <DefaultButton
                                        text="Clear"
                                        onClick={e => {
                                            e.preventDefault();
                                            formRenderProps.onFormReset();
                                            this.setState({ formResponse: NewJobPostingFormStatus.New });
                                            this._clearExtraTemplates(formRenderProps);
                                        }}
                                    />
                                </Stack>
                            </div>
                            <div style={{ padding: "5px" }}>
                                <Card>
                                    <CardBody>
                                        {
                                            formRenderProps.valueGetter("JobTitle") && formRenderProps.valueGetter('Division') ?
                                                <CardTitle>
                                                    {FormatTitle(formRenderProps.valueGetter('JobTitle'), formRenderProps.valueGetter('Division'))}
                                                </CardTitle> :
                                                <CardSubtitle>Please enter a Division, and Job Title.</CardSubtitle>
                                        }
                                        <div style={{ margin: "5px" }}>
                                            {
                                                this.state.formResponse === NewJobPostingFormStatus.Success &&
                                                <MessageBar messageBarType={MessageBarType.success} isMultiline={true}>
                                                    <div>Your Job Posting folder has successfully been created!</div>
                                                    <div>
                                                        {
                                                            this.state.documentSetPath &&
                                                            <div>
                                                                <b>NEXT STEP: </b><Link href={this.state.documentSetPath} target="_blank" underline>Click Here to work on the Requisition and Job Posting.</Link>
                                                            </div>
                                                        }
                                                    </div>
                                                </MessageBar>
                                            }
                                            {
                                                this.state.formResponse === NewJobPostingFormStatus.Failed &&
                                                <MessageBar messageBarType={MessageBarType.error}>
                                                    Error! Unable to submit form. Please contact <Link href="mailto:helpdesk@clarington.net?subject=Careers Job Posting Form" target="_blank" underline>helpdesk@clarington.net</Link>
                                                </MessageBar>
                                            }
                                            {
                                                this.state.formResponse === NewJobPostingFormStatus.DuplicateName &&
                                                <MessageBar messageBarType={MessageBarType.warning}>
                                                    <div>Duplicate Name.  Please Change Division or Job Title.</div>
                                                    <div>
                                                        {
                                                            this.state.documentSetPath &&
                                                            <Link href={this.state.documentSetPath} target="_blank" underline>Click Here to View {FormatTitle(formRenderProps.valueGetter("JobTitle"), formRenderProps.valueGetter("Division"))}</Link>
                                                        }
                                                    </div>
                                                </MessageBar>
                                            }
                                        </div>
                                    </CardBody>
                                </Card>
                            </div>
                        </FormElement>
                    )}
                />
                <PackageSolutionVersion />
            </div>
        );
    }
}
