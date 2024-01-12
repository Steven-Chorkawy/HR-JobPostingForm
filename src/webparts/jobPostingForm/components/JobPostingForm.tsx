import * as React from 'react';
import type { IJobPostingFormProps } from './IJobPostingFormProps';
import LoadingComponent from '../../../MyComponents/LoadingComponent';
import { IJobPostingFormState } from './IJobPostingFormState';
import { IDepartments } from '../../../interfaces/IDepartments';
import { GetDivisions, RemoveDuplicateDivisions } from '../../../HelperMethods/MyHelperMethod';
import NewJobPostingForm from '../../../MyComponents/NewJobPostingForm';

export default class JobPostingForm extends React.Component<IJobPostingFormProps, IJobPostingFormState> {
  constructor(props: IJobPostingFormProps) {
    super(props);
    this.state = {
      departmentLibraries: [],
      divisions: undefined,
      departments: []
    };
  }

  public render(): React.ReactElement<IJobPostingFormProps> {
    return this.state.departmentLibraries && this.state.divisions ?
      <NewJobPostingForm {...this.props} {...this.state} /> :
      <LoadingComponent />
  }

  public componentDidMount() {
    GetDivisions().then((departments: IDepartments[]) => {
      this.setState({
        departments: departments,
        departmentLibraries: departments.map(d => d.name),
        divisions: RemoveDuplicateDivisions(departments)
      });
    });
  }
}
