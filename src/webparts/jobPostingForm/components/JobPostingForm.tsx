import * as React from 'react';
import type { IJobPostingFormProps } from './IJobPostingFormProps';
import LoadingComponent from '../../../MyComponents/LoadingComponent';
import { IJobPostingFormState } from './IJobPostingFormState';
import { IDepartments } from '../../../interfaces/IDepartments';
import { GetDivisions, RemoveDuplicateDivisions } from '../../../HelperMethods/MyHelperMethod';

export default class JobPostingForm extends React.Component<IJobPostingFormProps, IJobPostingFormState> {
  constructor(props: IJobPostingFormProps) {
    super(props);
    this.state = {
      departmentLibraries: undefined,
      divisions: undefined,
      departments: []
    };
  }

  public render(): React.ReactElement<IJobPostingFormProps> {
    return this.state.departmentLibraries && this.state.divisions ?
      <div>hello world</div> :
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
