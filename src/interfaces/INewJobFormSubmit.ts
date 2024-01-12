export interface INewJobFormSubmit {
    Department: string;
    Division: string;
    // TODO: Determine the Title format. 
    Title: string;      // [JobTitle] - [Division] - [Date?]
    JobTitle: string;   // Title that the user provided.
    PartTimePosition: boolean;
    TemplateFiles: any[];   // Extra templates that the user has selected.
}