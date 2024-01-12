import { WebPartContext } from "@microsoft/sp-webpart-base";
import { FormCustomizerContext, ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { SPFI, SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/lists";
import "@pnp/sp/security/list";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/content-types";
import { IDepartments } from "../interfaces/IDepartments";
import { MyLibraries } from "../enums/MyLibraries";
import { PermissionKind } from "@pnp/sp/security";
import { INewJobFormSubmit } from "../interfaces/INewJobFormSubmit";
import { IFolderAddResult } from "@pnp/sp/folders";
import { IContentTypeInfo } from "@pnp/sp/content-types";

let _sp: SPFI;
export const setSP = (context: WebPartContext | ListViewCommandSetContext | FormCustomizerContext): SPFI => {
    if (_sp === null && context !== null) {
        _sp = spfi().using(SPFx(context));
    }
    return _sp;
};

export const getSP = (): SPFI => {
    return _sp;
}

/**
 * 01/05/2024 - Ticket: https://clarington.freshservice.com/a/tickets/35141
 * When this project was created the template documents were in the root of the templates library (https://claringtonnet.sharepoint.com/sites/Careers/JobPostingTemplates).
 * HR staff have created a new folder (Master Templates - Employee Requisition & Job Posting) within the templates library and have moved the template files into the new folder.
 */
export const MASTER_TEMPLATE_FOLDER_NAME = "Master Templates - Employee Requisition & Job Posting";


// Check if the current user has edit access to a given library.
export const DoesUserHaveAccessToLibrary = async (libraryName: string): Promise<boolean> => await _sp.web.lists.getByTitle(libraryName).currentUserHasPermissions(PermissionKind.AddListItems);


/**
 * Get a list of libraries that the current user has edit access in.
 * @returns Array of library names. 
 */
export const GetUsersDepartmentLibraries = async (): Promise<string[]> => {
    const libraries = Object.keys(MyLibraries).map(key => MyLibraries[key as keyof typeof MyLibraries]);
    const output: string[] = [];

    for (let i = 0; i < libraries.length; i++) {
        try {
            // Replacing sp.web.lists.getByTitle(libraries[i]).userHasPermissions(user.LoginName, PermissionKind.AddListItems) with DoesUserHaveAccessToLibrary(libraries[i])
            if (libraries[i] !== MyLibraries.JobPostingTemplates && await DoesUserHaveAccessToLibrary(libraries[i])) {
                output.push(libraries[i]);
            }
        } catch (error) {
            // User does not have access to the library, or it has been deleted. 
        }
    }
    return output;
};

/**
 * Get a list of division names from one or more library choice columns. 
 * @param libraries A list of library names that the user has access to.
 * @returns Array of division names.
 */
export const GetDivisions = async (): Promise<IDepartments[]> => {
    const usersDepartments = await GetUsersDepartmentLibraries();
    const departments: IDepartments[] = [];
    for (let i = 0; i < usersDepartments.length; i++) {
        try {
            const divisions = await _sp.web.lists.getByTitle(usersDepartments[i]).fields.getByInternalNameOrTitle('Division').select('Choices')();
            departments.push({
                name: usersDepartments[i],
                divisions: divisions.Choices
            });
        } catch (error) {
            // If division column doesn't exist in library then we can skip it. 
        }
    }

    return departments;
};

export const RemoveDuplicateDivisions = (departments: IDepartments[]): string[] => {
    let output = [];
    const divisionList: any = [];
    departments.map(d => { divisionList.push(...d.divisions as any); });
    //This filters out duplicate divisions.
    output = divisionList.filter((element: any, index: number) => { return divisionList.indexOf(element) === index; });
    return output;
};

/**
 * Get a list of all the template documents found.
 * @returns A list of template files found.
 */
export const GetTemplateDocuments = async (): Promise<any> => {
    const templateLibrary = await _sp.web.lists.getByTitle(MyLibraries.JobPostingTemplates)
        .select('Title', 'RootFolder/ServerRelativeUrl')
        .expand('RootFolder')();

    /**
     * 01/05/2024 - Ticket: https://clarington.freshservice.com/a/tickets/35141
     * When this project was created the template documents were in the root of the templates library (https://claringtonnet.sharepoint.com/sites/Careers/JobPostingTemplates).
     * HR staff have created a new folder (Master Templates - Employee Requisition & Job Posting) within the templates library and have moved the template files into the new folder.
     */
    const TEMPLATE_FOLDER = `${templateLibrary.RootFolder.ServerRelativeUrl}/${MASTER_TEMPLATE_FOLDER_NAME}`;
    try {
        const templateFolder: any = await _sp.web.getFolderByServerRelativePath(TEMPLATE_FOLDER).expand("Folders, Files")();
        const output = templateFolder.Files
        return output;
    } catch (error) {
        console.error(error);
        console.log('Failed to locate template files!!!');
        return [];
    }
}

export const RemoveStringFromDepartmentName = (departmentName: string): string => {
    // This string is present in the name of the libraries.  Removing it is needed in some situations.
    const REMOVE_THIS_STRING = " - Job Files";
    return departmentName.replace(REMOVE_THIS_STRING, "");
}

/**
 * Check in the JobPostingTemplates library to see if there is a Document Set for the given department.
 * @param departmentName Name of the document set we are looking for.  Expecting " - Job Files" to be present.  It will be removed.
 */
export const CheckForTemplateDocumentSet = async (departmentName: string): Promise<boolean> => {
    const templateLibrary = await _sp.web.lists.getByTitle(MyLibraries.JobPostingTemplates).select('Title', 'RootFolder/ServerRelativeUrl').expand('RootFolder')();
    const parsedDepartmentName = RemoveStringFromDepartmentName(departmentName);
    const output = await (await _sp.web.getFolderByServerRelativePath(`${templateLibrary.RootFolder.ServerRelativeUrl}/${parsedDepartmentName}`).select('Exists')()).Exists;

    return output;
};

export const FormatTitle = (jobTitle: string, division: string): string => `${jobTitle} - ${division} - ${new Date().toISOString().slice(0, 10)}`;


export const FormatDocumentSetPath = async (departmentName: string, title: string): Promise<string> => {
    const library = await _sp.web.lists.getByTitle(departmentName).select('Title', 'RootFolder/ServerRelativeUrl').expand('RootFolder')();
    return `${library.RootFolder.ServerRelativeUrl}/${title}`;
};

export const CheckForExistingDocumentSetByServerRelativePath = async (serverRelativePath: string): Promise<boolean> => {
    return await (await _sp.web.getFolderByServerRelativePath(serverRelativePath).select('Exists')()).Exists;
};


export const CheckForExistingDocumentSet = async (title: string, departmentName: string): Promise<boolean> => {
    const FOLDER_NAME = await FormatDocumentSetPath(departmentName, title);
    return await CheckForExistingDocumentSetByServerRelativePath(FOLDER_NAME);
};

export const GetLibraryContentTypes = async (departmentName: string): Promise<string> => {
    const library = await _sp.web.lists.getByTitle(departmentName);
    const output = (await library.contentTypes()).find((f: IContentTypeInfo) => (f.Group === "Custom Content Types" || f.Group === "Job Posting Content Types") && f.StringId.includes('0x0120'))?.StringId;
    return output ? output : "";
};

export const CreateDocumentSet = async (input: INewJobFormSubmit): Promise<void> => {
    let newFolderResult: IFolderAddResult;
    const FOLDER_NAME = await FormatDocumentSetPath(input.Department, input.Title);
    let libraryDocumentSetContentTypeId;

    try {
        libraryDocumentSetContentTypeId = await GetLibraryContentTypes(input.Department);
        if (!libraryDocumentSetContentTypeId) {
            throw new Error("Error! Cannot get content type for library.");
        }

        // Because sp.web.folders.add overwrites existing folder I have to do a manual check.
        if (await CheckForExistingDocumentSetByServerRelativePath(FOLDER_NAME)) {
            throw new Error(`Error! Cannot Create new Document Set. Duplicate Name detected. "${FOLDER_NAME}"`);
        }

        newFolderResult = await _sp.web.folders.addUsingPath(FOLDER_NAME);
    } catch (error) {
        console.error(error);
        throw error;
    }

    const newFolderProperties = await _sp.web.getFolderByServerRelativePath(newFolderResult.data.ServerRelativeUrl).listItemAllFields();

    await _sp.web.lists.getByTitle(input.Department).items.getById(newFolderProperties.ID).update({
        ContentTypeId: libraryDocumentSetContentTypeId,
        ///HTML_x0020_File_x0020_Type: "SharePoint.DocumentSet",
        PartTimePosition: input.PartTimePosition,
        Division: input.Division,
        ApprovalStatus: "New"
    });

    try {
        // input.TemplateFiles might be undefined.  CopyTemplateDocuments will determine how it is handled.
        await CopyTemplateDocuments(input.Department, input.Title, input.TemplateFiles);
    }
    catch (error) {
        console.log('Failed to copy template documents.');
        console.log(error);
        throw error;
    }
};

export const CopyTemplateDocuments = async (department: string, documentSetName: string, extraTemplateFiles?: any[]): Promise<void> => {
    const destinationLibrary = await _sp.web.lists.getByTitle(department)
        .select('Title', 'RootFolder/ServerRelativeUrl')
        .expand('RootFolder')();

    // let templateLibrary = await sp.web.lists.getByTitle(MyLibraries.JobPostingTemplates)
    //     .select('Title', 'RootFolder/ServerRelativeUrl')
    //     .expand('RootFolder')
    //     .get();


    // const TEMPLATE_FOLDER = `${templateLibrary.RootFolder.ServerRelativeUrl}/${MASTER_TEMPLATE_FOLDER_NAME}`;
    // let library: any = await sp.web.getFolderByServerRelativeUrl(TEMPLATE_FOLDER).expand("Folders, Files").get();

    const destinationUrl = `${destinationLibrary.RootFolder.ServerRelativeUrl}/${documentSetName}`;

    const templateFiles = await GetTemplateDocuments();

    templateFiles.forEach((value: any) => {
        if (extraTemplateFiles !== undefined && value.Name.includes('Requisition')) {
            _sp.web.getFileByServerRelativePath(value.ServerRelativeUrl).copyTo(`${destinationUrl}/${value.Name}`, false);

        }
        if (extraTemplateFiles === undefined) {
            _sp.web.getFileByServerRelativePath(value.ServerRelativeUrl).copyTo(`${destinationUrl}/${value.Name}`, false);
        }
    });

    // Copy over any template files the user has selected.
    if (extraTemplateFiles) {
        extraTemplateFiles.forEach(value => {
            _sp.web.getFileByUrl(value.fileAbsoluteUrl).copyTo(`${destinationUrl}/${value.fileName}`, false);
        });
    }
};

/**
 * Get a list of characters that cannot be in Folder or Document titles.  These characters are determined by SharePoint and Adobe Sign.
 * '"', '*', ':', '<', '>', '?', '/', '\\', '|' are invalid according to SharePoint. 
 * '#' is invalid because it causes Adobe Sign to crash.
 * @returns An array of invalid characters.
 */
export const GET_INVALID_CHARACTERS: Array<string> = ['"', '*', ':', '<', '>', '?', '/', '\\', '|', '#'];