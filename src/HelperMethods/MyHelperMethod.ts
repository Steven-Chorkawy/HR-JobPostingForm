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
    let output = [];

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
    let usersDepartments = await GetUsersDepartmentLibraries();
    let departments: IDepartments[] = [];
    for (let i = 0; i < usersDepartments.length; i++) {
        try {
            let divisions = await _sp.web.lists.getByTitle(usersDepartments[i]).fields.getByInternalNameOrTitle('Division').select('Choices')();
            departments.push({
                name: usersDepartments[i],
                divisions: divisions["Choices"]
            });
        } catch (error) {
            // If division column doesn't exist in library then we can skip it. 
        }
    }

    return departments;
};

export const RemoveDuplicateDivisions = (departments: IDepartments[]): string[] => {
    let output = [];
    let divisionList: any = [];
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
    let templateLibrary = await _sp.web.lists.getByTitle(MyLibraries.JobPostingTemplates)
        .select('Title', 'RootFolder/ServerRelativeUrl')
        .expand('RootFolder')();

    /**
     * 01/05/2024 - Ticket: https://clarington.freshservice.com/a/tickets/35141
     * When this project was created the template documents were in the root of the templates library (https://claringtonnet.sharepoint.com/sites/Careers/JobPostingTemplates).
     * HR staff have created a new folder (Master Templates - Employee Requisition & Job Posting) within the templates library and have moved the template files into the new folder.
     */
    const TEMPLATE_FOLDER = `${templateLibrary.RootFolder.ServerRelativeUrl}/${MASTER_TEMPLATE_FOLDER_NAME}`;
    try {
        let templateFolder: any = await _sp.web.getFolderByServerRelativePath(TEMPLATE_FOLDER).expand("Folders, Files")();
        debugger;
        let output = templateFolder.Files
        debugger;
        return output;
    } catch (error) {
        console.error(error);
        console.log('Failed to locate template files!!!');
        return [];
    }
}

export const RemoveStringFromDepartmentName = (departmentName: string) => {
    // This string is present in the name of the libraries.  Removing it is needed in some situations.
    const REMOVE_THIS_STRING = " - Job Files";
    return departmentName.replace(REMOVE_THIS_STRING, "");
}

/**
 * Check in the JobPostingTemplates library to see if there is a Document Set for the given department.
 * @param departmentName Name of the document set we are looking for.  Expecting " - Job Files" to be present.  It will be removed.
 */
export const CheckForTemplateDocumentSet = async (departmentName: string): Promise<boolean> => {
    let templateLibrary = await _sp.web.lists.getByTitle(MyLibraries.JobPostingTemplates).select('Title', 'RootFolder/ServerRelativeUrl').expand('RootFolder')();
    let parsedDepartmentName = RemoveStringFromDepartmentName(departmentName);

    let output = await (await _sp.web.getFolderByServerRelativePath(`${templateLibrary.RootFolder.ServerRelativeUrl}/${parsedDepartmentName}`).select('Exists')()).Exists;

    return output;
};

/**
 * Get a list of characters that cannot be in Folder or Document titles.  These characters are determined by SharePoint and Adobe Sign.
 * '"', '*', ':', '<', '>', '?', '/', '\\', '|' are invalid according to SharePoint. 
 * '#' is invalid because it causes Adobe Sign to crash.
 * @returns An array of invalid characters.
 */
export const GET_INVALID_CHARACTERS: Array<string> = ['"', '*', ':', '<', '>', '?', '/', '\\', '|', '#'];