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
