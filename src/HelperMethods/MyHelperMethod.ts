import { WebPartContext } from "@microsoft/sp-webpart-base";
import { FormCustomizerContext, ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { SPFI, SPFx, spfi } from "@pnp/sp";

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


