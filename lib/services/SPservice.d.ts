import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
export default class SPService {
    private _context;
    constructor(_context: WebPartContext | ApplicationCustomizerContext);
    /**
     * Get List Items
     * @param options
     */
    getListItems(filterText: string, listId: string, internalColumnName: string, webUrl?: string): Promise<any[]>;
}
