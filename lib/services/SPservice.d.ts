import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
export default class SPService {
    private _context;
    constructor(_context: WebPartContext | ApplicationCustomizerContext);
    /**
     * Get List Items
     *
     */
    getListItems(filterText: string, listId: string, internalColumnName: string, webUrl?: string): Promise<any[]>;
    getListItemAttachments(listId: string, itemId: number, webUrl?: string): Promise<any[]>;
    deleteAttachment(fileName: string, listId: string, itemId: number, webUrl?: string): Promise<void>;
}
