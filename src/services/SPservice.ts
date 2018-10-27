import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { sp, Web } from "@pnp/sp";

export default class SPService {
  constructor(private _context: WebPartContext | ApplicationCustomizerContext) {
    sp.setup({
      spfxContext: this._context
    });
  }
  /**
   * Get List Items
   *
   */
  public async getListItems(filterText: string, listId: string, internalColumnName: string, webUrl?: string): Promise<any[]> {
    let filter = `startswith(${internalColumnName},'${filterText}')`;
    let returnItems: any[];
    console.log(
      `Page context url ${this._context.pageContext.web.absoluteUrl}`
    );
    let spWeb: Web;
    if (typeof webUrl === undefined) {
      spWeb = new Web(webUrl);
    } else {
      spWeb = new Web(this._context.pageContext.web.absoluteUrl);
    }
    try {
      returnItems = await spWeb.lists
        .getById(listId)
        .items.select("Id", internalColumnName)
        .filter(filter)
        .get();
      return Promise.resolve(returnItems);
    } catch (error) {
      return Promise.reject(error);
    }
  }

  // Get ListAttachments
  public async getListItemAttachments(
    listId: string,
    itemId: number,
    webUrl?: string
  ): Promise<any[]> {
    let returnFiles: any[];
    let spWeb: Web;
    if (typeof webUrl === undefined) {
      spWeb = new Web(webUrl);
    } else {
      spWeb = new Web(this._context.pageContext.web.absoluteUrl);
    }
    try {
      let files = await spWeb.lists
        .getById(listId)
        .items.getById(itemId)
        .attachmentFiles.get();
      return Promise.resolve(files);
    } catch (error) {
      return Promise.reject(error);
    }
  }

  // delete attachement
  public async deleteAttachment(
    fileName: string,
    listId: string,
    itemId: number,
    webUrl?: string
  ): Promise<void> {
    let spWeb: Web;
    if (typeof webUrl === undefined) {
      spWeb = new Web(webUrl);
    } else {
      spWeb = new Web(this._context.pageContext.web.absoluteUrl);
    }
    try {
      await spWeb.lists
        .getById(listId)
        .items.getById(itemId)
        .attachmentFiles.getByName(fileName)
        .delete();
      return;
    } catch (error) {
      return Promise.reject(error);
    }
  }
}
