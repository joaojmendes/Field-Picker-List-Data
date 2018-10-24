import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
export interface IFieldAttachmentsRendererProps {
  listId: string;
  itemId: number;
  className?: string;
  webUrl?:string;
  disabled?: boolean;
  context: WebPartContext |  ApplicationCustomizerContext;
}
