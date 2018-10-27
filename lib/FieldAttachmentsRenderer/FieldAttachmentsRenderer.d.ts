/// <reference types="react" />
import * as React from 'react';
import { IFieldAttachmentsRendererProps } from './IFieldAttachmentsRendererProps';
export declare const DOCICONURL_XLSX = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/XLSX_16x3.png";
export declare const DOCICONURL_DOCX = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/DOCX_16x3.png";
export declare const DOCICONURL_PPTX = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/PPTX_16x3.png";
export declare const DOCICONURL_MPPX = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/MPPX_16x3.png";
export declare const DOCICONURL_PHOTO = "https://sonaesystems.sharepoint.com/sites/prm/HtmlWebParts/raidiDetail/build/static/css/PHOTO.png";
export declare const DOCICONURL_PDF = "https://sonaesystems.sharepoint.com/sites/prm/HtmlWebParts/raidiDetail/build/static/css/PDF.png";
export declare const DOCICONURL_TXT = "https://sonaesystems.sharepoint.com/sites/prm/HtmlWebParts/raidiDetail/build/static/css/TXT.png";
export declare class FieldAttachmentsRenderer extends React.Component<IFieldAttachmentsRendererProps, any> {
    private _spservice;
    private previewImages;
    constructor(props: IFieldAttachmentsRendererProps);
    private _loadAttachments();
    componentDidMount(): void;
    render(): JSX.Element;
    private _closeDialog(e);
    private _onUploadFile(file);
    private _onDeleteAttachment(_file);
}
export default FieldAttachmentsRenderer;
