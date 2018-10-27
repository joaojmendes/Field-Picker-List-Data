var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, } from 'office-ui-fabric-react/lib/Button';
import { DocumentCard, 
// DocumentCardActivity,
DocumentCardActions, DocumentCardPreview, DocumentCardTitle } from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import SPservice from "../services/SPservice";
// Links to Icons
export var DOCICONURL_XLSX = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/XLSX_16x3.png";
export var DOCICONURL_DOCX = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/DOCX_16x3.png";
export var DOCICONURL_PPTX = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/PPTX_16x3.png";
export var DOCICONURL_MPPX = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/MPPX_16x3.png";
export var DOCICONURL_PHOTO = "https://sonaesystems.sharepoint.com/sites/prm/HtmlWebParts/raidiDetail/build/static/css/PHOTO.png";
export var DOCICONURL_PDF = "https://sonaesystems.sharepoint.com/sites/prm/HtmlWebParts/raidiDetail/build/static/css/PDF.png";
export var DOCICONURL_TXT = "https://sonaesystems.sharepoint.com/sites/prm/HtmlWebParts/raidiDetail/build/static/css/TXT.png";
var FieldAttachmentsRenderer = (function (_super) {
    __extends(FieldAttachmentsRenderer, _super);
    function FieldAttachmentsRenderer(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            file: '',
            showDialog: false,
            dialogMessage: '',
            Documents: []
        };
        // Get SPService Factory
        _this._spservice = new SPservice(_this.props.context);
        // registo de event handlers
        //
        _this._onDeleteAttachment = _this._onDeleteAttachment.bind(_this);
        _this._closeDialog = _this._closeDialog.bind(_this);
        _this._onUploadFile = _this._onUploadFile.bind(_this);
        return _this;
    }
    // Load Item Attachments
    FieldAttachmentsRenderer.prototype._loadAttachments = function () {
        return __awaiter(this, void 0, void 0, function () {
            var files, _i, files_1, _file, _iconUrl, _fileTypes, _fileExtention, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.previewImages = [];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._spservice.getListItemAttachments(this.props.listId, this.props.itemId, this.props.webUrl)];
                    case 2:
                        files = _a.sent();
                        for (_i = 0, files_1 = files; _i < files_1.length; _i++) {
                            _file = files_1[_i];
                            _iconUrl = '';
                            _fileTypes = _file.ServerRelativeUrl.split('.');
                            _fileExtention = _fileTypes[1];
                            switch (_fileExtention) {
                                case 'XLSX':
                                    _iconUrl = DOCICONURL_XLSX;
                                    break;
                                case 'DOCX':
                                    _iconUrl = DOCICONURL_DOCX;
                                    break;
                                case 'PPTX':
                                    _iconUrl = DOCICONURL_PPTX;
                                    break;
                                case 'TXT':
                                    _iconUrl = '';
                                    break;
                                case 'MPPX':
                                    _iconUrl = DOCICONURL_MPPX;
                                    break;
                                case 'PDF':
                                    _iconUrl = DOCICONURL_PDF;
                                    break;
                                case 'TXT':
                                    _iconUrl = DOCICONURL_TXT;
                                    break;
                                case 'jpg':
                                    _iconUrl = DOCICONURL_PHOTO;
                                    break;
                                case 'png':
                                    _iconUrl = DOCICONURL_PHOTO;
                                    break;
                                case 'gif':
                                    _iconUrl = DOCICONURL_PHOTO;
                                    break;
                                default:
                                    _iconUrl = '';
                                    break;
                            }
                            this.previewImages.push({
                                name: _file.FileName,
                                previewImageSrc: _file.ServerRelativeUrl,
                                iconSrc: _iconUrl,
                                imageFit: ImageFit.cover,
                                width: 200,
                                height: 100,
                            });
                        }
                        ;
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        this.setState({
                            showDialog: true,
                            // tslint:disable-next-line:max-line-length
                            dialogMessage: 'Error on read file Attachments. Error: ' + error_1.message
                        });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    // Run befor render component
    FieldAttachmentsRenderer.prototype.componentDidMount = function () {
        this._loadAttachments();
    };
    // Render Attachments
    FieldAttachmentsRenderer.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", null,
            this.state.Documents.map(function (_file, i) {
                return (React.createElement("div", { className: "DocumentCard", style: { marginTop: 15 } },
                    React.createElement(DocumentCard, { onClickHref: _file.ServerRelativeUrl },
                        React.createElement(DocumentCardPreview, { previewImages: [_this.previewImages[i]] }),
                        React.createElement(DocumentCardTitle, { title: _file.FileName, shouldTruncate: true }),
                        React.createElement(DocumentCardActions, { actions: [
                                {
                                    iconProps: {
                                        iconName: 'Delete',
                                        title: 'Delete',
                                    },
                                    title: 'Delete',
                                    text: 'Delete',
                                    disabled: _this.props.disabled,
                                    className: _this.props.disabled ? 'documentAction-disabled' : 'documentAction',
                                    onClick: function (ev) {
                                        ev.preventDefault();
                                        ev.stopPropagation();
                                        _this._onDeleteAttachment(_file.FileName);
                                    }
                                },
                            ] }))));
            }),
            React.createElement(Dialog, { isOpen: this.state.showDialog, type: DialogType.normal, onDismiss: this._closeDialog, title: "Attachments", subText: this.state.dialogMessage, isBlocking: true },
                React.createElement(DialogFooter, null,
                    React.createElement(PrimaryButton, { onClick: this._closeDialog }, "OK")))));
    };
    // close dialog
    FieldAttachmentsRenderer.prototype._closeDialog = function (e) {
        //
        e.preventDefault();
        this.setState({
            showDialog: false,
            dialogMessage: '',
        });
    };
    // On UploadFIle
    FieldAttachmentsRenderer.prototype._onUploadFile = function (file) {
        //
        this._loadAttachments();
    };
    FieldAttachmentsRenderer.prototype._onDeleteAttachment = function (_file) {
        var _this = this;
        // Delete Attachment
        this._spservice.deleteAttachment(_file.Name, this.props.listId, this.props.itemId, this.props.webUrl)
            .then(function () {
            _this.setState({
                showDialog: true,
                dialogMessage: 'File ' + _file + ' Deleted.',
            });
            _this._loadAttachments();
        })
            .catch(function (reason) {
            _this.setState({
                showDialog: true,
                // tslint:disable-next-line:max-line-length
                dialogMessage: 'Error on delete file: ' + _file + 'Error: ' + reason
            });
        });
    };
    return FieldAttachmentsRenderer;
}(React.Component));
export { FieldAttachmentsRenderer };
export default FieldAttachmentsRenderer;
//# sourceMappingURL=FieldAttachmentsRenderer.js.map