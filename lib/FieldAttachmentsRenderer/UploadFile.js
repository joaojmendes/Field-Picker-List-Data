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
import * as React from 'react';
import * as $ from 'jquery';
import * as pnp from 'sp-pnp-js';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, } from 'office-ui-fabric-react/lib/Button';
import { CommandButton } from 'office-ui-fabric-react/lib/Button';
//import { Link } from 'office-ui-fabric-react/lib/Link';
var UploadFile = (function (_super) {
    __extends(UploadFile, _super);
    function UploadFile(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            file: '',
            imagePreviewUrl: '',
            showDialog: false,
            dialogMessage: '',
        };
        _this._closeDialog = _this._closeDialog.bind(_this);
        return _this;
    }
    UploadFile.prototype._handleSubmit = function (e) {
        e.preventDefault();
        //  
        $('#file-picker').trigger('click');
    };
    UploadFile.prototype._handleImageChange = function (e) {
        var _this = this;
        e.preventDefault();
        var reader = new FileReader();
        var file = e.target.files[0];
        reader.onloadend = function () {
            _this.setState({
                file: file,
                imagePreviewUrl: reader.result
            });
            // Add attachement
            var item = pnp.sp.web.lists.getById(_this.props.ListGuid)
                .items.getById(_this.props.RaidId);
            item.attachmentFiles.add(file.name, file)
                .then(function (v) {
                _this.setState({
                    showDialog: true,
                    // tslint:disable-next-line:max-line-length
                    dialogMessage: 'File: ' + file.name + ' Uploaded.'
                });
                _this.props.onFileUpload();
            })
                .catch(function (reason) {
                _this.setState({
                    showDialog: true,
                    // tslint:disable-next-line:max-line-length
                    dialogMessage: 'File: ' + file.name + 'Not Uploaded. Error: ' + reason
                });
            });
        };
        reader.readAsDataURL(file);
    };
    UploadFile.prototype.render = function () {
        var _this = this;
        var imagePreviewUrl = this.state.imagePreviewUrl;
        var $imagePreview = null;
        var _button = null;
        if (imagePreviewUrl) {
            $imagePreview = (React.createElement("img", { src: imagePreviewUrl }));
        }
        else {
            $imagePreview = (React.createElement("div", { className: "previewText" }, "Please select an file to Preview"));
        }
        if (this.props.IconButton === false) {
            _button = (React.createElement(PrimaryButton, { onClick: function (e) { return _this._handleSubmit(e); } }, "Upload a File"));
        }
        else {
            _button = (React.createElement(CommandButton, { "data-automation-id": 'Upload', iconProps: { iconName: 'Upload' }, onClick: function (e) { return _this._handleSubmit(e); }, className: "upload-file", disabled: this.props.Disabled }, "Upload a File"));
        }
        // render compomente
        return (React.createElement("div", null,
            React.createElement("input", { id: "file-picker", className: "ms-TextField-field", style: { display: 'none' }, type: "file", onChange: function (e) { return _this._handleImageChange(e); } }),
            React.createElement("div", { style: { textAlign: 'center', marginTop: 25, marginBottom: 25 } }, _button),
            React.createElement("div", null),
            React.createElement(Dialog, { isOpen: this.state.showDialog, type: DialogType.normal, onDismiss: this._closeDialog, title: "Upload File", subText: this.state.dialogMessage, isBlocking: true },
                React.createElement(DialogFooter, null,
                    React.createElement(PrimaryButton, { onClick: this._closeDialog }, "OK")))));
    };
    // close dialog
    UploadFile.prototype._closeDialog = function (e) {
        //  
        e.preventDefault();
        this.setState({
            showDialog: false,
            dialogMessage: '',
        });
    };
    return UploadFile;
}(React.Component));
export { UploadFile };
export default UploadFile;
//# sourceMappingURL=UploadFile.js.map