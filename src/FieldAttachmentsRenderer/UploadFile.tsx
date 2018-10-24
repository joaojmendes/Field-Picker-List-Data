import * as React from 'react';
import * as $ from 'jquery';
import * as pnp from 'sp-pnp-js';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, } from 'office-ui-fabric-react/lib/Button';
import { CommandButton } from 'office-ui-fabric-react/lib/Button';
//import { Link } from 'office-ui-fabric-react/lib/Link';
export class UploadFile extends React.Component<any, any> {
    constructor(props: any) {
        super(props);
        this.state = {
            file: '',
            imagePreviewUrl: '',
            showDialog: false,
            dialogMessage: '',
        };
        this._closeDialog = this._closeDialog.bind(this);
    }

    _handleSubmit(e: any) {
        e.preventDefault();
        //  
        $('#file-picker').trigger('click');
    }

    _handleImageChange(e: any) {
        e.preventDefault();

        let reader = new FileReader();
        let file = e.target.files[0];

        reader.onloadend = () => {
            this.setState({
                file: file,
                imagePreviewUrl: reader.result
            });
            // Add attachement
            let item = pnp.sp.web.lists.getById(this.props.ListGuid)
                .items.getById(this.props.RaidId);

            item.attachmentFiles.add(file.name, file)
                .then(v => {
                    this.setState({
                        showDialog: true,
                        // tslint:disable-next-line:max-line-length
                        dialogMessage: 'File: ' + file.name + ' Uploaded.'
                    });
                    this.props.onFileUpload();
                })
                .catch((reason: any) => {
                    this.setState({
                        showDialog: true,
                        // tslint:disable-next-line:max-line-length
                        dialogMessage: 'File: ' + file.name + 'Not Uploaded. Error: ' + reason
                    });
                });

        }
        reader.readAsDataURL(file)
    }

    render() {
        let { imagePreviewUrl } = this.state;
        let $imagePreview = null;
        let _button = null;
        if (imagePreviewUrl) {
            $imagePreview = (<img src={imagePreviewUrl} />);
        } else {
            $imagePreview = (<div className="previewText">Please select an file to Preview</div>);
        }

        if (this.props.IconButton === false) {
            _button = (
                < PrimaryButton
                    onClick={(e) => this._handleSubmit(e)}>Upload a File
                </ PrimaryButton>
            );
        } else {
            _button = (
                <CommandButton
                    data-automation-id='Upload'
                    iconProps={{ iconName: 'Upload' }}
                    onClick={(e) => this._handleSubmit(e)}
                    className="upload-file"
                    disabled={this.props.Disabled}
                >
                    Upload a File
        </CommandButton>
            );
        }
        // render compomente
        return (
            <div>
                <input id="file-picker" className="ms-TextField-field" style={{ display: 'none' }}
                    type="file"
                    onChange={(e) => this._handleImageChange(e)} />
                <div style={{ textAlign: 'center', marginTop: 25, marginBottom: 25 }}>
                    {_button}
                </div>
                <div>
                    { /*$imagePreview */}
                </div>
                <Dialog
                    isOpen={this.state.showDialog}
                    type={DialogType.normal}
                    onDismiss={this._closeDialog}
                    title="Upload File"
                    subText={this.state.dialogMessage}
                    isBlocking={true}>
                    <DialogFooter>
                        <PrimaryButton onClick={this._closeDialog}>OK</PrimaryButton>
                    </DialogFooter>
                </Dialog>
            </div>
        );
    }

    // close dialog
    private _closeDialog(e: any) {
        //  
        e.preventDefault();
        this.setState({
            showDialog: false,
            dialogMessage: '',
        });

    }

}
export default UploadFile;