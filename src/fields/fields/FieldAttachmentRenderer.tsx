require('./FieldAttachmentRenderer.min.css');
require('react-dropzone-component/styles/filepicker.css');
import * as React from 'react';
import { IFieldProps, FormMode } from '../interfaces';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { BaseFieldRenderer } from './BaseFieldRenderer';
import { DropzoneComponent } from 'react-dropzone-component';
import { FormFieldsStore } from '../store';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { getFieldPropsByInternalName } from '../utils';

export class FieldAttachmentRenderer extends BaseFieldRenderer {
  public constructor(props: IFieldProps) {
    super(props);
    this.state = {
      ...this.state,
      currentValue: props.FormFieldValue,
      existingToDelete: props.AttachmentsExistingToDelete === undefined ? [] : props.AttachmentsExistingToDelete
    };
  }

  protected renderNewForm() {
    return this.renderAllForms();
  }

  protected renderEditForm() {
    return this.renderAllForms();
  }

  protected renderDispForm() {
    return this.renderAllForms();
  }

  private renderAllForms() {
    return (
      <React.Fragment>
        {this.getUploadPart()}
        {this.getExistingItemsPart()}
      </React.Fragment>
    );
  }

  private getExistingItemsPart(): JSX.Element {
    let attachmentItems = [];
    if (this.state.currentValue) {
      attachmentItems = this.state.currentValue;
    }

    return (
      <React.Fragment>
        {attachmentItems.map((a, i) => {
          let linkStyle = {};
          if (this.state.existingToDelete && this.state.existingToDelete.indexOf(a.FileName) !== -1) {
            linkStyle['textDecoration'] = 'line-through';
          }
          return <div key={`attachmentItemContainer_${i}`}>
            <Link key={`attachment_${i}`} href={a.ServerRelativeUrl} target='_blank' style={linkStyle}>{a.FileName}</Link>
            {this.props.CurrentMode !== FormMode.Display ?
            <IconButton
              key={`attachmentDelete_${i}`} onClick={this.onExistingFileDeleteClick}
              data={a.FileName} iconProps={{ iconName: 'Delete' }}
              style={{ verticalAlign: 'middle', height: '2em' }} /> :
            null}

          </div>;
        })}
      </React.Fragment>);
  }

  private getUploadPart(): JSX.Element {
    let uploadPart = null;
    if (this.props.CurrentMode === FormMode.New || this.props.CurrentMode === FormMode.Edit) {
      let componentConfig = {
        showFiletypeIcon: false,
        disablePreview: true,
        postUrl: 'fake'
      };
      let dzStyle = {
        // height: '55px',
        width: '100%',
        border: '2px black dashed',
        cursor: 'pointer',
        marginBottom: '15px'
      };

      let eventHandlers = {
        drop: this.onDrop,
        addedfile: this.onFileAdded
      };

      let djsConfig = {
        addRemoveLinks: true,
        autoProcessQueue: false
      };

      uploadPart =
        <div style={dzStyle}>
          <DropzoneComponent config={componentConfig} eventHandlers={eventHandlers} djsConfig={djsConfig}>
          </DropzoneComponent>
        </div>;
    }
    return uploadPart;
  }

  private onFileAdded = (file) => {
    FormFieldsStore.actions.addNewAttachmentInfo(file);
  }

  private onExistingFileDeleteClick = (ev) => {
    ev.preventDefault();

    const toDelete = ev.target.closest('button').attributes['data'].value;
    FormFieldsStore.actions.addOrRemoveExistingAttachmentDeletion(toDelete);

    const globalState = FormFieldsStore.actions.getState();
    let attachmentProps = getFieldPropsByInternalName(globalState.Fields, 'Attachments');
    if (attachmentProps) {
      this.setState({ existingToDelete: attachmentProps.AttachmentsExistingToDelete });
    }
  }

  private onDrop = () => {
    // ...
  }
}
