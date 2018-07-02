import * as React from 'react';
import { IFieldProps } from '../interfaces';
import { EditorState, convertToRaw, ContentState } from 'draft-js';
import { Editor } from 'react-draft-wysiwyg';
import draftToHtml from 'draftjs-to-html';
import htmlToDraft from 'html-to-draftjs';
import { BaseFieldRenderer } from './BaseFieldRenderer';
import './FieldMultilineTextRenderer.css';

export class FieldMultilineTextRenderer extends BaseFieldRenderer {
  public constructor(props: IFieldProps) {
    super(props);
    let val = null;
    if (props.FormFieldValue) {
      val = props.FormFieldValue;
    }

    let editorState = null;
    if (val !== null) {
      const contentBlock = htmlToDraft(val);
      if (contentBlock) {
        const contentState = ContentState.createFromBlockArray(contentBlock.contentBlocks);
        editorState = EditorState.createWithContent(contentState);
      }
    } else {
      editorState = EditorState.createEmpty();
    }

    this.state = {
      ...this.state,
      currentValue: editorState
    };
  }

  protected renderNewForm() {
    return this.renderAllForms(true);
  }

  protected renderEditForm() {
    return this.renderAllForms(true);
  }

  protected renderDispForm() {
    return this.renderAllForms(false);
  }

  private renderAllForms(editable: boolean) {
    return this.getEditorComponent(this.props.IsRichText, editable);
  }

  private getEditorComponent(isRichTextEnabled: boolean, isEditable: boolean): JSX.Element {
    const toolbarStyle = isEditable ? { } : { display: 'none' };
    let contentState = convertToRaw(this.state.currentValue.getCurrentContent());
    if (contentState.blocks.length > 1 && contentState.blocks[0].text === '') {
      contentState.blocks.splice(0, 1);
    }
    let boxStyle = {};
    if (isEditable) {
      boxStyle['border'] = '1px solid #f1f1f1';
    }

    let editorComponent: JSX.Element = null;
    if (isRichTextEnabled) {
      editorComponent = (
        <Editor
          wrapperClassName='wrapper-class'
          editorClassName='editor-class'
          toolbarClassName='toolbar-class'
          readOnly={!isEditable}
          toolbarStyle={toolbarStyle}
          initialContentState={contentState}
          onEditorStateChange={this.onChange}
        />);
    } else {
      editorComponent = (
        <Editor
          toolbar={{}}
          wrapperClassName='wrapper-class'
          editorClassName='editor-class'
          toolbarClassName='toolbar-class'
          readOnly={!isEditable}
          toolbarStyle={{ display: 'none' }}
          initialContentState={contentState}
          onEditorStateChange={this.onChange}
          stripPastedStyles={true}
        />);
    }
    return (<div style={boxStyle}>{editorComponent}</div>);
  }

  private onChange = (editorState) => {
    this.setState({ currentValue: editorState });
    let toSave = draftToHtml(convertToRaw(editorState.getCurrentContent()));
    if (toSave) {
      toSave = toSave.trim();
    }

    if (toSave && (toSave === '<p></p>' || toSave === '')) {
      toSave = null;
    }

    if (toSave && !this.props.IsRichText) {
      let d = document.createElement('div');
      d.innerHTML = toSave;
      toSave = (d.textContent || d.innerText);
    }

    this.trySetChangedValue(toSave);
  }
}
