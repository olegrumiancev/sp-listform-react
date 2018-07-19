import * as React from 'react';
import { IFieldProps, FormMode } from '../interfaces';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { BaseFieldRenderer } from './BaseFieldRenderer';
import './FieldUrlRenderer.css';
import { FormFieldsStore } from '../store';
import { getFieldPropsByInternalName } from '../utils';

export class FieldUrlRenderer extends BaseFieldRenderer {
  public constructor(props: IFieldProps) {
    super(props);

    let urlPart = undefined;
    let descPart = undefined;
    if (props.FormFieldValue) {
      if (typeof props.FormFieldValue === 'string') {
        let vals = this.getSplitValues(props.FormFieldValue);
        urlPart = vals[0];
        descPart = vals[1];
      } else {
        urlPart = props.FormFieldValue.Url;
        descPart = props.FormFieldValue.Description;
      }
    }

    this.state = {
      ...this.state,
      urlPart,
      descPart
    };

    FormFieldsStore.actions.addValidatorToField(this.validateUrlLength, props.InternalName);
  }

  public componentDidMount() {
    if (this.state.urlPart !== undefined) {
      this.trySetChangedValue(`${this.state.urlPart}, ${this.state.descPart}`);
    } else {
      this.trySetChangedValue('');
    }
  }

  protected renderNewForm() {
    return this.renderNewOrEditForm();
  }

  protected renderEditForm() {
    return this.renderNewOrEditForm();
  }

  protected renderDispForm() {
    if (this.state.urlPart) {
      if (this.props.UrlRenderAsPicture) {
        return (
          <div className='url-field-picture'>
            <a target='_blank' href={this.state.urlPart}>
              <img src={this.state.urlPart} alt={this.state.descPart} />
            </a>
          </div>
        );
      } else {
        return <a target='_blank' href={this.state.urlPart}>{this.state.descPart}</a>;
      }
    } else {
      return null;
    }
  }

  private renderNewOrEditForm() {
    return (
      <React.Fragment>
        <TextField
          onChanged={(newValue) => this.onValueChange(newValue, true)}
          value={this.state.urlPart == null ? '' : this.state.urlPart}
          placeholder={`Enter a URL`}
        />
        <TextField
          onChanged={(newValue) => this.onValueChange(newValue, false)}
          value={this.state.descPart == null ? '' : this.state.descPart}
          placeholder={`Enter display text`}
        />
      </React.Fragment>
    );
  }

  private getSplitValues = (val: string): string[] => {
    if (val == null || val.indexOf(', ') === -1) {
      return null;
    }

    let splitIndex = val.indexOf(', ');
    let urlPart = val.substring(0, splitIndex);
    let descPart = val.substring(splitIndex + 2);
    return [urlPart, descPart];
  }

  private onValueChange = (newValue: string, isUrl: boolean) => {
    if (isUrl) {
      this.setState({ urlPart: newValue }, () => {
        if (this.state.urlPart === null || this.state.urlPart === '') {
          this.trySetChangedValue('');
        } else {
          this.trySetChangedValue(`${this.state.urlPart}, ${this.state.descPart ? this.state.descPart : this.state.urlPart}`);
        }
      });
    } else {
      this.setState({ descPart: newValue }, () => {
        if (this.state.urlPart) {
          this.trySetChangedValue(`${this.state.urlPart}, ${this.state.descPart ? this.state.descPart : this.state.urlPart}`);
        }
      });
    }
  }

  private validateUrlLength = (internalName: string): string => {
    const globalState = FormFieldsStore.actions.getState();
    let fieldProps = getFieldPropsByInternalName(globalState.Fields, internalName);
    if (!fieldProps) {
      return `Could not find a field by internal name '${internalName}'`;
    }

    if (typeof fieldProps.FormFieldValue !== 'string') {
      return null;
    }
    let val = this.getSplitValues(fieldProps.FormFieldValue);
    if (!val) {
      return null;
    }

    const urlPart = val[0];
    const descPart = val[1];
    const urlValidityErrorText = `${fieldProps.Title} URL is in invalid format`;
    const urlVaidity = new RegExp(/https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)/gi);
    const urlLengthErrorText = `${fieldProps.Title} URL part must not be more than 255 characters`;
    const descLengthErrorText = `${fieldProps.Title} description part must not be more than 255 characters`;
    let combinedError = '';

    if (!urlPart.match(urlVaidity)) {
      combinedError += urlValidityErrorText;
    }
    if (urlPart.length > 255) {
      if (combinedError.length > 0) {
        combinedError += '. ';
      }
      combinedError += urlLengthErrorText;
    }
    if (descPart.length > 255) {
      if (combinedError.length > 0) {
        combinedError += '. ';
      }
      combinedError += descLengthErrorText;
    }
    return combinedError.length > 0 ? combinedError : null;
  }
}
