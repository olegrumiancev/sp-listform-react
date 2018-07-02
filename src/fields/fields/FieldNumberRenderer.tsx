import * as React from 'react';
import { IFieldProps } from '../interfaces';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { BaseFieldRenderer } from './BaseFieldRenderer';
import './FieldNumberRenderer.css';
import { FormFieldsStore } from '../store';
import { ValidationManager } from '../managers/ValidationManager';

export class FieldNumberRenderer extends BaseFieldRenderer {
  public constructor(props: IFieldProps) {
    super(props);
    let val = props.FormFieldValue;
    if (val && props.NumberIsPercent) {
      val = val * 100;
    }
    this.state = {
      ...this.state,
      currentValue: val,
      decimalEncountered: false
    };

    if (props.Max) {
      // let val = props.NumberIsPercent ? props.Max * 100 : props.Max;
      FormFieldsStore.actions.addValidatorToField(ValidationManager.defaultValidators.maxValue, props.InternalName, props.Max);
    }
    if (props.Min) {
      // let val = props.NumberIsPercent ? props.Min * 100 : props.Min;
      FormFieldsStore.actions.addValidatorToField(ValidationManager.defaultValidators.minValue, props.InternalName, props.Min);
    }
  }

  protected renderNewForm() {
    return this.renderNewOrEditForm();
  }

  protected renderEditForm() {
    return this.renderNewOrEditForm();
  }

  protected renderDispForm() {
    const percent = this.props.NumberIsPercent ? '%' : '';
    return (<Label>{this.state.currentValue}{percent}</Label>);
  }

  private renderNewOrEditForm() {
    if (this.props.NumberIsPercent) {
      return (
        <React.Fragment>
          <TextField
            onChanged={this.onChanged}
            onKeyPress={this.onKeypress}
            value={this.state.currentValue == null ? '' : this.state.currentValue}
            prefix='%'
          />
        </React.Fragment>
      );
    } else {
      return (
        <React.Fragment>
          <TextField
            onChanged={this.onChanged}
            onKeyPress={this.onKeypress}
            value={this.state.currentValue == null ? '' : this.state.currentValue}
          />
        </React.Fragment>
      );
    }
  }

  private onChanged = (newValue) => {
    let containsDecimal = false;
    let toSave = newValue;
    if (toSave === '') {
      toSave = null;
    } else {
      if (toSave.indexOf('.') !== -1) {
        containsDecimal = true;
      }
      toSave = parseFloat(toSave);
    }
    if (toSave) {
      if (this.props.NumberIsPercent) {
        toSave = toSave / 100;
      }
    }
    this.setState({
      currentValue: newValue,
      decimalEncountered: containsDecimal
    });
    this.trySetChangedValue(toSave);
  }

  private onKeypress = (ev) => {
    if (ev.key.match(/[0-9]|\./) === null) {
      ev.preventDefault();
    }

    if (ev.key.match(/\./) && this.state.decimalEncountered) {
      ev.preventDefault();
    }
  }
}
