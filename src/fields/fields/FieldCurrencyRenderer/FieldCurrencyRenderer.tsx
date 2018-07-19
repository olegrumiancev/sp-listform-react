import * as React from 'react';
import { IFieldProps } from '../../interfaces';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { BaseFieldRenderer } from '../BaseFieldRenderer';
import './FieldCurrencyRenderer.css';
import { FormFieldsStore } from '../../store';
import { ValidationManager } from '../../managers/ValidationManager';
import { getCurrency } from './localeCurrency';

export class FieldCurrencyRenderer extends BaseFieldRenderer {
  public constructor(props: IFieldProps) {
    super(props);
    let val = props.FormFieldValue;

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
    let val = this.state.currentValue ? this.state.currentValue : 0;
    let currencySymbol = '';
    let currencySymbolOnLeft = false;
    let spaceBetweenSymbol = true;
    let currencyObj = getCurrency(this.props.CurrencyLocaleId);
    if (currencyObj) {
      currencySymbol = currencyObj.symbol;
      currencySymbolOnLeft = currencyObj.symbolOnLeft;
      spaceBetweenSymbol = currencyObj.spaceBetweenAmountAndSymbol;
    }
    if (currencySymbolOnLeft) {
      val = `${currencySymbol}${spaceBetweenSymbol ? ' ' : ''}${val}`.trim();
    } else {
      val = `${val}${spaceBetweenSymbol ? ' ' : ''}${currencySymbol}`.trim();
    }
    return (<Label>{val}</Label>);
  }

  private renderNewOrEditForm() {
    let currencyObj = getCurrency(this.props.CurrencyLocaleId);
    return (
      <React.Fragment>
        <TextField
          onChanged={this.onChanged}
          onKeyPress={this.onKeypress}
          value={this.state.currentValue == null ? '' : this.state.currentValue}
          prefix={currencyObj ? currencyObj.symbol : undefined}
        />
      </React.Fragment>
    );
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
