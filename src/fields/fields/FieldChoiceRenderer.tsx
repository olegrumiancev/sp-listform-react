import * as React from 'react';
import { IFieldProps } from '../interfaces';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { BaseFieldRenderer } from './BaseFieldRenderer';
import './FieldChoiceRenderer.css';
import { TextField } from 'office-ui-fabric-react/lib/TextField';

export class FieldChoiceRenderer extends BaseFieldRenderer {
  public constructor(props: IFieldProps) {
    super(props);
    let vals = [];
    if (this.props.FormFieldValue != null) {
      if (this.props.IsMulti) {
        vals = this.props.FormFieldValue.results;
      } else {
        vals.push(this.props.FormFieldValue);
      }
    }

    let ownOption = null;
    for (let i of vals) {
      if (!this.props.Choices.includes(i)) {
        ownOption = i;
        break;
      }
    }

    this.state = {
      ...this.state,
      currentValue: vals,
      ownOption: ownOption
    };
  }

  protected renderNewForm() {
    return this.renderNewOrEditForm();
  }

  protected renderEditForm() {
    return this.renderNewOrEditForm();
  }

  protected renderDispForm() {
    if (this.props.IsMulti && this.props.FormFieldValue != null) {
      return (
        this.props.FormFieldValue.results.map((fv, i) => <Label key={`${this.props.InternalName}_${i}`}>{fv}</Label>)
      );
    }
    return (<Label>{this.props.FormFieldValue}</Label>);
  }

  private renderNewOrEditForm() {
    let mainStyle = {
      width: this.props.FillInChoice ? '45%' : '100%',
      minWidth: this.props.FillInChoice ? '300px' : '90%',
      display: 'inline-block'
    };
    let middleAuxStyle = this.props.FillInChoice ? {
      width: '9%',
      display: 'inline-block'
    } : {
      display: 'none'
    };
    let auxStyle = this.props.FillInChoice ? {
      width: '45%',
      display: 'inline-block',
      verticalAlign: 'top'
    } : {
      display: 'none'
    };

    let currentVal = this.state.currentValue as any[];
    return (
      <div>
        <div style={mainStyle}>
          <Dropdown
            key={`dropdown_${this.props.InternalName}`}
            multiSelect={this.props.IsMulti}
            onChanged={(newValue) => {
              this.setState({ ownOption: null }, () => {
                this.saveFieldDataInternal(newValue);
              });
            }}
            options={this.props.Choices.map(c => {
              return {
                key: c,
                text: c,
                selected: currentVal && currentVal.includes(c) };
            })}
            placeHolder={this.props.IsMulti ? 'Select options' : 'Select an option'}
          />
        </div>
        <div style={middleAuxStyle}>&nbsp;</div>
        <div style={auxStyle}>
            <TextField placeholder='Enter own option' value={this.state.ownOption == null ? '' : this.state.ownOption}
            onChanged={(newValue) => {
              if (newValue !== this.state.ownOption) {
                let currentVal = this.state.currentValue as any[];
                let newCurrentVal = [];
                for (let i of currentVal) {
                  if (i !== this.state.ownOption && i !== newValue) {
                    newCurrentVal.push(i);
                  }
                }
                if (newValue !== '' && newValue !== null) {
                  newCurrentVal.push(newValue);
                }
                this.setState({
                  ownOption: newValue,
                  currentValue: this.props.IsMulti || (newValue === '' || newValue === null) ? newCurrentVal : [newValue]
                }, () => {
                  if (this.props.IsMulti) {
                    this.trySetChangedValue({ results: this.state.currentValue });
                  } else {
                    this.trySetChangedValue(this.state.currentValue.length > 0 ? this.state.currentValue[0] : undefined);
                  }
                });
              }
            }} />
        </div>
      </div>
    );
  }

  private saveFieldDataInternal(newValue: any) {
    if (this.props.IsMulti) {
      this.setState({ currentValue: this.constructNewState(newValue.key, newValue.selected) }, () => {
        this.trySetChangedValue({ results: this.state.currentValue });
      });
    } else {
      this.setState({ currentValue: [newValue.key] }, () => {
        this.trySetChangedValue(this.state.currentValue.length > 0 ? this.state.currentValue[0] : undefined);
      });
    }
  }

  private constructNewState(value: string, toAdd: boolean): string[] {
    let result: string[] = this.state.currentValue;
    if (toAdd) {
      if (!(this.state.currentValue as string[]).includes(value)) {
        result = [value, ...this.state.currentValue];
      }
    } else {
      result = [];
      for (let i of this.state.currentValue) {
        if (i !== value) {
          result.push(i);
        }
      }
    }
    return result;
  }
}
