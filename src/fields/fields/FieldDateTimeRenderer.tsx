import 'rc-time-picker/assets/index.css';
import * as React from 'react';
import { IFieldProps, FormMode } from '../interfaces';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { BaseFieldRenderer } from './BaseFieldRenderer';
import * as moment from 'moment';
import TimePicker from 'rc-time-picker';
import './FieldDateTimeRenderer.css';

export class FieldDateTimeRenderer extends BaseFieldRenderer {
  private timeFormat = 'HH:mm';
  public constructor(props: IFieldProps) {
    super(props);

    let stateObj = this.getStateObjectFromISO(props.FormFieldValue);
    if (stateObj) {
      this.state = {
        ...this.state,
        ...stateObj
      };
    }
  }

  protected renderNewForm() {
    return this.renderNewOrEditForm();
  }

  protected renderEditForm() {
    return this.renderNewOrEditForm();
  }

  protected renderDispForm() {
    if (this.props.FormFieldValue) {
      const d = new Date(Date.parse(this.props.FormFieldValue));
      let result = d.toLocaleDateString();
      if (this.props.DateTimeIsTimePresent) {
        result += ` ${d.toLocaleTimeString()}`;
      }
      return (<Label>{result}</Label>);
    }
    return null;
  }

  private renderNewOrEditForm() {
    let datePickerStyle =
      this.props.DateTimeIsTimePresent ?
      { display: 'inline-block' } :
      { width: '100%', display: 'block' };
    return (
      <React.Fragment>
        <div style={datePickerStyle}>
          <DatePicker
            onSelectDate={this.onDateChange}
            value={this.state.currentDateValue == null ? null : this.state.currentDateValue}
          />
        </div>
        {!this.props.DateTimeIsTimePresent ? null :
          (<div style={{ width: '100px', display: 'inline-block' }}><TimePicker
            style={{ width: '50px', margin: '10px', display: 'inline-block' }}
            showSecond={false}
            defaultValue={this.state.currentTimeValue ? this.state.currentTimeValue : moment()}
            onChange={this.onTimeChange}
          /></div>)}
      </React.Fragment>
    );
  }

  private getStateObjectFromISO = (isoDate: string): Object => {
    if (isoDate) {
      let fullDateTime = new Date(Date.parse(isoDate));
      let datePart = null;
      let timePart = null;
      timePart = moment(isoDate);
      fullDateTime.setHours(0);
      fullDateTime.setMinutes(0);
      datePart = fullDateTime;
      return {
        currentDateValue: datePart,
        currentTimeValue: timePart
      };
    }
    return null;
  }

  private onDateChange = (newValue: Date) => {
    this.setState({ currentDateValue: newValue }, () => {
      this.trySetChangedValue(this.getCompositeDateForSaving());
    });
  }

  private onTimeChange = (newValue: Date) => {
    let val = newValue ? moment(newValue.toISOString()) : null;
    this.setState({ currentTimeValue: val }, () => {
      this.trySetChangedValue(this.getCompositeDateForSaving());
    });
  }

  private getCompositeDateForSaving = (): string => {
    let baseDate: Date = this.state.currentDateValue;
    if (!baseDate) {
      baseDate = new Date(Date.now());
    }
    if (this.props.DateTimeIsTimePresent && this.state.currentTimeValue) {
      let m: moment.Moment = this.state.currentTimeValue;
      baseDate.setHours(m.hours());
      baseDate.setMinutes(m.minutes());
    } else {
      baseDate.setHours(0);
      baseDate.setMinutes(0);
    }
    baseDate.setSeconds(0);
    baseDate.setMilliseconds(0);
    return baseDate.toISOString();
  }
}
