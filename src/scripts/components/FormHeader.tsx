import * as React from 'react';
import { FormMode, IListFormProps } from '../interfaces';
import { Label } from 'office-ui-fabric-react/lib/Label';

// import { IFieldProps } from '@rumiancev/sp-react-formfields/lib/interfaces';

// import { IFieldProps } from '@olegrumiancev/sp-react-formfields/lib/interfaces';
import { IFieldProps } from '../../fields/interfaces';

export default class FormHeader extends React.Component<{Fields: IFieldProps[], CurrentMode: number}, {}> {
  public constructor(props) {
    super(props);

  }

  public render() {
    if (!this.props.Fields) {
      return null;
    }

    let headerText;
    let titleFieldInfo = this.props.Fields.filter(f => f.InternalName === 'Title');

    if (this.props.CurrentMode === FormMode.New) {
      headerText = 'New item';
    } else {
      headerText = `form for ${titleFieldInfo == null || titleFieldInfo.length < 1 || titleFieldInfo[0].FormFieldValue == null ? '(no title)' : titleFieldInfo[0].FormFieldValue}`;
      headerText = `${this.props.CurrentMode === FormMode.Edit ? 'Edit' : 'Display'} ${headerText}`;
    }
    return (<Label className='formHeader'>{headerText}</Label>);
  }
}
