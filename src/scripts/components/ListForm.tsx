import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { IListFormProps, FormMode } from '../interfaces';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Label } from 'office-ui-fabric-react/lib/Label';
import FormHeader from './FormHeader';
import { formHelper } from '../utils/formHelper';
import { ItemUpdateResult, ItemAddResult, AttachmentFileInfo } from '@pnp/sp';

// import { IFieldProps, IFormManagerProps } from '@olegrumiancev/sp-react-formfields/lib/interfaces';
// import { FormField } from '@olegrumiancev/sp-react-formfields/lib/fields/FormField';
// import { FormFieldsStore } from '@olegrumiancev/sp-react-formfields/lib/store';
// import { FieldChoiceRenderer } from '@olegrumiancev/sp-react-formfields/lib/fields';

import { IFieldProps, IFormManagerProps } from '../../fields/interfaces';
import { FormField } from '../../fields/fields/FormField';
import { FormFieldsStore } from '../../fields/store';

// import { IFieldProps, IFormManagerProps } from '@rumiancev/sp-react-formfields/lib/interfaces';
// import { FormField } from '@rumiancev/sp-react-formfields/lib/fields/FormField';
// import { FormFieldsStore } from '@rumiancev/sp-react-formfields/lib/store';
// import { FieldChoiceRenderer } from '@rumiancev/sp-react-formfields/lib/fields';

export default class ListForm extends React.Component<IListFormProps, IListFormProps> {
  private localContext: SP.ClientContext;

  public constructor(props) {
    super(props);
    this.localContext = SP.ClientContext.get_current();
    initializeIcons();

    this.state = {
      ...props
    };
  }

  public render() {
    let ListFormStateless = FormFieldsStore.connect(state => {
      let enhancedState = {
        ...state,
        closeForm: this.closeForm,
        getButtonsByFormMode: this.getButtonsByFormMode
      };
      return enhancedState;
    })(ListFormInternal);
    return (
      <div className='formContainer'>
        <FormFieldsStore.Provider>
          <ListFormStateless />
        </FormFieldsStore.Provider>
      </div>
    );
  }

  public componentDidMount() {
    FormFieldsStore.actions.initStore(
      this.state.SpWebUrl, this.state.CurrentListId,
      this.state.CurrentMode, this.state.CurrentItemId);
  }

  private closeForm() {
    //  console.log('Closing the form.');
  }

  private getButtonsByFormMode(mode: number) {
    let commandBarItemSave = {
      className: 'ms-bgColor-neutral',
      key: 'save',
      name: 'Save',
      iconProps: {
        iconName: 'Save'
      },
      onClick: () => {
        // debugger;
        const isValid = FormFieldsStore.actions.validateForm();
        if (isValid) {
          FormFieldsStore.actions.saveFormData().then(res => {
            if (res.IsSuccessful) {
              // this.setState({ CurrentItemId: res.ItemId });
              FormFieldsStore.actions.setFormMode(FormMode.Display);
            } else {
              // we need to do show save error dialog
              // FormFieldsStore.actions.initStore(
              //   this.state.SpWebUrl, this.state.CurrentListId,
              //   this.state.CurrentMode, this.state.CurrentItemId);
              FormFieldsStore.actions.setFormMessage(res.Error ? res.Error.toString() : 'Error has occurred while saving, reload the page and try again', () => {
                // window.location.href = window.location.href;
              });
            }
          });
        } else {
          FormFieldsStore.actions.setShowValidationErrors(true);
        }
      }
    };

    let commandBarItemEdit = {
      className: 'ms-bgColor-neutral',
      key: 'edit',
      name: 'Edit',
      iconProps: {
        iconName: 'Edit'
      },
      onClick: () => {
        FormFieldsStore.actions.setFormMode(FormMode.Edit);
      }
    };

    return mode === FormMode.Display ? [commandBarItemEdit] : [commandBarItemSave];
  }
}

export const ListFormInternal = (props) => {
  // console.log(props);
  return <div>
  <FormHeader CurrentMode={props.CurrentMode as number} Fields={props.Fields} />
  <CommandBar isSearchBoxVisible={false} key='commandBar'
    items={props.getButtonsByFormMode(props.CurrentMode)}
    farItems={[
      {
        className: 'ms-bgColor-neutral',
        key: 'close',
        name: 'Close',
        iconProps: {
          iconName: 'RemoveFilter'
        },
        onClick: props.closeForm()
      }
    ]}
  />
  {props.IsLoading ?
    <div className='formContainer' style={{ padding: '5em' }}><Spinner title='Loading...' /></div> :
    <React.Fragment>
    {props.Fields.map(f => (
      <div className='formRow' key={`formRow_${f.InternalName}`}>
        <div className='rowLabel' key={`formLabelContainer_${f.InternalName}`}>
          <Label key={`label_${f.InternalName}`}>
            {f.Title}
            {f.IsRequired ? <span key={`label_required_${f.InternalName}`} style={{ color: 'red' }}> *</span> : null}
          </Label>
        </div>
        <div className='rowField' key={`formFieldContainer_${f.InternalName}`}>
          <FormField key={`formfield_${f.InternalName}`} InternalName={f.InternalName} FormMode={f.CurrentMode} />
          {/* <FormField key={`formfield_${f.InternalName}`} {...f} /> */}
        </div>
      </div>
    ))}
    </React.Fragment>
  }
</div>;
};
