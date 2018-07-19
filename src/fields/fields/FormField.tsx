import * as React from 'react';
import { IFormFieldProps, IFieldProps, IFormManagerProps } from '../interfaces';
import { FormFieldsStore } from '../store';
import { BaseFieldRenderer, FieldTextRenderer, FieldChoiceRenderer, FieldLookupRenderer, FieldUserRenderer } from './index';
import { FieldMultilineTextRenderer } from './FieldMultilineTextRenderer';
import { FieldNumberRenderer } from './FieldNumberRenderer';
import { FieldDateTimeRenderer } from './FieldDateTimeRenderer';
import { FieldAttachmentRenderer } from './FieldAttachmentRenderer';
import { FieldTaxonomyRenderer } from './FieldTaxonomyRenderer';
import { getFieldPropsByInternalName } from '../utils';
import { FieldBooleanRenderer } from './FieldBooleanRenderer';
import { FieldCurrencyRenderer } from './FieldCurrencyRenderer';
import { FieldUrlRenderer } from './FieldUrlRenderer';

export class FormField extends React.Component<IFormFieldProps, any> {
  constructor(props: IFormFieldProps) {
    super(props);
  }

  public render() {
    let ConnectedFormField = FormFieldsStore.connect((state: IFormManagerProps) => getFieldPropsByInternalName(state.Fields, this.props.InternalName))(SpecificFormField);
    return <ConnectedFormField />;
  }
}

const SpecificFormField = (fieldProps: IFieldProps) => {
  let defaultElement = (<BaseFieldRenderer {...fieldProps} key={fieldProps.InternalName} />);
  let onFieldDataChangeCallback = FormFieldsStore.actions.setFieldData;
  if (fieldProps.Type === 'Text') {
    return <FieldTextRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  if (fieldProps.Type === 'Note') {
    return <FieldMultilineTextRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  if (fieldProps.Type === 'Boolean') {
    return <FieldBooleanRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  if (fieldProps.Type === 'Number') {
    return <FieldNumberRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  if (fieldProps.Type === 'Currency') {
    return <FieldCurrencyRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  if (fieldProps.Type === 'URL') {
    return <FieldUrlRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  if (fieldProps.Type === 'DateTime') {
    return <FieldDateTimeRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  if (fieldProps.Type.match(/user/gi)) {
    return <FieldUserRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  if (fieldProps.Type.match(/choice/gi)) {
    return <FieldChoiceRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  if (fieldProps.Type.match(/lookup/gi)) {
    return <FieldLookupRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  if (fieldProps.Type.match(/TaxonomyFieldType/gi)) {
    return <FieldTaxonomyRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  if (fieldProps.Type.match(/attachments/gi)) {
    return <FieldAttachmentRenderer
      {...fieldProps}
      key={fieldProps.InternalName}
      saveChangedFieldData={onFieldDataChangeCallback}
    />;
  }
  return defaultElement;
};
