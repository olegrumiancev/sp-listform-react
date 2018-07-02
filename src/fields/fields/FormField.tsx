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

export class FormField extends React.Component<IFormFieldProps, any> {
  constructor(props: IFormFieldProps) {
    super(props);
  }

  public render() {
    return (
      <FormFieldsStore.Consumer mapStateToProps={(state) => state}>
      {(consumerState) => {
        let fieldProps = getFieldPropsByInternalName(consumerState.Fields, this.props.InternalName);
        if (fieldProps) {
          return (this.createFieldRenderer(fieldProps, FormFieldsStore.actions.setFieldData));
        } else {
          return null;
        }
      }}
      </FormFieldsStore.Consumer>
    );
  }

  private createFieldRenderer(fieldProps: IFieldProps, onFieldDataChangeCallback: (internalName: string, newValue: any) => void): JSX.Element {
    let defaultElement = (<BaseFieldRenderer {...fieldProps} key={fieldProps.InternalName} />);
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
    if (fieldProps.Type === 'Number') {
      return <FieldNumberRenderer
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
  }
}
