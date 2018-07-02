import { IFieldProps } from '../interfaces';
import { SPRest, sp } from '@pnp/sp';
import { FormMode } from '../../scripts/interfaces';

export const FieldPropsManager = {
  createFieldRendererPropsFromFieldMetadata: (fieldMetadata: any, formMode: number, spListItem: any, spRest: SPRest) => {
    if (fieldMetadata == null) {
      return null;
    }

    let fieldProps = {
      SchemaXml: new DOMParser().parseFromString(fieldMetadata.SchemaXml, 'text/xml'),
      CurrentMode: formMode,
      Title: fieldMetadata.Title,
      InternalName: fieldMetadata.InternalName,
      EntityPropertyName: fieldMetadata.EntityPropertyName,
      IsHidden: fieldMetadata.Hidden,
      IsRequired: fieldMetadata.Required,
      IsMulti: fieldMetadata.TypeAsString.match(/multi/gi),
      Type: fieldMetadata.TypeAsString,
      Description: fieldMetadata.Description,
      DefaultValue: fieldMetadata.DefaultValue,
      pnpSPRest: spRest == null ? sp : spRest,
      ValidationErrors: [],
      IsValid: true,
      Validators: []
    } as IFieldProps;

    if (spListItem != null && spListItem[fieldProps.InternalName] != null && spListItem[fieldProps.InternalName].__deferred == null) {
      fieldProps.FormFieldValue = spListItem[fieldProps.InternalName];
    }

    fieldProps = addFieldTypeSpecificProperties(fieldProps, fieldMetadata);
    console.log(fieldMetadata);
    return fieldProps;
  }
};

const addFieldTypeSpecificProperties = (fieldProps: IFieldProps, fieldMetadata: any): IFieldProps => {
  let result = fieldProps;
  switch (fieldProps.Type) {
    case 'Text':
      result = addTextFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'Note':
      result = addMultilineTextFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'Choice':
    case 'MultiChoice':
      result = addChoiceFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'Lookup':
    case 'LookupMulti':
      result = addLookupFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'User':
    case 'UserMulti':
      result = addUserFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'Number':
      result = addNumberFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'DateTime':
      result = addDateTimeFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'TaxonomyFieldType':
      result = addTaxonomyFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'Attachments':
      // console.log(fieldMetadata);
      break;
    default:
      break;
  }
  return result;
};

const addTaxonomyFieldProperties = (fieldProps: IFieldProps, fieldMetadata: any) => {
  fieldProps.TaxonomyAnchorId = fieldMetadata.AnchorId;
  fieldProps.TaxonomyIsOpen = fieldMetadata.Open;
  fieldProps.TaxonomyTermSetId = fieldMetadata.TermSetId;
  fieldProps.IsMulti = fieldMetadata.AllowMultipleValues;
  return fieldProps;
};

const addUserFieldProperties = (fieldProps: IFieldProps, fieldMetadata: any): IFieldProps => {
  if (fieldMetadata.SchemaXml.match(/UserSelectionMode="PeopleAndGroups"/gi)) {
    fieldProps.UserSelectionMode = 'PeopleAndGroups';
  } else {
    fieldProps.UserSelectionMode = 'PeopleOnly';
  }
  return fieldProps;
};

const addTextFieldProperties = (fieldProps: IFieldProps, fieldMetadata: any): IFieldProps => {
  if (fieldProps.CurrentMode === FormMode.New && fieldProps.DefaultValue) {
    fieldProps.FormFieldValue = fieldProps.DefaultValue;
  }
  if (fieldMetadata.MaxLength) {
    fieldProps.Max = fieldMetadata.MaxLength;
  }
  return fieldProps;
};

const addMultilineTextFieldProperties = (fieldProps: IFieldProps, fieldMetadata: any): IFieldProps => {
  fieldProps.IsRichText = fieldMetadata.RichText;
  return fieldProps;
};

const addDateTimeFieldProperties = (fieldProps: IFieldProps, fieldMetadata: any): IFieldProps => {
  fieldProps.DateTimeIsTimePresent = fieldMetadata.DisplayFormat === 1;
  if (fieldProps.CurrentMode === FormMode.New && fieldProps.DefaultValue) {
    if (fieldProps.DefaultValue.toLowerCase() === '[today]') {
      let date = new Date(Date.now());
      date.setHours(0);
      date.setMinutes(0);
      date.setHours(0);
      date.setMilliseconds(0);
      fieldProps.FormFieldValue = date.toISOString();
    } else if (fieldProps.DefaultValue.toLowerCase() === '[now]') {
      fieldProps.FormFieldValue = new Date(Date.now()).toISOString();
    }
  }
  return fieldProps;
};

const addNumberFieldProperties = (fieldProps: IFieldProps, fieldMetadata: any): IFieldProps => {
  fieldProps.NumberIsPercent = fieldMetadata.SchemaXml.match(/percentage="true"/gi) !== null;
  if (fieldProps.CurrentMode === FormMode.New && fieldProps.DefaultValue) {
    fieldProps.FormFieldValue = fieldProps.DefaultValue;
  }
  if (fieldMetadata.MaximumValue) {
    fieldProps.Max = fieldMetadata.MaximumValue;
  }
  if (fieldMetadata.MinimumValue) {
    fieldProps.Min = fieldMetadata.MinimumValue;
  }
  return fieldProps;
};

const addChoiceFieldProperties = (fieldProps: IFieldProps, fieldMetadata: any): IFieldProps => {
  fieldProps.Choices = fieldMetadata.Choices == null ? undefined : fieldMetadata.Choices.results;
  if (fieldProps.CurrentMode === FormMode.New && fieldProps.DefaultValue) {
    if (fieldProps.IsMulti) {
      fieldProps.FormFieldValue = { results: [fieldProps.DefaultValue] };
    } else {
      fieldProps.FormFieldValue = fieldProps.DefaultValue;
    }
  }
  fieldProps.FillInChoice = fieldMetadata.FillInChoice;
  return fieldProps;
};

const addLookupFieldProperties = (fieldProps: IFieldProps, fieldMetadata: any): IFieldProps => {
  fieldProps.LookupListId = fieldMetadata.LookupList == null ? undefined : fieldMetadata.LookupList,
  fieldProps.LookupWebId = fieldMetadata.LookupWebId == null ? undefined : fieldMetadata.LookupWebId,
  fieldProps.LookupField = fieldMetadata.LookupField == null || fieldMetadata.LookupField === '' ? 'Title' : fieldMetadata.LookupField;
  return fieldProps;
};
