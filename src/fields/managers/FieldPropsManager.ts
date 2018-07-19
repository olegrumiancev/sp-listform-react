import { IFieldProps } from '../interfaces';
import { SPRest, sp } from '@pnp/sp';
import { FormMode } from '../../scripts/interfaces';

export const FieldPropsManager = {
  createFieldRendererPropsFromFieldMetadata: async (fieldMetadata: any, formMode: number, currentListId: string, spListItem: any, spRest: SPRest) => {
    if (fieldMetadata == null) {
      return null;
    }

    let fieldProps = {
      SchemaXml: new DOMParser().parseFromString(fieldMetadata.SchemaXml, 'text/xml'),
      CurrentListId: currentListId,
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
      Validators: [],
      ShowValidationErrors: false
    } as IFieldProps;

    if (spListItem != null && spListItem[fieldProps.InternalName] != null && spListItem[fieldProps.InternalName].__deferred == null) {
      fieldProps.FormFieldValue = spListItem[fieldProps.InternalName];
    }

    fieldProps = await addFieldTypeSpecificProperties(fieldProps, fieldMetadata);
    console.log(fieldMetadata);
    return fieldProps;
  }
};

const addFieldTypeSpecificProperties = async (fieldProps: IFieldProps, fieldMetadata: any): Promise<IFieldProps> => {
  let result = fieldProps;
  switch (fieldProps.Type) {
    case 'Text':
      result = await addTextFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'Note':
      result = await addMultilineTextFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'Choice':
    case 'MultiChoice':
      result = await addChoiceFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'Lookup':
    case 'LookupMulti':
      result = await addLookupFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'User':
    case 'UserMulti':
      result = await addUserFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'Number':
      result = await addNumberFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'DateTime':
      result = await addDateTimeFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'TaxonomyFieldType':
    case 'TaxonomyFieldTypeMulti':
      result = await addTaxonomyFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'Currency':
      result = await addCurrencyFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'URL':
      result = await addUrlFieldProperties(fieldProps, fieldMetadata);
      break;
    case 'Attachments':
      // console.log(fieldMetadata);
      break;
    default:
      break;
  }
  return result;
};

const addUrlFieldProperties = async (fieldProps: IFieldProps, fieldMetadata: any): Promise<IFieldProps> => {
  fieldProps.UrlRenderAsPicture = fieldMetadata.DisplayFormat === 1;
  return fieldProps;
};

const addCurrencyFieldProperties = async (fieldProps: IFieldProps, fieldMetadata: any): Promise<IFieldProps> => {
  fieldProps.CurrencyLocaleId = fieldMetadata.CurrencyLocaleId;
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

const addTaxonomyFieldProperties = async (fieldProps: IFieldProps, fieldMetadata: any): Promise<IFieldProps> => {
  fieldProps.TaxonomyAnchorId = fieldMetadata.AnchorId;
  fieldProps.TaxonomyIsOpen = fieldMetadata.Open;
  fieldProps.TaxonomyTermSetId = fieldMetadata.TermSetId;
  // fieldProps.TaxonomyPaneState = 0;
  fieldProps.IsMulti = fieldMetadata.AllowMultipleValues;
  if (fieldMetadata.TextField) {
    // const currentListId = fieldProps.SchemaXml.documentElement.getAttribute('List').replace('{', '').replace('}', '');
    let relatedNoteField = await fieldProps.pnpSPRest.web.lists.getById(fieldProps.CurrentListId).fields.getById(fieldMetadata.TextField).usingCaching().get();
    fieldProps.TaxonomyUpdateFieldEntityPropertyName = relatedNoteField.EntityPropertyName;
    // fieldProps.EntityPropertyName = relatedNoteField.EntityPropertyName;
  }
  return fieldProps;
};

const addUserFieldProperties = async (fieldProps: IFieldProps, fieldMetadata: any): Promise<IFieldProps> => {
  if (fieldMetadata.SchemaXml.match(/UserSelectionMode="PeopleAndGroups"/gi)) {
    fieldProps.UserSelectionMode = 'PeopleAndGroups';
  } else {
    fieldProps.UserSelectionMode = 'PeopleOnly';
  }
  return fieldProps;
};

const addTextFieldProperties = async (fieldProps: IFieldProps, fieldMetadata: any): Promise<IFieldProps> => {
  if (fieldProps.CurrentMode === FormMode.New && fieldProps.DefaultValue) {
    fieldProps.FormFieldValue = fieldProps.DefaultValue;
  }
  if (fieldMetadata.MaxLength) {
    fieldProps.Max = fieldMetadata.MaxLength;
  }
  return fieldProps;
};

const addMultilineTextFieldProperties = async (fieldProps: IFieldProps, fieldMetadata: any): Promise<IFieldProps> => {
  fieldProps.IsRichText = fieldMetadata.RichText;
  return fieldProps;
};

const addDateTimeFieldProperties = async (fieldProps: IFieldProps, fieldMetadata: any): Promise<IFieldProps> => {
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

const addNumberFieldProperties = async (fieldProps: IFieldProps, fieldMetadata: any): Promise<IFieldProps> => {
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

const addChoiceFieldProperties = async (fieldProps: IFieldProps, fieldMetadata: any): Promise<IFieldProps> => {
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

const addLookupFieldProperties = async (fieldProps: IFieldProps, fieldMetadata: any): Promise<IFieldProps> => {
  fieldProps.LookupListId = fieldMetadata.LookupList == null ? undefined : fieldMetadata.LookupList,
  fieldProps.LookupWebId = fieldMetadata.LookupWebId == null ? undefined : fieldMetadata.LookupWebId,
  fieldProps.LookupField = fieldMetadata.LookupField == null || fieldMetadata.LookupField === '' ? 'Title' : fieldMetadata.LookupField;
  return fieldProps;
};
