import { IFieldProps, IValidationManager, IValidateFieldResult } from '../interfaces';
import { FormFieldsStore } from '../store';
import { getFieldPropsByInternalName } from '../utils';

const required = (internalName: string): string => {
  const globalState = FormFieldsStore.actions.getState();
  let fieldProps = getFieldPropsByInternalName(globalState.Fields, internalName);
  if (!fieldProps) {
    return `Could not find a field by internal name '${internalName}'`;
  }

  let valueIsArray = Array.isArray(fieldProps.FormFieldValue);
  if ((valueIsArray && fieldProps.FormFieldValue.length > 0) || (!valueIsArray && fieldProps.FormFieldValue)) {
    return null;
  } else {
    return `${fieldProps.Title} field is required`;
  }
};

const mustMatch = (internalName1: string, internalName2: string): string => {
  const globalState = FormFieldsStore.actions.getState();
  let fieldProps1 = getFieldPropsByInternalName(globalState.Fields, internalName1);
  let fieldProps2 = getFieldPropsByInternalName(globalState.Fields, internalName2);
  if (!fieldProps1) {
    return `Could not find a field by internal name '${internalName1}'`;
  }
  if (!fieldProps2) {
    return `Could not find a field by internal name '${internalName2}'`;
  }
  const errorText = `${fieldProps1.Title} does not match ${fieldProps2.Title}`;
  let valueIsArray = Array.isArray(fieldProps1.FormFieldValue);
  let isEqual = false;
  if (valueIsArray) {
    if (fieldProps1.FormFieldValue.length === fieldProps2.FormFieldValue.length) {
      isEqual = true;
      for (let i = 0; i < fieldProps1.FormFieldValue.length; i++) {
        if (fieldProps1.FormFieldValue[i] !== fieldProps2.FormFieldValue[i]) {
          isEqual = false;
          break;
        }
      }
    }
  } else {
    isEqual = fieldProps1.FormFieldValue === fieldProps2.FormFieldValue;
  }
  return isEqual ? null : errorText;
};

const minLength = (internalName: string, length: number): string => {
  const globalState = FormFieldsStore.actions.getState();
  let fieldProps = getFieldPropsByInternalName(globalState.Fields, internalName);
  if (!fieldProps) {
    return `Could not find a field by internal name '${internalName}'`;
  }

  let textLength = fieldProps.FormFieldValue ? fieldProps.FormFieldValue.length : 0;
  const errorText = `${fieldProps.Title} must be at least ${length} character(-s)`;
  return textLength >= length ? null : errorText;
};

const maxLength = (internalName: string, length: number): string => {
  const globalState = FormFieldsStore.actions.getState();
  let fieldProps = getFieldPropsByInternalName(globalState.Fields, internalName);
  if (!fieldProps) {
    return `Could not find a field by internal name '${internalName}'`;
  }

  let textLength = fieldProps.FormFieldValue ? fieldProps.FormFieldValue.length : 0;
  const errorText = `${fieldProps.Title} must not be more than ${length} character(-s)`;
  return textLength <= length ? null : errorText;
};

const maxValue = (internalName: string, maxValue: number): string => {
  const globalState = FormFieldsStore.actions.getState();
  let fieldProps = getFieldPropsByInternalName(globalState.Fields, internalName);
  if (!fieldProps) {
    return `Could not find a field by internal name '${internalName}'`;
  }

  let val = fieldProps.FormFieldValue;
  const errorText = `${fieldProps.Title} must not be more than ${fieldProps.NumberIsPercent ? maxValue * 100 : maxValue}`;
  if (!val) {
    return null;
  }
  return val <= maxValue ? null : errorText;
};

const minValue = (internalName: string, minValue: number): string => {
  const globalState = FormFieldsStore.actions.getState();
  let fieldProps = getFieldPropsByInternalName(globalState.Fields, internalName);
  if (!fieldProps) {
    return `Could not find a field by internal name '${internalName}'`;
  }

  let val = fieldProps.FormFieldValue;
  const errorText = `${fieldProps.Title} must be at least ${fieldProps.NumberIsPercent ? minValue * 100 : minValue}`;
  if (!val) {
    return null;
  }
  return val >= minValue ? null : errorText;
};

const validateField = (fieldProps: IFieldProps): IValidateFieldResult => {
  let result = {} as IValidateFieldResult;
  result.IsValid = true;
  result.ValidationErrors = null;

  if (fieldProps.Validators && fieldProps.Validators.length > 0) {
    let errors: string[] = fieldProps.Validators.reduce((prevErrors: string[], currentValidator) => {
      const localError = currentValidator();
      if (localError) {
        if (prevErrors == null) {
          prevErrors = [localError];
        } else {
          prevErrors.push(localError);
        }
      }
      return prevErrors;
    }, null);

    result.IsValid = errors == null;
    result.ValidationErrors = errors;
  }
  return result;
};

export const ValidationManager: IValidationManager = {
  defaultValidators: {
    required,
    mustMatch,
    minLength,
    maxLength,
    minValue,
    maxValue
  },
  validateField
};
