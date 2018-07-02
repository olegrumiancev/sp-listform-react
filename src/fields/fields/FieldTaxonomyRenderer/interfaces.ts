import { ITermSet, ITerm } from './ISPTermStorePickerService';
import { IFieldProps } from '../../interfaces';

export interface ITaxonomyPickerProps {
  label: string;
  panelTitle: string;
  allowMultipleSelections?: boolean;
  initialValues?: IPickerTerms;
  termsetNameOrID: string;
  anchorId?: string;
  isTermSetSelectable?: boolean;
  disabledTermIds?: string[];
  disableChildrenOfDisabledParents?: boolean;
  disabled?: boolean;
  onGetErrorMessage?: (value: IPickerTerms) => string | Promise<string>;

  /**
   * onChange Event
   */
  onChange?: (newValue?: IPickerTerms) => void;
}

/**
 * PropertyFieldTermPickerHost state interface
 */
export interface ITaxonomyPickerState {

  termSetAndTerms?: ITermSet;
  errorMessage?: string;
  openPanel?: boolean;
  loaded?: boolean;
  activeNodes?: IPickerTerms;
}

export interface ITermChanges {
  changedCallback: (term: ITerm, checked: boolean) => void;
  activeNodes?: IPickerTerms;
  disabledTermIds?: string[];
  disableChildrenOfDisabledParents?: boolean;
}

export interface ITermParentProps extends ITermChanges {
  termset: ITermSet;
  multiSelection: boolean;
  anchorId?: string;
  isTermSetSelectable?: boolean;

  autoExpand: () => void;
  termSetSelectedChange?: (termSet: ITermSet, isChecked: boolean) => void;
}

export interface ITermParentState {

  loaded?: boolean;
  expanded?: boolean;
}

export interface ITermProps extends ITermChanges {
  termset: string;
  term: ITerm;
  multiSelection: boolean;
  disabled: boolean;
}

export interface ITermState {
  selected?: boolean;
}

export interface ITermPickerState {
  terms: IPickerTerms;
}

export interface ITermPickerProps {
  // termPickerHostProps: ITaxonomyPickerProps;
  fieldProps: IFieldProps;
  disabled: boolean;
  value: IPickerTerms;
  allowMultipleSelections: boolean;
  isTermSetSelectable?: boolean;
  disabledTermIds?: string[];
  disableChildrenOfDisabledParents?: boolean;

  onChanged: (items: IPickerTerm[]) => void;
}

export interface IPickerTerm {
  name: string;
  key: string;
  path: string;
  termSet: string;
  termSetName?: string;
}

export interface IPickerTerms extends Array<IPickerTerm> { }

export interface IPropertyFieldTermPickerProps {
  label: string;
  panelTitle: string;
  allowMultipleSelections?: boolean;
  initialValues?: IPickerTerms;
  excludeSystemGroup?: boolean;
  limitByGroupNameOrID?: string;
  limitByTermsetNameOrID?: string;

  properties: any;
  key: string;
  disabled?: boolean;
  /**
   * The method is used to get the validation error message and determine whether the input value is valid or not.
   *
   *   When it returns string:
   *   - If valid, it returns empty string.
   *   - If invalid, it returns the error message string and the text field will
   *     show a red border and show an error message below the text field.
   *
   *   When it returns Promise<string>:
   *   - The resolved value is display as error message.
   *   - The rejected, the value is thrown away.
   *
   */
  onGetErrorMessage?: (value: IPickerTerms) => string | Promise<string>;
  /**
   * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
   * Default value is 200.
   */
  deferredValidationTime?: number;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
}
