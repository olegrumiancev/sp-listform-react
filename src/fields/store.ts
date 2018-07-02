import { sp, SPRest, List, AttachmentFileInfo, ItemUpdateResult, ItemAddResult } from '@pnp/sp';
import { initStore } from 'react-waterfall';
import { IFormManagerProps, FormMode, IFieldProps, ISaveItemResult, IFormManagerActions } from './interfaces';
import { handleError, getFieldPropsByInternalName } from './utils';
import { FieldPropsManager } from './managers/FieldPropsManager';
import * as React from 'react';
import { ValidationManager } from './managers/ValidationManager';
import { enhanceProvider } from './EnhancedProvider';

const store = {
  initialState: {
    SPWebUrl: null,
    CurrentMode: 0,
    CurrentListId: null,
    IsLoading: true
  } as IFormManagerProps,
  actions: {
    initStore: async (state: IFormManagerProps, sPWebUrl: string, currentListId: string, currentMode: number, currentItemId?: number): Promise<IFormManagerProps> => {
      configurePnp(sPWebUrl);

      let list = sp.web.lists.getById(currentListId);
      let listFields: any[] =
        await list
          .fields
          .filter('ReadOnlyField eq false and Hidden eq false and Title ne \'Content Type\'').get();

      let toSelect = [];
      let toExpand = [];
      for (let f of listFields) {
        if (f.TypeAsString.match(/user/gi)) {
          toSelect.push(`${f.EntityPropertyName}/Title`);
          toSelect.push(`${f.EntityPropertyName}/Id`);
          toExpand.push(f.EntityPropertyName);
        } else if (f.TypeAsString.match(/lookup/gi)) {
          toSelect.push(`${f.EntityPropertyName}/Title`);
          toSelect.push(`${f.EntityPropertyName}/Id`);
          if (f.LookupField) {
            toSelect.push(`${f.EntityPropertyName}/${f.LookupField}`);
          }
          toExpand.push(f.EntityPropertyName);
        } else {
          toSelect.push(f.EntityPropertyName);
        }
      }

      let fieldInfos = [];
      let eTag = '*';
      if (currentMode !== FormMode.New) {
        let itemMetadata = list.items.getById(currentItemId);
        let item = await itemMetadata.select(...toSelect).expand(...toExpand).get();
        eTag = item.__metadata.etag;
        let attachmentMetadata = await itemMetadata.attachmentFiles.get();
        // console.log(item);
        fieldInfos = listFields.map(fm => {
          return FieldPropsManager.createFieldRendererPropsFromFieldMetadata(fm, currentMode, item, sp);
        });
        if (item.Attachments) {
          // console.log(attachmentMetadata);
          fieldInfos.filter(f => f.InternalName === 'Attachments')[0].FormFieldValue = attachmentMetadata;
        }
      } else {
        fieldInfos = listFields.map(fm => {
          return FieldPropsManager.createFieldRendererPropsFromFieldMetadata(fm, currentMode, null, sp);
        });
      }

      return {
        PnPSPRest: sp,
        SPWebUrl: sPWebUrl,
        CurrentListId: currentListId,
        CurrentItemId: currentItemId,
        CurrentMode: currentMode,
        Fields: fieldInfos,
        IsLoading: false,
        ShowValidationErrors: false,
        ETag: eTag
      } as IFormManagerProps;
    },
    setFormMode: (state: IFormManagerProps, mode: number) => {
      state.CurrentMode = mode;
      state.Fields.forEach(f => f.CurrentMode = mode);
      return state;
    },
    setItemId: (state: IFormManagerProps, itemId: number) => {
      state.CurrentItemId = itemId;
      return state;
    },
    setLoading: (state: IFormManagerProps, isLoading: boolean) => {
      state.IsLoading = isLoading;
      return state;
    },
    setShowValidationErrors: (state: IFormManagerProps, show: boolean) => {
      state.ShowValidationErrors = show;
      return state;
    },
    setFieldData: (state: IFormManagerProps, internalName: string, newValue: any) => {
      let fieldProps = getFieldPropsByInternalName(state.Fields, internalName);
      if (fieldProps) {
        fieldProps.FormFieldValue = newValue;
      }
      return state;
    },
    setFieldValidationState: (state: IFormManagerProps, internalName: string, isValid: boolean, validationErrors: string[]) => {
      let fieldProps = getFieldPropsByInternalName(state.Fields, internalName);
      if (fieldProps) {
        fieldProps.IsValid = isValid;
        fieldProps.ValidationErrors = validationErrors;
      }
      return state;
    },
    addNewAttachmentInfo: (state: IFormManagerProps, fileInfo: any) => {
      let attachmentProps = getFieldPropsByInternalName(state.Fields, 'Attachments');
      if (attachmentProps) {
        if (!attachmentProps.AttachmentsNewToAdd) {
          attachmentProps.AttachmentsNewToAdd = [];
        }
        attachmentProps.AttachmentsNewToAdd.push(fileInfo);
      }
      return state;
    },
    removeNewAttachmentInfo: (state: IFormManagerProps, fileInfo: any) => {
      let attachmentProps = getFieldPropsByInternalName(state.Fields, 'Attachments');
      if (attachmentProps && attachmentProps.AttachmentsNewToAdd) {
        attachmentProps.AttachmentsNewToAdd = attachmentProps.AttachmentsNewToAdd.filter(a => a.name !== fileInfo.name);
      }
      return state;
    },
    addOrRemoveExistingAttachmentDeletion: (state: IFormManagerProps, attachmentName: string) => {
      let attachmentProps = getFieldPropsByInternalName(state.Fields, 'Attachments');
      if (!attachmentProps.AttachmentsExistingToDelete) {
        attachmentProps.AttachmentsExistingToDelete = [];
      }

      if (attachmentProps.AttachmentsExistingToDelete.indexOf(attachmentName) !== -1) {
        attachmentProps.AttachmentsExistingToDelete = attachmentProps.AttachmentsExistingToDelete.filter(a => a !== attachmentName);
      } else {
        attachmentProps.AttachmentsExistingToDelete.push(attachmentName);
      }

      // console.log(state);
      return state;
    },
    clearHelperAttachmentProperties: (state: IFormManagerProps) => {
      let attachmentProps = getFieldPropsByInternalName(state.Fields, 'Attachments');
      if (attachmentProps) {
        attachmentProps.AttachmentsExistingToDelete = null;
        attachmentProps.AttachmentsNewToAdd = null;
      }
      return state;
    },
    setFieldPropValue: (state: IFormManagerProps, internalName: string, propName: string, propValue: any) => {
      let fieldProps = getFieldPropsByInternalName(state.Fields, internalName);
      if (fieldProps) {
        fieldProps[propName] = propValue;
      }
      return state;
    },
    addValidatorToField: (state: IFormManagerProps, validator: Function, internalName: string, ...validatorParams: any[]) => {
      let fieldProps = getFieldPropsByInternalName(state.Fields, internalName);
      if (fieldProps) {
        if (!fieldProps.Validators) {
          fieldProps.Validators = [];
        }
        // fieldProps.Validators.push(validator);
        fieldProps.Validators.push((): string => {
          return validator(internalName, ...validatorParams);
        });
      }
      return state;
    },
    clearValidatorsFromField: (state: IFormManagerProps, internalName: string) => {
      let fieldProps = getFieldPropsByInternalName(state.Fields, internalName);
      if (fieldProps) {
        fieldProps.Validators = [];
      }
      return state;
    },
    validateForm: (state: IFormManagerProps) => {
      // debugger;
      if (state.Fields) {
        state.Fields.forEach(f => {
          let result = ValidationManager.validateField(f);
          f.IsValid = result.IsValid;
          f.ValidationErrors = result.ValidationErrors;
        });
      }
      return state;
    },
    setFormMessage: (state: IFormManagerProps, message: string, callback: (globalState: IFormManagerProps) => void) => {
      if (message === null || message === '') {
        state.GlobalMessage = null;
      } else {
        state.GlobalMessage = {
          Text: message,
          DialogCallback: callback
        };
      }
      return state;
    }
  }
};

const initedStore = initStore(store);

const getFieldControlValuesForPost = (): Object => {
  const state = initedStore.getState();
  let toReturn = {};
  for (let fp of state.Fields) {
    if (fp.InternalName === 'Attachments') {
      continue;
    }
    if (fp.Type.match(/user/gi) || fp.Type.match(/lookup/gi)) {
      let result = null;
      if (fp.FormFieldValue != null) {
        if (!fp.IsMulti) {
          result = parseInt(fp.FormFieldValue.Id);
        } else {
          if (fp.FormFieldValue.results != null && fp.FormFieldValue.results.length > 0) {
            result = { results: fp.FormFieldValue.results.map(r => r.Id) };
          } else {
            result = { results: [] };
          }
        }
      } else {
        if (!fp.IsMulti) {
          result = 0;
        } else {
          result = { results: [] };
        }
      }
      toReturn[`${fp.EntityPropertyName}Id`] = result;
    } else {
      // if (fp.FormFieldValue) {
      //  toReturn[fp.EntityPropertyName] = fp.FormFieldValue;
      // }
      // toReturn[fp.EntityPropertyName] = fp.FormFieldValue == null ? undefined : fp.FormFieldValue;
      toReturn[fp.EntityPropertyName] = fp.FormFieldValue;
    }
  }
  return toReturn;
};

const getNewAttachmentsToSave = (): Promise<AttachmentFileInfo[]> => {
  let toReturn: Promise<AttachmentFileInfo[]> = new Promise<AttachmentFileInfo[]>((resolve, reject) => {
    const state = initedStore.getState();
    let filtered = state.Fields.filter(f => f.InternalName === 'Attachments');
    let attachmentProps: IFieldProps = filtered && filtered.length > 0 ? filtered[0] : null;
    if (attachmentProps.AttachmentsNewToAdd) {
      let individualFilePromises: Promise<AttachmentFileInfo>[] = [];
      attachmentProps.AttachmentsNewToAdd.forEach(na => {
        let individualFilePromise = new Promise<AttachmentFileInfo>((individualPromiseResolve, individualPromiseReject) => {
          const reader = new FileReader();
          reader.onload = () => {
            const fileAsBinaryString = reader.result;
            individualPromiseResolve({
              name: na.name,
              content: fileAsBinaryString
            } as AttachmentFileInfo);
          };
          reader.onabort = () => individualPromiseResolve(null);
          reader.onerror = () => individualPromiseResolve(null);
          reader.readAsBinaryString(na);
        });
        individualFilePromises.push(individualFilePromise);
      });
      Promise.all(individualFilePromises).then((attFileInfos: AttachmentFileInfo[]) => {
        resolve(attFileInfos);
      }).catch(e => {
        resolve(null);
      });
    } else {
      resolve(null);
    }
  });
  return toReturn;
};

const validateForm = (): boolean => {
  initedStore.actions.validateForm();
  let globalState = FormFieldsStore.actions.getState();
  return globalState.Fields.filter(f => !f.IsValid).length === 0;
};

const saveFormData = async (): Promise<ISaveItemResult> => {
  let toResolve = {} as ISaveItemResult;
  try {
    const globalState = FormFieldsStore.actions.getState();
    let formDataRegularFields = FormFieldsStore.actions.getFieldControlValuesForPost();
    let itemCollection = globalState.PnPSPRest.web.lists.getById(globalState.CurrentListId).items;
    let action: Promise<ItemUpdateResult | ItemAddResult> = null;
    if (globalState.CurrentMode === FormMode.New) {
      action = itemCollection.add(formDataRegularFields);
    } else {
      action = itemCollection.getById(globalState.CurrentItemId).update(formDataRegularFields, globalState.ETag);
    }

    try {
      let res: ItemAddResult | ItemUpdateResult = await action;
      toResolve.IsSuccessful = true;
      if (res.data.Id) {
        toResolve.ItemId = parseInt(res.data.Id);
      } else {
        toResolve.ItemId = globalState.CurrentItemId;
      }
      // once we have item id - need to set this to global state
      initedStore.actions.setItemId(toResolve.ItemId);
      let attachmentProps = getFieldPropsByInternalName(globalState.Fields, 'Attachments');
      if (attachmentProps) {
        // upload attachments, if needed
        let attachments: AttachmentFileInfo[] = await getNewAttachmentsToSave();
        if (attachments !== null && attachments.length > 0) {
          let list = globalState.PnPSPRest.web.lists.getById(globalState.CurrentListId);
          let addMultipleResult = await list.items.getById(toResolve.ItemId).attachmentFiles.addMultiple(attachments);

          // add new attachment file data to global state
          let attachmentData =
            await globalState.PnPSPRest.web.lists.getById(globalState.CurrentListId)
            .items.getById(toResolve.ItemId).attachmentFiles.get();
          attachmentProps.FormFieldValue = attachmentData;
        }

        // remove attachments, if needed
        if (attachmentProps.AttachmentsExistingToDelete && attachmentProps.AttachmentsExistingToDelete.length > 0) {
          await globalState.PnPSPRest.web
            .lists.getById(globalState.CurrentListId)
            .items.getById(toResolve.ItemId)
            .attachmentFiles.deleteMultiple(...attachmentProps.AttachmentsExistingToDelete);
          if (attachmentProps.FormFieldValue) {
            attachmentProps.FormFieldValue = attachmentProps.FormFieldValue.filter(v => !attachmentProps.AttachmentsExistingToDelete.includes(v.FileName));
          }
        }
      }
      initedStore.actions.clearHelperAttachmentProperties();
    } catch (e) {
      // realistically this is liklely to indicate problems with network or concurrency
      // console.log(e);
      toResolve.IsSuccessful = false;
      toResolve.Error = e.message.match(/precondition/gi) ? 'Save conflict: current changes would override recent edit(-s) made since this form was opened. Please reload the page and try again.' : e.message;
      toResolve.ItemId = -1;
    }
  } catch (e) {
    // console.log(e);
    toResolve.IsSuccessful = false;
    toResolve.Error = e.message;
    toResolve.ItemId = -1;
  }
  return toResolve;
};

const saveFormDataExternal = async (): Promise<ISaveItemResult> => {
  initedStore.actions.setLoading(true);
  let res: ISaveItemResult = await saveFormData();
  initedStore.actions.setLoading(false);
  return res;
};

const loadingEnabledStateChange = (action, ...args: any[]) => {
  initedStore.actions.setLoading(true);
  action(...args);
  initedStore.actions.setLoading(false);
};

const setFormModeExternal = (formMode: number) => {
  initedStore.actions.setLoading(true);
  initedStore.actions.setFormMode(formMode);
  initedStore.actions.setLoading(false);
};

const configurePnp = (webUrl: string) => {
  sp.setup({
    sp: {
      headers: {
        Accept: 'application/json;odata=verbose'
      },
      baseUrl: webUrl
    }
  });
};

export const FormFieldsStore = {
  // Provider: initedStore.Provider,
  Provider: enhanceProvider(initedStore.Provider),
  Consumer: initedStore.Consumer,
  actions: {
    getState: initedStore.getState,
    initStore: initedStore.actions.initStore,
    setLoading: initedStore.actions.setLoading,
    setFormMode: (arg) => { loadingEnabledStateChange(initedStore.actions.setFormMode, arg); },
    setItemId: initedStore.actions.setItemId,
    setFieldData: initedStore.actions.setFieldData,
    addNewAttachmentInfo: initedStore.actions.addNewAttachmentInfo,
    // addNewAttachmentInfo: (arg) => { loadingEnabledStateChange(initedStore.actions.addNewAttachmentInfo, arg); },
    removeNewAttachmentInfo: initedStore.actions.removeNewAttachmentInfo,
    // removeNewAttachmentInfo: (arg) => { loadingEnabledStateChange(initedStore.actions.removeNewAttachmentInfo, arg); },
    // addOrRemoveExistingAttachmentDeletion: initedStore.actions.addOrRemoveExistingAttachmentDeletion,
    addOrRemoveExistingAttachmentDeletion: (arg) => { loadingEnabledStateChange(initedStore.actions.addOrRemoveExistingAttachmentDeletion, arg); },
    clearHelperAttachmentProperties: initedStore.actions.clearHelperAttachmentProperties,
    // clearHelperAttachmentProperties: () => { loadingEnabledStateChange(initedStore.actions.clearHelperAttachmentProperties); },
    getFieldControlValuesForPost,
    getNewAttachmentsToSave,
    saveFormData: saveFormDataExternal,
    validateForm,
    setShowValidationErrors: initedStore.actions.setShowValidationErrors,
    addValidatorToField: initedStore.actions.addValidatorToField,
    setFieldValidationState: initedStore.actions.setFieldValidationState,
    clearValidatorsFromField: initedStore.actions.clearValidatorsFromField,
    setFieldPropValue: initedStore.actions.setFieldPropValue,
    setFormMessage: initedStore.actions.setFormMessage
  } as IFormManagerActions
};
