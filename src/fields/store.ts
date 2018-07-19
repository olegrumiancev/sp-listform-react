import { sp, SPRest, List, AttachmentFileInfo, ItemUpdateResult, ItemAddResult } from '@pnp/sp';
import createStore from 'react-waterfall';
import { IFormManagerProps, FormMode, IFieldProps, ISaveItemResult, IFormManagerActions } from './interfaces';
import { handleError, getFieldPropsByInternalName } from './utils';
import { FieldPropsManager } from './managers/FieldPropsManager';
import * as React from 'react';
import { ValidationManager } from './managers/ValidationManager';
import { enhanceProvider } from './EnhancedProvider';
// const deasync = require('deasync');
// const deasync = require('synchronize');

let exposedState: IFormManagerProps = null;

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

const storeConfig = {
  initialState: {
    SPWebUrl: null,
    CurrentMode: 0,
    CurrentListId: null,
    IsLoading: true
  } as IFormManagerProps,
  actionsCreators: {
    initStore: async (state: IFormManagerProps, actions: IFormManagerActions, sPWebUrl: string, currentListId: string, currentMode: number, currentItemId?: number): Promise<IFormManagerProps> => {
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
        console.log(item);
        for (const fm of listFields) {
          fieldInfos.push(await FieldPropsManager.createFieldRendererPropsFromFieldMetadata(fm, currentMode, currentListId, item, sp));
        }
        // fieldInfos = listFields.map(fm => {
        //   return await FieldPropsManager.createFieldRendererPropsFromFieldMetadata(fm, currentMode, item, sp);
        // });
        if (item.Attachments) {
          // console.log(attachmentMetadata);
          fieldInfos.filter(f => f.InternalName === 'Attachments')[0].FormFieldValue = attachmentMetadata;
        }
      } else {
        // fieldInfos = listFields.map(fm => {
        //   return FieldPropsManager.createFieldRendererPropsFromFieldMetadata(fm, currentMode, null, sp);
        // });
        for (const fm of listFields) {
          fieldInfos.push(await FieldPropsManager.createFieldRendererPropsFromFieldMetadata(fm, currentMode, currentListId, null, sp));
        }
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
    setFormMode: async (state: IFormManagerProps, actions: IFormManagerActions, mode: number) => {
      console.log(state);

      state.CurrentMode = mode;
      state.Fields.forEach(f => f.CurrentMode = mode);
      return { ...state };
    },
    setItemId: async (state: IFormManagerProps, actions: IFormManagerActions, itemId: number) => {
      state.CurrentItemId = itemId;
      return { ...state };
    },
    setLoading: async (state: IFormManagerProps, actions: IFormManagerActions, isLoading: boolean) => {
      state.IsLoading = isLoading;
      return { ...state };
    },
    setEtag: async (state: IFormManagerProps, actions: IFormManagerActions, etag: string) => {
      state.ETag = etag;
      return { ...state };
    },
    setShowValidationErrors: async (state: IFormManagerProps, actions: IFormManagerActions, show: boolean) => {
      state.ShowValidationErrors = show;
      state.Fields = state.Fields.map(f => {
        f.ShowValidationErrors = show;
        return f;
      });
      return { ...state };
    },
    setFieldData: async (state: IFormManagerProps, actions: IFormManagerActions, internalName: string, newValue: any) => {
      let fieldProps = getFieldPropsByInternalName(state.Fields, internalName);
      if (fieldProps) {
        fieldProps.FormFieldValue = newValue;
      }
      return { ...state };
    },
    setFieldValidationState: async (state: IFormManagerProps, actions: IFormManagerActions, internalName: string, isValid: boolean, validationErrors: string[]) => {
      let fieldProps = getFieldPropsByInternalName(state.Fields, internalName);
      if (fieldProps) {
        fieldProps.IsValid = isValid;
        fieldProps.ValidationErrors = validationErrors;
      }
      return { ...state };
    },
    addNewAttachmentInfo: async (state: IFormManagerProps, actions: IFormManagerActions, fileInfo: any) => {
      let attachmentProps = getFieldPropsByInternalName(state.Fields, 'Attachments');
      if (attachmentProps) {
        if (!attachmentProps.AttachmentsNewToAdd) {
          attachmentProps.AttachmentsNewToAdd = [];
        }
        attachmentProps.AttachmentsNewToAdd.push(fileInfo);
      }
      return { ...state };
    },
    removeNewAttachmentInfo: async (state: IFormManagerProps, actions: IFormManagerActions, fileInfo: any) => {
      let attachmentProps = getFieldPropsByInternalName(state.Fields, 'Attachments');
      if (attachmentProps && attachmentProps.AttachmentsNewToAdd) {
        attachmentProps.AttachmentsNewToAdd = attachmentProps.AttachmentsNewToAdd.filter(a => a.name !== fileInfo.name);
      }
      return { ...state };
    },
    addOrRemoveExistingAttachmentDeletion: async (state: IFormManagerProps, actions: IFormManagerActions, attachmentName: string) => {
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
      return { ...state };
    },
    clearHelperAttachmentProperties: async (state: IFormManagerProps) => {
      let attachmentProps = getFieldPropsByInternalName(state.Fields, 'Attachments');
      if (attachmentProps) {
        attachmentProps.AttachmentsExistingToDelete = null;
        attachmentProps.AttachmentsNewToAdd = null;
      }
      return { ...state };
    },
    setFieldPropValue: async (state: IFormManagerProps, actions: IFormManagerActions, internalName: string, propName: string, propValue: any) => {
      let fieldProps = getFieldPropsByInternalName(state.Fields, internalName);
      if (fieldProps) {
        fieldProps[propName] = propValue;
      }
      return { ...state };
    },
    addValidatorToField: async (state: IFormManagerProps, actions: IFormManagerActions, validator: Function, internalName: string, ...validatorParams: any[]) => {
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
      return { ...state };
    },
    clearValidatorsFromField: async (state: IFormManagerProps, actions: IFormManagerActions, internalName: string) => {
      let fieldProps = getFieldPropsByInternalName(state.Fields, internalName);
      if (fieldProps) {
        fieldProps.Validators = [];
      }
      return { ...state };
    },
    validateForm: async (state: IFormManagerProps) => {
      // debugger;
      if (state.Fields) {
        state.Fields.forEach(f => {
          let result = ValidationManager.validateField(f);
          f.IsValid = result.IsValid;
          f.ValidationErrors = result.ValidationErrors;
        });
      }
      return { ...state };
    },
    setFormMessage: async (state: IFormManagerProps, actions: IFormManagerActions, message: string, callback: (globalState: IFormManagerProps) => void) => {
      if (message === null || message === '') {
        state.GlobalMessage = null;
      } else {
        state.GlobalMessage = {
          Text: message,
          DialogCallback: callback
        };
      }
      return { ...state };
    }
  }
};

// const initedStore = initStore(storeConfig);
const initedStore = createStore(storeConfig);

// subscribe to store
initedStore.subscribe((action, state, args) => {
  // console.log(`subscriber, action: `, action, `state: `, state, `args: `, args);
  // console.log(new Date().toISOString());
  exposedState = state;
});

const getFieldControlValuesForPost = async (): Promise<Object> => {
  const state = // FormFieldsStore.actions.getState();
    exposedState;
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
    } else if (fp.Type.match(/taxonomy/gi)) {
      let result = null;
      let validField = fp.InternalName;
      if (fp.FormFieldValue && fp.FormFieldValue.length > 0) {
        if (fp.IsMulti) {
          result = fp.FormFieldValue.map(term => `-1;#${term.name}|${term.key}`).join(';#') + ';';
          validField = fp.TaxonomyUpdateFieldEntityPropertyName;
        } else {
          let term = fp.FormFieldValue[0];
          result = {
            __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
            Label: term.name,
            TermGuid: term.key,
            WssId: -1
          };
        }
      }
      toReturn[validField] = result;
    } else {
      // if (fp.FormFieldValue) {
      //  toReturn[fp.EntityPropertyName] = fp.FormFieldValue;
      // }
      // toReturn[fp.EntityPropertyName] = fp.FormFieldValue == null ? undefined : fp.FormFieldValue;
      toReturn[fp.EntityPropertyName] = fp.FormFieldValue;
    }
  }
  // console.log(toReturn);
  return toReturn;
};

const getFieldControlValuesForValidatedUpdate = async (): Promise<any[]> => {
  const state = // FormFieldsStore.actions.getState();
    exposedState;
  let toReturn = [];
  for (let fp of state.Fields) {
    let fieldValue = null;
    if (fp.InternalName === 'Attachments') {
      continue;
    }
    if (fp.Type.match(/lookup/gi)) {
    // if (fp.Type.match(/lookup/gi) || fp.Type.match(/user/gi)) {
      if (fp.FormFieldValue != null) {
        if (!fp.IsMulti) {
          fieldValue = fp.FormFieldValue.Id.toString();
        } else {
          if (fp.FormFieldValue.results != null && fp.FormFieldValue.results.length > 0) {
            fieldValue = fp.FormFieldValue.results.map(r => `${r.Id};#`).join(';#');
          }
        }
      }
    } else if (fp.Type.match(/user/gi)) {
      if (fp.FormFieldValue != null) {
        if (!fp.IsMulti) {
          fieldValue = `[${JSON.stringify({ Key: fp.FormFieldValue.key })}]`;
        } else {
          if (fp.FormFieldValue.results != null && fp.FormFieldValue.results.length > 0) {
            let results = fp.FormFieldValue.results.map(r => {
              return JSON.stringify({ Key: r.key });
            }).join(',');
            fieldValue = `[${results}]`;
          }
        }
      }
    } else if (fp.Type.match(/taxonomy/gi)) {
      if (fp.FormFieldValue && fp.FormFieldValue.length > 0) {
        fieldValue = fp.FormFieldValue.map(term => `${term.name}|${term.key}`).join(';');
      }
    } else if (fp.Type.match(/multichoice/gi)) {
      if (fp.FormFieldValue && fp.FormFieldValue.results && fp.FormFieldValue.results.length > 0) {
        fieldValue = fp.FormFieldValue.results.join(';#');
      }
    } else if (fp.Type.match(/datetime/gi)) {
      let d = fp.FormFieldValue === null || fp.FormFieldValue === undefined ? new Date(1900, 0, 1) : new Date(Date.parse(fp.FormFieldValue));
      fieldValue = d.format('dd/MM/yyyy HH:mm');
    } else if (fp.Type.match(/number/gi)) {
      if (fp.FormFieldValue) {
        if (fp.NumberIsPercent) {
          fieldValue = (fp.FormFieldValue * 100).toString();
        } else {
          fieldValue = fp.FormFieldValue;
        }
      }
    } else {
      fieldValue = fp.FormFieldValue;
    }

    if (fieldValue === undefined || fieldValue === null) {
      fieldValue = null;
    } else {
      fieldValue = fieldValue.toString();
    }

    toReturn.push({
      ErrorMessage: null,
      FieldName: fp.EntityPropertyName,
      FieldValue: fieldValue,
      HasException: false
    });
  }
  // console.log(toReturn);
  return toReturn;
};

// const getNewAttachmentsToSave = (): Promise<AttachmentFileInfo[]> => {
//   let toReturn: Promise<AttachmentFileInfo[]> = new Promise<AttachmentFileInfo[]>((resolve, reject) => {
//     const state = FormFieldsStore.actions.getState();
//     let filtered = state.Fields.filter(f => f.InternalName === 'Attachments');
//     let attachmentProps: IFieldProps = filtered && filtered.length > 0 ? filtered[0] : null;
//     if (attachmentProps.AttachmentsNewToAdd) {
//       let individualFilePromises: Promise<AttachmentFileInfo>[] = [];
//       attachmentProps.AttachmentsNewToAdd.forEach(na => {
//         let individualFilePromise = new Promise<AttachmentFileInfo>((individualPromiseResolve, individualPromiseReject) => {
//           const reader = new FileReader();
//           reader.onload = () => {
//             const fileAsBinaryString = reader.result;
//             individualPromiseResolve({
//               name: na.name,
//               content: fileAsBinaryString
//             } as AttachmentFileInfo);
//           };
//           reader.onabort = () => individualPromiseResolve(null);
//           reader.onerror = () => individualPromiseResolve(null);
//           reader.readAsBinaryString(na);
//         });
//         individualFilePromises.push(individualFilePromise);
//       });
//       Promise.all(individualFilePromises).then((attFileInfos: AttachmentFileInfo[]) => {
//         resolve(attFileInfos);
//       }).catch(e => {
//         resolve(null);
//       });
//     } else {
//       resolve(null);
//     }
//   });
//   return toReturn;
// };

const getNewAttachmentsToSave = (): Promise<AttachmentFileInfo[]> => {
  let toReturn: Promise<AttachmentFileInfo[]> = new Promise<AttachmentFileInfo[]>((resolve, reject) => {
    const state = // FormFieldsStore.actions.getState();
      exposedState;
    let filtered = state.Fields.filter(f => f.InternalName === 'Attachments');
    let attachmentProps: IFieldProps = filtered && filtered.length > 0 ? filtered[0] : null;
    if (attachmentProps.AttachmentsNewToAdd) {
      let individualFilePromises: Promise<AttachmentFileInfo>[] = [];
      attachmentProps.AttachmentsNewToAdd.forEach(na => {
        let individualFilePromise = new Promise<AttachmentFileInfo>((individualPromiseResolve, individualPromiseReject) => {
          const reader = new FileReader();
          reader.onload = () => {
            const res = reader.result;
            individualPromiseResolve({
              name: na.name,
              content: res
            } as AttachmentFileInfo);
          };
          reader.onabort = () => individualPromiseResolve(null);
          reader.onerror = () => individualPromiseResolve(null);
          reader.readAsArrayBuffer(na);
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
  let globalState = // FormFieldsStore.actions.getState();
    exposedState;
  let isValid = true;
  if (globalState && globalState.Fields) {
    isValid = globalState.Fields.filter(f => !f.IsValid).length === 0;
  }
  return isValid;
};

const saveFormData = async (): Promise<ISaveItemResult> => {
  let toResolve = {} as ISaveItemResult;
  try {
    const globalState = // FormFieldsStore.actions.getState();
      exposedState;
    let formDataRegularFields =
      await getFieldControlValuesForValidatedUpdate();
      // await FormFieldsStore.actions.getFieldControlValuesForPost();

    let itemCollection = globalState.PnPSPRest.web.lists.getById(globalState.CurrentListId).items;
    let action: Promise<any> = null;

    // rewite adding as regular add with no properties and then + validateupdate
    let currentEtag = globalState.ETag;
    let currentItemId = globalState.CurrentItemId;
    if (globalState.CurrentMode === FormMode.New) {
      // action = itemCollection.add(formDataRegularFields);

      // action = globalState.PnPSPRest.web.lists.getById(globalState.CurrentListId).addValidateUpdateItemUsingPath(formDataRegularFields);
      let initialAdding: ItemAddResult = await itemCollection.add();
      console.log(initialAdding);
      if (initialAdding && initialAdding.data && initialAdding.data.Id) {
        currentItemId = parseInt(initialAdding.data.Id);
        console.log(currentItemId);
        FormFieldsStore.actions.setItemId(currentItemId);
      }
    }
      // action = itemCollection.getById(globalState.CurrentItemId).update(formDataRegularFields, globalState.ETag);

    action = itemCollection.getById(currentItemId).configure({
      headers: {
        // 'If-Match': `${globalState.ETag}`
        'If-Match': `${currentEtag}`
      }
    }).validateUpdateListItem(formDataRegularFields);

    // try {
      // debugger;
      // let res: ItemAddResult | ItemUpdateResult = await action;
    let res = await action;
    if (res.ValidateUpdateListItem.results.some(f => f.HasException)) {
      let errors = res.ValidateUpdateListItem.results.reduce((prev, current) => {
        if (current.HasException) {
          let props = getFieldPropsByInternalName(globalState.Fields, current.FieldName);
          prev.push(`${props.Title}: ${current.ErrorMessage}`);
        }
        return prev;
      }, []).join('<br />');
      throw new Error(errors);
    }

    try {
      // console.log(res);
      toResolve.IsSuccessful = true;

      // try assigning item id
      if (res && res.data && res.data.Id) {
        toResolve.ItemId = parseInt(res.data.Id);
      } else {
        toResolve.ItemId = currentItemId;
      }

      // try assigning new etag
      if (res && res.data && res.data['odata.etag']) {
        toResolve.ETag = res.data['odata.etag'];
      } else {
        toResolve.ETag = globalState.ETag;
      }

      // console.log(toResolve);

      // once we have item id - need to set this to global state
      // initedStore.actions.setItemId(toResolve.ItemId);
      // debugger;
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
      toResolve.Error = e.message.match(/precondition/gi) ? 'Save conflict - current changes would override recent edit(-s) made since this form was opened. Please reload the page and try again.' : e.message;
      toResolve.ItemId = -1;
      toResolve.ETag = null;
    }
  } catch (e) {
    // console.log(e);
    toResolve.IsSuccessful = false;
    toResolve.Error = e.toString();
    toResolve.ItemId = -1;
    toResolve.ETag = null;
  }
  return toResolve;
};

const saveFormDataExternal = async (): Promise<ISaveItemResult> => {
  initedStore.actions.setLoading(true);
  let res: ISaveItemResult = await saveFormData();
  if (res.IsSuccessful) {
    initedStore.actions.setEtag(res.ETag);
    initedStore.actions.setItemId(res.ItemId);
  }
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

export const FormFieldsStore = {
  // Provider: initedStore.Provider,
  Provider: enhanceProvider(initedStore.Provider),
  // Consumer: initedStore.Consumer,
  connect: initedStore.connect,
  actions: {
    getState: () => {
      return exposedState;
    },
    initStore: initedStore.actions.initStore,
    setLoading: initedStore.actions.setLoading,
    setFormMode: (arg) => { loadingEnabledStateChange(initedStore.actions.setFormMode, arg); },
    setItemId: initedStore.actions.setItemId,
    setFieldData: initedStore.actions.setFieldData,
    addNewAttachmentInfo: initedStore.actions.addNewAttachmentInfo,
    removeNewAttachmentInfo: initedStore.actions.removeNewAttachmentInfo,
    addOrRemoveExistingAttachmentDeletion: (arg) => { loadingEnabledStateChange(initedStore.actions.addOrRemoveExistingAttachmentDeletion, arg); },
    clearHelperAttachmentProperties: initedStore.actions.clearHelperAttachmentProperties,
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
  // tslint:disable-next-line:one-line
  } as IFormManagerActions
};
