import { SPRest } from '@pnp/sp';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
// import { IFieldProps } from '@rumiancev/sp-react-formfields/lib/interfaces';

export interface IFormMode {
  New: number,
  Display: number,
  Edit: number
}

export const FormMode: IFormMode = {
  New: 1, Display: 2, Edit: 3
};

export interface IListFormProps {
  CurrentListId?: string;
  CurrentItemId?: number;
  SpWebUrl?: string;
  CurrentMode: number;
  IsLoading?: boolean;
  IsSaving?: boolean;
  pnpSPRest?: SPRest;
}

export interface IFieldUpdate {
  fieldInternalName: string;
  newValue: any;
}

export interface IPeoplePickerState {
  peopleList: IPersonaProps[];
  mostRecentlyUsed: IPersonaProps[];
  currentSelectedItems: any[];
}

export const getQueryString = (url, field) => {
  let href = url ? url : window.location.href;
  let reg = new RegExp('[?&]' + field + '=([^&#]*)', 'i');
  let s = reg.exec(href);
  return s ? s[1] : null;
};

export const executeSPQuery = async (ctx: SP.ClientRuntimeContext): Promise<any> => {
  let promise = new Promise<any>((resolve, reject) => {
    ctx.executeQueryAsync((sender, args) => {
      resolve();
    }, (sender, args) => {
      reject(args.get_message());
    });
  });
  return promise;
};
