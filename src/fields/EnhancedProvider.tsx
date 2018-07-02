import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { FormFieldsStore } from './store';
import { IFormManagerProps } from './interfaces';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

export const enhanceProvider = (InitialProviderComponent) => {
  return class extends React.Component {
    constructor(props) {
      super(props);
    }

    public render() {
      return (
        <InitialProviderComponent {...this.props}>
          <FormFieldsStore.Consumer mapStateToProps={(state) => state}>
            {(state: IFormManagerProps) => {
              if (!state || !state.GlobalMessage) {
                return null;
              }

              return (<Dialog
                hidden={false}
                onDismiss={() => FormFieldsStore.actions.setFormMessage(null)}
                dialogContentProps={{
                  type: DialogType.normal,
                  title: 'Message',
                  subText: state && state.GlobalMessage ? state.GlobalMessage.Text : ''
                }}
                modalProps={{
                  isBlocking: false,
                  containerClassName: 'ms-dialogMainOverride'
                }}
              >
                <DialogFooter>
                  <PrimaryButton onClick={() => {
                    if (state.GlobalMessage.DialogCallback) {
                      state.GlobalMessage.DialogCallback(state);
                    }
                    FormFieldsStore.actions.setFormMessage(null);
                  }} text='OK' />
                  {/* <DefaultButton onClick={this._closeDialog} text='Cancel' /> */}
                </DialogFooter>
              </Dialog>);
            }}
          </FormFieldsStore.Consumer>
          {this.props.children}
        </InitialProviderComponent>
      );
    }
  };
};
