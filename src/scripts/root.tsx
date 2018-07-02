require('core-js/shim');
import * as React from 'react';
import { sp } from '@pnp/sp';
import { setupPnp } from './utils/odata';
import ListForm from './components/ListForm';
import { FormMode, getQueryString, executeSPQuery, IListFormProps } from './interfaces';

export class RootInternal extends React.Component<{}, {formProps: IListFormProps}> {
  private localContext = null;

  public constructor(props) {
    super(props);
    this.localContext = SP.ClientContext.get_current();
    this.state = {
      formProps: null
    };
  }

  public componentDidMount() {
    const promise = this.createInitialProps();
    promise.then(formProps => {
      this.setState({ formProps });
    });
  }

  public render() {
    return (
      <div>
        {this.state.formProps == null ? null : <ListForm {...this.state.formProps} />}
      </div>
    );
  }

  private createInitialProps = async (): Promise<IListFormProps> => {
    let currentWeb = this.localContext.get_web();
    this.localContext.load(currentWeb);
    await executeSPQuery(this.localContext);
    let webUrl = currentWeb.get_url();

    setupPnp(webUrl);

    return {
      pnpSPRest: sp,
      Fields: [],
      CurrentMode: this.getFormMode(),
      CurrentListId: this.getCurrentListId(),
      CurrentItemId: this.getCurrentItemId(),
      SpWebUrl: webUrl,
      IsLoading: true
    } as IListFormProps;
  }

  private getFormMode = () => {
    let fm = getQueryString(null, 'fm');
    if (fm != null) {
      return parseInt(fm);
    }
    return FormMode.New;
  }

  private getCurrentListId = () => {
    return getQueryString(null, 'listid');
  }

  private getCurrentItemId = () => {
    let itemid = getQueryString(null, 'itemid');
    if (itemid == null) {
      return 0;
    } else {
      return parseInt(itemid);
    }
  }
}

export default (<RootInternal />);
