import * as React from 'react';
import { IFieldProps, FormMode } from '../interfaces';
import { BaseFieldRenderer } from './BaseFieldRenderer';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { NormalPeoplePicker, IBasePicker, ValidationState, BasePeoplePicker } from 'office-ui-fabric-react/lib/Pickers';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { WebEnsureUserResult, PeoplePickerEntity, PrincipalSource, PrincipalType, PrincipalInfo } from '@pnp/sp';
import './FieldUserRenderer.css';
import { handleError } from '../utils';

export class FieldUserRenderer extends BaseFieldRenderer {
  // private pp: IBasePicker<IPersonaProps> = null;
  private isFieldMounted = false;
  public constructor(props: IFieldProps) {
    super(props);
    let vals = [];
    if (this.props.FormFieldValue != null) {
      if (this.props.IsMulti) {
        vals = this.props.FormFieldValue.results;
      } else {
        vals.push(this.props.FormFieldValue);
      }
    }

    let selectedValues: IPersonaProps[] = [];
    for (let v of vals) {
      if (v != null && v.Id !== 0) {
        selectedValues.push({
          text: v.Title,
          id: v.Id.toString()
        });
      }
    }

    this.state = {
      peopleList: [],
      currentSelectedItems: selectedValues,
      mostRecentlyUsed: [],
      spGroupRestriction: null
    };
  }

  public componentWillUnmount() {
    this.isFieldMounted = false;
  }

  public componentDidMount() {
    this.isFieldMounted = true;
    let spGroupRestriction = this.props.SchemaXml.documentElement.getAttribute('UserSelectionScope');
    if (spGroupRestriction !== '0') {
      this.props.pnpSPRest.web.siteGroups.getById(parseInt(spGroupRestriction)).get().then(res => {
        try {
          if (this.isFieldMounted) {
            this.setState({ spGroupRestriction: res.LoginName });
          }
        } catch (e) {
          // ...
        }
      });
    }

    // get the login names of selected user values
    if (this.state.currentSelectedItems && this.state.currentSelectedItems.length > 0) {
      let promises: Promise<any>[] = [];
      let newValues = [];
      for (let sv of this.state.currentSelectedItems) {
        let p = this.props.pnpSPRest.web.siteUsers.getById(sv.id).get();
        promises.push(p);
        p.then(r => {
          console.log(r);
          newValues.push({
            id: r.Id,
            key: r.PrincipalType === 4 ? r.Title : r.LoginName,
            text: r.Title
          });
          // console.log(r);
        });
      }
      Promise.all(promises).then(() => {
        if (this.isFieldMounted) {
          this.setState({ currentSelectedItems: newValues }, () => {
            this.saveDataInternal();
          });
        }
      });
    }
    // this.saveDataInternal();
  }

  protected renderNewForm() {
    return this.renderNewOrEditForm();
  }

  protected renderEditForm() {
    return this.renderNewOrEditForm();
  }

  protected renderDispForm() {
    if (this.props.FormFieldValue == null) {
      return null;
    }

    return (
      <div>
      {this.state.currentSelectedItems.map((m, i) => {
        return <Label key={`${m.id}_${i}`}>{m.text}</Label>;
      })}
      </div>
    );
  }

  private renderNewOrEditForm() {
    return (<div>
      <NormalPeoplePicker
      itemLimit={this.props.IsMulti ? undefined : 1}
      defaultSelectedItems={this.state.currentSelectedItems}
      onResolveSuggestions={this.onFilterChanged}
      getTextFromItem={this.getTextFromItem}
      pickerSuggestionsProps={{
        searchingText: 'Searching more...'
      }}
      className={'ms-PeoplePicker'}
      key={`${this.props.InternalName}_normalpicker`}
      onRemoveSuggestion={this.onRemoveSuggestion}
      onValidateInput={this.validateInput}
      removeButtonAriaLabel={'Remove'}
      onChange={this.onItemsChange}
      searchingText={'Searching...'}
      inputProps={{
        placeholder: 'Enter a name or email address'
      }}
      resolveDelay={300}
      // componentRef={(p) => { this.pp = p; }}
      />
      </div>);
  }

  private getTextFromItem(persona: IPersonaProps): string {
    return persona.text as string;
  }

  private validateInput = (input: string): ValidationState => {
    if (input.indexOf('@') !== -1) {
      return ValidationState.valid;
    } else if (input.length > 1) {
      return ValidationState.warning;
    } else {
      return ValidationState.invalid;
    }
  }

  private onItemsChange = (items: any[]): void => {
    // some items will not contain local Id - need to call ensureUser
    let result = [];
    let promises: Promise<WebEnsureUserResult>[] = [];
    for (let entry of items) {
      console.log(entry);
      if (entry.id === -1) {
        let pp = this.props.pnpSPRest.web.ensureUser(entry.key);
        pp.catch(e => {
          // handleError(e);
        });
        promises.push(pp);
        pp.then((r: WebEnsureUserResult) => {
          result.push({
            text: r.data.Title,
            id: r.data.Id.toString(),
            key: r.data.PrincipalType === 4 ? r.data.Title : r.data.LoginName // r.data.LoginName
          });
        });
      } else {
        let selected = this.state.currentSelectedItems.filter(i => i.id.toString() === entry.id);
        console.log(selected);
        if (selected && selected.length > 0) {
          selected[0].id = selected[0].id.toString();
          result.push(selected[0]);
        }
      }
    }
    Promise.all(promises).then(() => {
      this.setState({ currentSelectedItems: result }, () => {
        this.saveDataInternal();
      });
    }).catch(e => {
      // ...
    });
  }

  private onRemoveSuggestion = (item: IPersonaProps): void => {
    const { peopleList, mostRecentlyUsed: mruState } = this.state;
    const indexPeopleList: number = peopleList.indexOf(item);
    const indexMostRecentlyUsed: number = mruState.indexOf(item);

    if (indexPeopleList >= 0) {
      const newPeople: IPersonaProps[] = peopleList.slice(0, indexPeopleList).concat(peopleList.slice(indexPeopleList + 1));
      this.setState({ peopleList: newPeople });
    }

    if (indexMostRecentlyUsed >= 0) {
      const newSuggestedPeople: IPersonaProps[] = mruState.slice(0, indexMostRecentlyUsed).concat(mruState.slice(indexMostRecentlyUsed + 1));
      this.setState({ mostRecentlyUsed: newSuggestedPeople });
    }
  }

  private onFilterChanged = (filterText: string, currentPersonas: IPersonaProps[],
                               limitResults?: number): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (filterText) {
      let scope: PrincipalType = PrincipalType.User;
      if (this.props.UserSelectionMode === 'PeopleAndGroups') {
        scope |= PrincipalType.SecurityGroup;
      }

      let p = new Promise<IPersonaProps[]>((resolve) => {
        let searchQuery = this.props.pnpSPRest.utility.searchPrincipals(
          filterText,
          scope,
          PrincipalSource.All,
          this.state.spGroupRestriction ? this.state.spGroupRestriction : '',
          10);
        searchQuery.then((entries: any) => {
          let result = [];
          if (entries && entries.SearchPrincipalsUsingContextWeb &&
              entries.SearchPrincipalsUsingContextWeb.results && entries.SearchPrincipalsUsingContextWeb.results.length > 0) {
            console.log(entries.SearchPrincipalsUsingContextWeb.results);
            result = entries.SearchPrincipalsUsingContextWeb.results.map(e => ({
              text: e.DisplayName,
              id: e.PrincipalId,
              key: e.PrincipalType === 4 ? e.Email : e.LoginName
              // key: e.LoginName
            }));
          }
          resolve(result);
        }).catch(e => resolve([]));
      });

      p.catch(e => {
        // console.log(e);
      });

      return p;
    } else {
      return [];
    }
  }

  private saveDataInternal = () => {
    let result = this.state.currentSelectedItems.map((persona) => {
      return {
        Title: persona.text,
        Id: persona.id,
        key: persona.key
      };
    });

    if (this.props.IsMulti) {
      if (result && result.length > 0) {
        result = { results: result };
      } else {
        result = null;
      }
    } else {
      if (this.state.currentSelectedItems.length > 0) {
        result = {
          Title: this.state.currentSelectedItems[0].text,
          Id: this.state.currentSelectedItems[0].id,
          key: this.state.currentSelectedItems[0].key
        };
      } else {
        result = null;
      }
    }
    this.trySetChangedValue(result);
  }
}
