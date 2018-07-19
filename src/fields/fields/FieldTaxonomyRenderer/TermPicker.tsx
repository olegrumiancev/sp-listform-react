import * as React from 'react';
import { BasePicker, IBasePickerProps, IPickerItemProps } from 'office-ui-fabric-react/lib/Pickers';
import { IPickerTerm, ITermPickerProps, ITermPickerState } from './interfaces';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export class TermBasePicker extends BasePicker<IPickerTerm, IBasePickerProps<IPickerTerm>> { }

export default class TermPicker extends React.Component<ITermPickerProps, ITermPickerState> {
  private tagPicker: TermBasePicker = null;
  constructor(props: any) {
    super(props);
    this.onRenderItem = this.onRenderItem.bind(this);
    this.onRenderSuggestionsItem = this.onRenderSuggestionsItem.bind(this);
    this.onFilterChanged = this.onFilterChanged.bind(this);
    this.onGetTextFromItem = this.onGetTextFromItem.bind(this);

    this.state = {
      terms: this.props.value
    };

  }

  public componentWillReceiveProps(newProps: any) {
    // debugger;
    let isUpdateNeeded = false;

    // compare
    if (this.state.terms === undefined || newProps.value === undefined) {
      isUpdateNeeded = true;
    } else if (this.state.terms.length === newProps.value.length) {
      for (let i = 0; i < newProps.value.length; i++) {
        if (this.state.terms[i].key !== newProps.value[i].key) {
          isUpdateNeeded = true;
          break;
        }
      }
    } else {
      isUpdateNeeded = true;
    }

    if (isUpdateNeeded) {
      let toSet = newProps.value === undefined || newProps.value === null ? [] : newProps.value;
      this.setState({ terms: toSet }, () => {
        if (this.tagPicker) {
          this.tagPicker.setState({ items: this.state.terms });
        }
      });
    }
  }

  public render(): JSX.Element {
    return (
      <div>
        <TermBasePicker
          disabled={this.props.disabled}
          onResolveSuggestions={this.onFilterChanged}
          onRenderSuggestionsItem={this.onRenderSuggestionsItem}
          getTextFromItem={this.onGetTextFromItem}
          onRenderItem={this.onRenderItem}
          defaultSelectedItems={this.state.terms}
          onChange={items => {
            this.setState({ terms: items });
            this.props.onChanged(items);
          }}
          itemLimit={!this.props.allowMultipleSelections ? 1 : undefined}
          className={`termBasePicker`}
          ref={(r) => this.tagPicker = r}
        />
      </div>
    );
  }

  protected onRenderItem(term: IPickerItemProps<IPickerTerm>) {
    return (
      <div className={`pickedTermRoot`}
           key={term.index}
           data-selection-index={term.index}
           data-is-focusable={!term.disabled && true}>
        <span className={`pickedTermText`}>{term.item.name}</span>
        {
          !term.disabled && (
            <span className={`pickedTermCloseIcon`}
              onClick={term.onRemoveItem}>
              <Icon iconName='Cancel' />
            </span>
          )
        }
      </div>
    );
  }

  /**
   * Renders the suggestions in the picker
   */
  protected onRenderSuggestionsItem(term: IPickerTerm) {
    let termParent = term.termSetName;
    let termTitle = `${term.name} [${term.termSetName}]`;
    if (term.path.indexOf(';') !== -1) {
      let splitPath = term.path.split(';');
      termParent = splitPath[splitPath.length - 2];
      splitPath.pop();
      termTitle = `${term.name} [${term.termSetName}:${splitPath.join(':')}]`;
    }
    return (
      <div className={`termSuggestion`} title={termTitle}>
        <div>{term.name}</div>
        {/* <div className={`termSuggestionSubTitle`}> {'TaxonomyPickerInLabel'} {termParent ? termParent : 'TaxonomyPickerTermSetLabel'}</div> */}
        <div className={`termSuggestionSubTitle`}> {'in'} {termParent ? termParent : 'TaxonomyPickerTermSetLabel'}</div>
      </div>
    );
  }

  /**
   * When Filter Changes a new search for suggestions
   */
  private async onFilterChanged(filterText: string, tagList: IPickerTerm[]): Promise<IPickerTerm[]> {
    if (filterText !== '') {
      // let termsService = new SPTermStorePickerService(this.props.fieldProps);
      // let terms: IPickerTerm[] = await termsService.searchTermsByName(filterText);
      // // Check if the termset can be selected
      // if (this.props.isTermSetSelectable) {
      //   // Retrieve the current termset
      //   const termSet = await termsService.getTermSet();
      //   // Check if termset was retrieved and if it contains the filter value
      //   if (termSet && termSet.Name.toLowerCase().indexOf(filterText.toLowerCase()) === 0) {
      //     // Add the termset to the suggestion list
      //     terms.push({
      //       key: termsService.cleanGuid(termSet.Id),
      //       name: termSet.Name,
      //       path: '',
      //       termSet: termsService.cleanGuid(termSet.Id)
      //     });
      //   }
      // }

      // Filter out the terms which are already set
      const filteredTerms = this.props.allTerms.filter(t => t.name.toLowerCase().indexOf(filterText.toLowerCase()) !== -1);

      // const { disabledTermIds, disableChildrenOfDisabledParents } = this.props;
      // for (const term of terms) {
      //   let canBePicked = true;

      //   // Check if term is not disabled
      //   if (disabledTermIds && disabledTermIds.length > 0) {
      //     // Check if current term need to be disabled
      //     if (disabledTermIds.indexOf(term.key) !== -1) {
      //       canBePicked = false;
      //     } else {
      //       // Check if child terms need to be disabled
      //       if (disableChildrenOfDisabledParents) {
      //         // Check if terms were already retrieved
      //         if (!this.allTerms) {
      //           this.allTerms = (await termsService.getAllTerms(this.props.fieldProps.TaxonomyTermSetId)).Terms;
      //         }

      //         // Check if there are terms retrieved
      //         if (this.allTerms.length > 0) {
      //           // Find the disabled parents
      //           const disabledParents = this.allTerms.filter(t => disabledTermIds.indexOf(t.key) !== -1);
      //           // Check if disabled parents were found
      //           if (disabledParents && disabledParents.length > 0) {
      //             // Check if the current term lives underneath a disabled parent
      //             const findTerm = disabledParents.filter(pt => term.path.indexOf(pt.path) !== -1);
      //             if (findTerm && findTerm.length > 0) {
      //               canBePicked = false;
      //             }
      //           }
      //         }
      //       }
      //     }
      //   }

      //   if (canBePicked) {
      //     // Only retrieve the terms which are not yet tagged
      //     if (tagList.filter(tag => tag.key === term.key).length === 0) {
      //       filteredTerms.push(term);
      //     }
      //   }
      // }
      return filteredTerms;
    } else {
      return Promise.resolve([]);
    }
  }

  /**
   * gets the text from an item
   */
  private onGetTextFromItem(item: any): any {
    return item.name;
  }
}
