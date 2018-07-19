import * as React from 'react';
import { PrimaryButton, DefaultButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import TermPicker from './TermPicker';
import { IPickerTerms, IPickerTerm } from './interfaces';
import TermParent from './TermParent';
import FieldErrorMessage from './ErrorMessage';
import SPTermStorePickerService from './SPTermStorePickerService';
import { IFieldProps } from '../../interfaces';
import { BaseFieldRenderer } from '../BaseFieldRenderer';
import './TaxonomyPicker.css';

/**
 * Image URLs / Base64
 */
export const COLLAPSED_IMG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAA8AAAAUCAYAAABSx2cSAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAABh0RVh0U29mdHdhcmUAUGFpbnQuTkVUIHYzLjEwcrIlkgAAAIJJREFUOE/NkjEKwCAMRdu7ewZXJ/EqHkJwE9TBCwR+a6FLUQsRwYBTeD8/35wADnZVmPvY4OOYO3UNbK1FKeUWH+fRtK21hjEG3vuhQBdOKUEpBedcV6ALExFijJBSIufcFBjCVSCEACEEqpNvBmsmT+3MTnvqn/+O4+1vdtv7274APmNjtuXVz6sAAAAASUVORK5CYII='; // /_layouts/15/images/MDNCollapsed.png
export const EXPANDED_IMG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAA8AAAAUCAYAAABSx2cSAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAABh0RVh0U29mdHdhcmUAUGFpbnQuTkVUIHYzLjEwcrIlkgAAAFtJREFUOE9j/P//PwPZAKSZXEy2RrCLybV1CGjetWvX/46ODqBLUQOXoJ9BGtXU1MCYJM0wjZGRkaRpRtZIkmZ0jSRpBgUOzJ8wmqwAw5eICIb2qGYSkyfNAgwAasU+UQcFvD8AAAAASUVORK5CYII='; // /_layouts/15/images/MDNExpanded.png
export const GROUP_IMG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAC9SURBVDhPY2CgNXh1qEkdiJ8D8X90TNBuJM0V6IpBhoHFgIxebKYTIwYzAMNpxGhGdsFwNoBgNEFjAWsYgOSKiorMgPgbEP/Hgj8AxXpB0Yg1gQAldYuLix8/efLkzn8s4O7du9eAan7iM+DV/v37z546der/jx8/sJkBdhVOA5qbm08ePnwYrOjQoUOkGwDU+AFowLmjR4/idwGukAYaYAkMgxfPnj27h816kDg4DPABoAI/IP6DIxZA4l0AOd9H3QXl5+cAAAAASUVORK5CYII='; // /_layouts/15/Images/EMMGroup.png
export const TERMSET_IMG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACaSURBVDhPrZLRCcAgDERdpZMIjuQA7uWH4CqdxMY0EQtNjKWB0A/77sxF55SKMTalk8a61lqCFqsLiwKac84ZRUUBi7MoYHVmAfjfjzE6vJqZQfie0AcwBQVW8ATi7AR7zGGGNSE6Q2cyLSPIjRswjO7qKhcPDN2hK46w05wZMcEUIG+HrzzcrRsQBIJ5hS8C9fGAPmRwu/9RFxW6L8CM4Ry8AAAAAElFTkSuQmCC'; // /_layouts/15/Images/EMMTermSet.png
export const TERM_IMG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACzSURBVDhPY2AYNKCoqIgTiOcD8X8S8F6wB4Aa1IH4akNDw+mPHz++/E8EuHTp0jmQRSDNCcXFxa/XrVt3gAh9KEpgBvx/9OjRLVI1g9TDDYBp3rlz5//Kysr/IJoYgGEASPPatWsbQDQxAMOAbdu2gZ0FookBcAOePHlyhxgN6GqQY+Hdhg0bDpJqCNgAaDrQAnJuNDY2nvr06dMbYgw6e/bsabgBUEN4yEiJ2wdNViLfIQC3sTh2vtJcswAAAABJRU5ErkJggg==';

export class FieldTaxonomyRenderer extends BaseFieldRenderer {
  private isFieldMounted = false;
  private termsService: SPTermStorePickerService;
  private previousValues: IPickerTerms = [];
  private cancel: boolean = true;

  constructor(props: IFieldProps) {
    super(props);

    let unprocessedCurrentValue = [];
    if (props.FormFieldValue) {
      if (props.IsMulti) {
        unprocessedCurrentValue = props.FormFieldValue.results;
      } else {
        unprocessedCurrentValue = [props.FormFieldValue];
      }
    }

    this.state = {
      currentValue: [],
      unprocessedInitialValue: unprocessedCurrentValue,
      termSetAndTerms: null,
      loaded: false,
      openPanel: false,
      errorMessage: ''
    };

    this.onOpenPanel = this.onOpenPanel.bind(this);
    this.onClosePanel = this.onClosePanel.bind(this);
    this.onSave = this.onSave.bind(this);
    this.termsChanged = this.termsChanged.bind(this);
    this.termsFromPickerChanged = this.termsFromPickerChanged.bind(this);
    this.trySetValue = this.trySetValue.bind(this);
  }

  public componentDidMount() {
    this.isFieldMounted = true;
    if (!this.state.termSetAndTerms) {
      this.termsService = new SPTermStorePickerService(this.props);
      this.termsService.getAllTerms(this.props.TaxonomyTermSetId).then((response) => {
        if (response !== null) {
          // console.log(response);
          // console.log(this.props.FormFieldValue);
          // debugger;

          if (this.isFieldMounted) {
            this.setState({
              termSetAndTerms: response,
              loaded: true
            }, () => {
              this.trySetProcessedValue();
            });
          }
        }
      });
    } else {
      this.trySetProcessedValue();
    }
  }

  public componentWillUnmount() {
    this.isFieldMounted = false;
  }

  protected renderNewForm(props: IFieldProps) {
    return this.renderNewOrEditForm(props);
  }

  protected renderEditForm(props: IFieldProps) {
    return this.renderNewOrEditForm(props);
  }

  protected renderDispForm() {
    if (this.props.FormFieldValue && this.props.FormFieldValue.length > 0) {
      return (
        <React.Fragment>
          {this.props.FormFieldValue.map(v => <div key={v.key}>{v.name}</div>)}
        </React.Fragment>
      );
    }
    return null;
  }

  protected renderNewOrEditForm(props: IFieldProps): JSX.Element {
    return (
      <div>
        {
          /* Show spinner in the panel while retrieving terms */
          !this.state.loaded ?
            <Spinner type={SpinnerType.normal} /> :
            <table className={`termFieldTable`}>
              <tbody>
                <tr>
                  <td className={`termFieldRowPicker`}>
                    <TermPicker
                      disabled={false}
                      fieldProps={props}
                      allTerms={this.state.termSetAndTerms ? this.state.termSetAndTerms.Terms : []}
                      value={this.state.currentValue}
                      isTermSetSelectable={false}
                      onChanged={this.termsFromPickerChanged}
                      allowMultipleSelections={props.IsMulti}
                      disabledTermIds={null}
                      disableChildrenOfDisabledParents={null} />
                  </td>
                  <td className={`termFieldRowIcon`}>
                    <IconButton iconProps={{ iconName: 'Tag' }} onClick={this.onOpenPanel} />
                  </td>
                </tr>
              </tbody>
            </table>
        }

        <FieldErrorMessage errorMessage={this.state.errorMessage} />

        <Panel
          isOpen={this.state.openPanel}
          hasCloseButton={true}
          onDismiss={this.onClosePanel}
          isLightDismiss={true}
          type={PanelType.medium}
          headerText={this.state.termSetAndTerms ? this.state.termSetAndTerms.Name : ''}
          onRenderFooterContent={() => {
            return (
              <div className={`actions`}>
                <PrimaryButton iconProps={{ iconName: 'Save' }} text='Save' value='Save' onClick={this.onSave} />
                <DefaultButton iconProps={{ iconName: 'Cancel' }} text='Cancel' value='Cancel' onClick={this.onClosePanel} />
              </div>
            );
          }}>
          {
            this.state.openPanel && this.state.loaded && this.state.termSetAndTerms && (
              <div key={this.state.termSetAndTerms.Id} >
                {/* <h3>{this.state.termSetAndTerms.Name}</h3> */}
                <TermParent
                  anchorId={props.TaxonomyAnchorId}
                  autoExpand={null}
                  termset={this.state.termSetAndTerms}
                  isTermSetSelectable={false}
                  activeNodes={this.state.currentValue}
                  disabledTermIds={null}
                  disableChildrenOfDisabledParents={null}
                  changedCallback={this.termsChanged}
                  multiSelection={props.IsMulti ? true : false} />
              </div>
            )
          }
        </Panel>
      </div >
    );
  }

  private trySetProcessedValue() {
    if (!this.state.termSetAndTerms) {
      return;
    }

    let currentValue = this.state.currentValue;
    if (this.state.unprocessedInitialValue) {
      currentValue = this.state.unprocessedInitialValue.reduce((prevResults, iv) => {
        let terms = this.state.termSetAndTerms.Terms.filter(t => t.key === iv.TermGuid || t.key === iv.key);
        if (terms && terms.length > 0) {
          prevResults.push(terms[0]);
        }
        return prevResults;
      }, []);
      if (currentValue.length === 0) {
        currentValue = this.props.FormFieldValue;
      }
    } else {
      currentValue = this.props.FormFieldValue;
    }

    if (this.isFieldMounted) {
      this.setState({ currentValue }, () => {
        this.trySetValue(this.state.currentValue);
      });
    }
  }

  private onOpenPanel(): void {
    // Store the current code value
    this.previousValues = [...this.state.currentValue];
    this.cancel = true;
    this.setState({ openPanel: true });
  }

  private onClosePanel(): void {
    let newState: any = {
      openPanel: false
    };

    // Check if the property has to be reset
    if (this.cancel) {
      newState.currentValue = this.previousValues;
    }
    this.setState(newState, () => {
      this.trySetValue(this.state.currentValue);
    });
  }

  private onSave(): void {
    this.cancel = false;
    this.onClosePanel();
  }

  private termsChanged(term: IPickerTerm, checked: boolean): void {
    let currentValue = this.state.currentValue;
    if (typeof term === 'undefined' || term === null) {
      return;
    }

    // Term item to add to the active nodes array
    const termItem = term;

    // Check if the term is checked or unchecked
    if (checked) {
      // Check if it is allowed to select multiple terms
      if (this.props.IsMulti) {
        // Add the checked term
        currentValue.push(termItem);
        // Filter out the duplicate terms
        // activeNodes = uniqBy(activeNodes, 'key');
      } else {
        // Only store the current selected item
        currentValue = [termItem];
      }
    } else {
      // Remove the term from the list of active nodes
      currentValue = currentValue.filter(item => item.key !== term.key);
    }
    // Sort all active nodes
    // activeNodes = sortBy(activeNodes, 'path');
    // Update the current state
    this.setState({
      currentValue: currentValue
    });
  }

  private termsFromPickerChanged(terms: IPickerTerms) {
    this.setState({ currentValue: terms }, () => {
      this.trySetValue(terms);
    });
  }

  private trySetValue(terms: IPickerTerms) {
    let toSet = [];
    if (terms) {
      for (const term of terms) {
        toSet.push(`-1;#${term.name}|${term.key}`);
      }
    }
    // this.trySetChangedValue(toSet.join(';#'));
    this.trySetChangedValue(terms);
  }
}
