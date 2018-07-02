import * as React from 'react';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import { ITermParentProps, ITermParentState } from './interfaces';
import { EXPANDED_IMG, COLLAPSED_IMG, TERMSET_IMG, TERM_IMG } from './TaxonomyPicker';
import * as TermService from './ISPTermStorePickerService';
import Term from './Term';

import styles from './TaxonomyPicker.module.scss';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';

/**
 * Term Parent component, represents termset or term if anchorId
 */
export default class TermParent extends React.Component<ITermParentProps, ITermParentState> {

  private _terms: TermService.ITerm[];
  private _anchorName: string;

  constructor(props: ITermParentProps) {
    super(props);

    this._terms = this.props.termset.Terms;
    this.state = {
      loaded: true,
      expanded: true
    };
    this._handleClick = this._handleClick.bind(this);
  }

  /**
   * componentWillMount
   */
  public componentWillMount() {
    // fix term depth if anchroid for rendering
    if (this.props.anchorId) {
      const anchorTerm = this._terms.filter(t => t.Id.toLowerCase() === this.props.anchorId.toLowerCase()).shift();
      if (anchorTerm) {
        this._anchorName = anchorTerm.Name;
        let anchorTerms: TermService.ITerm[] = this._terms.filter(t => t.PathOfTerm.substring(0, anchorTerm.PathOfTerm.length) === anchorTerm.PathOfTerm && t.Id !== anchorTerm.Id);

        anchorTerms = anchorTerms.map(term => {
          term.PathDepth = term.PathDepth - anchorTerm.PathDepth;
          return term;
        });

        this._terms = anchorTerms;
      }
    }
  }

  /**
   * Default React render method
   */
  public render(): JSX.Element {
    // Specify the inline styling to show or hide the termsets
    const styleProps: React.CSSProperties = {
      display: this.state.expanded ? 'block' : 'none'
    };

    let termElm: JSX.Element = <div />;

    // Check if the terms have been loaded
    if (this.state.loaded) {
      if (this._terms.length > 0) {
        let disabledPaths = [];
        termElm = (
          <div style={styleProps}>
            {
              this._terms.map(term => {
                let disabled = false;
                if (this.props.disabledTermIds && this.props.disabledTermIds.length > 0) {
                  // Check if the current term ID exists in the disabled term IDs array
                  disabled = this.props.disabledTermIds.indexOf(term.Id) !== -1;
                  if (disabled) {
                    // Push paths to the disabled list
                    disabledPaths.push(term.PathOfTerm);
                  }
                }

                if (this.props.disableChildrenOfDisabledParents) {
                  // Check if parent is disabled
                  const parentPath = disabledPaths.filter(p => term.PathOfTerm.indexOf(p) !== -1);
                  disabled = parentPath && parentPath.length > 0;
                }

                return <Term key={term.Id} term={term} termset={this.props.termset.Id} activeNodes={this.props.activeNodes} changedCallback={this.props.changedCallback} multiSelection={this.props.multiSelection} disabled={disabled} />;
              })
            }
          </div>
        );
      } else {
        termElm = <div className={`${styles.listItem} ${styles.term}`}>{`TaxonomyPickerNoTerms`}</div>;
      }
    } else {
      termElm = <Spinner type={SpinnerType.normal} />;
    }

    return (
      <div>
        <div className={`${styles.listItem} ${styles.termset} ${(!this.props.anchorId && this.props.isTermSetSelectable) ? styles.termSetSelectable : ''}`} onClick={this._handleClick}>
          <img src={this.state.expanded ? EXPANDED_IMG : COLLAPSED_IMG} alt={`TaxonomyPickerExpandTitle`} title={`TaxonomyPickerExpandTitle`} />
          <img src={this.props.anchorId ? TERM_IMG : TERMSET_IMG} alt={`TaxonomyPickerMenuTermSet`} title={`TaxonomyPickerMenuTermSet`} />
          {
            this.props.anchorId ?
              this._anchorName :
              this.props.termset.Name
          }
        </div>
        <div style={styleProps}>
          {termElm}
        </div>
      </div>
    );
  }

  /**
   * Handle the click event: collapse or expand
   */
  private _handleClick() {
    this.setState({
      expanded: !this.state.expanded
    });
  }
}
