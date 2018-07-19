import * as React from 'react';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import { ITermParentProps, ITermParentState, IPickerTerm } from './interfaces';
import { EXPANDED_IMG, COLLAPSED_IMG, TERMSET_IMG, TERM_IMG } from './TaxonomyPicker';
import Term from './Term';

/**
 * Term Parent component, represents termset or term if anchorId
 */
export default class TermParent extends React.Component<ITermParentProps, ITermParentState> {

  private _terms: IPickerTerm[];
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

  public componentWillMount() {
    // fix term depth if anchroid for rendering
    if (this.props.anchorId) {
      const anchorTerm = this._terms.filter(t => t.key.toLowerCase() === this.props.anchorId.toLowerCase()).shift();
      if (anchorTerm) {
        this._anchorName = anchorTerm.name;
        let anchorTerms = this._terms.filter(t => t.path.substring(0, anchorTerm.path.length) === anchorTerm.path && t.key !== anchorTerm.key);

        anchorTerms = anchorTerms.map(term => {
          // term.path = (parseInt(term.path) - parseInt(anchorTerm.path)).toString();
          term.pathDepth = term.pathDepth = anchorTerm.pathDepth;
          return term;
        });

        this._terms = anchorTerms;
      }
    }
  }

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
                  disabled = this.props.disabledTermIds.indexOf(term.key) !== -1;
                  if (disabled) {
                    // Push paths to the disabled list
                    disabledPaths.push(term.path);
                  }
                }

                if (this.props.disableChildrenOfDisabledParents) {
                  // Check if parent is disabled
                  const parentPath = disabledPaths.filter(p => term.path.indexOf(p) !== -1);
                  disabled = parentPath && parentPath.length > 0;
                }

                return <Term key={term.key} term={term} termset={this.props.termset.Id} activeNodes={this.props.activeNodes}
                  changedCallback={this.props.changedCallback} multiSelection={this.props.multiSelection} disabled={disabled} />;
              })
            }
          </div>
        );
      } else {
        termElm = <div className={`${'listItem'} ${'term'}`}>{`No terms found`}</div>;
      }
    } else {
      termElm = <Spinner type={SpinnerType.normal} />;
    }

    return (
      <div>
        <div className={`${'listItem'} ${'termset'} ${(!this.props.anchorId && this.props.isTermSetSelectable) ? 'termSetSelectable' : ''}`} onClick={this._handleClick}>
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
