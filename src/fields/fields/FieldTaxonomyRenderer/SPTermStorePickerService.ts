import * as Taxonomy from '@pnp/sp-taxonomy';
import { IPickerTerms, IPickerTerm, IPickerTermSet } from './interfaces';
import { IFieldProps } from '../../interfaces';
import { registerDefaultFontFaces } from '@uifabric/styling/lib';

export default class SPTermStorePickerService {
  constructor(private props: IFieldProps) {
    // this.clientServiceUrl = this.context.pageContext.web.absoluteUrl + '/_vti_bin/client.svc/ProcessQuery';
  }

  public async getAllTerms(termsetId: string): Promise<IPickerTermSet> {
    let termSet: Taxonomy.ITermSetData & Taxonomy.ITermSet =
      await Taxonomy.taxonomy.getDefaultSiteCollectionTermStore().getTermSetById(termsetId)
        .usingCaching()
        .get();
    let result: IPickerTermSet = null;
    if (termSet) {
      result = {
        Id: termsetId,
        Name: termSet.Name,
        Description: termSet.Description,
        CustomSortOrder: termSet.CustomSortOrder
      };

      let retrievedTerms: any[] =
        await Taxonomy.taxonomy.getDefaultSiteCollectionTermStore().getTermSetById(termsetId)
          .terms
          .select('Id', 'Name', 'PathOfTerm', 'IsDeprecated', 'Parent', 'CustomSortOrder')
          .usingCaching()
          .get();
      if (retrievedTerms) {
        // console.log(retrievedTerms);
        result.Terms = retrievedTerms.map(rt => {
          return {
            key: this.cleanGuid(rt.Id),
            parentId: rt.Parent ? this.cleanGuid(rt.Parent.Id) : null,
            customSortOrder: rt.CustomSortOrder ? rt.CustomSortOrder : null,
            name: rt.Name,
            path: rt.PathOfTerm,
            pathDepth: rt.PathOfTerm ? rt.PathOfTerm.split(';').length : 1,
            termSet: termsetId,
            termSetName: result.Name,
            isDeprecated: rt.IsDeprecated
          } as IPickerTerm;
        });
        // result.Terms = result.Terms.sort(this.sortTerms);
        // console.log(result.Terms);
      }
    }

    return this.sortTermsInTermSetByHierarchy(result);
  }

  /**
   * Retrieve all terms that starts with the searchText
   * @param searchText
   */
  public searchTermsByName(searchText: string): Promise<IPickerTerm[]> {
    return this.searchTermsByTermSet(searchText, this.props.TaxonomyTermSetId);
  }

  /**
   * Clean the Guid from the Web Service response
   * @param guid
   */
  public cleanGuid(guid: string): string {
    if (guid !== undefined) {
      return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
    } else {
      return '';
    }
  }

  /**
   * Get the term set ID by its name
   * @param termstore
   * @param termset
   */
  // private getTermSetId(termstore: TermService.ITermStore[], termsetName: string): TermService.ITermSet {
  //   if (termstore && termstore.length > 0 && termsetName) {
  //     // Get the first term store
  //     const ts = termstore[0];
  //     // Check if the term store contains groups
  //     if (ts.Groups && ts.Groups._Child_Items_) {
  //       for (const group of ts.Groups._Child_Items_) {
  //         // Check if the group contains term sets
  //         if (group.TermSets && group.TermSets._Child_Items_) {
  //           for (const termSet of group.TermSets._Child_Items_) {
  //             // Check if the term set is found
  //             if (termSet.Name === termsetName) {
  //               return termSet;
  //             }
  //           }
  //         }
  //       }
  //     }
  //   }

  //   return null;
  // }

  /**
   * Searches terms for the given term set
   * @param searchText
   * @param termsetId
   */
  private searchTermsByTermSet(searchText: string, termSetId: string, termLimit: number = 10): Promise<IPickerTerm[]> {
    return new Promise<IPickerTerm[]>(async resolve => {
      let allTerms = await this.getAllTerms(termSetId);
      let filteredTerms = allTerms.Terms.filter(t => t.name.match(new RegExp(searchText, 'gi')));
      if (filteredTerms.length > termLimit) {
        filteredTerms = filteredTerms.slice(0, termLimit - 1);
      }
      resolve(filteredTerms);
    });
  }

  private isGuid(strGuid: string): boolean {
    return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(strGuid);
  }

  /**
   * Sort the terms by their path
   * @param a term 2
   * @param b term 2
   */
  // private sortTerms(a: IPickerTerm, b: IPickerTerm) {
  //   if (a.pathDepth < b.pathDepth) {
  //     return -1;
  //   }
  //   if (a.pathDepth > b.pathDepth) {
  //     return 1;
  //   }
  //   return 0;
  // }

  private sortTermsInTermSetByHierarchy(termSet: IPickerTermSet): IPickerTermSet {
    if (!termSet) {
      return termSet;
    }

    let mockRootLevelTerm: IPickerTerm = {
      key: null,
      name: '',
      customSortOrder: termSet.CustomSortOrder,
      isDeprecated: false,
      parentId: null,
      path: '',
      pathDepth: 0,
      termSet: ''
    };

    // debugger;
    let sortedTerms: IPickerTerm[] = [];
    let toProcess = this.getSortedTermsForAParent(termSet.Terms, mockRootLevelTerm);
    while (toProcess.length > 0) {
      let item = toProcess[0];
      sortedTerms.push(item);
      let currentItemChildren = this.getSortedTermsForAParent(termSet.Terms, item);
      if (currentItemChildren.length > 0) {
        toProcess.splice(0, 1, ...currentItemChildren);
      } else {
        toProcess.splice(0, 1);
      }
    }
    termSet.Terms = sortedTerms;
    return termSet;
  }

  private getSortedTermsForAParent(allUnsortedTerms: IPickerTerm[], parentPickerTerm: IPickerTerm): IPickerTerm[] {
    let results = allUnsortedTerms.filter(f => f.parentId === parentPickerTerm.key);
    if (parentPickerTerm.customSortOrder) {
      const orderedTermIds = parentPickerTerm.customSortOrder.split(':');
      let sortedResults = [];
      for (let id of orderedTermIds) {
        const res = results.filter(r => r.key === id);
        if (res.length > 0) {
          sortedResults.push(res[0]);
        }
      }
      return sortedResults;
    } else {
      return results;
    }
  }
}
