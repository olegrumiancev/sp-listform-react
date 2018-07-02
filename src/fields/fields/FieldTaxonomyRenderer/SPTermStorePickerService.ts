/**
 * DISCLAIMER
 *
 * As there is not yet an OData end-point for managed metadata, this service makes use of the ProcessQuery end-points.
 * The service will get updated once the APIs are in place for managing managed metadata.
 */

import * as Taxonomy from '@pnp/sp-taxonomy';
import { IPickerTerms, IPickerTerm } from './interfaces';
import * as TermService from './ISPTermStorePickerService';
import { IFieldProps } from '../../interfaces';

/**
 * Service implementation to manage term stores in SharePoint
 */
export default class SPTermStorePickerService {
  // private taxonomySession: string;
  // private formDigest: string;
  // private clientServiceUrl: string;

  /**
   * Service constructor
   */
  constructor(private props: IFieldProps) {
    // this.clientServiceUrl = this.context.pageContext.web.absoluteUrl + '/_vti_bin/client.svc/ProcessQuery';
  }

  /**
   * Gets the collection of term stores in the current SharePoint env
   */
  public getTermStores(): Promise<TermService.ITermStore[]> {
    let result = new Promise<TermService.ITermStore[]>(async (resolve) => {
      let termStores: TermService.ITermStore[] = [];
      let siteCollTermStore = await Taxonomy.taxonomy.getDefaultSiteCollectionTermStore().get();
      let tStore = {
        Id: siteCollTermStore.Id,
        Name: siteCollTermStore.Name
      } as TermService.ITermStore;
      let groupSiteColl = await siteCollTermStore.getSiteCollectionGroup(false).get();
      if (groupSiteColl) {
        tStore.Groups = {
          _Child_Items_: [],
          _ObjectType_: 'Groups'
        };

        let transformedTermSets: TermService.ITermSet[] = [];
        tStore.Groups._Child_Items_.push({
          Id: groupSiteColl.Id,
          IsSystemGroup: groupSiteColl.IsSystemGroup,
          Name: groupSiteColl.Name,
          TermSets: {
            _Child_Items_: transformedTermSets
          } as TermService.ITermSets,
          _ObjectType_: 'Group',
          _ObjectIdentity_: ''
        });

        let termSets: (Taxonomy.ITermSetData & Taxonomy.ITermSet)[] = await groupSiteColl.termSets.get();
        if (termSets) {
          transformedTermSets.push(...termSets.map(ts => {
            return {
              Id: ts.Id,
              Description: ts.Description,
              Name: ts.Name,
              Names: ts.Names,
              _ObjectType_: 'TermSet',
              _ObjectIdentity_: ''
            } as TermService.ITermSet;
          }));
        }
      }

      termStores.push(tStore);
      resolve(termStores);
    });
    return result;
  }

  /**
   * Gets the current term set
   */
  public async getTermSet(): Promise<TermService.ITermSet> {
    const termStore = await this.getTermStores();
    return this.getTermSetId(termStore, this.props.TaxonomyTermSetId);
  }

  /**
   * Retrieve all terms for the given term set
   * @param termset
   */
  public async getAllTerms(termsetId: string): Promise<IPickerTerm[]> {
    // let termsetId: string = termset;
    let result = new Promise<IPickerTerm[]>(async resolve => {
      let retrievedTerms: (Taxonomy.ITermData & Taxonomy.ITerm)[] = await Taxonomy.taxonomy.getDefaultSiteCollectionTermStore().getTermSetById(termsetId).terms.get();
      if (retrievedTerms) {
        let toReturn = retrievedTerms.map(rt => {
          return {
            key: rt.Id,
            name: rt.Name,
            path: rt.PathOfTerm,
            termSet: termsetId
          } as IPickerTerm;
        });
        resolve(toReturn);
      }
      resolve(null);
    });
    return result;
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
  private getTermSetId(termstore: TermService.ITermStore[], termsetName: string): TermService.ITermSet {
    if (termstore && termstore.length > 0 && termsetName) {
      // Get the first term store
      const ts = termstore[0];
      // Check if the term store contains groups
      if (ts.Groups && ts.Groups._Child_Items_) {
        for (const group of ts.Groups._Child_Items_) {
          // Check if the group contains term sets
          if (group.TermSets && group.TermSets._Child_Items_) {
            for (const termSet of group.TermSets._Child_Items_) {
              // Check if the term set is found
              if (termSet.Name === termsetName) {
                return termSet;
              }
            }
          }
        }
      }
    }

    return null;
  }

  /**
   * Searches terms for the given term set
   * @param searchText
   * @param termsetId
   */
  private searchTermsByTermSet(searchText: string, termSetId: string, termLimit: number = 10): Promise<IPickerTerm[]> {
    return new Promise<IPickerTerm[]>(async resolve => {
      let allTerms = await this.getAllTerms(termSetId);
      let filteredTerms = allTerms.filter(t => t.name.match(new RegExp(searchText, 'gi')));
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
  private _sortTerms(a: TermService.ITerm, b: TermService.ITerm) {
    if (a.PathOfTerm < b.PathOfTerm) {
      return -1;
    }
    if (a.PathOfTerm > b.PathOfTerm) {
      return 1;
    }
    return 0;
  }
}
