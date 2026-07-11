import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { getGraphClient } from '../../../common/services/graphService';
import { IPerson } from '../../../common/models';

/** A directory entry: an IPerson plus the resolved company domain and avatar gradient. */
export interface IDirectoryPerson extends IPerson {
  domain: string; // email/UPN domain, lower-cased (used by the company filter chips)
  gradient?: string;
}

/** Filter configuration sourced from the central Admin settings. */
export interface IDirectoryConfig {
  domains: string[]; // allowed company domains (lower-cased); empty = show all
  exclude: string[]; // case-insensitive substrings matched on name / email to hide accounts
  pageSize: number; // how many people to show in the default (no-search) view
}

/** Shape of a Microsoft Graph user object (subset selected). */
interface IGraphUser {
  displayName?: string;
  jobTitle?: string;
  department?: string;
  mail?: string;
  userPrincipalName?: string;
  businessPhones?: string[];
  officeLocation?: string;
}

interface IGraphUsersResponse {
  value?: IGraphUser[];
}

/** Extracts the lower-cased domain from an email / UPN. */
function domainOf(address: string | undefined): string {
  if (!address) {
    return '';
  }
  const at: number = address.lastIndexOf('@');
  return at === -1 ? '' : address.substring(at + 1).toLowerCase();
}

function mapUser(u: IGraphUser): IDirectoryPerson {
  const phones: string[] = u.businessPhones || [];
  const address: string | undefined = u.mail || u.userPrincipalName;
  return {
    displayName: u.displayName || '',
    jobTitle: u.jobTitle,
    department: u.department,
    email: address,
    phone: phones.length > 0 ? phones[0] : undefined,
    location: u.officeLocation,
    domain: domainOf(address)
  };
}

/**
 * True when the person should be hidden per the configured exclude substrings.
 * Matches across name, email, position (jobTitle) and department so accounts
 * like "Service Account" are filtered whichever field carries the term.
 */
function isExcluded(person: IDirectoryPerson, exclude: string[]): boolean {
  if (!exclude || exclude.length === 0) {
    return false;
  }
  const haystack: string = `${person.displayName} ${person.email || ''} ${person.jobTitle || ''} ${person.department || ''}`.toLowerCase();
  return exclude.some((term) => term && haystack.indexOf(term.toLowerCase()) !== -1);
}

/** True when the person's domain is allowed (empty allow-list = everyone). */
function inAllowedDomains(person: IDirectoryPerson, domains: string[]): boolean {
  if (!domains || domains.length === 0) {
    return true;
  }
  return domains.indexOf(person.domain) !== -1;
}

/** Applies the domain allow-list and exclude rules to a raw user list. */
function applyConfig(users: IGraphUser[], config: IDirectoryConfig): IDirectoryPerson[] {
  return users
    .filter((u) => !!u.displayName)
    .map(mapUser)
    .filter((p) => inAllowedDomains(p, config.domains) && !isExcluded(p, config.exclude));
}

const PERSON_SELECT: string = 'displayName,jobTitle,department,mail,userPrincipalName,businessPhones,officeLocation';

export const FALLBACK_PEOPLE: IDirectoryPerson[] = [];

/** Escapes a user-supplied term for use inside a Graph $search "field:term" clause. */
function sanitizeTerm(term: string): string {
  return term.replace(/["\\]/g, ' ').trim();
}

/**
 * Default (no-search) listing. Pulls a generous page from Graph, applies the
 * domain allow-list and exclude rules, sorts by name and caps to the page size.
 * Falls back to the design dataset on error or empty result. Requires
 * User.Read.All.
 */
export async function getPeople(context: WebPartContext, config: IDirectoryConfig): Promise<IDirectoryPerson[]> {
  try {
    const client: MSGraphClientV3 = await getGraphClient(context);
    const response: IGraphUsersResponse = await client
      .api('/users')
      .select(PERSON_SELECT)
      .top(200)
      .get();

    const users: IGraphUser[] = response && response.value ? response.value : [];
    const people: IDirectoryPerson[] = applyConfig(users, config).sort((a, b) =>
      a.displayName.localeCompare(b.displayName)
    );

    if (people.length === 0) {
      return [];
    }
    const limit: number = config.pageSize > 0 ? config.pageSize : people.length;
    return people.slice(0, limit);
  } catch {
    return [];
  }
}

/**
 * Full-directory search. Uses Microsoft Graph `$search` (with eventual
 * consistency) so it queries the whole tenant directory rather than a single
 * fetched page, then applies the same domain / exclude filtering. Returns up to
 * 100 matches. Requires User.Read.All.
 */
export async function searchPeople(
  context: WebPartContext,
  query: string,
  config: IDirectoryConfig
): Promise<IDirectoryPerson[]> {
  const term: string = sanitizeTerm(query);
  if (term.length < 2) {
    return [];
  }
  try {
    const client: MSGraphClientV3 = await getGraphClient(context);
    const fields: string[] = ['displayName', 'mail', 'userPrincipalName', 'jobTitle', 'department', 'givenName', 'surname'];
    const search: string = fields.map((f) => `"${f}:${term}"`).join(' OR ');
    const response: IGraphUsersResponse = await client
      .api('/users')
      .header('ConsistencyLevel', 'eventual')
      .query({ $count: 'true' })
      .search(search)
      .select(PERSON_SELECT)
      .top(100)
      .get();

    const users: IGraphUser[] = response && response.value ? response.value : [];
    return applyConfig(users, config).sort((a, b) => a.displayName.localeCompare(b.displayName));
  } catch {
    return [];
  }
}
