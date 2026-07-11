import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';

export interface ICurrentUser {
  displayName: string;
  initials: string;
}

/**
 * Resolves the signed-in user's display name and initials via Microsoft Graph
 * (requires the User.Read permission, requested in package-solution.json).
 * Falls back gracefully to the SPFx page context value if Graph is unavailable.
 */
export class GraphUserService {
  constructor(private readonly graphFactory: MSGraphClientFactory) {}

  public async getCurrentUser(fallbackDisplayName: string): Promise<ICurrentUser> {
    try {
      const client: MSGraphClientV3 = await this.graphFactory.getClient('3');
      const me: { displayName?: string; givenName?: string; surname?: string } = await client
        .api('/me')
        .select('displayName,givenName,surname')
        .get();

      const displayName: string = me.displayName || fallbackDisplayName;
      return {
        displayName,
        initials: GraphUserService.toInitials(me.givenName, me.surname, displayName)
      };
    } catch {
      return {
        displayName: fallbackDisplayName,
        initials: GraphUserService.toInitials(undefined, undefined, fallbackDisplayName)
      };
    }
  }

  private static toInitials(givenName?: string, surname?: string, fullName?: string): string {
    if (givenName && surname) {
      return (givenName.charAt(0) + surname.charAt(0)).toUpperCase();
    }

    const source: string = (fullName || '').trim();
    if (!source) {
      return '?';
    }

    const parts: string[] = source.split(/\s+/).filter((p) => p.length > 0);
    if (parts.length === 1) {
      return parts[0].substring(0, 2).toUpperCase();
    }

    return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
  }
}
