import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import '@pnp/sp/security';
import '@pnp/sp/fields';
import { PermissionKind } from '@pnp/sp/security';

/** A single content record shown in a panel's list. */
export interface IContentRow {
  Id: number;
  Title: string;
  [key: string]: unknown;
}

/**
 * True when the current user may add items to the given list — the gate for the
 * contributor dashboard. Falls back to a web-level AddListItems check, then false.
 */
export async function canContribute(context: WebPartContext, listTitle: string): Promise<boolean> {
  const sp = getSP(context);
  try {
    const perms = await sp.web.lists.getByTitle(listTitle).getCurrentUserEffectivePermissions();
    return sp.web.hasPermissions(perms, PermissionKind.AddListItems);
  } catch {
    try {
      const perms = await sp.web.getCurrentUserEffectivePermissions();
      return sp.web.hasPermissions(perms, PermissionKind.AddListItems);
    } catch {
      return false;
    }
  }
}

/** True when the named list exists on the current site. */
export async function listExists(context: WebPartContext, listTitle: string): Promise<boolean> {
  try {
    await getSP(context).web.lists.getByTitle(listTitle).select('Id')();
    return true;
  } catch {
    return false;
  }
}

/** Returns the choices configured on a choice column, or [] if unavailable. */
export async function getChoices(context: WebPartContext, listTitle: string, fieldName: string): Promise<string[]> {
  try {
    const field: { Choices?: string[] } = await getSP(context)
      .web.lists.getByTitle(listTitle)
      .fields.getByInternalNameOrTitle(fieldName)
      .select('Choices')();
    return field.Choices || [];
  } catch {
    return [];
  }
}

/** Reads the most recent rows for a content list. */
export async function getItems(
  context: WebPartContext,
  listTitle: string,
  selectFields: string[],
  orderByField: string,
  top: number = 50
): Promise<IContentRow[]> {
  try {
    const fields: string[] = selectFields.indexOf('Id') === -1 ? ['Id', ...selectFields] : selectFields;
    return (await getSP(context)
      .web.lists.getByTitle(listTitle)
      .items.select(...fields)
      .orderBy(orderByField, false)
      .top(top)()) as IContentRow[];
  } catch {
    return [];
  }
}

export async function addItem(context: WebPartContext, listTitle: string, data: Record<string, unknown>): Promise<void> {
  await getSP(context).web.lists.getByTitle(listTitle).items.add(data);
}

export async function updateItem(
  context: WebPartContext,
  listTitle: string,
  id: number,
  data: Record<string, unknown>
): Promise<void> {
  await getSP(context).web.lists.getByTitle(listTitle).items.getById(id).update(data);
}

export async function deleteItem(context: WebPartContext, listTitle: string, id: number): Promise<void> {
  await getSP(context).web.lists.getByTitle(listTitle).items.getById(id).delete();
}
