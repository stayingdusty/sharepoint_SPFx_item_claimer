import { Log } from '@microsoft/sp-core-library';
import { Dialog } from '@microsoft/sp-dialog';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  IListViewCommandSetListViewUpdatedParameters,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'ClaimQueueItemCommandSetStrings';

const LOG_SOURCE = 'ClaimQueueItemCommandSet';

export interface IClaimQueueItemCommandSetProperties {
  claimFieldInternalName?: string;
}

interface IClaimableListItem {
  Id: number;
  Title?: string;
  '@odata.etag'?: string;
  [key: string]: unknown;
}

interface IClaimFieldDefinition {
  InternalName?: string;
  Title?: string;
  TypeAsString?: string;
  Hidden?: boolean;
  ReadOnlyField?: boolean;
}

interface IClaimFieldResponse {
  value?: IClaimFieldDefinition[];
  d?: {
    results?: IClaimFieldDefinition[];
  };
}

interface ISharePointItemResponse<T> {
  d?: T;
}

export default class ClaimQueueItemCommandSet extends BaseListViewCommandSet<IClaimQueueItemCommandSetProperties> {
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ClaimQueueItemCommandSet');
    return Promise.resolve();
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const claimCommand: Command | undefined = this.tryGetCommand('CLAIM_ITEM');

    if (claimCommand) {
      claimCommand.visible = event.selectedRows?.length === 1;
    }
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'CLAIM_ITEM':
        await this._claimSelectedItem(event.selectedRows[0]);
        break;
      default:
        throw new Error(`Unknown command: ${event.itemId}`);
    }
  }

  private async _claimSelectedItem(selectedRow: RowAccessor): Promise<void> {
    const itemId = this._getSelectedItemId(selectedRow);
    const currentUserId = Number(this.context.pageContext.legacyPageContext?.userId || 0);

    if (!currentUserId) {
      throw new Error('Could not resolve the current SharePoint user ID.');
    }

    try {
      const claimFieldInternalName: string = await this._resolveClaimFieldInternalName(itemId);
      const itemResponse = await this._getJsonResponse(this._getReadItemUrl(itemId, claimFieldInternalName));

      if (!itemResponse.ok) {
        throw new Error(`Could not load the current item. Status ${itemResponse.status}.`);
      }

      const itemPayload = (await itemResponse.json()) as IClaimableListItem | ISharePointItemResponse<IClaimableListItem>;
      const item: IClaimableListItem = this._unwrapSharePointItem(itemPayload);
      const assignedUserLabel: string | undefined = this._getAssignedUserLabel(item, claimFieldInternalName);

      if (this._hasAssignee(item, claimFieldInternalName)) {
        await Dialog.alert(
          assignedUserLabel
            ? `Already taken by ${assignedUserLabel}.`
            : strings.AlreadyTakenMessage
        );
        return;
      }

      const etag: string = itemResponse.headers.get('ETag') || item['@odata.etag'] || '*';
      const claimResult = await this._updateClaimedItem(itemId, claimFieldInternalName, currentUserId, etag);

      if (claimResult === 'success') {
        await Dialog.alert(strings.SuccessMessage);
        window.location.reload();
        return;
      }

      await Dialog.alert(strings.AlreadyTakenMessage);
      return;
    } catch (error) {
      const message: string = error instanceof Error ? error.message : strings.UnexpectedErrorMessage;
      await Dialog.alert(`${strings.UnexpectedErrorMessage}\n\n${message}`);
    }
  }

  private _getSelectedItemId(selectedRow: RowAccessor): number {
    const rawValue: unknown = selectedRow.getValueByName('ID') || selectedRow.getValueByName('Id');
    const itemId = Number(rawValue);

    if (!itemId) {
      throw new Error('A SharePoint list item must be selected before it can be claimed.');
    }

    return itemId;
  }

  private _getReadItemUrl(itemId: number, claimFieldInternalName: string): string {
    const listId: string = this._getListId();
    const selectClause: string = [
      'Id',
      'Title',
      `${claimFieldInternalName}Id`,
      `${claimFieldInternalName}/Id`,
      `${claimFieldInternalName}/Title`,
      `${claimFieldInternalName}/EMail`
    ].join(',');

    return `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})?$select=${selectClause}&$expand=${claimFieldInternalName}`;
  }

  private _getUpdateItemUrl(itemId: number): string {
    const listId: string = this._getListId();
    return `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})`;
  }

  private async _resolveClaimFieldInternalName(itemId: number): Promise<string> {
    const configuredFieldName: string | undefined = this.properties.claimFieldInternalName?.trim();

    if (configuredFieldName) {
      return configuredFieldName;
    }

    const preferredFieldNames: string[] = this._getPreferredClaimFieldNames();
    const attemptedFieldNames: string[] = [];

    for (const fieldName of preferredFieldNames) {
      attemptedFieldNames.push(fieldName);

      if (await this._canReadClaimField(itemId, fieldName)) {
        return fieldName;
      }
    }

    const availableFields = await this._getAvailableClaimFields();
    const availableFieldNames: string[] = availableFields
      .map((field) => field.InternalName)
      .filter((value): value is string => Boolean(value));

    let namedMatch: IClaimFieldDefinition | undefined;

    for (const field of availableFields) {
      const fieldName = field.InternalName;

      if (fieldName && attemptedFieldNames.indexOf(fieldName) < 0) {
        attemptedFieldNames.push(fieldName);

        if (await this._canReadClaimField(itemId, fieldName)) {
          return fieldName;
        }
      }

      const searchText = `${field.InternalName || ''} ${field.Title || ''}`;

      if (!namedMatch && /assigned|claim|owner/i.test(searchText)) {
        namedMatch = field;
      }
    }

    if (namedMatch && namedMatch.InternalName) {
      return namedMatch.InternalName;
    }

    if (availableFieldNames.length === 1) {
      return availableFieldNames[0];
    }

    if (availableFieldNames.length > 1) {
      throw new Error(
        `Could not resolve the claim field automatically. Set claimFieldInternalName to one of these Person or Group fields: ${availableFieldNames.join(', ')}.`
      );
    }

    throw new Error(
      `Could not find a writable Person or Group column to store claims. Tried: ${attemptedFieldNames.join(', ')}.`
    );
  }

  private _getPreferredClaimFieldNames(): string[] {
    return [
      'Assigned_To',
      'AssignedTo',
      'Assigned_x0020_To'
    ].filter((value, index, array): value is string => Boolean(value) && array.indexOf(value) === index);
  }

  private async _updateClaimedItem(
    itemId: number,
    claimFieldInternalName: string,
    currentUserId: number,
    etag: string
  ): Promise<'success' | 'alreadyTaken'> {
    const updateResponse = await this.context.spHttpClient.post(
      this._getUpdateItemUrl(itemId),
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Content-Type': 'application/json;odata=nometadata',
          'IF-MATCH': etag,
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify({
          [`${claimFieldInternalName}Id`]: currentUserId
        })
      }
    );

    if (updateResponse.ok) {
      return 'success';
    }

    if (updateResponse.status === 412) {
      return 'alreadyTaken';
    }

    if (updateResponse.status === 400 || updateResponse.status === 406) {
      const verificationResult = await this._verifyClaimOutcome(itemId, claimFieldInternalName, currentUserId);

      if (verificationResult) {
        return verificationResult;
      }
    }

    throw new Error(`Claim update failed. Status ${updateResponse.status}.`);
  }

  private async _verifyClaimOutcome(
    itemId: number,
    claimFieldInternalName: string,
    currentUserId: number
  ): Promise<'success' | 'alreadyTaken' | undefined> {
    const verificationResponse = await this._getJsonResponse(this._getReadItemUrl(itemId, claimFieldInternalName));

    if (!verificationResponse.ok) {
      return undefined;
    }

    const verificationPayload = (await verificationResponse.json()) as IClaimableListItem | ISharePointItemResponse<IClaimableListItem>;
    const verificationItem: IClaimableListItem = this._unwrapSharePointItem(verificationPayload);
    const assignedUserId = this._getAssignedUserId(verificationItem, claimFieldInternalName);

    if (assignedUserId === currentUserId) {
      return 'success';
    }

    if (assignedUserId > 0) {
      return 'alreadyTaken';
    }

    return undefined;
  }

  private async _canReadClaimField(itemId: number, claimFieldInternalName: string): Promise<boolean> {
    const itemResponse = await this._getJsonResponse(this._getReadItemUrl(itemId, claimFieldInternalName));

    if (itemResponse.ok) {
      return true;
    }

    if (itemResponse.status === 400 || itemResponse.status === 404) {
      return false;
    }

    throw new Error(`Could not validate the claim field "${claimFieldInternalName}". Status ${itemResponse.status}.`);
  }

  private async _getAvailableClaimFields(): Promise<IClaimFieldDefinition[]> {
    const listId: string = this._getListId();
    const fieldsResponse = await this._getJsonResponse(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/fields?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField`
    );

    if (!fieldsResponse.ok) {
      throw new Error(`Could not load the list fields. Status ${fieldsResponse.status}.`);
    }

    const payload = (await fieldsResponse.json()) as IClaimFieldResponse;
    const fields: IClaimFieldDefinition[] = payload.value || payload.d?.results || [];

    return fields.filter((field) => {
      return Boolean(field.InternalName) && field.TypeAsString === 'User' && !field.Hidden && !field.ReadOnlyField;
    });
  }

  private async _getJsonResponse(url: string): Promise<SPHttpClientResponse> {
    const acceptValues: string[] = [
      'application/json;odata.metadata=none',
      'application/json;odata=nometadata',
      'application/json;odata=verbose'
    ];

    let lastResponse: SPHttpClientResponse | undefined;

    for (const acceptValue of acceptValues) {
      const response = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: acceptValue
          }
        }
      );

      lastResponse = response;

      if (response.ok || response.status !== 406) {
        return response;
      }
    }

    if (!lastResponse) {
      throw new Error('Could not issue the SharePoint request.');
    }

    return lastResponse;
  }

  private _unwrapSharePointItem<T>(payload: T | ISharePointItemResponse<T>): T {
    return (payload as ISharePointItemResponse<T>).d || (payload as T);
  }

  private _getListId(): string {
    const listId: string | undefined = this.context.pageContext.list?.id?.toString();

    if (!listId) {
      throw new Error('This command can only run from a SharePoint list view.');
    }

    return listId;
  }

  private _hasAssignee(item: IClaimableListItem, claimFieldInternalName: string): boolean {
    return this._getAssignedUserId(item, claimFieldInternalName) > 0;
  }

  private _getAssignedUserId(item: IClaimableListItem, claimFieldInternalName: string): number {
    const assigneeId: unknown = item[`${claimFieldInternalName}Id`];

    if (typeof assigneeId === 'number') {
      return assigneeId;
    }

    if (typeof assigneeId === 'string') {
      const parsedValue = Number(assigneeId.trim());
      return isNaN(parsedValue) ? 0 : parsedValue;
    }

    return 0;
  }

  private _getAssignedUserLabel(item: IClaimableListItem, claimFieldInternalName: string): string | undefined {
    const assignee = item[claimFieldInternalName] as { Title?: string; EMail?: string } | undefined;
    return assignee?.Title || assignee?.EMail;
  }
}
