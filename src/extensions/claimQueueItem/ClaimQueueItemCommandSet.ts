import { Log } from '@microsoft/sp-core-library';
import { Dialog } from '@microsoft/sp-dialog';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  IListViewCommandSetListViewUpdatedParameters,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient } from '@microsoft/sp-http';

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
    const claimFieldInternalName: string = this.properties.claimFieldInternalName || 'Assigned_To';
    const currentUserId = Number(this.context.pageContext.legacyPageContext?.userId || 0);

    if (!currentUserId) {
      throw new Error('Could not resolve the current SharePoint user ID.');
    }

    try {
      const itemResponse = await this.context.spHttpClient.get(
        this._getReadItemUrl(itemId, claimFieldInternalName),
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      );

      if (!itemResponse.ok) {
        throw new Error(`Could not load the current item. Status ${itemResponse.status}.`);
      }

      const item: IClaimableListItem = (await itemResponse.json()) as IClaimableListItem;
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
      const updateResponse = await this.context.spHttpClient.post(
        this._getUpdateItemUrl(itemId),
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
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
        await Dialog.alert(strings.SuccessMessage);
        window.location.reload();
        return;
      }

      if (updateResponse.status === 412) {
        await Dialog.alert(strings.AlreadyTakenMessage);
        return;
      }

      throw new Error(`Claim update failed. Status ${updateResponse.status}.`);
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

  private _getListId(): string {
    const listId: string | undefined = this.context.pageContext.list?.id?.toString();

    if (!listId) {
      throw new Error('This command can only run from a SharePoint list view.');
    }

    return listId;
  }

  private _hasAssignee(item: IClaimableListItem, claimFieldInternalName: string): boolean {
    const assigneeId: unknown = item[`${claimFieldInternalName}Id`];

    if (typeof assigneeId === 'number') {
      return assigneeId > 0;
    }

    if (typeof assigneeId === 'string') {
      return assigneeId.trim().length > 0 && assigneeId !== '0';
    }

    return Boolean(assigneeId);
  }

  private _getAssignedUserLabel(item: IClaimableListItem, claimFieldInternalName: string): string | undefined {
    const assignee = item[claimFieldInternalName] as { Title?: string; EMail?: string } | undefined;
    return assignee?.Title || assignee?.EMail;
  }
}
