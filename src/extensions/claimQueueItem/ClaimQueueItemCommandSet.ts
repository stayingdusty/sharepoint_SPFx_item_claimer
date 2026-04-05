import { Log } from '@microsoft/sp-core-library';
import { Dialog } from '@microsoft/sp-dialog';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'ClaimQueueItemCommandSetStrings';

const LOG_SOURCE = 'ClaimQueueItemFieldCustomizer';

export interface IClaimQueueItemFieldCustomizerProperties {
  claimFieldInternalName?: string;
}

interface IClaimableListItem {
  Id: number;
  Title?: string;
  '@odata.etag'?: string;
  [key: string]: unknown;
}

interface ISharePointItemResponse<T> {
  d?: T;
}

export default class ClaimQueueItemFieldCustomizer extends BaseFieldCustomizer<IClaimQueueItemFieldCustomizerProperties> {
  private readonly _cellDisposers = new WeakMap<HTMLDivElement, () => void>();

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized for field "${this._getClaimFieldInternalName()}"`);
    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    this._disposeCell(event.domElement);
    event.domElement.innerHTML = '';
    event.domElement.style.whiteSpace = 'normal';

    const assignedDisplayValue = this._getAssignedDisplayValue(event.fieldValue);

    if (assignedDisplayValue) {
      this._renderTextValue(event.domElement, assignedDisplayValue);
      return;
    }

    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = strings.CommandLabel;
    button.setAttribute('aria-label', strings.CommandLabel);
    button.style.padding = '2px 8px';
    button.style.border = '1px solid #0078d4';
    button.style.borderRadius = '4px';
    button.style.background = '#ffffff';
    button.style.color = '#0078d4';
    button.style.cursor = 'pointer';
    button.style.fontSize = '12px';
    button.style.lineHeight = '18px';

    const onClick = async (clickEvent: MouseEvent): Promise<void> => {
      clickEvent.preventDefault();
      clickEvent.stopPropagation();

      button.disabled = true;
      button.style.cursor = 'default';
      button.textContent = strings.ClaimingLabel;

      try {
        const claimResult = await this._claimItem(this._getItemId(event));

        if (claimResult.status === 'success') {
          button.textContent = strings.ClaimedLabel;
          await Dialog.alert(strings.SuccessMessage);
          window.location.reload();
          return;
        }

        await Dialog.alert(
          claimResult.assignedUserLabel
            ? `Already taken by ${claimResult.assignedUserLabel}.`
            : strings.AlreadyTakenMessage
        );
        window.location.reload();
      } catch (error) {
        const message: string = error instanceof Error ? error.message : strings.UnexpectedErrorMessage;
        button.disabled = false;
        button.style.cursor = 'pointer';
        button.textContent = strings.CommandLabel;
        await Dialog.alert(`${strings.UnexpectedErrorMessage}\n\n${message}`);
      }
    };

    button.addEventListener('click', onClick);
    event.domElement.appendChild(button);

    this._cellDisposers.set(event.domElement, () => {
      button.removeEventListener('click', onClick);
    });
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    this._disposeCell(event.domElement);
    super.onDisposeCell(event);
  }

  private _disposeCell(domElement: HTMLDivElement): void {
    const disposer = this._cellDisposers.get(domElement);

    if (disposer) {
      disposer();
      this._cellDisposers.delete(domElement);
    }
  }

  private _renderTextValue(domElement: HTMLDivElement, value: string): void {
    domElement.innerHTML = '';

    const span = document.createElement('span');
    span.textContent = value;
    domElement.appendChild(span);
  }

  private _getItemId(event: IFieldCustomizerCellEventParameters): number {
    const rawValue: unknown = event.listItem.getValueByName('ID') || event.listItem.getValueByName('Id');
    const itemId = Number(rawValue);

    if (!itemId) {
      throw new Error('Could not resolve the SharePoint list item ID.');
    }

    return itemId;
  }

  private _getClaimFieldInternalName(): string {
    return this.context.field.internalName || this.properties.claimFieldInternalName?.trim() || 'Assigned_To';
  }

  private _getAssignedDisplayValue(fieldValue: unknown): string | undefined {
    if (fieldValue === null || fieldValue === undefined) {
      return undefined;
    }

    if (typeof fieldValue === 'string') {
      const trimmedValue = fieldValue.trim();
      return trimmedValue || undefined;
    }

    if (typeof fieldValue === 'number') {
      return fieldValue > 0 ? String(fieldValue) : undefined;
    }

    if (Array.isArray(fieldValue)) {
      const values: string[] = fieldValue
        .map((entry) => this._getAssignedDisplayValue(entry))
        .filter((value): value is string => Boolean(value));

      return values.length > 0 ? values.join(', ') : undefined;
    }

    if (typeof fieldValue === 'object') {
      const candidate = fieldValue as Record<string, unknown>;
      const orderedKeys: string[] = [
        'title',
        'Title',
        'lookupValue',
        'LookupValue',
        'displayName',
        'name',
        'EMail',
        'email',
        'value'
      ];

      for (const key of orderedKeys) {
        const value = candidate[key];

        if (typeof value === 'string' && value.trim()) {
          return value.trim();
        }
      }
    }

    return undefined;
  }

  private async _claimItem(itemId: number): Promise<{ status: 'success' | 'alreadyTaken'; assignedUserLabel?: string }> {
    const currentUserId = Number(this.context.pageContext.legacyPageContext?.userId || 0);
    const claimFieldInternalName = this._getClaimFieldInternalName();

    if (!currentUserId) {
      throw new Error('Could not resolve the current SharePoint user ID.');
    }

    const itemResponse = await this._getJsonResponse(this._getReadItemUrl(itemId, claimFieldInternalName));

    if (!itemResponse.ok) {
      throw new Error(`Could not load the current item. Status ${itemResponse.status}.`);
    }

    const itemPayload = (await itemResponse.json()) as IClaimableListItem | ISharePointItemResponse<IClaimableListItem>;
    const item: IClaimableListItem = this._unwrapSharePointItem(itemPayload);
    const assignedUserLabel: string | undefined = this._getAssignedUserLabel(item, claimFieldInternalName);

    if (this._hasAssignee(item, claimFieldInternalName)) {
      return {
        status: 'alreadyTaken',
        assignedUserLabel
      };
    }

    const etag: string = itemResponse.headers.get('ETag') || item['@odata.etag'] || '*';
    const status = await this._updateClaimedItem(itemId, claimFieldInternalName, currentUserId, etag);

    return { status };
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
      throw new Error('This field customizer can only run from a SharePoint list view.');
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
