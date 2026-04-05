# sharepoint_SPFx_item_claimer

An SPFx **ListView Command Set** that adds a `Claim` button to a SharePoint work-queue list and uses **ETag-based optimistic concurrency** to prevent accidental double-claims or item stealing during UI sync lag.

## ✅ Claim flow

1. A user selects a queue item and clicks `Claim`.
2. The extension fetches the latest item state from SharePoint.
3. It checks whether `Assigned_To` is already populated.
4. If blank, it sends a `MERGE` update with the current item `ETag` in `IF-MATCH`.
5. If the update succeeds, the user sees a success message.
6. If SharePoint returns `412 Precondition Failed` or the item is already assigned, the user sees an "already taken" message.

## 📁 Key files

- `src/extensions/claimQueueItem/ClaimQueueItemCommandSet.ts` — command logic and REST/ETag handling
- `src/extensions/claimQueueItem/ClaimQueueItemCommandSet.manifest.json` — command definition
- `sharepoint/assets/elements.xml` — registers the command for generic SharePoint lists

## ⚙️ Configuration

By default, the command now tries the most common SharePoint person-field internal names automatically, including:

```text
AssignedTo
Assigned_x0020_To
Assigned_To
```

If your list uses a different **Person or Group** field, set `claimFieldInternalName` in `sharepoint/assets/elements.xml` to that internal name and then repackage the solution.

## 🔧 Build locally

Use the supported Node version first:

```bash
source /usr/local/share/nvm/nvm.sh
nvm use 18.20.4
npm install
npx gulp bundle
npm run package
```

The deployable package is created at:

```text
sharepoint/solution/sharepoint-spfx-item-claimer.sppkg
```

## 🚀 Deploy to SharePoint

1. Upload the `.sppkg` to your SharePoint App Catalog.
2. Deploy the app.
3. Add the app to the site that hosts the work queue list.
4. Open the list and use the new `Claim` command from the command bar or context menu.
