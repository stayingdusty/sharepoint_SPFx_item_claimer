# sharepoint_SPFx_item_claimer

An SPFx **Field Customizer** that adds an inline `Claim` button to the `Assigned_To` column in a SharePoint work-queue list and uses **ETag-based optimistic concurrency** to prevent accidental double-claims or item stealing during UI sync lag.

## ✅ Claim flow

1. A user sees a blank `Assigned_To` cell with an inline `Claim` button.
2. Clicking `Claim` fetches the latest item state from SharePoint.
3. The extension checks whether `Assigned_To` is already populated.
4. If blank, it sends a `MERGE` update with the current item `ETag` in `IF-MATCH`.
5. If the update succeeds, the user sees a success message.
6. If SharePoint returns `412 Precondition Failed` or the item is already assigned, the user sees an "already taken" message.

## 📁 Key files

- `src/extensions/claimQueueItem/ClaimQueueItemCommandSet.ts` — inline field rendering and REST/ETag claim logic
- `src/extensions/claimQueueItem/ClaimQueueItemCommandSet.manifest.json` — SPFx field customizer manifest
- `sharepoint/assets/elements.xml` — SharePoint packaging/provisioning assets
- `scripts/attach-field-customizer.ps1` — optional PowerShell helper to bind the field customizer to an existing field

## ⚙️ Configuration

The current solution is verified against a SharePoint list with:

```text
Site URL: https://forgeweldapps.sharepoint.com/sites/test_site
List title: test_list
Field internal name: Assigned_To
```

For existing SharePoint columns, the field customizer must be associated to the field **after** the app is deployed to the site. That one-time binding updates the field's:

- `ClientSideComponentId`
- `ClientSideComponentProperties`

Without that step, the inline button will not render even if the app package is installed.

## 🔧 Build locally

This SPFx 1.20 project must be built with a supported Node runtime. In this dev container, the verified working command is:

```bash
npx -y node@20 ./node_modules/gulp/bin/gulp.js bundle --ship && \
npx -y node@20 ./node_modules/gulp/bin/gulp.js package-solution --ship
```

If you're using `nvm` locally, Node `18.x` or `20.x` in the supported SPFx range also works.

The deployable package is created at:

```text
sharepoint/solution/sharepoint-spfx-item-claimer.sppkg
```

## 🚀 Deploy to SharePoint

### 1) Build the package

```bash
npx -y node@20 ./node_modules/gulp/bin/gulp.js bundle --ship && \
npx -y node@20 ./node_modules/gulp/bin/gulp.js package-solution --ship
```

### 2) Upload and deploy the app

1. Upload `sharepoint/solution/sharepoint-spfx-item-claimer.sppkg` to the SharePoint App Catalog.
2. Click **Deploy**.
3. In the target site, go to **Site contents** and add/update the app.
4. If SharePoint shows **Get it** or **Update available**, complete that step first.

### 3) Bind the field customizer to the existing field

#### Option A — Browser console ✅ verified working

Open the target SharePoint site, press **F12**, open the **Console**, and run:

```js
(async () => {
  const listTitle = "test_list";
  const fieldInternalName = "Assigned_To";
  const componentId = "fae2eec7-5401-4ea3-a4a8-9958dd98721f";
  const componentProperties = JSON.stringify({
    claimFieldInternalName: "Assigned_To"
  });

  const digestRes = await fetch("/sites/test_site/_api/contextinfo", {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" }
  });
  const digestJson = await digestRes.json();
  const digest = digestJson.FormDigestValue;

  const url = `/sites/test_site/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields/getbyinternalnameortitle('${fieldInternalName}')`;

  const res = await fetch(url, {
    method: "POST",
    headers: {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": digest,
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE"
    },
    body: JSON.stringify({
      ClientSideComponentId: componentId,
      ClientSideComponentProperties: componentProperties
    })
  });

  if (res.ok) {
    console.log("✅ Field customizer attached successfully.");
  } else {
    console.error("❌ Failed:", res.status, await res.text());
  }
})();
```

#### Option B — PowerShell helper

```powershell
pwsh ./scripts/attach-field-customizer.ps1 \
  -SiteUrl "https://forgeweldapps.sharepoint.com/sites/test_site" \
  -ListTitle "test_list" \
  -FieldInternalName "Assigned_To"
```

### 4) Refresh and test

- hard refresh the list page with `Ctrl+Shift+R`
- blank `Assigned_To` cells should now show the inline **`Claim`** button
- clicking the button should claim the item safely using ETag concurrency checks

