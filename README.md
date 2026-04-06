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

This solution intentionally does **not** package a SharePoint Feature `elements.xml`. SharePoint Online rejects the attempted `CustomAction` registration for `ClientSideExtension.ListViewFieldCustomizer`; the supported deployment path is to install the app package and then bind the field customizer to the target field.

## 🔧 Build locally

This SPFx 1.20 project must be built with a supported Node runtime. In this dev container, the verified working command is:

```bash
npm run package:ship
```

`npm run package:ship` automatically increments the patch version on each run before creating the `.sppkg`.

If you're using `nvm` locally, Node `18.x` or `20.x` in the supported SPFx range also works.

The deployable package is created at:

```text
sharepoint/solution/sharepoint-spfx-item-claimer.sppkg
```

## 🚀 Deploy to SharePoint

### 1) Build the package

```bash
npm run package:ship
```

### 2) Upload and deploy the app

1. Upload `sharepoint/solution/sharepoint-spfx-item-claimer.sppkg` to the SharePoint App Catalog.
2. Click **Deploy**.
3. In the target site, go to **Site contents** and add/update the app.
4. If SharePoint shows **Get it** or **Update available**, complete that step first.

### 3) Bind the field customizer to the existing field

#### Option A — Browser console ✅ verified working

Open the **target list view page** in SharePoint, press **F12**, open the **Console**, and run the script below.
It uses the current list context dynamically, so no `listTitle` is required.

```js
(async () => {
  const listIdRaw = _spPageContextInfo.pageListId;
  const fieldInternalName = "Assigned_To"; // change only if your target field differs
  const componentId = "fae2eec7-5401-4ea3-a4a8-9958dd98721f";
  const componentProperties = JSON.stringify({
    claimFieldInternalName: fieldInternalName
  });

  if (!listIdRaw) {
    throw new Error("No list context found. Open a list view page and try again.");
  }

  const listId = String(listIdRaw).replace(/[{}]/g, "");
  const webRel = (_spPageContextInfo.webServerRelativeUrl || "").replace(/\/$/, "");
  const apiBase = webRel + "/_api";

  const digestRes = await fetch(apiBase + "/contextinfo", {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" }
  });
  const digest = (await digestRes.json()).FormDigestValue;

  const fieldUrl =
    apiBase +
    "/web/lists(guid'" +
    listId +
    "')/fields/getbyinternalnameortitle('" +
    encodeURIComponent(fieldInternalName) +
    "')";

  const res = await fetch(fieldUrl, {
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

  console.log(res.ok ? "✅ Field customizer attached successfully." : `❌ Failed ${res.status}: ${await res.text()}`);
})();
```

> **Tip:** For many AMYN lists, keep one script and run it from each target list page. In most cases, you only need to confirm `fieldInternalName`.

#### Optional — Unbind via browser console (current list)

```js
(async () => {
  const listIdRaw = _spPageContextInfo.pageListId;
  const fieldInternalName = "Assigned_To";

  if (!listIdRaw) {
    throw new Error("No list context found. Open a list view page and try again.");
  }

  const listId = String(listIdRaw).replace(/[{}]/g, "");
  const webRel = (_spPageContextInfo.webServerRelativeUrl || "").replace(/\/$/, "");
  const apiBase = webRel + "/_api";

  const digestRes = await fetch(apiBase + "/contextinfo", {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" }
  });
  const digest = (await digestRes.json()).FormDigestValue;

  const fieldUrl =
    apiBase +
    "/web/lists(guid'" +
    listId +
    "')/fields/getbyinternalnameortitle('" +
    encodeURIComponent(fieldInternalName) +
    "')";

  const res = await fetch(fieldUrl, {
    method: "POST",
    headers: {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": digest,
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE"
    },
    body: JSON.stringify({
      ClientSideComponentId: null,
      ClientSideComponentProperties: null
    })
  });

  if (!res.ok) {
    const retry = await fetch(fieldUrl, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest,
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE"
      },
      body: JSON.stringify({
        ClientSideComponentId: "00000000-0000-0000-0000-000000000000",
        ClientSideComponentProperties: null
      })
    });

    if (!retry.ok) {
      console.error(`❌ Failed ${retry.status}: ${await retry.text()}`);
      return;
    }
  }

  console.log("✅ Field customizer unbound.");
})();
```

### 4) Refresh and test

- hard refresh the list page with `Ctrl+Shift+R`
- blank `Assigned_To` cells should now show the inline **`Claim`** button
- clicking the button should claim the item safely using ETag concurrency checks

