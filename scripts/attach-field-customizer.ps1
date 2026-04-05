param(
  [Parameter(Mandatory = $true)]
  [string]$SiteUrl,

  [Parameter(Mandatory = $true)]
  [string]$ListTitle,

  [string]$FieldInternalName = 'Assigned_To',
  [string]$ComponentId = 'fae2eec7-5401-4ea3-a4a8-9958dd98721f',
  [string]$ComponentProperties = '{"claimFieldInternalName":"Assigned_To"}'
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
  throw "PnP.PowerShell is not installed. Install it first with: Install-Module PnP.PowerShell -Scope CurrentUser"
}

Write-Host "Connecting to $SiteUrl..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

$field = Get-PnPField -List $ListTitle -Identity $FieldInternalName

if (-not $field) {
  throw "Could not find field '$FieldInternalName' on list '$ListTitle'."
}

$field.ClientSideComponentId = [Guid]$ComponentId
$field.ClientSideComponentProperties = $ComponentProperties
$field.Update()
Invoke-PnPQuery

Write-Host "Attached field customizer $ComponentId to $ListTitle/$FieldInternalName" -ForegroundColor Green
Write-Host "Refresh the list page with Ctrl+Shift+R." -ForegroundColor Yellow
