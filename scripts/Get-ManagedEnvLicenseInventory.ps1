<#
.SYNOPSIS
    Pre-flight user and license entitlement inventory for a Power Platform
    environment, ahead of enabling Managed Environments.

.DESCRIPTION
    Joins three data sources that Microsoft does not join for you:

      1. Dataverse  - who is ACTUALLY in the environment (systemuser table),
                      filtered to license-consuming access modes only.
      2. Entra ID   - what licenses each of those users actually holds,
                      including licenses inherited via group-based assignment.
      3. Rules      - classification of each license against the documented
                      Managed Environments qualifying list.

    Every user lands in exactly one bucket:

      Covered          - holds a qualifying premium or D365 Enterprise license
      NotQualifying    - D365 Pro, M365 seeded rights, Developer Plan, free/viral
      Unclassified     - license SKU not in the rules table -> REVIEW MANUALLY
      NoLicense        - in the environment, holds nothing at all

    The script is READ-ONLY. It changes nothing in Dataverse or Entra.

.PARAMETER EnvironmentUrl
    The Dataverse environment URL to inventory.
    e.g. https://contoso.crm.dynamics.com
    <<< REPLACE THE PLACEHOLDER DEFAULT BELOW, OR PASS AT RUNTIME >>>

.PARAMETER OutputPath
    Folder for the output workbook / CSVs. Defaults to the script folder.

.PARAMETER IncludeDisabled
    Include disabled Dataverse users. Off by default - they do not consume
    licenses, but you may want them for an access-review.

.PARAMETER AllAccessModes
    Include ALL access modes rather than just Read-Write (0).
    Off by default. Non-interactive / administrative / stub users and
    application users do not consume user licenses and would inflate the count.

.PARAMETER GraphPageAll
    Page every user in the tenant once and join locally (default, fastest for
    environments with many users). Switch off to look users up individually -
    slower per user, but far less data pulled on very large tenants with a
    small environment population.

.EXAMPLE
    .\Get-ManagedEnvLicenseInventory.ps1 -EnvironmentUrl "https://contoso.crm.dynamics.com"

.EXAMPLE
    .\Get-ManagedEnvLicenseInventory.ps1 `
        -EnvironmentUrl "https://contoso.crm.dynamics.com" `
        -OutputPath "C:\Reports" -IncludeDisabled

.NOTES
    AUTH
      Uses Azure CLI for both tokens (zero module install):
          az login
      The Azure CLI first-party client generally holds the delegated Graph
      permissions needed to read users and licenses. If the Graph call returns
      403, sign in with an account holding at least one of:
          Directory.Read.All / User.Read.All / Organization.Read.All
      You must also be a Dataverse user in the target environment with
      privileges to read the systemuser table (System Administrator or
      equivalent read access).

    EXCEL OUTPUT
      If the ImportExcel module is present, writes a multi-tab .xlsx.
      Otherwise falls back to CSVs automatically. To get the workbook:
          Install-Module ImportExcel -Scope CurrentUser

    KNOWN LIMITS - READ THESE
      * Power Apps PER APP licenses are allocated to the ENVIRONMENT, not to
        the user, so they cannot be seen in Entra per-user data. If you use
        per-app plans, check PPAC > Licensing > Power Apps > Environments and
        treat this report's "NotQualifying" bucket as an upper bound.
      * Pay-as-you-go meters likewise sit at environment level and are invisible
        here. Same caveat applies.
      * This report shows ENTITLEMENT, not USAGE. A user may be unlicensed and
        never launch an app - in which case they may need removing rather than
        licensing. Corroborate against PPAC > Licensing > Power Apps >
        Download Reports > "Active users" before you buy anything.
      * D365 Pro / Professional: RESOLVED as NOT qualifying. The Power Platform
        Licensing Guide (August 2026), p.9 footnote 2, states plainly:
        "Dynamics 365 Pro does not have use rights for Power Apps or Power
        Pages." The same guide, p.24 footnote 1, confirms the Managed
        Environments entitlement covers standalone licenses only and expressly
        excludes "the limited Power Apps, Power Automate and Power Pages use
        rights that come with select Dynamics 365 and Microsoft 365 licenses."
        Sales Professional / Customer Service Professional users therefore need
        a per-app plan, PAYG meter, or a premium/Enterprise license.
#>

[CmdletBinding()]
param(
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    # REPLACE THIS PLACEHOLDER with your environment URL, or pass -EnvironmentUrl
    [string] $EnvironmentUrl = "https://REPLACE-ME.crm.dynamics.com",
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    [string] $OutputPath,
    [switch] $IncludeDisabled,
    [switch] $AllAccessModes,
    [bool]   $GraphPageAll = $true
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# 0. Preflight
# ---------------------------------------------------------------------------

if ($EnvironmentUrl -match 'REPLACE-ME') {
    throw "EnvironmentUrl is still the placeholder. Edit the default in the param block or pass -EnvironmentUrl 'https://yourorg.crm.dynamics.com'."
}

$EnvironmentUrl = $EnvironmentUrl.TrimEnd('/')

if (-not $OutputPath) {
    $OutputPath = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
}
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    throw "Azure CLI ('az') not found on PATH. Install it, then run 'az login'."
}

function Get-Token {
    param([Parameter(Mandatory)][string] $Resource)

    $raw = az account get-access-token --resource $Resource --output json 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to acquire a token for '$Resource'. Run 'az login' and try again.`n$raw"
    }
    return ($raw | ConvertFrom-Json).accessToken
}

function Invoke-Paged {
    <#
      Follows @odata.nextLink for both Dataverse and Graph, accumulating .value
    #>
    param(
        [Parameter(Mandatory)][string] $Uri,
        [Parameter(Mandatory)][hashtable] $Headers,
        [string] $Label = 'records'
    )

    $all  = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    $page = 0

    while ($next) {
        $page++
        Write-Verbose "  page $page -> $next"
        $resp = Invoke-RestMethod -Method Get -Uri $next -Headers $Headers
        if ($resp.PSObject.Properties.Name -contains 'value' -and $resp.value) {
            $all.AddRange([object[]]$resp.value)
        }
        Write-Host ("    ...{0,7} {1}" -f $all.Count, $Label) -NoNewline:$false
        $next = if ($resp.PSObject.Properties.Name -contains '@odata.nextLink') {
            $resp.'@odata.nextLink'
        } else { $null }
    }
    return $all
}

Write-Host ""
Write-Host "Managed Environments - pre-enablement license inventory" -ForegroundColor Cyan
Write-Host ("Environment : {0}" -f $EnvironmentUrl)
Write-Host ("Run at      : {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
Write-Host ""

# ---------------------------------------------------------------------------
# 1. Dataverse - the population actually in the environment
# ---------------------------------------------------------------------------

Write-Host "[1/4] Reading Dataverse systemuser..." -ForegroundColor Yellow

$dvToken = Get-Token -Resource $EnvironmentUrl
$dvHeaders = @{
    Authorization      = "Bearer $dvToken"
    Accept             = 'application/json'
    'OData-MaxVersion' = '4.0'
    'OData-Version'    = '4.0'
    Prefer             = 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
}

# accessmode: 0 Read-Write | 1 Admin | 2 Read | 3 Support | 4 Non-interactive
#             5 Delegated Admin | 6 Bulk Load Data Import | 7 Bulk Load Data Import
# Only Read-Write (0) consumes a standard user license.
$filters = @()
if (-not $IncludeDisabled)  { $filters += 'isdisabled eq false' }
if (-not $AllAccessModes)   { $filters += 'accessmode eq 0' }
# Exclude application users (S2S) - they never consume user licenses.
$filters += 'applicationid eq null'

$select = 'systemuserid,fullname,internalemailaddress,domainname,' +
          'azureactivedirectoryobjectid,isdisabled,accessmode,' +
          'islicensed,isintegrationuser,createdon,lastaccessedtime'

$dvUri = "$EnvironmentUrl/api/data/v9.2/systemusers?`$select=$select"
if ($filters.Count) { $dvUri += "&`$filter=" + ($filters -join ' and ') }
$dvUri += '&$count=true'

$dvUsers = Invoke-Paged -Uri $dvUri -Headers $dvHeaders -Label 'Dataverse users'

Write-Host ("      {0} license-consuming users in environment" -f $dvUsers.Count) -ForegroundColor Green

if ($dvUsers.Count -eq 0) {
    throw "No users returned. Check the environment URL and that your account can read systemuser."
}

# ---------------------------------------------------------------------------
# 2. Entra ID - SKU catalogue + per-user assignments (incl. group-based)
# ---------------------------------------------------------------------------

Write-Host "[2/4] Reading Entra ID license data..." -ForegroundColor Yellow

$graphToken = Get-Token -Resource 'https://graph.microsoft.com'
$gHeaders = @{
    Authorization    = "Bearer $graphToken"
    Accept           = 'application/json'
    ConsistencyLevel = 'eventual'
}

# 2a. Tenant SKU catalogue - ground truth for SkuId -> SkuPartNumber
$skus = Invoke-Paged -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus' `
                     -Headers $gHeaders -Label 'tenant SKUs'

$skuMap = @{}
foreach ($s in $skus) { $skuMap[[string]$s.skuId] = $s.skuPartNumber }

Write-Host ("      {0} SKUs in tenant catalogue" -f $skuMap.Count) -ForegroundColor Green

# 2b. User license assignments.
#     licenseAssignmentStates exposes assignedByGroup - the field that catches
#     group-based licensing, the single most common source of a miscount.
$gSelect = 'id,userPrincipalName,displayName,accountEnabled,assignedLicenses,licenseAssignmentStates'

$graphUsers = @{}

if ($GraphPageAll) {
    $gUri  = "https://graph.microsoft.com/v1.0/users?`$select=$gSelect&`$top=999"
    $all   = Invoke-Paged -Uri $gUri -Headers $gHeaders -Label 'Entra users'
    foreach ($u in $all) { $graphUsers[[string]$u.id] = $u }
}
else {
    $ids = $dvUsers |
        Where-Object { $_.azureactivedirectoryobjectid } |
        ForEach-Object { [string]$_.azureactivedirectoryobjectid } |
        Sort-Object -Unique

    $i = 0
    foreach ($id in $ids) {
        $i++
        if ($i % 50 -eq 0) { Write-Host "    ...$i / $($ids.Count) looked up" }
        try {
            $u = Invoke-RestMethod -Method Get -Headers $gHeaders `
                 -Uri "https://graph.microsoft.com/v1.0/users/$id`?`$select=$gSelect"
            $graphUsers[[string]$u.id] = $u
        }
        catch {
            Write-Warning "  Graph lookup failed for $id : $($_.Exception.Message)"
        }
    }
}

Write-Host ("      {0} Entra users retrieved" -f $graphUsers.Count) -ForegroundColor Green

# ---------------------------------------------------------------------------
# 3. Classification rules
# ---------------------------------------------------------------------------
#
# Matched on SkuPartNumber (the stable, human-readable identifier) rather than
# GUIDs, which vary and rot. Anything unmatched is deliberately surfaced as
# 'Unclassified' rather than silently assumed non-qualifying - an inventory
# that quietly guesses is worse than one that admits it does not know.
#
# Rules are evaluated in order; first match wins per license.
# ---------------------------------------------------------------------------

$rules = @(
    # --- Power Platform standalone premium ------------------------------
    @{ Pattern = '^POWERAPPS_PER_USER';                Bucket='Covered';        Family='Power Apps Premium' }
    @{ Pattern = '^POWERAPPS_PER_APP';                 Bucket='Covered';        Family='Power Apps per app (env-allocated - verify in PPAC)' }
    @{ Pattern = '^FLOW_PER_USER|^POWERAUTOMATE_';     Bucket='Covered';        Family='Power Automate Premium' }
    @{ Pattern = '^FLOW_PER_BUSINESS|^FLOW_PER_FLOW|FLOW_BUSINESS_PROCESS'; Bucket='Covered'; Family='Power Automate per flow (env-allocated)' }
    @{ Pattern = 'POWER_AUTOMATE_PROCESS|AUTOMATE_PROCESS|HOSTED_PROCESS|HOSTED_RPA'; Bucket='Covered'; Family='Power Automate Process / Hosted Process' }
    @{ Pattern = '^POWERAPPS_PORTALS|^POWERPAGES';     Bucket='Covered';        Family='Power Pages' }
    @{ Pattern = 'VIRTUAL_AGENT|COPILOT_STUDIO|^POWER_VIRTUAL_AGENT'; Bucket='Covered'; Family='Copilot Studio' }

    # --- Dynamics 365 Pro / Professional --------------------------------
    # ORDERING IS LOAD-BEARING: these MUST precede the Enterprise rules.
    # Some Professional SKU part numbers embed "ENTERPRISE" (e.g.
    # DYN365_ENTERPRISE_SALES_PROFESSIONAL) and would otherwise be matched by
    # '^DYN365_ENTERPRISE_SALES' and wrongly classed as Covered.
    # RESOLVED as NOT qualifying - Power Platform Licensing Guide (Aug 2026)
    # p.9 fn2: "Dynamics 365 Pro does not have use rights for Power Apps or
    # Power Pages." p.24 fn1 confirms Managed Environments requires standalone
    # licenses and excludes limited D365-bundled use rights.
    @{ Pattern = 'SALES_PRO(FESSIONAL)?';                           Bucket='NotQualifying'; Family='D365 Sales Professional - NO Power Apps rights' }
    @{ Pattern = 'CUSTOMER_SERVICE_PRO(FESSIONAL)?|^D365_CS_PRO';   Bucket='NotQualifying'; Family='D365 Customer Service Professional - NO Power Apps rights' }

    # --- Dynamics 365 Enterprise / Premium / Team Members ---------------
    @{ Pattern = '^DYN365_ENTERPRISE_SALES|^D365_SALES_ENT';        Bucket='Covered'; Family='D365 Sales Enterprise' }
    @{ Pattern = '^D365_SALES_PREMIUM|SALES_PREMIUM';               Bucket='Covered'; Family='D365 Sales Premium' }
    @{ Pattern = '^DYN365_ENTERPRISE_CUSTOMER_SERVICE|^D365_CS_ENT';Bucket='Covered'; Family='D365 Customer Service Enterprise' }
    @{ Pattern = 'CUSTOMER_SERVICE_PREMIUM|^D365_CS_PREMIUM';       Bucket='Covered'; Family='D365 Customer Service Premium' }
    @{ Pattern = 'FIELD_SERVICE';                                   Bucket='Covered'; Family='D365 Field Service' }
    @{ Pattern = '^DYN365_FINANCE|^D365_FINANCE';                   Bucket='Covered'; Family='D365 Finance' }
    @{ Pattern = '^DYN365_SCM|^D365_SCM|SUPPLY_CHAIN';              Bucket='Covered'; Family='D365 Supply Chain Management' }
    @{ Pattern = 'PROJECT_OPERATIONS|^DYN365_PROJECT';              Bucket='Covered'; Family='D365 Project Operations' }
    @{ Pattern = '^DYN365_COMMERCE|^D365_COMMERCE|^DYN365_RETAIL';  Bucket='Covered'; Family='D365 Commerce' }
    @{ Pattern = 'HUMAN_RESOURCES|^DYN365_HR|^D365_HR';             Bucket='Covered'; Family='D365 Human Resources' }
    @{ Pattern = 'BUSCENTRAL|BUSINESS_CENTRAL';                     Bucket='Covered'; Family='D365 Business Central' }
    @{ Pattern = 'TEAM_MEMBERS';                                    Bucket='Covered'; Family='D365 Team Members' }
    @{ Pattern = 'INTELLIGENT_ORDER|^DYN365_IOM';                   Bucket='Covered'; Family='D365 Intelligent Order Management' }
    @{ Pattern = 'CUSTOMER_INSIGHTS|^DYN365_CI';                    Bucket='Covered'; Family='D365 Customer Insights' }

    # --- Dynamics 365 Pro / Professional --------------------------------
    # (see the Pro rules above the Enterprise block - ordering is load-bearing)

    # --- Explicitly non-qualifying --------------------------------------
    @{ Pattern = '^POWERAPPS_DEV';                                  Bucket='NotQualifying'; Family='Power Apps Developer Plan (excluded)' }
    @{ Pattern = 'OPERATIONS_ACTIVITY|_DEVICE$|^DYN365_OPERATIONS_DEVICE'; Bucket='NotQualifying'; Family='D365 Operations Activity / Device (not qualifying)' }
    @{ Pattern = '^POWERAPPS_INDIVIDUAL|^POWERAPPS_VIRAL|^FLOW_FREE|^POWERFLOW_VIRAL'; Bucket='NotQualifying'; Family='Free / viral Power Platform' }
    @{ Pattern = '^SPE_E3|^SPE_E5|^SPE_F1|^ENTERPRISEPACK|^ENTERPRISEPREMIUM|^STANDARDPACK|^DESKLESSPACK|^M365_F1|^O365_'; Bucket='NotQualifying'; Family='M365/O365 seeded rights only' }
    @{ Pattern = '^POWER_BI';                                       Bucket='NotQualifying'; Family='Power BI (does not grant Power Apps rights)' }
    @{ Pattern = '^EMS|^AAD_|^ENTRA';                               Bucket='NotQualifying'; Family='Entra / EMS (not Power Platform)' }
)

function Get-LicenseClass {
    param([string] $SkuPartNumber)
    foreach ($r in $rules) {
        if ($SkuPartNumber -match $r.Pattern) {
            return [pscustomobject]@{ Bucket = $r.Bucket; Family = $r.Family }
        }
    }
    return [pscustomobject]@{ Bucket = 'Unclassified'; Family = 'UNKNOWN - review manually' }
}

# Bucket precedence: best outcome a user holds wins.
$rank = @{ 'Covered' = 4; 'Unclassified' = 2; 'NotQualifying' = 1; 'NoLicense' = 0 }

# ---------------------------------------------------------------------------
# 4. Join, classify, emit
# ---------------------------------------------------------------------------

Write-Host "[3/4] Joining and classifying..." -ForegroundColor Yellow

$report        = [System.Collections.Generic.List[object]]::new()
$unknownSkus   = [System.Collections.Generic.HashSet[string]]::new()

foreach ($dv in $dvUsers) {

    $aadId = if ($dv.azureactivedirectoryobjectid) { [string]$dv.azureactivedirectoryobjectid } else { $null }
    $g     = if ($aadId -and $graphUsers.ContainsKey($aadId)) { $graphUsers[$aadId] } else { $null }

    $skuNames      = @()
    $qualifying    = @()
    $groupAssigned = @()
    $bestBucket    = 'NoLicense'

    if ($g -and $g.assignedLicenses) {
        # Map group-assigned SKUs so we can annotate provenance.
        $byGroup = @{}
        if ($g.PSObject.Properties.Name -contains 'licenseAssignmentStates' -and $g.licenseAssignmentStates) {
            foreach ($st in $g.licenseAssignmentStates) {
                if ($st.assignedByGroup) { $byGroup[[string]$st.skuId] = $st.assignedByGroup }
            }
        }

        foreach ($lic in $g.assignedLicenses) {
            $sid  = [string]$lic.skuId
            $name = if ($skuMap.ContainsKey($sid)) { $skuMap[$sid] } else { "UNKNOWN_SKU($sid)" }
            $cls  = Get-LicenseClass -SkuPartNumber $name

            if ($cls.Bucket -eq 'Unclassified') { [void]$unknownSkus.Add($name) }

            $tag = if ($byGroup.ContainsKey($sid)) { "$name [group]" } else { $name }
            $skuNames += $tag
            if ($byGroup.ContainsKey($sid)) { $groupAssigned += $name }

            if ($cls.Bucket -eq 'Covered') { $qualifying += $cls.Family }

            if ($rank[$cls.Bucket] -gt $rank[$bestBucket]) { $bestBucket = $cls.Bucket }
        }
    }

    $action = switch ($bestBucket) {
        'Covered'        { 'None - compliant' }
        'ContextLimited' { 'REVIEW - reclassify; see rules table' }
        'Unclassified'   { 'REVIEW - unrecognized SKU, classify manually' }
        'NotQualifying'  { 'ACTION - assign premium license, or remove from environment' }
        'NoLicense'      { 'ACTION - no license at all; likely stale user or needs licensing' }
    }

    $report.Add([pscustomobject]@{
        FullName          = $dv.fullname
        Email             = $dv.internalemailaddress
        UPN               = if ($g) { $g.userPrincipalName } else { $dv.domainname }
        Bucket            = $bestBucket
        Action            = $action
        QualifyingVia     = ($qualifying | Sort-Object -Unique) -join '; '
        AllLicenses       = ($skuNames | Sort-Object -Unique) -join '; '
        GroupAssigned     = ($groupAssigned | Sort-Object -Unique) -join '; '
        EntraAccountEnabled = if ($g) { $g.accountEnabled } else { $null }
        FoundInEntra      = [bool]$g
        DataverseDisabled = $dv.isdisabled
        AccessMode        = $dv.'accessmode@OData.Community.Display.V1.FormattedValue'
        DvIsLicensedFlag  = $dv.islicensed
        LastAccessed      = $dv.lastaccessedtime
        CreatedOn         = $dv.createdon
        SystemUserId      = $dv.systemuserid
        AadObjectId       = $aadId
    })
}

# --- Summaries --------------------------------------------------------------

$summary = $report | Group-Object Bucket | ForEach-Object {
    [pscustomobject]@{
        Bucket = $_.Name
        Users  = $_.Count
        Share  = '{0:P1}' -f ($_.Count / $report.Count)
    }
} | Sort-Object { -$rank[$_.Bucket] }

$licenseRollup = $report |
    Where-Object { $_.AllLicenses } |
    ForEach-Object { $_.AllLicenses -split '; ' } |
    ForEach-Object { $_ -replace ' \[group\]$','' } |
    Group-Object |
    Sort-Object Count -Descending |
    ForEach-Object {
        $c = Get-LicenseClass -SkuPartNumber $_.Name
        [pscustomobject]@{
            SkuPartNumber = $_.Name
            Users         = $_.Count
            Bucket        = $c.Bucket
            Family        = $c.Family
        }
    }

$actionList = $report |
    Where-Object { $_.Bucket -in @('NotQualifying','NoLicense','Unclassified') } |
    Sort-Object @{e={ $rank[$_.Bucket] }}, FullName

# --- Console readout --------------------------------------------------------

Write-Host ""
Write-Host "RESULT" -ForegroundColor Cyan
Write-Host "------"
foreach ($row in $summary) {
    $color = switch ($row.Bucket) {
        'Covered'        { 'Green' }
        'ContextLimited' { 'Yellow' }
        'Unclassified'   { 'Yellow' }
        default          { 'Red' }
    }
    Write-Host ("  {0,-16} {1,6} users  ({2})" -f $row.Bucket, $row.Users, $row.Share) -ForegroundColor $color
}
Write-Host ""
Write-Host ("  Users needing attention before enabling Managed Environments: {0}" -f $actionList.Count) `
    -ForegroundColor $(if ($actionList.Count) { 'Red' } else { 'Green' })

if ($unknownSkus.Count) {
    Write-Host ""
    Write-Host "  Unrecognized SKUs - add these to the `$rules table:" -ForegroundColor Yellow
    $unknownSkus | Sort-Object | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
}

# --- Write output -----------------------------------------------------------

Write-Host ""
Write-Host "[4/4] Writing output..." -ForegroundColor Yellow

$stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$envSlug = ([uri]$EnvironmentUrl).Host.Split('.')[0]
$base  = Join-Path $OutputPath ("ManagedEnv_LicenseInventory_{0}_{1}" -f $envSlug, $stamp)

$meta = [pscustomobject]@{
    EnvironmentUrl     = $EnvironmentUrl
    GeneratedUtc       = (Get-Date).ToUniversalTime().ToString('u')
    UsersInScope       = $report.Count
    IncludeDisabled    = [bool]$IncludeDisabled
    AllAccessModes     = [bool]$AllAccessModes
    NeedsAttention     = $actionList.Count
    UnrecognizedSkus   = ($unknownSkus | Sort-Object) -join '; '
    Caveat_PerApp      = 'Per-app plans and PAYG meters are environment-allocated and invisible to this report. Verify in PPAC > Licensing > Power Apps > Environments.'
    Caveat_Usage       = 'This is entitlement, not usage. Cross-check PPAC > Licensing > Power Apps > Download Reports > Active users before purchasing.'
    Caveat_Pro         = 'D365 Pro/Professional does NOT qualify. Power Platform Licensing Guide (Aug 2026) p.9 fn2: "Dynamics 365 Pro does not have use rights for Power Apps or Power Pages."'
}

$haveExcel = [bool](Get-Module -ListAvailable -Name ImportExcel)

if ($haveExcel) {
    Import-Module ImportExcel -ErrorAction Stop
    $xlsx = "$base.xlsx"
    Remove-Item $xlsx -ErrorAction SilentlyContinue

    $summary       | Export-Excel -Path $xlsx -WorksheetName 'Summary'      -AutoSize -BoldTopRow -FreezeTopRow
    $actionList    | Export-Excel -Path $xlsx -WorksheetName 'Action List'  -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter
    $report        | Export-Excel -Path $xlsx -WorksheetName 'All Users'    -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter
    $licenseRollup | Export-Excel -Path $xlsx -WorksheetName 'License Mix'  -AutoSize -BoldTopRow -FreezeTopRow
    $meta          | Export-Excel -Path $xlsx -WorksheetName 'Run Notes'    -AutoSize -BoldTopRow

    Write-Host ("      Workbook: {0}" -f $xlsx) -ForegroundColor Green
}
else {
    $report        | Export-Csv "$base`_AllUsers.csv"    -NoTypeInformation -Encoding UTF8
    $actionList    | Export-Csv "$base`_ActionList.csv"  -NoTypeInformation -Encoding UTF8
    $summary       | Export-Csv "$base`_Summary.csv"     -NoTypeInformation -Encoding UTF8
    $licenseRollup | Export-Csv "$base`_LicenseMix.csv"  -NoTypeInformation -Encoding UTF8
    $meta          | Export-Csv "$base`_RunNotes.csv"    -NoTypeInformation -Encoding UTF8

    Write-Host ("      CSVs written to: {0}" -f $OutputPath) -ForegroundColor Green
    Write-Host "      (Install-Module ImportExcel -Scope CurrentUser for a single workbook)" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "Done. Review the Action List before enabling Managed Environments." -ForegroundColor Cyan
Write-Host ""
