[CmdletBinding()]
param(
    # service principal object Id
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$mySpId,

    # sharepoint site display names (or URL names)
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string[]]$siteName,

    [string[]]$rolesNeeded = @('Sites.Selected'),

    # permission level granted on each site
    [ValidateSet('read', 'write')]
    [string]$accessLevel = 'write'
)

$ErrorActionPreference = 'Stop'

$graphAppId = '00000003-0000-0000-c000-000000000000'
$requiredScopes = @('Application.Read.All', 'AppRoleAssignment.ReadWrite.All', 'Sites.FullControl.All')

$context = Get-MgContext
if (-not $context) {
    throw 'Not connected to Microsoft Graph. Run Connect-MgGraph before running this script.'
}
$missingScopes = @($requiredScopes | Where-Object { $_ -notin $context.Scopes })
if ($missingScopes.Count -gt 0) {
    throw "Missing required permissions: $($missingScopes -join ', ')"
}

try {
    $mySp = Get-MgServicePrincipal -ServicePrincipalId $mySpId -ErrorAction Stop
} catch {
    throw "Failed to get service principal '$mySpId': $_"
}
if (-not $mySp) {
    throw "Service principal '$mySpId' not found"
}
$mySpAppId = $mySp.AppId

try {
    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -ErrorAction Stop
} catch {
    throw "Failed to get Microsoft Graph service principal: $_"
}
if (-not $graphSp) {
    throw "Microsoft Graph service principal (appId '$graphAppId') not found in this tenant"
}
$appRoles = @($graphSp | Select-Object -ExpandProperty AppRoles `
    | Where-Object { $_.Value -iin $rolesNeeded })

$missingRoles = @($rolesNeeded | Where-Object { $_ -notin $appRoles.Value })
if ($missingRoles.Count -gt 0) {
    Write-Warning "Microsoft Graph app roles not found: $($missingRoles -join ', ')"
}
if ($appRoles.Count -eq 0) {
    throw "None of the requested app roles were found on Microsoft Graph: $($rolesNeeded -join ', ')"
}

foreach ($ar in $appRoles) {
    try {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $mySpId `
            -AppRoleId $ar.Id -PrincipalId $mySpId -ResourceId $graphSp.Id -ErrorAction Stop
    } catch {
        if ("$_" -match 'already exists') {
            Write-Warning "App role '$($ar.Value)' is already assigned to service principal '$mySpId'"
        } else {
            throw "Failed to assign app role '$($ar.Value)' to service principal '$mySpId': $_"
        }
    }
}

try {
    $sites = Get-MgSite -All -ErrorAction Stop
} catch {
    throw "Failed to get SharePoint sites: $_"
}
if (-not $sites) {
    throw 'No SharePoint sites returned'
}

$pbody = @{
    roles = @( $accessLevel.ToLower() )
    grantedToIdentities = @(
        @{
            application = @{
                id = $mySpAppId
                displayName = $mySp.DisplayName
            }
        }
    )
}

$succeeded = [System.Collections.Generic.List[string]]::new()
$failed = [System.Collections.Generic.List[string]]::new()
$notFound = [System.Collections.Generic.List[string]]::new()
$ambiguous = [System.Collections.Generic.List[string]]::new()

foreach ($name in $siteName) {
    $matchedSites = @($sites | Where-Object { $_.DisplayName -eq $name -or $_.Name -eq $name })
    if ($matchedSites.Count -eq 0) {
        Write-Warning "No SharePoint site found with name '$name'"
        $notFound.Add($name)
        continue
    }
    if ($matchedSites.Count -gt 1) {
        Write-Warning ("'$name' matches $($matchedSites.Count) sites, skipping: " +
            ($matchedSites.WebUrl -join ', '))
        $ambiguous.Add($name)
        continue
    }
    foreach ($site in $matchedSites) {
        try {
            New-MgSitePermission -SiteId $site.Id -BodyParameter $pbody -ErrorAction Stop | Out-Null
            $succeeded.Add($site.WebUrl)
        } catch {
            Write-Warning "Failed to grant access on site '$($site.WebUrl)': $_"
            $failed.Add($site.WebUrl)
        }
    }
}

Write-Host ''
Write-Host "Granted access on $($succeeded.Count) site(s):"
$succeeded | ForEach-Object { Write-Host "  $_" }
if ($failed.Count -gt 0) {
    Write-Host "Failed on $($failed.Count) site(s):"
    $failed | ForEach-Object { Write-Host "  $_" }
}
if ($notFound.Count -gt 0) {
    Write-Host "No site found for $($notFound.Count) name(s):"
    $notFound | ForEach-Object { Write-Host "  $_" }
}
if ($ambiguous.Count -gt 0) {
    Write-Host "Multiple sites matched $($ambiguous.Count) name(s), skipped:"
    $ambiguous | ForEach-Object { Write-Host "  $_" }
}

if ($failed.Count -gt 0 -or $notFound.Count -gt 0 -or $ambiguous.Count -gt 0) {
    exit 1
}
