# entra-sp-sharepoint-access

A PowerShell script that grants an Entra ID service principal (app registration) access to specific SharePoint Online sites using the **`Sites.Selected`** model, so the app gets access only to the sites you name and not to the whole tenant.

## What it does

`main.ps1` does two things:

1. **Assigns a Microsoft Graph app role to the service principal.** By default this is `Sites.Selected`. It is assigned as an application permission with admin consent already given. If the role is already assigned, the script warns and carries on.
2. **Grants the app `read` or `write` permission (default `write`) on each matching SharePoint site.** It does this by calling `POST /sites/{site-id}/permissions` for each site.

Steps in order:

1. Checks that a Microsoft Graph session exists (`Get-MgContext`) and that it has the permissions it needs. If there is no session or any permission is missing, it stops with an error. The script does **not** sign in by itself.
2. Looks up the target service principal by its object ID.
3. Looks up the Microsoft Graph service principal (appId `00000003-0000-0000-c000-000000000000`) in the tenant, then finds the requested app roles on it.
4. Assigns each role found to the target service principal.
5. Gets every SharePoint site (`Get-MgSite -All`). For each `-siteName` value, it matches sites whose **display name or URL name exactly equals** that value (ignoring case).
6. Grants the app the chosen access level on every matching site, then prints a summary of sites that succeeded, sites that failed, names that matched no site, and names that matched more than one site.

## Prerequisites

- PowerShell 7+ (Windows PowerShell 5.1 should also work)
- Microsoft Graph PowerShell SDK modules:
  ```powershell
  Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Sites -Scope CurrentUser
  ```
- **An existing Microsoft Graph session.** Run `Connect-MgGraph` before the script. This is normally done as a separate automation service principal using app-only auth. That session must have these permissions:
  - `Application.Read.All`
  - `AppRoleAssignment.ReadWrite.All`
  - `Sites.FullControl.All`

  For example:
  ```powershell
  Connect-MgGraph -TenantId <tenant-id> -ClientId <automation-app-id> -CertificateThumbprint <thumbprint> -NoWelcome
  ```
- The **object ID** of the target service principal. This is the Enterprise Application object ID, *not* the app registration's Application (client) ID.

## Usage

```powershell
./main.ps1 -mySpId <service-principal-object-id> -siteName <site>[,<site>...] [-rolesNeeded <role>[,<role>...]] [-accessLevel read|write]
```

| Parameter      | Required | Default            | Description |
|----------------|----------|--------------------|-------------|
| `-mySpId`      | Yes      |                    | Object ID of the service principal that will receive access. |
| `-siteName`    | Yes      |                    | One or more site display names or URL names. Each is matched **exactly** (ignoring case). |
| `-rolesNeeded` | No       | `Sites.Selected`   | Microsoft Graph app role(s) to assign to the service principal. |
| `-accessLevel` | No       | `write`            | Permission granted on each site. Must be `read` or `write`. |

### Examples

Grant access to a single site:

```powershell
./main.ps1 -mySpId 11111111-2222-3333-4444-555555555555 -siteName 'Finance'
```

Grant read-only access to several sites:

```powershell
./main.ps1 -mySpId 11111111-2222-3333-4444-555555555555 -siteName 'Finance','HR-Portal','Marketing' -accessLevel read
```

### Exit codes

| Code | Meaning |
|------|---------|
| `0`  | Every name matched exactly one site, and access was granted on all of them. |
| `1`  | At least one site grant failed, or at least one name matched no site or more than one site. |

If a fatal error occurs (there is no Graph session or it is missing permissions, the service principal is not found, no roles match, or listing sites fails), the script throws and stops before granting anything on sites.

## Things to be aware of

- **Names are matched exactly, but display names are not unique.** `-siteName 'HR'` matches only a site whose display name or URL name is `HR`, not `HR-Portal`. However, two sites can have the same display name, and one site's display name can equal another site's URL name. When a name matches more than one site, the script **skips that name, grants nothing for it**, lists the matching URLs in a warning, and exits with code `1`. To target one specific site, use its URL name (the part after `/sites/`).
- **Only `read` and `write` can be chosen.** Graph also supports `fullcontrol` and `manage` for `Sites.Selected`, but the script deliberately rejects them.
- **Running it again is mostly safe.** An existing app role assignment is detected and skipped. A site permission grant is sent again each time you run the script.
- **Site discovery depends on `Get-MgSite -All`.** This relies on the authenticated principal being able to list every site in the tenant, which works with app-only `Sites.FullControl.All`.

## Verifying access

Once the script has run, the app can authenticate with its own credentials (client credentials flow) and call:

```http
GET https://graph.microsoft.com/v1.0/sites/{site-id}
```

This should succeed for the granted sites and return `403` for all other sites. To see who has access to a site:

```powershell
Get-MgSitePermission -SiteId <site-id>
```
