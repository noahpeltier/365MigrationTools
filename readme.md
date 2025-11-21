# Migration Project Toolkit

This repository contains the `MigrationToolkit` PowerShell module. Import it once per shell and you get a project-aware workflow for provisioning Microsoft Graph application registrations (source/target tenants), caching their credentials, and connecting later without prompts.

## Quick Start

```powershell
# Load/refresh the module
Import-Module "$PSScriptRoot/MigrationToolkit.psm1" -Force

# Create a new project (folder, config.json, profiles/)
New-MigrationProject -Name "PmP to Spectrum tenant" -Path "$PWD/Projects"

# Activate the project (sets env vars like a venv)
Set-ActiveMigrationProject -ProjectPath "$PWD/Projects/PmP-to-Spectrum-tenant"

# Provision target tenant credentials (runs the New-MgCustomApp cmdlet under the covers)
Add-ProjectTargetTenant -TenantId 'cae47bb1-6b8e-4af1-9590-2adedf9701dc'

Add-ProjectSourceTenant -TenantId '7b3b52f6-ce55-427e-95b1-12ae5d8c33fb'

# Later, connect silently when running migration scripts
Connect-ProjectTargetTenant -Silent
Connect-ProjectSourceTenant -Silent
```

## Command Reference

- `Import-Module <path>/MigrationToolkit.psm1`: loads the module. Place the folder under `$env:PSModulePath` to import it globally.
- `New-MigrationProject -Name <string> [-Path <root>] [-Force]`: creates the folder, `Config.json`, and `profiles/` cache. Returns a project context object.
- `Get-MigrationProject [-ProjectPath <path>]`: summarize project details and tenant entries (defaults to the active project or current directory).
- `Export-MigrationProject [-ProjectPath <path>] [-DestinationPath <zip>] [-Overwrite]`: packages Config.json, profiles, and certificates (to PFX with a password prompt) into a zip for portability.
- `Import-MigrationProject -ZipPath <file> [-DestinationRoot <path>] [-Force]`: restores a project from an export, rewrites profile paths, and imports PFX certs.
- `Set-ActiveMigrationProject -ProjectPath <path>` / `Clear-ActiveMigrationProject`: acts like `python -m venv activate`—sets `MG_PROFILE_DIRECTORY`/`MG_CONFIG_PATH` for the current shell so any helper uses that project automatically.
- `Add-ProjectTargetTenant` / `Add-ProjectSourceTenant`: wraps `New-MgCustomApp`, ensuring Graph is disconnected first, collecting tenant metadata (`tenantDisplayName`, default domain), writing credentials to `Config.json`, and disconnecting again. Accepts `-AuthType Certificate|ClientCredential` and overrides like `-Permissions` or `-AppDisplayName`.
- Need Exchange ManageAsApp? Add `-EnableExchangeManageAsApp` to `Add-ProjectTargetTenant` / `Add-ProjectSourceTenant` to include that role and Exchange connection example.
- `Connect-ProjectTargetTenant` / `Connect-ProjectSourceTenant`: connects to the stored project entry (defaults to the active project, but you can pass `-ProjectPath`). Add `-Silent` to suppress verbose output, or `-PassThruProfile` to inspect the JSON.
- `Connect-ProjectTargetExchange` / `Connect-ProjectSourceExchange`: connect to Exchange Online using the stored app registration (certificate thumbprint) and `defaultDomain` from `Config.json`. Requires `-EnableExchangeManageAsApp` when provisioning.
- Aliases: `cptt`/`cpst` for target/source Graph connects, `cpte`/`cpse` for target/source Exchange connects.
- `Invoke-MigrateUser -Source <upn> -Target <upn> [-Licenses <skuIds>] -Password <string> [-BlockSignin]`: clones a user from source to target, copying basic profile fields, setting a password (force change on next sign-in), optionally assigning licenses, and blocking sign-in if requested. Throws if the target UPN already exists.
- `Invoke-DiscoverSourceAccounts [-ProjectPath <path>] [-OutputPath <file>]`: connects to the source tenant, pulls all users, looks up mailbox type via Exchange (if available), and writes `DiscoveredAccounts.csv` into the project by default.
- `Invoke-DiscoverTargetAccounts [-ProjectPath <path>] [-OutputPath <file>]`: same as source discovery but against the target tenant; default output `DiscoveredTargetAccounts.csv`.
- Discovery commands support `-UseSqlite` (requires the `mySQLite` module) to also upsert results into a project-local `Discovery.sqlite` for later matching/reporting.
- `Invoke-DiscoverSourceMailboxStatistics [-ProjectPath <path>] -UseSqlite`: collect mailbox stats (size, item count, archive size, last logon, forwarding, hold flags) into `MailboxesSource` within `Discovery.sqlite` and track progress in `Tasks`. Requires Exchange connectivity and SQLite enabled.
- Async wrappers (thread jobs): `Start-AccountDiscoveryJob -Scope Source|Target [-ProjectPath] [-OutputPath] [-UseSqlite]` and `Start-CollectMailboxStatisticsJob [-ProjectPath] -UseSqlite` start background jobs that import the module and run the corresponding discovery commands. With `-UseSqlite`, job progress is tracked in the `Tasks` table inside `Discovery.sqlite`.
- `Get-MigrationTask [-ProjectPath <path>] [-TaskId <id>] [-Name <string>]`: read task rows from `Discovery.sqlite` and show totals/processed/status/percent for jobs that ran with `-UseSqlite`.
- `New-MgCustomApp`: exposed directly for advanced scenarios (manual app creation without the project scaffolding). It still honors `MG_PROFILE_DIRECTORY` and `MG_CONFIG_PATH`.
- Low-level helpers (`Connect-Source`, `Connect-Target`, `Set-MgProfileDirectory`, `Set-MgConfigPath`, etc.) stay exported if you need to build custom scripts.

### Example Workflow (full)

```powershell
Import-Module "$PSScriptRoot/MigrationToolkit.psm1" -Force

# 1. Create and activate project
$project = New-MigrationProject -Name 'Test Migration' -Path "$PWD/Projects"
Set-ActiveMigrationProject -ProjectPath $project.ProjectPath

# 2. Provision tenants (authenticates interactively in a browser once per tenant)
Add-ProjectTargetTenant -TenantId '00000000-0000-0000-0000-000000000000'
Add-ProjectSourceTenant -TenantId '11111111-1111-1111-1111-111111111111'

# 3. Use cached credentials in automation
Connect-ProjectTargetTenant
# ... run migration commands ...
Connect-ProjectSourceTenant

# 4. Switch or deactivate when done
Clear-ActiveMigrationProject
```

### Multiple Projects / Custom Paths

You can manage many projects side-by-side; each `Config.json` lives under its project root. Either `Set-ActiveMigrationProject` before calling helpers or pass `-ProjectPath` explicitly:

```powershell
Connect-ProjectSourceTenant -ProjectPath "$PWD/Projects/PmP-to-Spectrum-tenant"
Add-ProjectSourceTenant -ProjectPath "$PWD/Projects/AnotherMigration" -TenantId '2222...'
```

To reuse an existing target tenant in a new project without reprovisioning, import the target entry (and cached profile):

```powershell
Add-ProjectTargetTenant -ProjectPath "$PWD/Projects/NewMigration" -ImportProjectPath "$PWD/Projects/PmP-to-Spectrum-tenant"
```

For direct control, work with the config JSON yourself:

```powershell
$ctx = Get-MigrationProjectContext -ProjectPath "$PWD/Projects/PmP-to-Spectrum-tenant"
Connect-Source -ConfigPath $ctx.ConfigPath -TenantKey 'sourceTenant'
```

When you need additional tenants for the same project (e.g., multiple sources), run `Add-ProjectSourceTenant` again with a different `-ProjectPath` or extend the module to accept alternate tenant keys. Each run rewrites the relevant section of `Config.json`, so consider keeping backups under version control.
