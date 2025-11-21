# SQLite-backed discovery/matching plan

Goal: optional SQLite storage (via MySQLite module) for richer discovery, matching, and planning, while keeping CSV flow intact.

## Step 1: Baseline DB plumbing
- Add project-local DB path (`Discovery.sqlite`) and helper `Get-MigrationDatabase` that ensures the DB exists and required tables are created.
- Add MySQLite dependency check (clear error if not installed/loaded).
- Tables (initial):
  - `AccountsSource` (objectId PK, sourceUpn, displayName, mail, proxyAddresses, assignedLicenses, mailboxType, usageLocation, lastSeenUtc).
  - `AccountsTarget` (objectId PK, targetUpn, displayName, mail, proxyAddresses, assignedLicenses, mailboxType, usageLocation, lastSeenUtc).
  - `MailboxesSource` / `MailboxesTarget` (objectId PK, recipientType, size, itemCount, archiveSize, lastLogon, forwarding, holdFlags, lastSeenUtc).
  - `Metadata` (schemaVersion).

## Step 2: Persist discovery into DB (keep CSV optional)
- Extend discovery commands:
  - `Invoke-DiscoverSourceAccounts -UseSqlite`: upsert users into `AccountsSource`; capture EXO recipient info into `MailboxesSource` if Exchange is reachable. Continue to produce CSV when `-OutputPath` is set (default `DiscoveredAccounts.csv`).
  - Add `Invoke-DiscoverTargetAccounts -UseSqlite` mirroring source.
- Add a helper `Get-DiscoveryDatabaseContext` or reuse `Get-MigrationDatabase` to return DB path and ensure tables exist.
- Prefer INSERT OR REPLACE for upsert keyed on `objectId`.
- Keep proxy addresses and license SKUs as comma-separated text columns.
- Record `lastSeenUtc` on each discovery run.

## Future Steps (not in this batch)
- Matching workflow: `Find-AccountMatches`, `Set-AccountMatch`, `Get-MigrationPlan`, `matches` table.
- Migration integration: `Invoke-MigrateUser -FromPlan`, update matchStatus to migrated.
- Reports/exports: CSV/XLSX from DB, optional ImportExcel support.
- DB maintenance: reset/backup, schema versioning.
