 A tidy split that still feels PowerShell-ish:

  - Core module (MigrationToolkit.psm1) only re-exports; real code lives in submodules under src/.
  - src/Config.psm1: project context, active project state, config read/write, default paths.
  - src/Database.psm1: SQLite bootstrap/migrations, table access helpers (query/exec wrappers), typed getters for each table, export helpers.
  - src/Auth.psm1: connect/provision/import tenant profiles (source/target), Exchange connect helpers.
  - src/Discovery.psm1: discover source/target accounts + mailbox stats; product sheet import helpers.
  - src/Matching.psm1: displayName matching, reconciliation/status normalization.
  - src/Migration.psm1: migrate single/batch users, target UPN resolution rules, license application, SetProperties merge.
  - src/Tasks.psm1: background wrappers (Start-Project*Task), task row lifecycle (create/update/fail/complete), shared job utilities.
  - src/Exports.psm1: export/get commands for tables, CSV output.
  - Cross-cutting: a small Utils.psm1 for error handling/logging, Graph client setup, and retries.

  Mechanics:

  - Keep each submodule self-contained, importing the minimal neighbors (e.g., Migration imports Auth/Database/Config).
  - Add Pester tests per submodule (mocks for Graph/SQLite) and a simple build script to bundle/re-export.
  - Consider putting shared models/constants (table names, status values) in a single Constants.psm1 to avoid string drift.