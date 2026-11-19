# Testing

The test projects target .NET Framework 4.8 and use xUnit 2.9.3 with the Visual Studio test adapter. Both use `packages.config` and the Visual Studio Test Platform packages rather than an SDK-style project definition.

## Unit tests

`FinancialDataLab.UnitTests` tests `FinancialCalculationService` without external runtime services: it does not require Excel, SQL Server, or Docker. NuGet packages and the Visual Studio test adapter still need to be restored for discovery/execution. Current cases cover positive and negative percentage changes, unchanged values, and the exception when the previous value is zero.

## Integration tests

`FinancialDataLab.IntegrationTests` exercises the concrete ADO.NET repository against SQL Server. It checks full retrieval, numbered-page values, a keyset page continuation without duplicate IDs, invalid page-number validation, and a 100-row page from the synthetic dataset. These tests require a reachable local database; they do not mock `SqlConnection` or `SqlDataReader`.

The tests locate the solution root and read `Docker\.env`, extracting the `MSSQL_SA_PASSWORD` line. They connect to `127.0.0.1,1433`, database `FinancialDataLab`, using the `sa` login. Set up that file and the database as described in [setup.md](setup.md). The test data comes from:

1. `database\001_create_schema.sql` — database, table, and index.
2. `database\002_seed_data.sql` — three `EURUSD` rows dated November 16-18, 2026, with the values asserted by the tests.
3. `database\003_seed_pagination_data.sql` — 300 synthetic `EURUSD_DEMO` observations for pagination tests and the add-in.

The seed data is fixed test data, not a live financial feed. The test suite assumes it exists and is not designed to initialize the database automatically.

## Run from Visual Studio

1. Open `FinancialDataLab.slnx` in Visual Studio and allow NuGet packages to restore.
2. Build the solution to enable test discovery.
3. Open **Test > Test Explorer**.
4. Run `FinancialDataLab.UnitTests` independently; no local SQL Server, Excel, or Docker service is needed (the test packages and adapter must be available).
5. To run `FinancialDataLab.IntegrationTests`, start the Compose SQL Server container, execute the three SQL scripts in order, and ensure the test process can read `Docker\.env`. Select and run the integration test project or its tests in Test Explorer.

## Current limitations and possible improvements

- Coverage is small and focused on a single calculation service and the SQL repository. Excel COM behavior, ribbon validation/messages, worksheet contents/formatting, and failure paths are not currently covered by automated tests.
- The integration suite depends on a local Docker SQL Server, a fixed port, a specific `.env` file location/format, and pre-seeded database state. Failures due to missing local setup are not automatically skipped or provisioned.
- SQL assertions rely on fixed sample data; tests can interfere if that database is changed manually or shared with unrelated work.
- The repository offers both OFFSET/FETCH and keyset paging, but the tests use the current fixed sample set and do not establish comparative performance or concurrency guarantees.
- There is no documented CI workflow or automated test-database lifecycle in the solution.

Possible follow-ups include isolating integration data per test run, injecting test configuration instead of reading a local file, adding coverage for empty/boundary/error cases and Excel behavior behind a seam, and adding CI orchestration. These are suggestions, not existing capabilities.
