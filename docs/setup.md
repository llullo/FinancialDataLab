# Windows developer setup

These instructions use the checked-in solution and scripts. They require a Windows development machine with Visual Studio, desktop Excel, and Docker Desktop. The SQL service is exposed only on the local loopback address by the Compose file.

## Prerequisites

- Visual Studio with support for C#/.NET Framework projects and the VSTO/Office development components required to load and build the add-in.
- The .NET Framework 4.8 Developer Pack/targeting pack.
- Microsoft Excel desktop and the Visual Studio Tools for Office (VSTO) Runtime.
- Docker Desktop running with Linux containers available for the SQL Server image.
- SQL Server Management Studio (SSMS), used below to execute the database scripts.

Exact Visual Studio component names can vary by installer release. If the VSTO project is unavailable or cannot resolve Office references, verify the Office/VSTO development components and runtime are installed; do not replace the target framework or project references as a setup workaround.

## Start SQL Server

From the repository root in PowerShell:

```powershell
Copy-Item Docker\.env.example Docker\.env
```

Edit `Docker\.env` locally and replace the `MSSQL_SA_PASSWORD` placeholder with a strong password that satisfies SQL Server's password policy. This file is ignored by Git. Keep it local and do not paste the password into source files, screenshots, issue reports, or commits.

Start and check the container:

```powershell
docker compose --env-file Docker\.env -f Docker\compose.yaml up -d
docker compose --env-file Docker\.env -f Docker\compose.yaml ps
```

SQL Server can take a short time to become ready. If connection attempts fail, inspect its startup log:

```powershell
docker compose --env-file Docker\.env -f Docker\compose.yaml logs sqlserver
```

The Compose file publishes port `1433` on `127.0.0.1` and stores SQL Server data in a named Docker volume. Stopping the service with `docker compose --env-file Docker\.env -f Docker\compose.yaml down` retains that volume and its database.

## Create and populate the database

In SSMS, connect to server `127.0.0.1,1433` with SQL Server Authentication and user `sa`, using the password from the local `.env` file. If SSMS rejects the container's local certificate, open the connection security options and enable **Trust server certificate** for this local development connection. Do not use this trust setting as a production TLS policy.

Open and execute these scripts one at a time, in order:

1. `database\001_create_schema.sql`
2. `database\002_seed_data.sql`
3. `database\003_seed_pagination_data.sql`

The first script creates `FinancialDataLab`, the `dbo.FinancialObservations` table, and its indicator/date/ID index. The second inserts the three `EURUSD` rows used by integration tests. The third adds 300 synthetic `EURUSD_DEMO` rows spanning 100 dates, with three rows per date; these rows drive pagination tests and the add-in demo. The seed scripts include duplicate-avoidance checks for the expected sample data; note that the `EURUSD` script skips its whole insert if any `EURUSD` row already exists.

## Local connection configuration

There are two consumers of the local password, configured differently:

- **Docker Compose and integration tests:** `Docker\.env` contains `MSSQL_SA_PASSWORD`. Compose passes this value to the SQL Server container, and integration tests read this file directly.
- **Excel add-in:** The ribbon reads `FINANCIALDATALAB_SQL_PASSWORD` from the current Windows user's environment variables. Create or update that user variable in Windows **Environment Variables** settings to use the same local password. Restart Excel after changing it (and restart Visual Studio if it was already open before the change).

Do not put the password in C# code or the repository. A user environment variable and the local `.env` file are convenient local-development configuration, not a managed secrets vault; the project currently connects as `sa`. The ribbon also sets `TrustServerCertificate=true` for its localhost connection, which is not an appropriate general production setting.

## Open, build, and run Excel

1. Open `FinancialDataLab.slnx` in Visual Studio.
2. Restore NuGet packages if Visual Studio prompts; the test projects use `packages.config`.
3. Set `FinancialDataLab` (the VSTO add-in project) as the startup project and build the solution.
4. With SQL Server running, the database seeded, and the per-user password configured, press **F5** to start debugging. Visual Studio should launch Excel with the add-in.
5. Open a workbook. On the **Financial Data Lab** tab, choose **Load Observations**. The ribbon currently requests indicator `EURUSD_DEMO` in pages of 100, from midnight November 1, 2026 through midnight March 1, 2027. The repository's end-date comparison is inclusive, but the midnight end value does not include later times on March 1. Use **Previous Page** and **Next Page** to navigate. Output is written to a worksheet named `FDL Data`.

## Local VSTO manifest signing

The VSTO project has `SignManifests` enabled and points `ManifestKeyFile` to `FinancialDataLab_TemporaryKey.pfx` in the VSTO project directory. `.pfx` files are ignored by Git, so that key file is not included in a clean checkout. Since the build signs the VSTO manifests, the missing key can prevent a successful build, not only add-in launch.

On the developer machine, open the `FinancialDataLab` project properties in Visual Studio and go to **Signing**. Keep **Sign the ClickOnce manifests** enabled and use **Create Test Certificate** to create a local development certificate, or select an existing developer test certificate. Confirm Visual Studio has a valid local key file configured for the manifest signing settings (the project key path is `FinancialDataLab_TemporaryKey.pfx`); save the project properties and rebuild. Visual Studio may update the project's local certificate path or thumbprint when creating or selecting a certificate. Do not commit project-file changes made only for local signing, and never commit or distribute the `.pfx` or use a self-signed development certificate for public release signing.

## Run tests

Use **Test > Test Explorer** in Visual Studio. Build the solution first for xUnit discovery. The unit test project runs without Excel or SQL Server. Integration tests require the running and seeded database plus `Docker\.env`; see [testing.md](testing.md) for the actual test coverage and limitations.
