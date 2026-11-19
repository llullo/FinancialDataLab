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

## Building from a Clean Clone

A clean-clone build and F5 launch were verified in Visual Studio 2026. Use `main` as the default and recommended branch for cloning, building, and running the released project. The `develop` branch is used for ongoing integration work; individual changes are developed on `feature/*` branches.

1. Clone the repository. Git checks out the default branch (`main`); if you already have the repository locally, switch to it explicitly:

   ```powershell
   git clone https://github.com/llullo/FinancialDataLab.git
   cd FinancialDataLab
   git switch main
   ```

2. Open `FinancialDataLab.slnx` in Visual Studio 2026 and restore the required NuGet packages. The test projects use `packages.config`; restore them through Visual Studio/NuGet if they are not restored automatically.
3. In Solution Explorer, right-click the `FinancialDataLab` VSTO project and select **Properties**, then open **Signing**.
4. Keep **Sign the ClickOnce manifests** enabled. Click **Create Test Certificate** to generate a local development certificate. Do not enable assembly signing unless your development or release process specifically requires it.
5. Keep the generated `.pfx` private and excluded from Git. Creating the certificate may modify `FinancialDataLab.csproj`, including its manifest certificate thumbprint and/or local key reference. These machine-specific changes are for local clean-clone validation and must not be committed.
6. Rebuild the solution. If Visual Studio changed the startup project, right-click the `FinancialDataLab` VSTO project and select **Set as Startup Project**.
7. Press **F5**. Excel should launch with the add-in; verify that the **Financial Data Lab** ribbon tab appears.

This test certificate and manifest configuration are for local development/debugging only. They do not produce a production-ready signed VSTO installer. Distribution requires a separately managed release-signing certificate and an appropriate deployment/publisher trust process; do not distribute the local test `.pfx`.

After the ribbon appears, open a workbook to try the data actions. To retrieve observations, also start and seed the local SQL Server as described above and configure `FINANCIALDATALAB_SQL_PASSWORD` as described in [Local connection configuration](#local-connection-configuration). The ribbon currently requests indicator `EURUSD_DEMO` in pages of 100, from midnight November 1, 2026 through midnight March 1, 2027. The repository's end-date comparison is inclusive, but the midnight end value does not include later times on March 1. Output is written to a worksheet named `FDL Data`.

## Run tests

Use **Test > Test Explorer** in Visual Studio. Build the solution first for xUnit discovery. The unit test project runs without Excel or SQL Server. Integration tests require the running and seeded database plus `Docker\.env`; see [testing.md](testing.md) for the actual test coverage and limitations.
