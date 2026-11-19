# Architecture

FinancialDataLab is divided into an Excel/VSTO host, an application project, and an infrastructure project. The split keeps SQL commands out of the model and repository contract, but the current host still constructs the concrete repository itself; this is a small layered solution rather than a fully injected composition root.

## VSTO presentation layer

`FinancialDataLab/FinancialDataRibbon.cs` supplies three ribbon buttons: **Load Observations**, **Previous Page**, and **Next Page**. Load starts at page 1; the navigation buttons read the current page from cell `F2` on the active workbook's `FDL Data` worksheet and request an adjacent numbered page. The handler checks for an active workbook, reads the password from the current user's `FINANCIALDATALAB_SQL_PASSWORD` environment variable, constructs the connection settings and repository, and requests `EURUSD_DEMO` for the fixed November 1, 2026 to March 1, 2027 interval with a page size of 100. It calls `GetPage`, which uses SQL Server `OFFSET/FETCH`; the repository's keyset method is not used by the ribbon.

`FinancialDataLab/Excel/ExcelObservationExporter.cs` owns worksheet interaction. It creates the `FDL Data` sheet if needed or reuses it, clears `A1:D501`, writes the four data columns starting at `A1`, and formats the heading, date, and value columns. It stores current page and row count in `F1:F2` and `G1:G2`, then autofits columns `A:G` and activates the sheet. It does not clear unrelated worksheet cells; in particular, its data clear range is `A1:D501`. If a page query returns no rows, the ribbon displays a message and does not call the exporter, so the prior worksheet data and page marker remain. Workbook and worksheet objects are Excel COM objects used from the ribbon/Excel interaction path.

## Application services, models, and repository contract

`FinancialDataLab.Application/Models/FinancialObservation.cs` is the observation model (`Id`, indicator code, observation date, and decimal value). `Interfaces/IFinancialObservationRepository.cs` declares operations for retrieving all matching observations, an OFFSET/FETCH page, and the next keyset page. Its signatures expose `IReadOnlyList<FinancialObservation>` rather than SQL or Excel types.

`Services/FinancialCalculationService.cs` calculates percentage change and throws when the previous value is zero. It has no repository or Excel dependency. The current ribbon workflow does not call this service.

## Infrastructure and ADO.NET

`FinancialDataLab.Infrastructure/Repositories/FinancialObservationRepository.cs` implements the application repository interface with `System.Data.SqlClient`. It validates the indicator, date range, and paging arguments; opens a `SqlConnection`; builds commands with explicitly typed SQL parameters; and maps `SqlDataReader` columns into observation objects. Connections, commands, and readers are disposed with `using` scopes. Reads are ordered by `ObservationDate, Id` for deterministic paging.

The SQL schema and supporting index are in `database/001_create_schema.sql`. The index keys are `(IndicatorCode, ObservationDate, Id)` and include `Value`, matching the main filter/order columns.

## Dependency direction and responsibilities

- The Application project defines the observation model, service, and repository abstraction; it has no dependency on the Excel host or SQL repository implementation.
- Infrastructure references Application and implements its repository interface.
- The VSTO project references Application and Infrastructure. Its ribbon directly constructs `FinancialObservationRepository`, so the host is responsible for current composition and connection-string creation; no dependency-injection container is present.
- The Excel exporter accepts observations and writes them to Excel; it does not query SQL Server.

This keeps data access and worksheet formatting in distinct classes, while leaving a future option to introduce injected composition or move date/indicator selection out of the ribbon.

## Excel export and `Range.Value2`

The exporter builds an `object[,]` containing a header row and the requested observations, then assigns that array to the destination range through `Range.Value2`. A single rectangular assignment avoids one COM call per data cell. Formatting and metadata are applied with additional range operations.

The database and application model use `DECIMAL(19,6)`/`decimal`, but the exporter casts the value to `double` for Excel. Excel stores numeric values using floating-point semantics and supports about 15 significant decimal digits; formatting to six decimal places does not make the underlying Excel number an exact decimal. The exporter also sends the observation date as an OLE Automation date number.

## Screenshots

The repository contains screenshots of the ribbon and two sample data pages:

- [Excel ribbon](images/excel-ribbon.png)
- [First exported page](images/excel-first-page.png)
- [Second exported page](images/excel-pagination.png)

A database-setup screenshot has not yet been created; its planned path is `docs/images/database-setup.png`.

See [technical notes](technical-notes.md) for data access, pagination, and open design questions.
