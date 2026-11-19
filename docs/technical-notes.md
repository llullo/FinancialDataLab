
# FinancialDataLab — Technical Notes

These notes describe the current implementation, not a general production recommendation. See [architecture](architecture.md), [setup](setup.md), and [testing](testing.md) for system-level context and developer instructions.

## Architecture and responsibilities

- The VSTO ribbon handles Excel-facing actions and currently selects a fixed indicator, date interval, and page size.
- Application contains the observation model, calculation service, and `IFinancialObservationRepository` contract for full, numbered-page, and keyset-page reads.
- Infrastructure implements the repository with ADO.NET and SQL Server.
- The host references both Application and Infrastructure and constructs the concrete repository directly. Application has no dependency on either SQL Server or Excel, but there is not a dependency-injection container.
- The calculation service is currently independent of the ribbon's retrieval/export workflow.

## ADO.NET data access and resource lifetime

- `SqlConnection` represents a connection to SQL Server. Opening a connection obtains a connection from the provider's pool when pooling is enabled (the default for `System.Data.SqlClient`); disposing/closing it returns the logical connection to that pool rather than necessarily tearing down the physical socket. Pooling is keyed in part by the connection string.
- `SqlCommand` executes the repository's SQL text. The implementation uses named parameters with explicit SQL types and sizes, not string concatenation for values.
- `SqlDataReader` reads rows sequentially while the command and connection are open. The repository calls `GetOrdinal` once per selected column before iterating, then uses typed getters such as `GetInt32` and `GetDecimal`.
- `using` scopes dispose the connection, command, and reader even when execution or mapping throws. This is important both for timely release of pooled connections and for closing the active reader.
- `AddWithValue` is not used in the repository. If adding a parameter, prefer `Parameters.Add(name, SqlDbType, size)` (or an equivalent explicit type/precision/scale configuration) and assign its value. `AddWithValue` infers a SQL type and length from the runtime value, which can create type/length mismatches, implicit conversions, or inconsistent query plans. Keep the chosen SQL types aligned with the schema (`VARCHAR(50)`, `DATETIME2`, `INT`, and `DECIMAL(19,6)`).

## Collections, memory, and immutability

- Repository methods return `IReadOnlyList<FinancialObservation>`, which exposes read-only collection operations to callers. It does not guarantee an immutable collection or immutable elements: in this implementation the result is backed by a `List<T>`, and callers could potentially cast it back or mutate the observation objects.
- The repository materializes the complete result page into a `List<T>` before returning. Memory usage is therefore proportional to the number of returned rows; page size bounds the page methods, while `GetByIndicator` has no page limit.

## Financial values and Excel numeric precision

- The SQL column is `DECIMAL(19,6)`, and ADO.NET maps it to .NET `decimal`, which is appropriate for base-10 financial quantities requiring fixed decimal scale.
- Excel numeric cells use floating-point representation. `ExcelObservationExporter` casts the `decimal` value to `double` and writes it through `Range.Value2`; Excel retains about 15 significant digits. Six-place number formatting controls display, not the underlying precision, so decimal values should not be assumed to remain exact after export. Keep calculation and rounding policy separate from display formatting.
- The percentage-change service rejects a zero previous value. Behavior for negative starting values is mathematically defined by the formula but needs a domain policy before interpreting the result in financial terms.

## Pagination

- `GetPage` uses SQL Server `OFFSET ... FETCH`, with ordering by `ObservationDate, Id`. The ID tie-breaker makes ordering deterministic when multiple rows share a date. Offset paging is straightforward for numbered pages, but SQL Server may need to walk past preceding rows for deep offsets, and inserts/deletes between separate page requests can shift page boundaries.
- `GetNextPage` uses keyset pagination with `TOP (@PageSize)` and continues after the last row using the composite cursor `(ObservationDate, Id)`. Both cursor values must be supplied together. The predicate is `ObservationDate > @LastDate OR (ObservationDate = @LastDate AND Id > @LastId)`, matching the sort order and handling multiple observations on one date.
- The composite index `(IndicatorCode, ObservationDate, Id) INCLUDE (Value)` supports the indicator filter, ordering/cursor, and selected value in the sample schema. No execution-plan or performance benchmark has been recorded, so this is not a measured performance claim.
- Both paging methods are implemented, but the ribbon currently calls `GetPage` and keeps the page number in workbook cells; it does not use a keyset cursor.

## Date-range filtering

The repository applies inclusive comparisons (`ObservationDate >= @FromDate` and `ObservationDate <= @ToDate`) to a `DATETIME2` column. A `DateTime` whose time component is midnight used as `toDate` will not include later observations on that same calendar date. The ribbon currently uses fixed dates and supplies a midnight end value; it does not expose date-range selection in the UI. Decide whether the application means date-only days or timestamp intervals before changing this behavior. A half-open interval (`>= start` and `< day-after-end`) is a common date-only option, but is not implemented here.

## Excel export and COM behavior

`ExcelObservationExporter` constructs a two-dimensional `object[,]` and assigns it to a rectangular `Range.Value2`, avoiding one COM write per observation cell. It then applies formatting, metadata, autofit, and sheet activation with additional COM calls. The exporter clears the managed `A1:D501` area; do not assume it clears unrelated workbook content.

Excel's Office object model is COM-based and is expected to be called on Excel's owning UI/STA thread. The ribbon currently performs synchronous database access and Excel writes from the button-click flow. Moving SQL work off-thread may improve UI responsiveness, but Excel COM objects must not be accessed from an arbitrary worker thread; an asynchronous design would need to marshal worksheet operations back to Excel's UI thread and define cancellation/error behavior.

## Testing

- xUnit is used with .NET Framework 4.8.
- Unit tests exercise percentage-change behavior without Excel or SQL Server.
- Integration tests exercise the concrete ADO.NET repository against pre-seeded SQL Server in Docker. They read the password from local `Docker\\.env`; the add-in instead reads a per-user environment variable.
- Current coverage, fixture data, and limitations are documented in [testing.md](testing.md).

## Outstanding design decisions

- Should filtering be date-only or timestamp-based, and should the end date be inclusive or represented as a half-open interval?
- Should indicator, date range, and page size be configurable in Excel rather than fixed in the ribbon?
- Should the UI use keyset pagination instead of numbered OFFSET/FETCH pages, and should a cursor be represented as a dedicated type?
- What dataset size, SQL Server execution plans, and memory measurements are representative before making performance claims?
- Should repository access become asynchronous, and how should Excel COM operations be marshaled to its UI thread?
- Should integration tests use isolated database state and injected test configuration rather than shared seed rows and a local `.env` file?
- Is CI automation in scope?
