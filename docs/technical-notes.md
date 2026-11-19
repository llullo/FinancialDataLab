
# FinancialDataLab — Technical Notes

## Architecture
- VSTO handles Excel interactions and presentation.
- Application contains services, models and repository contracts.
- Infrastructure implements data access with ADO.NET.
- Dependencies point towards Application, avoiding circular references.

## Data access
- SqlConnection manages database connections and uses connection pooling.
- SqlCommand executes parameterized SQL.
- SqlDataReader reads results sequentially.
- GetOrdinal resolves column positions outside the row loop.
- using ensures disposal of connections, commands and readers.

## Collections and memory
- IReadOnlyList<T> exposes read-only collection operations but does not guarantee immutability.
- Materializing all rows into List<T> increases memory usage with the result count.
- Keyset pagination limits results and continues using a composite cursor.

## Financial calculations
- decimal is suitable for many exact base-10 financial values.
- Rounding and presentation should be separated from calculations.
- Percentage change requires a documented policy for zero and negative starting values.

## Testing
- xUnit is used with .NET Framework 4.8.
- Unit tests run without Excel or SQL Server.
- Integration tests exercise the real ADO.NET repository against SQL Server in Docker.

## Open decisions
- Date-only versus date-time filtering.
- Cursor abstraction and larger pagination test datasets.
- SQL Server execution-plan and memory benchmarks.
- Async database access and Excel COM thread affinity.
- Local test configuration and CI automation.
