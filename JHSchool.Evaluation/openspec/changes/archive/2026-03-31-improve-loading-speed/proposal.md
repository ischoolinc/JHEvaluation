## Why

The application experiences significant performance bottlenecks during startup, batch operations, and large-scale data loads, leading to unresponsive UI and long wait times. Improving performance is critical for user productivity and system scalability.

## What Changes

- **SQL Optimization**: Replace inefficient `UNION ALL` statements with the `VALUES` clause in batch operations to improve query performance and reduce database server load.
- **Startup Sequence Improvement**: Defer non-critical cache synchronization until it is actually needed, improving initial launch time.
- **Background Cache Rebuilding**: Move data-intensive lookup table rebuilding logic to background threads to prevent UI-blocking operations.
- **XML Parsing Optimization**: Streamline the parsing of large XML datasets to reduce CPU and memory consumption.
- **Event Debouncing**: Implement debouncing for high-frequency events that trigger cache reloads to prevent redundant and expensive operations.

## Capabilities

### New Capabilities
- `query-optimization`: Optimized SQL generation for batch data handling.
- `cache-efficiency`: Enhanced background processing and delayed synchronization of system-wide caches.

### Modified Capabilities
<!-- No requirement changes to existing specs as none exist. -->

## Impact

- `Program.cs`: Entry point and batch deletion logic.
- `TCInstruct.cs`: Teacher-course association cache.
- `SCAttend.cs`: Student attendance count cache.
- `ScoreCalculator.cs`: Performance-critical rounding logic for score calculations.
- `ProgramPlan.cs`: Course planning data retrieval.
