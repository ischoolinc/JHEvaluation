## ADDED Requirements

### Requirement: Optimized batch SQL generation
The system SHALL use the PostgreSQL `VALUES` clause within a Common Table Expression (CTE) to handle batch course operations, replacing the legacy `UNION ALL` pattern.

#### Scenario: Generate batch delete SQL
- **WHEN** user selects multiple courses for deletion
- **THEN** the system generates a single SQL query using `WITH course_data_row(id) AS (VALUES (id1), (id2), ...)`

### Requirement: Efficient dc_bind_key count check
The system SHALL use the `VALUES` clause when checking for existing `dc_bind_key` records for multiple course IDs.

#### Scenario: Check dc_bind_key count for selection
- **WHEN** the system calculates the number of associated `dc_bind_key` records for a selection of courses
- **THEN** it performs a single optimized query using the `VALUES` clause to identify the courses
