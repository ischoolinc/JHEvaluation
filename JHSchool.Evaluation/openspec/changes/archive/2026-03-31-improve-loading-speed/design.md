## Context

The application currently performs numerous synchronous operations during startup and batch data handling. Large-scale datasets cause significant UI lag (up to several seconds) when lookup tables are rebuilt or when batch SQL queries are generated using thousands of `UNION ALL` statements.

## Goals / Non-Goals

**Goals:**
- Reduce application startup time by deferring non-critical data loading.
- Improve batch operation performance by optimizing SQL generation.
- Eliminate UI freezes during cache updates by moving logic to background threads.
- Reduce redundant server calls through event debouncing.

**Non-Goals:**
- Complete architectural rewrite of the cache management system.
- Database schema changes (the focus is on application-level and query-level optimization).
- Optimization of unrelated modules not identified as bottlenecks.

## Decisions

### 1. SQL Optimization: `VALUES` Clause vs. `UNION ALL`
**Decision:** Use the `VALUES` clause in `WITH` common table expressions (CTEs) for batch course processing.
**Rationale:** The current `SELECT ... UNION ALL` approach is exponentially slower to parse and execute by the database as the number of items increases. `VALUES` is more concise and specifically designed for batch data input in PostgreSQL.
**Alternatives:** Temporary tables (requires multiple steps and can be complex to manage in a stateless query model).

### 2. Lazy Cache Synchronization
**Decision:** Defer `SyncAllBackground()` for `TCInstruct`, `ProgramPlan`, `ScoreCalcRule`, and `AssessmentSetup` at startup.
**Rationale:** Initializing all caches simultaneously creates a resource contention at startup. By deferring synchronization until the first time the data is accessed (e.g., when a relevant UI component loads), the application becomes interactive much faster.
**Alternatives:** Sequential background synchronization (still consumes resources early and may not be necessary if the user doesn't access certain features).

### 3. Background Cache Rebuilding with Thread Safety
**Decision:** Move `_course_teachers` and `_teacher_courses` rebuilding logic in `TCInstruct` to the `GetAllData` method (background thread) and use atomic reference swaps.
**Rationale:** This removes the heavy XML-to-Dictionary processing from the UI thread. Atomic swaps ensure that UI components always see a consistent state without needing explicit locks for reading.
**Alternatives:** BackgroundWorker in the UI class (leads to more complex UI code and harder-to-manage state).

### 4. Event Debouncing for Cache Reloads
**Decision:** Implement a 1-2 second debounce timer for `ReloadCourseStudentCount` and `JH_CourseTeacherChange` events.
**Rationale:** In batch operations, events can fire hundreds of times in rapid succession. Debouncing ensures that only one server request is made after the burst of activity completes.
**Alternatives:** Manual suppression flags (error-prone and harder to maintain than a general debouncing pattern).

## Risks / Trade-offs

- [Risk] Data may be slightly stale if the user interacts with a feature immediately before a deferred sync completes.  Mitigation: Provide visual "Loading..." indicators in relevant UI columns.
- [Risk] Debouncing adds a small (1-2s) delay before the latest data is reflected.  Mitigation: This is an acceptable trade-off for significantly better system stability and responsiveness.
- [Risk] Background threads accessing shared state.  Mitigation: Use thread-safe collections and atomic reference updates.
