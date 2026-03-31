## ADDED Requirements

### Requirement: Deferred startup cache synchronization
The system SHALL not automatically sync all data caches (`TCInstruct`, `ProgramPlan`, `ScoreCalcRule`, `AssessmentSetup`) at application startup. Synchronization SHALL only occur when the data is first requested or a specific refresh event is triggered.

#### Scenario: Application launch
- **WHEN** the application starts
- **THEN** the initial eager `SyncAllBackground()` calls for non-critical caches are omitted, reducing the "time-to-interactive"

### Requirement: Background rebuilding of teacher-course lookups
The system SHALL rebuild the `_course_teachers` and `_teacher_courses` lookup dictionaries on a background thread during data retrieval. The system SHALL use atomic reference swapping to ensure UI components always access a consistent version of the lookup tables.

#### Scenario: Full cache load
- **WHEN** the `TCInstruct` cache performs a `GetAllData` operation
- **THEN** it builds its associated lookup dictionaries before returning the result to the main thread

### Requirement: Optimized XML parsing for teacher-course data
The system SHALL use direct `SelectSingleNode` access when parsing `TCInstruct` records to minimize CPU and memory overhead.

#### Scenario: Parse TCInstruct records
- **WHEN** the system iterates through XML records from the `GetTCInstruct` service
- **THEN** it retrieves child node values directly without creating a new `DSXmlHelper` instance for every record

### Requirement: Event debouncing for high-frequency cache refreshes
The system SHALL implement a 1-second debounce for `ReloadCourseStudentCount` and a 2-second debounce for the `JH_CourseTeacherChange` event listener.

#### Scenario: Rapid sequence of student attendance updates
- **WHEN** multiple `JHSCAttend` records are updated in a short interval (e.g., during batch operations)
- **THEN** the `ReloadCourseStudentCount` method is called only once after 1 second of inactivity

#### Scenario: Batch course teacher changes
- **WHEN** the `JH_CourseTeacherChange` event fires multiple times within 2 seconds
- **THEN** the `TCInstruct.SyncAllBackground()` method is called only once
