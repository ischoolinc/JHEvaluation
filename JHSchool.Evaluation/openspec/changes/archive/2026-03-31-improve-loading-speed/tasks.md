## 1. Batch Operation Optimization

- [x] 1.1 Update `DeleteCourseWithLogSQL` in `Program.cs` to use the `VALUES` clause for `course_data_row`.
- [x] 1.2 Update `GetDcBindKeyCountByCourseIDs` in `Program.cs` to use the `VALUES` clause for `course_data_row`.
- [x] 1.3 Verify that the SQL generates correctly and that batch course deletion still works as expected.

## 2. Startup Performance

- [x] 2.1 Comment out eager `SyncAllBackground()` calls for `TCInstruct`, `ProgramPlan`, `ScoreCalcRule`, and `AssessmentSetup` in `Program.Main`.
- [x] 2.2 Verify that the application startup time is noticeably improved.
- [x] 2.3 Ensure that the Teacher column in the course list still loads on-demand (shows "Loading..." and then populates).

## 3. Background Cache Processing

- [x] 3.1 Modify `TCInstruct.GetAllData` to rebuild `_course_teachers` and `_teacher_courses` lookups on the background thread.
- [x] 3.2 Update `TCInstruct.GetData` to also benefit from background lookup table management.
- [x] 3.3 Remove the redundant lookup rebuilding logic from the `TCInstruct.ItemLoaded` event on the UI thread.
- [x] 3.4 Implement atomic reference swapping for the lookup dictionaries to ensure thread safety.

## 4. Cache Update Efficiency

- [x] 4.1 Optimize XML parsing in `TCInstruct.cs` to use direct `SelectSingleNode` access instead of creating new `DSXmlHelper` instances.
- [x] 4.2 Optimize the `TCInstruct.ItemUpdated` event to selectively update lookup tables rather than performing full scans.

## 5. Event Debouncing

- [x] 5.1 Implement a 1-second debounce timer for `SCAttend.ReloadCourseStudentCount`.
- [x] 5.2 Implement a 2-second debounce timer for the `JH_CourseTeacherChange` event listener in `Program.cs`.
- [x] 5.3 Verify that redundant server calls are eliminated during batch operations (e.g., adding several students to a course).

## 6. Score Calculation Optimization

- [x] 6.1 Optimize the `GetRoundScore` method in `ScoreCalculator.cs` to avoid slow string conversions and redundant math operations.
- [x] 6.2 Verify that the rounding results remain identical to the legacy implementation.
