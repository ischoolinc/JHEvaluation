# Task: Add "Year PR Percentage" block by cloning the existing "Class PR Percentage" block

## Target file
- `PrintForm.cs`

## Goal
Create a new code section:
```csharp
// 定期年PR百分比 start
// ... (new code cloned and adapted from '定期班PR百分比')
// 定期年PR百分比 end
```
Use the existing block:
```csharp
// 定期班PR百分比 start
... (existing)
// 定期班PR百分比 end
```
as the template, then refactor all **班/Class** semantics to **年/Year** **inside the new block only**.

---

## Steps

1) **Locate template block**
- Find the block delimited by:
  - `// 定期班PR百分比 start`
  - `// 定期班PR百分比 end`
- Select (but do not modify) everything from `start` to `end`.

2) **Create new block**
- Immediately **below** the class block (or at the most appropriate nearby location), insert a **new** region delimited by:
  ```csharp
  // 定期年PR百分比 start
  // 定期年PR百分比 end
  ```
- Paste the copied contents **between** these two new markers.

3) **Constrain edits to the new block only**
- All replacements and renames below must occur **only within**
  `// 定期年PR百分比 start` ~ `// 定期年PR百分比 end`.

4) **Systematic rename (lexical) inside the new block**
Perform these **case-sensitive** replacements, in this order:

- Identifiers/words:
  - `Class` → `Year`
  - `class` → `year` (for local vars and lambda params)
- Chinese labels/comments (keep meaning consistent):
  - `班` → `年`
  - `班級` → `年級` (if it appears)
- Common identifier examples to transform (apply if present):
  - `CalculateClassPRPercent` → `CalculateYearPRPercent`
  - `BuildClassPRSeries` → `BuildYearPRSeries`
  - `classDict` → `yearDict`
  - `classScores` → `yearScores`
  - `classPRMap` → `yearPRMap`
  - `lbClassPR` / `lblClassPR` → `lbYearPR` / `lblYearPR`
  - `ClassPRTitle` → `YearPRTitle`
  - `classId`/`classID` → `yearKey`
  - `className` → `yearName`

> Keep the original class block intact; only the new year block should have these renames.

5) **Adjust grouping/aggregation semantics (logic)**
If the class block groups by Class, change the **new** year block to group by **Year**:
- Examples (apply if present):
  - `GroupBy(x => x.ClassID)` → `GroupBy(x => x.GradeYear)` or `GroupBy(x => x.Year)`
  - `var key = rec.ClassID;` → `var key = rec.GradeYear;`
  - `rec.ClassName` → `rec.GradeYear.ToString()` (or a human-readable year label)
  - Any dictionary keyed by ClassID → keyed by Year (e.g., `Dictionary<int, ...>`)

If helper accessors exist:
- Replace `GetClassPRMap()` → `GetYearPRMap()` (create a year-version if needed by cloning the class version)
- Replace `FetchClassScores(...)` → `FetchYearScores(...)`
- Replace `BuildClassPRChart(...)` → `BuildYearPRChart(...)`

6) **UI label / title strings**
Inside the new block, update UI strings (if any):
- `"定期班 PR 百分比"` → `"定期年 PR 百分比"`
- `"班"`/`"班級"` → `"年"`/`"年級"`

7) **Data source selectors**
If the original block filters by class:
- Change predicates to filter/aggregate by year:
  - `where s.ClassID == ...` → `where s.GradeYear == ...`
  - Ensure any joins keyed by class now use the year key.
  - If there is no year field, derive year from the available model (e.g., `Student.GradeYear`, `Class.GradeYear`, or a precomputed mapping).

8) **Chart/Report bindings (if present)**
- Duplicate any series/binding logic and bind to the **year-based** data set.
- Axis/category labels should show **year** (e.g., `1年`, `2年`, `3年`) rather than class names.

9) **Safety checks**
- Ensure variable names in the new block do **not** collide with the original block (scope-safe).
- Avoid accidental global replacements—limit changes strictly to the `// 定期年PR百分比` block.

10) **Build & quick validation**
- Build the solution; fix any missing method references by cloning the class-based helpers into year-based equivalents within relevant classes/modules.
- If unit/UI tests exist for class PR, create parallel tests for year PR with analogous assertions.

---

## Optional (if helpers don’t exist yet)
- Create the following year-based helpers by cloning the class-based ones and renaming inside:
  - `CalculateYearPRPercent(...)`
  - `GetYearPRMap(...)`
  - `BuildYearPRChart(...)`
  - `FetchYearScores(...)`

---

## Acceptance Criteria
- `PrintForm.cs` contains a **new** block:
  ```csharp
  // 定期年PR百分比 start
  ... year-based logic and labels ...
  // 定期年PR百分比 end
  ```
- The class block remains unchanged.
- The year block compiles and runs, producing PR% aggregated **by year** (not by class).
- UI/labels within the year block read “年/Year” semantics.
