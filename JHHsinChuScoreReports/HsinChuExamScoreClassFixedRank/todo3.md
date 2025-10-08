# Task: Add "Year PR Percentage Fill" block by cloning the existing "Class PR Percentage Fill" block

## Target file
- `PrintForm.cs`

## Goal
Create a new code section for **Year-based PR fill** by cloning the existing **Class-based PR fill** block.

**New block to add:**
```csharp
// 定期年PR百分比填值 start
// ... (new code cloned and adapted from '定期班PR百分比填值')
// 定期年PR百分比填值 end
```

**Template block (do NOT modify):**
```csharp
// 定期班PR百分比填值 start
... (existing)
// 定期班PR百分比填值 end
```
Use the template as reference, but all edits are restricted to the new "年/Year" block only.

---

## Steps

1) **Locate template block**
   - Find the region delimited by:
     - `// 定期班PR百分比填值 start`
     - `// 定期班PR百分比填值 end`
   - Select (but do **not** modify) everything between `start` and `end`.

2) **Create the new block**
   - Immediately **below** the class fill block (or the most appropriate nearby location), insert:
     ```csharp
     // 定期年PR百分比填值 start
     // 定期年PR百分比填值 end
     ```
   - Paste the copied template **between** these two markers.

3) **Constrain changes to the new block only**
   - All replacements below are limited to the region:
     `// 定期年PR百分比填值 start` ~ `// 定期年PR百分比填值 end`.

4) **Lexical renames (case-sensitive) in the new block**
   - Identifiers/words:
     - `Class` → `Year`
     - `class` → `year` (for local variables / lambda parameters)
   - Chinese labels/comments:
     - `班` → `年`
     - `班級` → `年級` (if it appears)

   - Common identifiers to transform if present (examples; apply as found):
     - `FillClassPR` → `FillYearPR`
     - `WriteClassPRValues` → `WriteYearPRValues`
     - `classPRValue` → `yearPRValue`
     - `classPRMap` → `yearPRMap`
     - `classScores` → `yearScores`
     - `classDict` → `yearDict`
     - `classId`/`classID` → `yearKey`
     - `className` → `yearName`
     - `lblClassPR`/`lbClassPR` → `lblYearPR`/`lbYearPR`
     - `ClassPRTitle` → `YearPRTitle`

5) **Adjust data selection & grouping for "fill" logic**
   - If the class block fetches/aggregates by **Class**, the new block must fetch/aggregate by **Year**:
     - `GroupBy(x => x.ClassID)` → `GroupBy(x => x.GradeYear)` or `GroupBy(x => x.Year)`
     - `var key = rec.ClassID;` → `var key = rec.GradeYear;`
     - `rec.ClassName` → `rec.GradeYear.ToString()` (or a readable year label)
     - Any dictionary keyed by ClassID → keyed by **year** (e.g., `Dictionary<int, T>`).

   - If the fill writes into UI/report cells indexed by class:
     - Switch to year-based index/loop:
       - `foreach (var kv in classPRMap)` → `foreach (var kv in yearPRMap)`
       - Category labels in UI/report should be `1年`, `2年`, `3年`, ...

6) **UI strings / labels**
   - Update any displayed strings inside the new block:
     - `"定期班 PR 百分比"` → `"定期年 PR 百分比"`
     - `"班"`/`"班級"` → `"年"`/`"年級"`

7) **Helper methods parity (if referenced)**
   - If the class block calls helpers, clone and rename equivalents for year logic:
     - `GetClassPRMap(...)` → `GetYearPRMap(...)`
     - `FetchClassScores(...)` → `FetchYearScores(...)`
     - `BuildClassPRSeries(...)` → `BuildYearPRSeries(...)`
     - `FillClassPR(...)` → `FillYearPR(...)`
   - Ensure the **year** versions are invoked **only** inside the new block.

8) **Safety checks**
   - Ensure no variable name conflicts with the original block (scope-safe).
   - Avoid global replace: verify edits occur only within the new year block.
   - Confirm that any LINQ joins or keys formerly using `ClassID` now use `GradeYear` (or the concrete field representing year in your domain).

9) **Build & quick validation**
   - Build the solution.
   - If any missing references occur, create the year-based helper variants by cloning the class-based ones with the same body but year keys.
   - Run the form/page and confirm the new **Year PR** values appear where expected.

---

## Acceptance Criteria
- `PrintForm.cs` contains a **new** block:
  ```csharp
  // 定期年PR百分比填值 start
  ... year-based PR fill logic and labels ...
  // 定期年PR百分比填值 end
  ```
- The original **Class PR fill** block remains unchanged.
- The new block compiles and correctly fills PR% **by year** (not by class).
- UI/report strings within the new block reflect **年/Year** semantics.

---

## Notes for maintainers
- If your actual year field differs (e.g., `Student.GradeYear`, `Class.GradeYear`, `Score.YearKey`), adapt the key accordingly.
- Keep both blocks co-located for future diffs; this eases review and maintenance.
