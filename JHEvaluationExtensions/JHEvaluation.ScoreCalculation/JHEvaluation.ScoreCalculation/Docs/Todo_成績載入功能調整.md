# Todo: Disable Unused Score Calculation Functions

## Goal

Modify `Program.cs` to disable the following 3 unused functions by commenting out their menu registration / Click event code.

### Functions to Disable

1. 學生端：`加總學習領域文字描述`
2. 教務批次：`批次計算學習領域成績`
3. 教務批次：`批次加總學習領域文字描述`

Do not delete the related classes.

Do not change any other score calculation logic.

---

# 1. Disable Student Function: 加總學習領域文字描述

Current code:

```csharp
mb["加總學習領域文字描述"].Click += delegate
{ new DomainTextScoreSum(NLDPanels.Student.SelectedSource).ShowDialog(); };
```

Change to:

```csharp
// 暫停使用：加總學習領域文字描述
//mb["加總學習領域文字描述"].Click += delegate
//{ new DomainTextScoreSum(NLDPanels.Student.SelectedSource).ShowDialog(); };
```

This function belongs to:

```text
學生
→ 教務
→ 成績作業
→ 加總學習領域文字描述
```

---

# 2. Disable Batch Function: 批次計算學習領域成績

Current code:

```csharp
mbAdmin["批次計算學習領域成績"].Click += delegate
{ new LearningDomainScoreCalculateByGradeyear().ShowDialog(); };
```

Change to:

```csharp
// 暫停使用：批次計算學習領域成績
//mbAdmin["批次計算學習領域成績"].Click += delegate
//{ new LearningDomainScoreCalculateByGradeyear().ShowDialog(); };
```

This function belongs to:

```text
教務作業
→ 批次作業/檢視
→ 成績作業
→ 批次計算學習領域成績
```

---

# 3. Disable Batch Function: 批次加總學習領域文字描述

Current code:

```csharp
mbAdmin["批次加總學習領域文字描述"].Click += delegate
{ new DomainTextScoreSumByGradeyear().ShowDialog(); };
```

Change to:

```csharp
// 暫停使用：批次加總學習領域文字描述
//mbAdmin["批次加總學習領域文字描述"].Click += delegate
//{ new DomainTextScoreSumByGradeyear().ShowDialog(); };
```

This function belongs to:

```text
教務作業
→ 批次作業/檢視
→ 成績作業
→ 批次加總學習領域文字描述
```

---

# Functions That Must Remain Unchanged

The following functions must continue working.

## Student Functions

Keep:

```csharp
mb["計算科目成績"].Click += delegate
{ new SubjectScoreCalculate(NLDPanels.Student.SelectedSource).ShowDialog(); };

mb["計算領域成績"].Click += delegate
{ new DomainScoreCalculate(NLDPanels.Student.SelectedSource).ShowDialog(); };

mb["計算學習領域成績"].Click += delegate
{ new LearningDomainScoreCalculate(NLDPanels.Student.SelectedSource).ShowDialog(); };
```

Especially:

```text
計算學習領域成績
```

must remain available.

Only:

```text
加總學習領域文字描述
```

is being disabled on the student side.

---

## Batch Functions

Keep:

```csharp
mbAdmin["批次計算科目成績"].Click += delegate
{ new SubjectScoreCalculateByGradeyear().ShowDialog(); };

mbAdmin["批次計算領域成績"].Click += delegate
{ new DomainScoreCalculateByGradeyear().ShowDialog(); };
```

Do not modify these functions.

---

# Expected Result

After modification:

| Function | Result |
|---|---|
| 計算科目成績 | Keep |
| 計算領域成績 | Keep |
| 計算學習領域成績 | Keep |
| 加總學習領域文字描述 | Disable |
| 批次計算科目成績 | Keep |
| 批次計算領域成績 | Keep |
| 批次計算學習領域成績 | Disable |
| 批次加總學習領域文字描述 | Disable |

---

# Important Scope Control

Do NOT:

- Delete `DomainTextScoreSum`.
- Delete `LearningDomainScoreCalculateByGradeyear`.
- Delete `DomainTextScoreSumByGradeyear`.
- Modify the internal calculation logic of those classes.
- Modify `SubjectScoreCalculate`.
- Modify `DomainScoreCalculate`.
- Modify `LearningDomainScoreCalculate`.
- Modify `SubjectScoreCalculateByGradeyear`.
- Modify `DomainScoreCalculateByGradeyear`.
- Modify ACL / permission codes.
- Modify `mbAdmin.Enable`.
- Modify student selection event handling.
- Modify graduation-related functions.
- Modify semester-history functions.
- Refactor unrelated `Program.cs` code.

Only comment out the three specified function registration blocks.

---

# Verification

After modification, compile the project and confirm there are no build errors.

Verify the student-side score menu:

```text
學生
→ 教務
→ 成績作業
```

The following functions should remain usable:

```text
計算科目成績
計算領域成績
計算學習領域成績
```

The following function should no longer be registered by this module:

```text
加總學習領域文字描述
```

Verify the batch score menu:

```text
教務作業
→ 批次作業/檢視
→ 成績作業
```

The following functions should remain usable:

```text
批次計算科目成績
批次計算領域成績
```

The following functions should no longer be registered by this module:

```text
批次計算學習領域成績
批次加總學習領域文字描述
```

If any disabled menu item still appears after removing these registrations, first confirm whether another module also registers the same menu item. Do not change unrelated modules without evidence.

---

# Completion Record

After completing and testing the modification, create:

```text
成績載入功能調整0914.md
```

Record:

1. Modified file name.
2. The three disabled functions.
3. Exact registration code that was commented out.
4. Functions intentionally kept unchanged.
5. Confirmation that no related calculation classes were deleted.
6. Build result.
7. Student menu verification result.
8. Batch menu verification result.
9. Whether any of the disabled menu items are registered by another module.
10. Any issues discovered during testing.