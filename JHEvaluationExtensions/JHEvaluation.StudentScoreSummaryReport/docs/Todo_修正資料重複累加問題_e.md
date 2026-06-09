## Goal

Use the best fix, `ClearReportCache()`, to resolve duplicated report data generation in the English version of the Student Score Certification report.

After the modification is completed, record all changes in:

```text
校成績證明書(英文新版)調整0609.md
```

## Scope

Modify only the duplicated cached data issue in the English report.

Do not change unrelated report logic, data query logic, template merge logic, score calculation logic, absence counting rules, export behavior, or UI behavior.

## Target file

```text
PrintForm_StudentScoreCertificattion_English.cs
```

## Problem

When printing the English Student Score Certification report multiple times from the same form instance, some report data may be duplicated.

The most visible issue is absence statistics being accumulated repeatedly.

The cause is that several report data dictionaries are class-level fields and are not cleared before each print job.

Example:

```csharp
Dictionary<string, List<K12.Data.AttendanceRecord>> ar_dict = new Dictionary<string, List<K12.Data.AttendanceRecord>>();
```

`ar_dict` stores absence records by student ID.

When the report is printed again without closing the form, the program queries attendance records again and appends them to the existing `ar_dict`.

This causes the same absence records to be counted multiple times.

## Root cause

The English report contains several form-level cache dictionaries, such as:

```csharp
sr_dict
_PhotoPDict
shr_dict
ar_dict
jssr_dict
gsr_dict
urr_dict
msr_dict
```

These dictionaries persist while the form instance remains open.

When the user prints again, the data loading process runs again, but previous cached data still exists.

This can cause duplicated data, especially in `ar_dict`, because absence statistics are calculated by iterating attendance records and incrementing absence counts.

## Required fix

Use the best fix:

Create a centralized `ClearReportCache()` method and call it at the beginning of each report generation process.

## Task 1: Add ClearReportCache method

Add the following method inside the English report form class:

```csharp
private void ClearReportCache()
{
    sr_dict.Clear();
    _PhotoPDict.Clear();
    shr_dict.Clear();
    ar_dict.Clear();
    jssr_dict.Clear();
    gsr_dict.Clear();
    urr_dict.Clear();
    msr_dict.Clear();
}
```

Purpose:

* Clear all form-level cached report dictionaries before each print job.
* Prevent duplicate data from previous print operations.
* Avoid fixing only `ar_dict` while leaving similar cache risks in other dictionaries.

## Task 2: Call ClearReportCache at the beginning of MasterWorker_DoWork

Find the method:

```csharp
private void MasterWorker_DoWork(object sender, DoWorkEventArgs e)
```

At the beginning of the method, after validating `StudentIDs.Count`, call:

```csharp
ClearReportCache();
```

Recommended placement:

```csharp
private void MasterWorker_DoWork(object sender, DoWorkEventArgs e)
{
    if (StudentIDs.Count <= 0)
    {
        Feedback("", -1);
        throw new ArgumentException("沒有任何學生資料可列印。");
    }

    // Clear cached report data before loading data for this print job.
    // This prevents duplicated data when printing multiple times from the same form instance.
    ClearReportCache();

    #region 抓取學生資料
    ...
}
```

## Task 3: Keep existing absence calculation unchanged

Do not change the absence counting logic unless necessary.

Keep existing logic such as:

```csharp
absence_dic[key] += 1;
```

The issue is not the absence counting rule itself.

The issue is duplicated source data in `ar_dict`.

## Task 4: Do not change unrelated logic

Do not change:

* Student data query logic
* Semester history query logic
* Semester score query logic
* Attendance query logic
* Reward/discipline query logic
* Template merge logic
* Absence mapping rules
* Absence category names
* Grade/year mapping logic
* English field naming
* English template layout
* UI behavior
* Export behavior
* File naming behavior

Only add cache clearing before each print job.

## Validation checklist

After modification, verify the following:

### Case 1: Print once

1. Open the English Student Score Certification report form.
2. Select students.
3. Print the report once.
4. Confirm absence statistics are correct.

Expected:

```text
Absence statistics are calculated normally.
```

### Case 2: Print twice without closing the form

1. Open the English report form.
2. Select the same students.
3. Print the report once.
4. Without closing the form, print again.

Expected:

```text
The second print result should be the same as the first print result.
Absence statistics should not double.
```

### Case 3: Print three times without closing the form

1. Print the same English report three times.
2. Compare the absence statistics.

Expected:

```text
The absence statistics should remain stable.
They should not increase each time.
```

### Case 4: Other report data

Verify that the following data still displays normally:

* Student basic information
* Semester history
* Semester scores
* Domain scores
* Graduate score data
* Update record data
* Merit/demerit data
* Photo data, if applicable

### Case 5: Regression check

Confirm:

```text
No compile errors.
No report layout changes.
No unrelated data behavior changes.
No English template field behavior changes.
```

## Completion record

After completing the modification, create or update:

```text
校成績證明書(英文新版)調整0609.md
```

Record the following:

````markdown
# 校成績證明書(英文新版)調整 0609

## Issue

When printing the English Student Score Certification report multiple times from the same form instance, some report data may be duplicated.

The most visible issue is absence statistics being accumulated repeatedly.

## Root cause

Several report data dictionaries are class-level fields.

They persist while the form remains open.

When the report is printed again, the program reloads data and appends records into existing dictionaries without clearing previous print data.

The main affected dictionary is:

```csharp
ar_dict
````

Because absence statistics are calculated by iterating `ar_dict` and incrementing absence counts.

## Fix

Added a centralized cache clearing method:

```csharp
private void ClearReportCache()
{
    sr_dict.Clear();
    _PhotoPDict.Clear();
    shr_dict.Clear();
    ar_dict.Clear();
    jssr_dict.Clear();
    gsr_dict.Clear();
    urr_dict.Clear();
    msr_dict.Clear();
}
```

Called `ClearReportCache()` at the beginning of `MasterWorker_DoWork`, after validating that `StudentIDs` is not empty.

## Result

Each print job now starts with clean report cache data.

Repeated printing from the same English report form instance no longer causes absence statistics or other cached report data to be duplicated.

## Notes

Only the cache clearing logic was added.

The following logic was not changed:

* Data query logic
* Absence counting rules
* Score calculation rules
* Template merge logic
* English template field behavior
* UI behavior
* Export behavior

## Validation

Verified:

1. First print result is normal.
2. Second print from the same form does not double absence statistics.
3. Third print from the same form remains stable.
4. Other report fields still display normally.
5. No compile errors.

```
```
