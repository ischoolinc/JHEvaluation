
## Goal

Use the best fix to resolve duplicated report data generation when printing the Student Score Certification report multiple times from the same form instance.

After the modification is completed, record all changes in:

```text
校成績證明書(新版)調整0609.md
```

## Scope

Modify only the duplicated cached data issue.

Do not change unrelated report logic, data query logic, template merge logic, score calculation logic, absence counting rules, or UI behavior.

## Target file

```text
PrintForm_StudentScoreCertificattion.cs
```

## Problem

When the user prints the report multiple times without closing the form, absence statistics may be accumulated repeatedly.

The main cause is that report data dictionaries are class-level fields and are not cleared before each print job.

For example:

```csharp
Dictionary<string, List<K12.Data.AttendanceRecord>> ar_dict = new Dictionary<string, List<K12.Data.AttendanceRecord>>();
```

`ar_dict` stores absence records by student ID.

During each print job, the program queries absence records again and appends them to `ar_dict`.

If `ar_dict` is not cleared before the next print job, the same absence records are added again, causing absence statistics to be counted multiple times.

## Root cause

The cached dictionaries are form-level fields.

They persist as long as the form instance remains open.

When the print button is clicked again, the data loading process runs again, but previous cached data still exists.

This can cause duplicated data in dictionaries such as:

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

The most visible issue is `ar_dict`, because absence statistics are calculated by incrementing counts from each `AttendanceRecord`.

## Required fix

Use the best fix:

Create a centralized cache clearing method and call it at the beginning of each report generation process.

## Task 1: Add ClearReportCache method

Add the following method inside the form class:

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

Keep existing rules such as:

```csharp
absence_dic[key] += 1;
```

The issue is not the counting logic itself.

The issue is duplicated source data in `ar_dict`.

## Task 4: Do not change unrelated logic

Do not change:

* Student data query logic
* Semester score query logic
* Attendance query logic
* Reward/discipline query logic
* Template merge logic
* Absence mapping rules
* Absence category names
* Grade/year mapping logic
* UI behavior
* Export behavior
* File naming behavior

Only add cache clearing before each print job.

## Validation checklist

After modification, verify the following:

### Case 1: Print once

1. Open the report form.
2. Select students.
3. Print the report once.
4. Confirm absence statistics are correct.

Expected:

```text
Absence statistics are calculated normally.
```

### Case 2: Print twice without closing the form

1. Open the report form.
2. Select the same students.
3. Print the report once.
4. Without closing the form, print again.

Expected:

```text
The second print result should be the same as the first print result.
Absence statistics should not double.
```

### Case 3: Print three times without closing the form

1. Print the same report three times.
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
```

## Completion record

After completing the modification, create or update:

```text
校成績證明書(新版)調整0609.md
```

Record the following:

````markdown
# 校成績證明書(新版)調整 0609

## Issue

When printing the Student Score Certification report multiple times from the same form instance, some report data may be duplicated.

The most visible issue is absence statistics being accumulated repeatedly.

## Root cause

Several report data dictionaries are class-level fields.

They persist while the form remains open.

When the report is printed again, the program reloads data and appends records into existing dictionaries without clearing the previous print data.

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

Repeated printing from the same form instance no longer causes absence statistics or other cached report data to be duplicated.

## Notes

Only the cache clearing logic was added.

The following logic was not changed:

* Data query logic
* Absence counting rules
* Score calculation rules
* Template merge logic
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
