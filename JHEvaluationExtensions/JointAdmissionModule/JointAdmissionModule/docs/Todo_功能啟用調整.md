## Objective

Update the report feature availability logic in `Program.cs` according to the current module mode.

Do not modify any unrelated business logic, report generation logic, permission behavior, event handling, or deployment mode detection.

## Target File

* `Program.cs`

## Current Mode Check

Use the existing mode condition:

```csharp
Mode == ModuleMode.KaoHsiung
```

## Required Behavior

### 1. When Running in Kaohsiung Mode

Condition:

```csharp
Mode == ModuleMode.KaoHsiung
```

Only the following report must be displayed and available:

* 台南區高中高職免試入學成績證明

The following reports must not be displayed:

* 高雄區高中高職免試入學成績證明
* 竹苗區高中高職免試入學成績證明
* 中投區高中高職免試入學成績證明

### 2. When Running in Non-Kaohsiung Mode

Condition:

```csharp
Mode != ModuleMode.KaoHsiung
```

The following reports must be displayed and available:

* 竹苗區高中高職免試入學成績證明
* 台南區高中高職免試入學成績證明
* 中投區高中高職免試入學成績證明

The following report must not be displayed:

* 高雄區高中高職免試入學成績證明

## Implementation Requirements

1. Review both report permission registration and report menu registration in `Program.Main()`.

2. Do not create or register the Kaohsiung report menu in either Kaohsiung mode or non-Kaohsiung mode.

3. Register the Tainan report menu in all modes.

4. Register the Hsinchu and Taichung report menus only when:

```csharp
Mode != ModuleMode.KaoHsiung
```

5. Preserve the existing report constructors and mode parameters:

```csharp
ModuleMode.HsinChu
ModuleMode.Tainan
ModuleMode.Taichung
```

6. Preserve the existing student selection check:

```csharp
K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0
```

7. Preserve the existing permission check:

```csharp
FISCA.Permission.UserAcl.Current[PermissionCodeRpt_HsinChu].Executable
```

8. Preserve all existing click-event behavior and report form invocation logic.

9. Do not modify:

* `DeployModeSetup()`
* `ModuleMode`
* Student detail items
* Context menu behavior
* Report content or calculation logic
* Existing permission codes unless required to prevent the hidden Kaohsiung report from being registered
* Any unrelated comments or legacy code

10. Keep the code changes minimal and limited to the report feature availability logic.

## Expected Availability Matrix

| Module Mode            | Kaohsiung Report | Hsinchu Report | Tainan Report | Taichung Report |
| ---------------------- | ---------------- | -------------- | ------------- | --------------- |
| `ModuleMode.KaoHsiung` | Hidden           | Hidden         | Available     | Hidden          |
| Non-Kaohsiung          | Hidden           | Available      | Available     | Available       |

## Verification

After modification, verify the following:

1. The project compiles without errors.
2. No duplicate menu event handlers are registered.
3. In Kaohsiung mode, only the Tainan report appears.
4. In non-Kaohsiung mode, the Hsinchu, Tainan, and Taichung reports appear.
5. The Kaohsiung report does not appear in either mode.
6. Report permissions still correctly control whether an available report button is enabled.
7. No unrelated program behavior has changed.

## Change Log

After completing the modification, create or update:

```text
高中高職免試入學成績證明調整.md
```

The change log must include:

* Modification date
* Modified file
* Original behavior
* New behavior
* Mode-based report availability matrix
* Summary of changed conditions
* Confirmation that no unrelated logic was modified
* Compilation or verification results
