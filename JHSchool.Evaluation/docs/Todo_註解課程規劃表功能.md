## Objective

Temporarily disable all user-facing Program Plan related features by commenting out the related UI registrations and event bindings.

Preserve the existing Program Plan data-loading and reading logic because other score-related functions may still depend on existing Program Plan assignments.

## Required Changes

### 1. Disable Program Plan Management

File: `Program.cs`

Comment out the following Program Plan management button settings and event binding:

* `教務作業 > 設定 > 課程規劃表`
* Permission key: `JHSchool.EduAdmin.Ribbon0050`
* Opening of `ProgramPlanManager`

Also comment out the related permission registration:

```csharp
ribbon.Add(new RibbonFeature("JHSchool.EduAdmin.Ribbon0050", "課程規劃表"));
```

### 2. Disable Creating Courses by Program Plan

File: `Program.cs`

Comment out the event binding for:

```text
班級 > 教務 > 班級開課 > 依課程規劃表開課
```

Do not disable the shared `班級開課` button or the `直接開課` function.

### 3. Disable Student and Class Program Plan Assignment

File: `ProgramPlan.cs`

Inside `SetupPresentation()`, comment out:

```csharp
AddAssignProgramPlanButtons();
```

This must disable both assignment functions:

* `班級 > 指定 > 課程規劃`
* `學生 > 指定 > 課程規劃`

Keep the Program Plan display fields in the student and class list panels unchanged.

### 4. Remove Assignment Permission Entries

File: `Program.cs`

Comment out the following permission registrations:

```csharp
ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0055", "課程規劃"));
ribbon.Add(new RibbonFeature("JHSchool.Class.Ribbon0055", "課程規劃"));
```

## Logic That Must Remain Unchanged

Do not comment out or modify:

```csharp
ProgramPlan.Instance.SyncAllBackground();
ProgramPlan.Instance.SetupPresentation();
```

Keep the following logic unchanged because other functions may still need to read existing Program Plan assignments:

* `GetProgramPlanRecord()`
* `FillProgramPlanRecord()`
* `GetProgramPlan()`
* Existing class `RefProgramPlanID` data
* Existing student `OverrideProgramPlanID` data
* Program Plan cache and data-query logic
* Direct course creation
* Other score-related functions

## Verification

Confirm the following after modification:

1. The Program Plan management button is no longer available.
2. `依課程規劃表開課` cannot be used.
3. `直接開課` remains available.
4. Students cannot be assigned an individual Program Plan.
5. Classes cannot be assigned a Program Plan.
6. Existing Program Plan assignments can still be read by the system.
7. Student and class Program Plan display fields still work.
8. The project builds successfully without errors.
9. No unrelated behavior is changed.

## Completion Record

After completing and verifying the changes, document all modifications and test results in:

```text
國小成績調整0615.md
```
