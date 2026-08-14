
## Objective

Adjust the default report template loading logic for the four built-in junior high assessment report templates.

When the report configuration is loaded, check whether each default configuration uses the same template as the corresponding embedded Resource template. If the saved template is missing or different, replace it with the Resource template and save it back.

After completing the change, record the implementation details and verification results in:

```text
評量成績(固定排名)調整0617.md
```

## Target File

```text
PrintForm.cs
```

## Background

The current code only assigns the default Resource templates when the default configurations are first created.

Current logic:

```csharp
switch (name)
{
    case "領域成績單":
        cn.Template = new Document(new MemoryStream(Properties.Resources.國中評量成績單樣板_領域成績單));
        break;

    case "科目成績單":
        cn.Template = new Document(new MemoryStream(Properties.Resources.國中評量成績單樣板_科目成績單));
        break;

    case "科目及領域成績單_領域組距":
        cn.Template = new Document(new MemoryStream(Properties.Resources.國中評量成績單樣板_科目及領域成績單_領域組距));
        break;

    case "科目及領域成績單_科目組距":
        cn.Template = new Document(new MemoryStream(Properties.Resources.國中評量成績單樣板_科目及領域成績單_科目組距));
        break;
}
```

Problem:

If the default Resource templates are updated later, existing saved configurations still use the old saved templates. This can cause missing merge field values, such as:

```text
科目平時評量加權平均
科目定期評量加權平均
```

Using "Change Mail Merge Template" works because it loads the selected Word file directly into `_Configure.Template`.

## Required Changes

### 1. Add a helper method to get the Resource template by default configuration name

Create a method similar to:

```csharp
private Document GetDefaultResourceTemplate(string configName)
```

It should return the correct Resource template for these four default configurations:

```text
領域成績單
科目成績單
科目及領域成績單_領域組距
科目及領域成績單_科目組距
```

Mapping:

```csharp
"領域成績單"
=> Properties.Resources.國中評量成績單樣板_領域成績單

"科目成績單"
=> Properties.Resources.國中評量成績單樣板_科目成績單

"科目及領域成績單_領域組距"
=> Properties.Resources.國中評量成績單樣板_科目及領域成績單_領域組距

"科目及領域成績單_科目組距"
=> Properties.Resources.國中評量成績單樣板_科目及領域成績單_科目組距
```

Return `null` for non-default or custom configurations.

### 2. Add a method to compare templates

Add a method similar to:

```csharp
private bool IsSameTemplateMergeFields(Document currentTemplate, Document resourceTemplate)
```

Recommended comparison:

* Compare mail merge field names.
* Normalize field names by trimming spaces.
* Sort field names before comparing.
* Compare duplicate field occurrences as well, not only distinct field names.

Do not compare the whole Word binary content because Word metadata may differ even when the visible template is effectively the same.

### 3. Add a method to sync default templates

Add a method similar to:

```csharp
private bool SyncDefaultTemplateFromResource(Configure conf)
```

Logic:

1. Get the Resource template by `conf.Name`.
2. If no Resource template exists, return false.
3. Make sure `conf.Template` is loaded.

   * If `conf.Template == null`, call `conf.Decode()`.
4. Compare `conf.Template` with the Resource template.
5. If the saved template is missing or merge fields differ:

   * Replace `conf.Template` with the Resource template.
   * Call `conf.Encode()`.
   * Call `conf.Save()`.
   * Return true.
6. If the template is the same, do nothing and return false.

### 4. Call the sync method when configurations are loaded

After `_ConfigureList` is loaded and before the UI uses the selected configuration, loop through all configurations:

```csharp
foreach (Configure conf in _ConfigureList)
{
    SyncDefaultTemplateFromResource(conf);
}
```

Only the four default configuration names should be affected.

Do not overwrite custom configurations.

### 5. Keep existing default creation logic

Do not remove the existing default creation logic.

When no default configuration exists, the system should still create the four default configurations using the Resource templates.

The new sync logic should handle existing saved configurations that were created before the Resource templates were updated.

## Important Notes

* Do not use the external Template folder as the source of truth.
* The embedded Resource templates should be the source of truth for these four default configurations.
* Do not affect user-created custom configurations.
* Do not change report calculation logic.
* Do not change merge field data generation logic unless required by compile errors.
* Do not change the "Change Mail Merge Template" behavior.
* Do not change direct print behavior except that it should use the updated saved default template after sync.

## Verification

Please verify:

1. Existing saved default configurations are checked on load.
2. If the saved default template has different merge fields from the Resource template, it is replaced and saved.
3. If the saved default template is already the same as the Resource template, it is not rewritten.
4. Custom configurations are not modified.
5. Direct print uses the updated default template.
6. "Change Mail Merge Template" still works.
7. The four default report types still load correctly:

   * `領域成績單`
   * `科目成績單`
   * `科目及領域成績單_領域組距`
   * `科目及領域成績單_科目組距`
8. Merge fields such as the following can be populated when present in the Resource templates:

   * `科目平時評量加權平均`
   * `科目定期評量加權平均`
9. The project builds successfully.

## Completion Record

After implementation and verification, update:

```text
評量成績(固定排名)調整0617.md
```

Include:

* Files changed.
* Methods added.
* Where the sync is called.
* Verification result.
* Any remaining notes or risks.
