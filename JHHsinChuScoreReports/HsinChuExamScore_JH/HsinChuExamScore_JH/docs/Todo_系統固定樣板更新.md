## Goal

Replace the Word templates used by the fixed-rank assessment score notification report.

Also rename the resource references from the old `新竹_...` resource names to clear junior-high resource names, so future maintenance is easier.

After the modification is completed, record all changes in:

```text
評量成績通知單(固定排名)調整0611.md
```

## Scope

Update only the Word template resources and related resource references.

Do not change any unrelated report logic.

Do not change:

* Score calculation logic
* Ranking calculation logic
* Grade calculation logic
* Data query logic
* Merge field generation logic
* Report configuration loading logic
* Report export logic
* UI behavior
* Sorting logic

## Target files / areas

Check and update:

```text
PrintForm.cs
Properties/Resources.resx
Properties/Resources.Designer.cs
Project resource files, if Word templates are stored as file references
```

## Required template replacement and resource rename mapping

Apply the following four replacements.

### 1. Subject score report

Old resource reference:

```csharp
Properties.Resources.新竹_科目成績單
```

New Word template file:

```text
國中評量成績單樣板(科目成績單).doc
```

New resource reference:

```csharp
Properties.Resources.國中評量成績單樣板_科目成績單
```

### 2. Domain score report

Old resource reference:

```csharp
Properties.Resources.新竹_領域成績單
```

New Word template file:

```text
國中評量成績單樣板(領域成績單).doc
```

New resource reference:

```csharp
Properties.Resources.國中評量成績單樣板_領域成績單
```

### 3. Subject and domain score report: subject distribution

Old resource reference:

```csharp
Properties.Resources.新竹_科目及領域成績單_科目組距
```

New Word template file:

```text
國中評量成績單樣板(科目及領域成績單_科目組距).doc
```

New resource reference:

```csharp
Properties.Resources.國中評量成績單樣板_科目及領域成績單_科目組距
```

### 4. Subject and domain score report: domain distribution

Old resource reference:

```csharp
Properties.Resources.新竹_科目及領域成績單_領域組距
```

New Word template file:

```text
國中評量成績單樣板(科目及領域成績單_領域組距).doc
```

New resource reference:

```csharp
Properties.Resources.國中評量成績單樣板_科目及領域成績單_領域組距
```

## Resource naming rule

Use C#-friendly resource names.

Do not use parentheses in resource property names.

Do not use names like:

```csharp
Properties.Resources.國中評量成績單樣板(科目成績單)
```

Use underscore-based names instead:

```csharp
Properties.Resources.國中評量成績單樣板_科目成績單
```

## Task 1: Backup current resource files

Before modifying resources, backup:

```text
Properties/Resources.resx
Properties/Resources.Designer.cs
```

Also backup existing Word template resource files if the project stores templates as file references.

## Task 2: Inspect Resources.resx

Open:

```text
Properties/Resources.resx
```

Search for the old resource names:

```text
新竹_科目成績單
新竹_領域成績單
新竹_科目及領域成績單_科目組距
新竹_科目及領域成績單_領域組距
```

Determine whether the resources are stored as file references or embedded binary/base64 content.

### Case A: File reference

Example:

```xml
<data name="新竹_科目成績單" type="System.Resources.ResXFileRef, System.Windows.Forms">
  <value>..\Resources\xxx.doc;System.Byte[], mscorlib</value>
</data>
```

If resources are stored as file references:

1. Add or replace the referenced Word files with the new junior-high Word templates.
2. Rename the resource names to the new junior-high resource names.
3. Confirm that `Resources.Designer.cs` regenerates the new byte[] properties.

### Case B: Embedded binary/base64 content

Example:

```xml
<data name="新竹_科目成績單" type="System.Byte[], mscorlib">
  <value>...</value>
</data>
```

If resources are embedded binary/base64:

1. Replace the embedded content with the new junior-high Word templates.
2. Rename the resource names to the new junior-high resource names.
3. Confirm that `Resources.Designer.cs` regenerates the new byte[] properties.

## Task 3: Update resource names and Word template content

Remove or replace the old four resources:

```text
新竹_科目成績單
新竹_領域成績單
新竹_科目及領域成績單_科目組距
新竹_科目及領域成績單_領域組距
```

Create or update these four new resources:

```text
國中評量成績單樣板_科目成績單
國中評量成績單樣板_領域成績單
國中評量成績單樣板_科目及領域成績單_科目組距
國中評量成績單樣板_科目及領域成績單_領域組距
```

Each new resource must use the correct Word file:

```text
國中評量成績單樣板_科目成績單
-> 國中評量成績單樣板(科目成績單).doc

國中評量成績單樣板_領域成績單
-> 國中評量成績單樣板(領域成績單).doc

國中評量成績單樣板_科目及領域成績單_科目組距
-> 國中評量成績單樣板(科目及領域成績單_科目組距).doc

國中評量成績單樣板_科目及領域成績單_領域組距
-> 國中評量成績單樣板(科目及領域成績單_領域組距).doc
```

## Task 4: Update PrintForm.cs default template switch-case

Find the default template switch-case in `PrintForm.cs`.

Old code is similar to:

```csharp
// 設預設樣板
switch (name)
{
    case "領域成績單":
        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_領域成績單));
        break;

    case "科目成績單":
        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目成績單));
        break;

    case "科目及領域成績單_領域組距":
        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目及領域成績單_領域組距));
        break;

    case "科目及領域成績單_科目組距":
        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目及領域成績單_科目組距));
        break;
}
```

Replace it with:

```csharp
// 設預設樣板
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

Do not change the switch-case names:

```csharp
case "領域成績單":
case "科目成績單":
case "科目及領域成績單_領域組距":
case "科目及領域成績單_科目組距":
```

Only change the `Properties.Resources...` references.

## Task 5: Search and replace all old resource references

Search the entire project for:

```text
Properties.Resources.新竹_科目成績單
Properties.Resources.新竹_領域成績單
Properties.Resources.新竹_科目及領域成績單_科目組距
Properties.Resources.新竹_科目及領域成績單_領域組距
```

Replace them with:

```text
Properties.Resources.國中評量成績單樣板_科目成績單
Properties.Resources.國中評量成績單樣板_領域成績單
Properties.Resources.國中評量成績單樣板_科目及領域成績單_科目組距
Properties.Resources.國中評量成績單樣板_科目及領域成績單_領域組距
```

Make sure no old `新竹_...` resource references remain in source code for this report unless they are intentionally used by another report.

## Task 6: Verify Resources.Designer.cs

Check:

```text
Properties/Resources.Designer.cs
```

Confirm these generated byte[] properties exist:

```csharp
internal static byte[] 國中評量成績單樣板_科目成績單
internal static byte[] 國中評量成績單樣板_領域成績單
internal static byte[] 國中評量成績單樣板_科目及領域成績單_科目組距
internal static byte[] 國中評量成績單樣板_科目及領域成績單_領域組距
```

Also confirm there are no accidental duplicate properties such as:

```text
國中評量成績單樣板_科目成績單1
國中評量成績單樣板_領域成績單1
```

## Task 7: Verify fallback template behavior

Search for fallback usage such as:

```csharp
Properties.Resources.新竹評量成績單樣板_固定排名_科目_領域__doc1
```

Do not change it unless the new junior-high templates are also intended to replace the fallback template.

If this fallback should remain unchanged, document that it was not modified.

If it should be replaced, update it explicitly and record the mapping.

## Task 8: Do not change report logic

Do not change:

```text
Score calculation logic
Ranking calculation logic
Grade calculation logic
Data query logic
Merge field generation logic
Report configuration loading logic
Report export logic
UI behavior
Sorting logic
```

Only update:

```text
Word template resources
Resource names
Related Properties.Resources references
```

## Task 9: Build validation

Run:

```text
Clean Solution
Rebuild Solution
```

Confirm:

```text
No compile errors.
No missing resource property errors.
No duplicate resource name errors.
No broken Resources.Designer.cs generation.
```

## Task 10: Runtime validation

Test these four report settings:

```text
領域成績單
科目成績單
科目及領域成績單_領域組距
科目及領域成績單_科目組距
```

For each setting, verify:

```text
The report can be generated successfully.
The correct junior-high Word template is used.
Merge fields are populated correctly.
The report layout is correct.
No unrelated calculation or report behavior is changed.
```

## Important note about existing saved report configurations

If existing saved report configurations are stored in UDT or database storage, they may still contain old Word templates.

Replacing and renaming default resources affects newly created default configurations or fallback/default template loading.

If the old template still appears during runtime testing, check whether the report is loading from an existing saved configuration instead of the updated default resource.

Do not change saved configuration behavior unless explicitly required.

## Completion record

After completing the modification, create or update:

```text
評量成績通知單(固定排名)調整0611.md
```

Record the following:

```markdown
# 評量成績通知單(固定排名)調整 0611

## Goal

Replace and rename the Word template resources used by the fixed-rank assessment score notification report.

## Template resource changes

Updated the following resource references:

1. `Properties.Resources.新竹_科目成績單`
   - Changed to `Properties.Resources.國中評量成績單樣板_科目成績單`
   - Uses `國中評量成績單樣板(科目成績單).doc`

2. `Properties.Resources.新竹_領域成績單`
   - Changed to `Properties.Resources.國中評量成績單樣板_領域成績單`
   - Uses `國中評量成績單樣板(領域成績單).doc`

3. `Properties.Resources.新竹_科目及領域成績單_科目組距`
   - Changed to `Properties.Resources.國中評量成績單樣板_科目及領域成績單_科目組距`
   - Uses `國中評量成績單樣板(科目及領域成績單_科目組距).doc`

4. `Properties.Resources.新竹_科目及領域成績單_領域組距`
   - Changed to `Properties.Resources.國中評量成績單樣板_科目及領域成績單_領域組距`
   - Uses `國中評量成績單樣板(科目及領域成績單_領域組距).doc`

## PrintForm.cs update

Updated the default template switch-case to use the new junior-high resource names.

The switch-case report names were not changed:

- 領域成績單
- 科目成績單
- 科目及領域成績單_領域組距
- 科目及領域成績單_科目組距

## Resources update

Updated:

- `Properties/Resources.resx`
- `Properties/Resources.Designer.cs`
- Word template resource files if stored as file references

## Notes

Only Word template resources and related resource references were changed.

The following logic was not changed:

- Score calculation logic
- Ranking calculation logic
- Data query logic
- Merge field generation logic
- Report configuration loading logic
- Report export logic
- UI behavior

## Validation

Verified the following report settings:

1. 領域成績單
2. 科目成績單
3. 科目及領域成績單_領域組距
4. 科目及領域成績單_科目組距

## Result

The report now uses clearly named junior-high Word template resources, making future maintenance easier.

## Important note

If old templates still appear during runtime testing, check whether existing saved report configurations are overriding the updated default resource templates.
```
