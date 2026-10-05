# 目標

調整 `ImportExamScore.cs` 的「匯入評量成績」驗證邏輯。

需求重點：

- 只放寬「平時評量」的驗證。
- 當 `評量名稱 = 平時評量` 時：
  - 不需要課程有 `RefAssessmentSetupID`。
  - 不需要檢查 `JHAEInclude` 評量設定。
  - 不需要檢查系統 Exam 是否存在「平時評量」。
  - 只要學生確實有修習該學年度、學期、課程，即允許匯入。
- 一般段考／其他評量的原本驗證邏輯必須完全保留。
- 不要修改平時評量目前的寫入方式。
- 修改完成後，將調整內容記錄在：
  `匯入評量成績調整1002.md`

---

# 修改檔案

`ImportExport/ImportExamScore.cs`

---

# 現況問題

目前在 `ValidateRow` 中，學生有修課後會進入「驗證評量是否存在」：

```csharp
if (string.IsNullOrEmpty(attendCourse.RefAssessmentSetupID))
{
    if (!e.ErrorFields.ContainsKey("無評量設定"))
        e.ErrorFields.Add("無評量設定", "課程(" + attendCourse.Name + ")無評量設定");
}
else
{
    if (!courseAe.ContainsKey(attendCourse.RefAssessmentSetupID))
    {
        if (!e.ErrorFields.ContainsKey("無評量設定"))
            e.ErrorFields.Add("無評量設定", "課程(" + attendCourse.Name + ")無評量設定");
    }
    else
    {
        bool examValid = false;

        foreach (JHAEIncludeRecord ae in courseAe[attendCourse.RefAssessmentSetupID])
        {
            if (!exams.ContainsKey(ae.RefExamID)) continue;

            if (exams[ae.RefExamID].Name == examName || examName == "平時評量")
                examValid = true;
        }

        if (!examValid)
        {
            if (!e.ErrorFields.ContainsKey("評量名稱無效"))
                e.ErrorFields.Add("評量名稱無效", "評量名稱(" + examName + ")不存在系統中");
        }
    }
}
```

目前即使：

```text
評量名稱 = 平時評量
```

只要：

```text
attendCourse.RefAssessmentSetupID
```

為空白，就會先出現：

```text
無評量設定
```

導致平時評量無法匯入。

---

# 調整需求

將「平時評量」從課程評量設定驗證中獨立出來。

建議邏輯：

```csharp
if (examName != "平時評量")
{
    // 保留原本一般評量的驗證
}
```

也就是：

```text
評量名稱 = 平時評量
    ↓
略過 RefAssessmentSetupID 驗證
    ↓
略過 courseAe 驗證
    ↓
略過 Exam 名稱驗證
    ↓
允許進入後續匯入流程
```

一般評量：

```text
第一次評量
第二次評量
第三次評量
其他正式評量
```

仍必須依照原本流程驗證：

```text
RefAssessmentSetupID
→ courseAe
→ Exam
```

---

# 建議修改方式

將原本：

```csharp
#region 驗證評量是否存在
...
#endregion
```

調整成類似：

```csharp
#region 驗證評量是否存在

// 平時評量不需要課程的段考評量設定。
// 只要學生有修習該課程，即允許匯入平時評量。
if (examName != "平時評量")
{
    if (string.IsNullOrEmpty(attendCourse.RefAssessmentSetupID))
    {
        if (!e.ErrorFields.ContainsKey("無評量設定"))
            e.ErrorFields.Add(
                "無評量設定",
                "課程(" + attendCourse.Name + ")無評量設定"
            );
    }
    else
    {
        if (!courseAe.ContainsKey(attendCourse.RefAssessmentSetupID))
        {
            if (!e.ErrorFields.ContainsKey("無評量設定"))
                e.ErrorFields.Add(
                    "無評量設定",
                    "課程(" + attendCourse.Name + ")無評量設定"
                );
        }
        else
        {
            bool examValid = false;

            foreach (JHAEIncludeRecord ae in courseAe[attendCourse.RefAssessmentSetupID])
            {
                if (!exams.ContainsKey(ae.RefExamID))
                    continue;

                if (exams[ae.RefExamID].Name == examName)
                {
                    examValid = true;
                    break;
                }
            }

            if (!examValid)
            {
                if (!e.ErrorFields.ContainsKey("評量名稱無效"))
                    e.ErrorFields.Add(
                        "評量名稱無效",
                        "評量名稱(" + examName + ")不存在系統中"
                    );
            }
        }
    }
}

#endregion
```

---

# 不要修改的部分

以下原本邏輯必須保留。

## 1. 學生驗證

仍需確認學生存在。

## 2. 修課驗證

仍需確認學生有修習：

```text
學年度
學期
課程名稱
```

對應的課程。

如果學生沒有修課，仍應出現：

```text
無修課記錄
```

## 3. 一般評量驗證

非「平時評量」時：

```text
第一次評量
第二次評量
第三次評量
...
```

仍必須檢查課程評量設定，不可一起放寬。

## 4. 平時評量寫入邏輯

不要修改既有：

```csharp
record.OrdinarilyScore
record.OrdinarilyEffort
JHSCAttend.Update(record)
```

相關邏輯。

也不要改成寫入 `JHSCETake`。

平時評量仍然寫入修課紀錄。

## 5. 一般評量寫入邏輯

一般段考仍維持：

```text
JHSCETake
```

新增／更新流程，不要修改。

## 6. 努力程度轉換

目前沒有「努力程度」欄位時，自動依照分數轉換努力程度的邏輯維持不變。

---

# 驗證案例

請至少確認以下案例。

### Case 1：沒有評量設定 + 平時評量

```text
課程 RefAssessmentSetupID = null
評量名稱 = 平時評量
分數評量 = 85
```

預期：

```text
可以匯入
不顯示「無評量設定」
OrdinarilyScore = 85
```

---

### Case 2：沒有評量設定 + 一般段考

```text
課程 RefAssessmentSetupID = null
評量名稱 = 第一次評量
```

預期：

```text
不可匯入
仍顯示「無評量設定」
```

---

### Case 3：有評量設定 + 正確段考名稱

```text
課程有評量設定
設定內有「第一次評量」
匯入評量名稱 = 第一次評量
```

預期：

```text
正常匯入
原本邏輯不變
```

---

### Case 4：有評量設定 + 不存在的段考名稱

```text
課程有評量設定
設定內沒有「第四次評量」
匯入評量名稱 = 第四次評量
```

預期：

```text
不可匯入
顯示「評量名稱無效」
```

---

### Case 5：沒有修課 + 平時評量

```text
學生沒有修習該課程
評量名稱 = 平時評量
```

預期：

```text
不可匯入
仍顯示「無修課記錄」
```

---

# 注意事項

- 本次只處理「平時評量不需要段考評量設定」。
- 不要大幅重構 `ImportExamScore.cs`。
- 不要改動其他匯入格式。
- 不要改動 RequiredFields。
- 不要改動一般段考寫入方式。
- 不要改動 `OrdinarilyScore` / `OrdinarilyEffort` 儲存位置。
- 優先採最小修改方式，避免影響既有使用者。

---

# 完成紀錄

修改完成後新增或更新：

`匯入評量成績調整1002.md`

至少記錄：

- 修改原因
- 原本平時評量為何會被「無評量設定」擋住
- 修改的驗證條件
- 平時評量改為不檢查 `RefAssessmentSetupID`
- 一般評量仍維持原本驗證
- 測試案例與測試結果
- 實際修改檔案名稱