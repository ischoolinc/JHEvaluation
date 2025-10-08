# TODO: 調整 Global.cs 年排名欄位順序 (for Cursor)

## 說明
請修改 `Global.cs` 中 `// 需要調整年排名 start` 至 `// 需要調整年排名 end` 的區塊，
根據下列規則重新排序「定期」欄位。

---

## 欄位順序規則

每一組「定期」欄位順序固定為：
1. 年排名  
2. 年百分比  
3. 年PR  

依序套用到以下四種指標：
- 總分(定期)
- 加權總分(定期)
- 平均(定期)
- 加權平均(定期)

---

## 修改範例

### 目前程式結構

```csharp
builder.InsertCell();
builder.Write("總分(定期)年排名");
builder.InsertCell();
builder.Write("加權總分(定期)年排名");
builder.InsertCell();
builder.Write("平均(定期)年排名");
builder.InsertCell();
builder.Write("加權平均(定期)年排名");

// PR、百分比欄位(定期)
builder.InsertCell();
builder.Write("總分年PR(定期)");
builder.InsertCell();
builder.Write("總分年百分比(定期)");
builder.InsertCell();
builder.Write("加權總分年PR(定期)");
builder.InsertCell();
builder.Write("加權總分年百分比(定期)");
builder.InsertCell();
builder.Write("平均年PR(定期)");
builder.InsertCell();
builder.Write("平均年百分比(定期)");
builder.InsertCell();
builder.Write("加權平均年PR(定期)");
builder.InsertCell();
builder.Write("加權平均年百分比(定期)");
```

---

### 修改後應改為

```csharp
// 需要調整年排名 start

// 總分(定期)
builder.InsertCell(); builder.Write("總分(定期)年排名");
builder.InsertCell(); builder.Write("總分年百分比(定期)");
builder.InsertCell(); builder.Write("總分年PR(定期)");

// 加權總分(定期)
builder.InsertCell(); builder.Write("加權總分(定期)年排名");
builder.InsertCell(); builder.Write("加權總分年百分比(定期)");
builder.InsertCell(); builder.Write("加權總分年PR(定期)");

// 平均(定期)
builder.InsertCell(); builder.Write("平均(定期)年排名");
builder.InsertCell(); builder.Write("平均年百分比(定期)");
builder.InsertCell(); builder.Write("平均年PR(定期)");

// 加權平均(定期)
builder.InsertCell(); builder.Write("加權平均(定期)年排名");
builder.InsertCell(); builder.Write("加權平均年百分比(定期)");
builder.InsertCell(); builder.Write("加權平均年PR(定期)");

// 需要調整年排名 end
```

---

### 對應 MERGEFIELD 範例

請同步修改學生列區塊中 `MERGEFIELD` 欄位順序：

```csharp
// 總分(定期)
builder.InsertCell(); builder.InsertField("MERGEFIELD 總分_定期年排名" + studCot + " \* MERGEFORMAT ", "SYR" + studCot + "");
builder.InsertCell(); builder.InsertField("MERGEFIELD 總分年百分比_定期" + studCot + " \* MERGEFORMAT ", "SPCT" + studCot + "");
builder.InsertCell(); builder.InsertField("MERGEFIELD 總分年PR_定期" + studCot + " \* MERGEFORMAT ", "SPR" + studCot + "");

// 加權總分(定期)
builder.InsertCell(); builder.InsertField("MERGEFIELD 加權總分_定期年排名" + studCot + " \* MERGEFORMAT ", "SAYR" + studCot + "");
builder.InsertCell(); builder.InsertField("MERGEFIELD 加權總分年百分比_定期" + studCot + " \* MERGEFORMAT ", "SAPCT" + studCot + "");
builder.InsertCell(); builder.InsertField("MERGEFIELD 加權總分年PR_定期" + studCot + " \* MERGEFORMAT ", "SAPR" + studCot + "");

// 平均(定期)
builder.InsertCell(); builder.InsertField("MERGEFIELD 平均_定期年排名" + studCot + " \* MERGEFORMAT ", "AYR" + studCot + "");
builder.InsertCell(); builder.InsertField("MERGEFIELD 平均年百分比_定期" + studCot + " \* MERGEFORMAT ", "APCT" + studCot + "");
builder.InsertCell(); builder.InsertField("MERGEFIELD 平均年PR_定期" + studCot + " \* MERGEFORMAT ", "APR" + studCot + "");

// 加權平均(定期)
builder.InsertCell(); builder.InsertField("MERGEFIELD 加權平均_定期年排名" + studCot + " \* MERGEFORMAT ", "AAYR" + studCot + "");
builder.InsertCell(); builder.InsertField("MERGEFIELD 加權平均年百分比_定期" + studCot + " \* MERGEFORMAT ", "AAPCT" + studCot + "");
builder.InsertCell(); builder.InsertField("MERGEFIELD 加權平均年PR_定期" + studCot + " \* MERGEFORMAT ", "AAPR" + studCot + "");
```

---

## 驗收條件
- [ ] 所有欄位順序符合「年排名 → 年百分比 → 年PR」規則。
- [ ] 總分(定期)、加權總分(定期)、平均(定期)、加權平均(定期) 四組欄位依序排列。
- [ ] MERGEFIELD 對應欄位順序一致。
