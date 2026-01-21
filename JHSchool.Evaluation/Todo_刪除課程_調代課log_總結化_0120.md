# Todo.md — 刪除課程「調代課對應 log」改為每課程一筆總結（0120）

## 目標
- 修正刪除課程時「調代課對應資料（dc_bind_key）」刪除 log 目前會產生多筆的問題。
- 調整後：**同一課程只寫入 1 筆總結 log**（可包含刪除筆數）。
- 修改完成後，將變更記錄整理到：`刪除課程與調代課對應資料調整0120.md`

## 範圍
- 專案：`JHSchool.Evaluation`
- 檔案：`Program.cs`
- 方法：`DeleteCourseWithLogSQL()`
- SQL 片段：CTE `insert_dc_bind_key_log`（與其依賴的 `delete_dc_bind_key`）

---

## Step 1 — 找到目前「調代課刪除 log」產生多筆的原因
- 開啟 `Program.cs`
- 定位 `DeleteCourseWithLogSQL()`
- 找到 SQL CTE：
  - `delete_dc_bind_key AS ( DELETE FROM dc_bind_key ... RETURNING dc_bind_key.* )`
  - `insert_dc_bind_key_log AS ( INSERT INTO log ... FROM delete_dc_bind_key )`
- 確認目前行為：`delete_dc_bind_key` 每刪一筆回傳一列，導致 `insert_dc_bind_key_log` 逐列 INSERT → 產生多筆 log。

驗收：已清楚標註目前多筆 log 的 SQL 驅動來源。

---

## Step 2 — 修改 SQL：將 insert_dc_bind_key_log 改為「每課程彙總一筆」
### 2.1 修改原則
- 不改動 `delete_dc_bind_key` 的刪除邏輯（避免影響資料刪除結果）。
- 僅調整 `insert_dc_bind_key_log` 的 SELECT 來源：
  - 由「逐筆」改為「GROUP BY ref_course_id」彙總。
- 建議在 description 加上「共刪除 N 筆」，提升稽核可讀性。

### 2.2 建議替換片段（整段替換 insert_dc_bind_key_log）
> 在 SQL 字串中，將原本的 `,insert_dc_bind_key_log AS(` 這段完整替換為下列版本。

```sql
,insert_dc_bind_key_log AS(
INSERT INTO log(
    actor
    , action_type
    , action
    , target_category
    , target_id
    , server_time
    , client_info
    , action_by
    , description
)
SELECT
    '{1}'::TEXT AS actor
    , 'Record' AS action_type
    , '課程_刪除' AS action
    , 'course'::TEXT AS target_category
    , d.ref_course_id AS target_id
    , now() AS server_time
    , '{2}' AS client_info
    , '刪除_課程_調代課' AS action_by
    , '課程調代課資料刪除（總結），課程ID「'
      || d.ref_course_id
      || '」，共刪除 '
      || d.del_cnt
      || ' 筆，使用者「{1}」' AS description
FROM (
    SELECT
        delete_dc_bind_key.ref_course_id,
        COUNT(*) AS del_cnt
    FROM delete_dc_bind_key
    GROUP BY delete_dc_bind_key.ref_course_id
) d
)
```

驗收：刪除同一課程但 `dc_bind_key` 有多筆時，log 僅新增 1 筆（每課程 1 筆）。

---

## Step 3 — 測試案例（必做）
### Case A：單一課程，dc_bind_key 有 0 筆
1. 選取 1 門課程（確認該課程沒有任何調代課對應資料）
2. 執行刪除課程
3. 檢查 log：
   - 不應產生 `刪除_課程_調代課` 這筆（因為沒有刪到 dc_bind_key）

預期：0 筆調代課刪除 log。

### Case B：單一課程，dc_bind_key 有 3 筆
1. 建立/選取 1 門課程，並確保 `dc_bind_key.ref_course_id = 課程ID` 有 3 筆
2. 執行刪除課程
3. 檢查 log：
   - `action_by = '刪除_課程_調代課'` 僅 1 筆
   - description 內含 `共刪除 3 筆`

預期：調代課刪除 log = 1 筆（總結）。

### Case C：一次刪除多門課程（每門都有多筆 dc_bind_key）
1. 一次選取 3 門課程刪除
2. 每門課 `dc_bind_key` 各有 N 筆
3. 檢查 log：
   - 每門課程各 1 筆調代課總結 log
   - 總筆數 = 課程數量（非 dc_bind_key 刪除總筆數）

預期：調代課刪除 log = 課程數量。

---

## Step 4 — 回歸確認（避免副作用）
- 確認 `delete_course`、`delete_sc_attend`、`delete_sce_take` 的刪除行為不受影響。
- 確認其他兩類 log：
  - `刪除_學生_課程評量成績`（仍可能多筆，屬逐筆評量成績刪除）
  - `刪除_課程_課程`（每課程 1 筆）
- 確認 UI 不會因 Course cache sync 而異常（原本最後有 `Course.Instance.SyncDataBackground(...)`）。

---

## Step 5 — 產出變更紀錄
在專案文件（或指定路徑）新增 / 更新：`刪除課程與調代課对应資料調整0120.md`，至少包含：
- 問題描述：為何原本會產生多筆 log（dc_bind_key 一對多）
- 調整說明：改為 GROUP BY ref_course_id 彙總寫入
- 測試結果：Case A/B/C 的實際結果（貼 log 關鍵欄位）
- 風險與取捨：不再逐筆記錄 dc_bind_key 明細，改為總結

---

## 交付驗收條件（Definition of Done）
- [ ] 同一課程刪除時，調代課刪除 log 僅 1 筆（若確實有刪到 dc_bind_key）。
- [ ] description 正確顯示刪除筆數（del_cnt）。
- [ ] Case A/B/C 測試完成且結果符合預期。
- [ ] `刪除課程與調代課對應資料調整0120.md` 已完成撰寫並包含測試證據。
