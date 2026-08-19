# 模組載入盤點

## 宿主限制

- 模組為 .NET Framework 4.8 class library，由 FISCA `[MainMethod]` 在宿主 UI 執行緒呼叫 `Program.Main()`。
- `FISCA.Presentation.MotherForm` 公開 `Form`、`TabGotFocus` 與 `LayoutResumed`，但未提供具明確模組語意的 module-loaded／UI-ready 契約。
- Ribbon、ACL、事件與擴充點註冊都會影響載入完成後的功能可發現性，必須在 `Main()` 返回前完成。
- `CheckStudentSemHistoryScoreRAT` 的建構不執行查詢；資料查詢僅在使用者執行檢查時發生。
- `DetailItemFeature(typeof(...))` 只註冊型別；三個明細元件的 UI 與資料工作在框架建立元件後才執行。
- 載入路徑沒有資料庫查詢、網路呼叫或報表產生工作，因此在取得實機量測前沒有可安全延後的顯著工作。

## 初始化階段

| 階段識別 | 同步工作 | 載入完成前必要 | 可延後判定 |
|---|---|---:|---|
| `deploy-mode` | 讀取部署參數並設定新竹／高雄模式 | 是 | 否，後續功能可能依賴 `Mode` |
| `student-detail-builder` | 註冊畢業成績 detail builder | 是 | 否，屬框架擴充點 |
| `score-menus` | 註冊學期歷程、學生與教務成績選單及 Click handler | 是 | 否，影響選單與事件可用性 |
| `graduation-reports` | 註冊畢業功能、預警報表、ACL 與 Click handler | 是 | 否，影響選單與權限 |
| `acl-selection-events` | 註冊權限功能及學生選取變更事件 | 是 | 否，影響權限與即時狀態 |
| `detail-items` | 註冊三個學生資料項目 | 是 | 否，屬框架擴充點 |
| `data-rationality` | 註冊學期歷程／成績合理性檢查器 | 是 | 否，屬框架擴充點；建構成本低 |
| `total` | 全部初始化總耗時 | — | — |

## 診斷機制

- 每個階段使用 `Stopwatch` 記錄毫秒、成功狀態與例外；例外在記錄後原樣重新拋出。
- 結果使用 `Trace` 本機輸出，類別為 `JHEvaluation.ScoreCalculation.Load`。
- 設定 `JH_SCORE_LOAD_METRICS_PATH` 時，會在總計時停止後一次附加至 CSV，不會同步呼叫遠端服務。
- `JH_SCORE_LOAD_VARIANT` 標記 `baseline`、`optimized` 等測試版本；未設定時為 `unspecified`。
- `JH_SCORE_LOAD_DIAGNOSTICS=0` 可關閉分段明細，仍保留 `total`，用來比較詳細量測開關的負擔。

## 冪等與失敗策略

- `InitializeModule()` 以單一鎖保護；成功後的重複呼叫直接返回，不會重複掛載 UI 或事件。
- 首次初始化失敗會保存例外。因部分框架註冊不可回復，後續呼叫不猜測性重試，而是回報先前失敗，避免重複註冊。
- 診斷輸出失敗只寫入 `Trace`，不改變模組載入原本的成功或失敗語意。
