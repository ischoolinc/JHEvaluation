# 模組載入效能驗收紀錄

## 測試設定

| 欄位 | 值 |
|---|---|
| 版本／commit | 待實機測試填寫 |
| 機器與作業系統 | 待實機測試填寫 |
| 宿主版本 | 待實機測試填寫 |
| 帳號與 ACL 條件 | 待實機測試填寫 |
| 資料條件 | 待實機測試填寫 |
| 冷啟動快取控制方式 | 待實機測試填寫 |
| 暖啟動定義 | 待實機測試填寫 |

## 收集方式

1. 設定 `JH_SCORE_LOAD_METRICS_PATH` 為可寫入的 CSV 絕對路徑。
2. 最佳化前設定 `JH_SCORE_LOAD_VARIANT=baseline`；最佳化後設定為 `optimized`。
3. 新竹及高雄模式分別收集至少 10 次冷啟動與 10 次暖啟動；冷暖資料可使用不同 CSV，避免混合統計。
4. 執行：

   ```powershell
   .\tools\Analyze-ModuleLoadMetrics.ps1 -Path <metrics.csv> -BaselineVariant baseline -CandidateVariant optimized -MinimumSamples 10
   ```

## 最佳化前基準

目前工作區只有模組 class library 與相依 DLL，沒有可啟動且已登入的 FISCA 宿主，因此無法在此環境產生有效的冷／暖啟動樣本。請在實際宿主依上述方式填入統計。

| 模式 | 冷／暖 | 樣本 | Total Median | Total P95 | 最大階段 | 占比 |
|---|---|---:|---:|---:|---|---:|
| HsinChu | 冷 | 待填 | 待填 | 待填 | 待填 | 待填 |
| HsinChu | 暖 | 待填 | 待填 | 待填 | 待填 | 待填 |
| KaoHsiung | 冷 | 待填 | 待填 | 待填 | 待填 | 待填 |
| KaoHsiung | 暖 | 待填 | 待填 | 待填 | 待填 | 待填 |

## 已採用的低風險改善

- 將單一大型 `Main()` 拆成七個可獨立量測的初始化階段。
- 加入同一宿主生命週期只初始化一次的保護，避免重複 UI、事件及擴充點註冊。
- 快取學生與教務成績選單、畢業功能按鈕、預警報表按鈕及同一階段重複 ACL 結果。
- 學生選取變更事件直接使用已快取的三個目標控制項，移除每次事件重新走訪多層 Ribbon 路徑。
- 未延後任何必要註冊：靜態盤點未發現資料庫、網路或重型計算，且宿主沒有正式 UI-ready 契約。

## 最佳化後與門檻

| 模式 | 冷／暖 | 樣本 | Baseline Median | Optimized Median | 改善率 | Baseline P95 | Optimized P95 | 通過 |
|---|---|---:|---:|---:|---:|---:|---:|---|
| HsinChu | 冷 | 待填 | 待填 | 待填 | 待填 | 待填 | 待填 | 待填 |
| HsinChu | 暖 | 待填 | 待填 | 待填 | 待填 | 待填 | 待填 | 待填 |
| KaoHsiung | 冷 | 待填 | 待填 | 待填 | 待填 | 待填 | 待填 | 待填 |
| KaoHsiung | 暖 | 待填 | 待填 | 待填 | 待填 | 待填 | 待填 | 待填 |

- Baseline median > 200 ms：optimized median 必須至少降低 20%，optimized P95 不得超過 baseline P95 的 110%。
- Baseline median <= 200 ms：optimized median 與 P95 均不得退化。

## 已知限制與回退

- 實際效能門檻必須在可登入且能切換兩種部署模式的宿主完成，不能以 class library 的編譯時間代替。
- 若功能驗證或實機數據顯示回歸，回退 `Program.InitializeModule()` 的階段化呼叫與 `RegistrationContext` 快取，恢復原同步註冊；`ModuleLoadDiagnostics` 可獨立保留以協助定位。
- CSV 輸出只供診斷；移除 `JH_SCORE_LOAD_METRICS_PATH` 即停止檔案寫入，`Trace` 診斷仍可使用。

## 工作區自動驗證結果

- Debug Build：通過（Visual Studio 18 MSBuild，0 errors）。
- Release Rebuild：通過（0 errors）。
- 編譯警告：11 個，與修改前 Debug baseline 相同；包含既有重複 using、未使用欄位／區域變數、`SemesterData` equality 及原有部署參數 API obsolete 警告，本次未新增警告。
- `ModuleLoadDiagnosticsSmokeTest`：通過成功紀錄、詳細量測關閉、失敗重新拋出、失敗狀態記錄、部署模式回填及不可寫路徑不影響初始化語意。
- Idempotent completed-state smoke test：連續兩次呼叫已完成狀態的 `Program.Main()` 均直接返回，沒有解析或呼叫任何宿主 UI API。
- `Analyze-ModuleLoadMetrics.ps1`：可解析 smoke-test CSV 並產出階段統計。
- 相依性檢查：僅新增 BCL (`System.Diagnostics`、`System.IO`) 程式碼，未新增外部套件、公開 API 或資料格式變更。
