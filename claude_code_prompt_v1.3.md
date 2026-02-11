# Claude Code プロンプト: NXCustomViewDxfExporter v1.3 改修

## 概要
既存のNXCustomViewDxfExporterアドオン（v1.2）に以下の3つの改修を加えてください。既存の処理ロジック（カスタムビュー取得→図面シート作成→DXFエクスポート→クリーンアップのループ）は変更しないでください。

## 改修内容

---

### 改修1: 日英自動切替の実装

NXの言語モードに応じて、全てのUI文字列（ダイアログ、メッセージ、ボタン）を日本語/英語に自動切替する機能を追加してください。

#### 言語検出の仕組み（優先度順）
1. **環境変数 `UGII_LANG`** を `Environment.GetEnvironmentVariable("UGII_LANG")` で取得
   - `"japanese"` → 日本語モード
   - `"english"` または その他 → 英語モード
2. **フォールバック**: `UGII_LANG` が未設定の場合、`System.Globalization.CultureInfo.CurrentUICulture` のNameが `"ja"` で始まるかで判定

#### 実装方法
- **コード内辞書（`Dictionary<string, string>`）** で日英メッセージペアを管理する
- `.resx` ファイルは使用しない（DLL単体で完結させるため）
- 起動時に一度だけ言語判定を行い、`bool isJapanese` フラグをセットする
- 全てのUI文字列をヘルパーメソッド経由で取得する

```csharp
// 実装例
private static bool isJapanese;
private static Dictionary<string, string> messagesJa;
private static Dictionary<string, string> messagesEn;

private static void InitializeLanguage()
{
    string lang = Environment.GetEnvironmentVariable("UGII_LANG");
    if (!string.IsNullOrEmpty(lang))
    {
        isJapanese = lang.Equals("japanese", StringComparison.OrdinalIgnoreCase);
    }
    else
    {
        isJapanese = System.Globalization.CultureInfo.CurrentUICulture.Name.StartsWith("ja");
    }
    
    // 辞書の初期化
    messagesJa = new Dictionary<string, string>
    {
        { "FolderSelectTitle", "DXF出力先フォルダを選択" },
        { "ProgressTitle", "DXFエクスポート中..." },
        { "ProgressMessage", "処理中: {0}（{1}/{2}）" },
        { "StopButton", "停止" },
        { "CompleteTitleSuccess", "DXFエクスポート完了" },
        { "CompleteTitleError", "DXFエクスポート完了（一部エラー）" },
        { "CompleteMessage", "DXFエクスポートが完了しました" },
        { "CompleteMessageError", "DXFエクスポートが完了しました（一部エラー）" },
        { "SuccessCount", "成功: {0} ファイル" },
        { "FailureCount", "失敗: {0} ファイル" },
        { "CancelCount", "中断: {0} ファイル（ユーザー停止）" },
        { "OutputPath", "出力先: {0}" },
        { "ErrorDetails", "エラー詳細:" },
        { "ErrorNoPartOpen", "パートが開かれていません" },
        { "ErrorNoCustomView", "カスタムビューが見つかりません" },
        { "OkButton", "OK" },
        // 必要に応じて追加
    };
    
    messagesEn = new Dictionary<string, string>
    {
        { "FolderSelectTitle", "Select DXF Output Folder" },
        { "ProgressTitle", "Exporting DXF..." },
        { "ProgressMessage", "Processing: {0} ({1}/{2})" },
        { "StopButton", "Stop" },
        { "CompleteTitleSuccess", "DXF Export Complete" },
        { "CompleteTitleError", "DXF Export Complete (with errors)" },
        { "CompleteMessage", "DXF export completed successfully" },
        { "CompleteMessageError", "DXF export completed with some errors" },
        { "SuccessCount", "Success: {0} files" },
        { "FailureCount", "Failed: {0} files" },
        { "CancelCount", "Cancelled: {0} files (user stopped)" },
        { "OutputPath", "Output: {0}" },
        { "ErrorDetails", "Error details:" },
        { "ErrorNoPartOpen", "No part is open" },
        { "ErrorNoCustomView", "No custom views found" },
        { "OkButton", "OK" },
        // 必要に応じて追加
    };
}

private static string GetMessage(string key)
{
    var dict = isJapanese ? messagesJa : messagesEn;
    return dict.ContainsKey(key) ? dict[key] : key;
}

private static string GetMessage(string key, params object[] args)
{
    return string.Format(GetMessage(key), args);
}
```

#### 切替対象の全UI要素

| UI要素 | 日本語 | 英語 |
|--------|--------|------|
| フォルダ選択タイトル | DXF出力先フォルダを選択 | Select DXF Output Folder |
| 進捗ダイアログタイトル | DXFエクスポート中... | Exporting DXF... |
| 進捗メッセージ | 処理中: ○○（3/9） | Processing: ○○ (3/9) |
| 停止ボタン | 停止 | Stop |
| 完了ダイアログタイトル（正常） | DXFエクスポート完了 | DXF Export Complete |
| 完了ダイアログタイトル（エラー） | DXFエクスポート完了（一部エラー） | DXF Export Complete (with errors) |
| 成功数表示 | 成功: N ファイル | Success: N files |
| 失敗数表示 | 失敗: N ファイル | Failed: N files |
| 中断数表示 | 中断: N ファイル（ユーザー停止） | Cancelled: N files (user stopped) |
| 出力先パス | 出力先: C:\... | Output: C:\... |
| エラー詳細 | エラー詳細: | Error details: |
| エラー（パート未オープン） | パートが開かれていません | No part is open |
| エラー（ビューなし） | カスタムビューが見つかりません | No custom views found |
| OKボタン | OK | OK |

※ 既存コードにハードコードされている全ての日本語文字列を洗い出し、全て辞書に移行すること。

---

### 改修2: 進捗表示・処理中断UIの変更

現在の進捗表示を、**独立したWinFormsダイアログ**に変更してください。NXのステータスバーやMessageBoxではなく、独自のFormを使用します。

#### ダイアログ仕様

**進捗ダイアログ（ProgressForm）の構成：**

```
┌──────────────────────────────────────┐
│ DXFエクスポート中...            [×]  │  ← タイトルバー（×ボタンは無効化）
├──────────────────────────────────────┤
│                                      │
│  処理中: 03_DWG_RIGHT（3/9）         │  ← ビュー名と進捗カウント
│                                      │
│  ████████████░░░░░░░░░  33%          │  ← プログレスバー + パーセント表示
│                                      │
│          [ 停止 ]                    │  ← 赤い停止ボタン
│                                      │
└──────────────────────────────────────┘
```

**詳細仕様：**
- `FormBorderStyle = FixedDialog`（サイズ変更不可）
- `MaximizeBox = false`, `MinimizeBox = false`
- `ControlBox = true` だが、`FormClosing` イベントで閉じるのをキャンセル（×ボタン無効化）
- `StartPosition = CenterScreen`
- `TopMost = true`（NXが最小化されていても常に前面に表示）
- `ShowInTaskbar = true`（タスクバーに表示）
- ダイアログサイズ: 約 400×200 ピクセル
- **停止ボタン**: 赤色背景 (`Color.FromArgb(220, 53, 69)`) に白文字
- **プログレスバー**: `ProgressBar` コントロールを使用

**更新メソッド：**
```csharp
public void UpdateProgress(string viewName, int current, int total)
{
    if (InvokeRequired)
    {
        Invoke(new Action(() => UpdateProgress(viewName, current, total)));
        return;
    }
    labelProgress.Text = GetMessage("ProgressMessage", viewName, current, total);
    progressBar.Maximum = total;
    progressBar.Value = current;
    labelPercent.Text = $"{(int)((double)current / total * 100)}%";
}
```

**停止フラグ：**
```csharp
public bool StopRequested { get; private set; } = false;

private void btnStop_Click(object sender, EventArgs e)
{
    StopRequested = true;
    btnStop.Enabled = false;
    btnStop.Text = isJapanese ? "停止中..." : "Stopping...";
}
```

**メインループとの連携：**
- ループ処理の各ビュー処理前に `progressForm.StopRequested` をチェック
- trueなら安全にループを離脱（現在処理中のビューは最後まで完了させる）
- 進捗ダイアログは別スレッドで表示するのではなく、**メインスレッドのForm**として作成し、ループ内で `Application.DoEvents()` を呼んでUIを応答させる

```csharp
// メインループ内での使用例
for (int i = 0; i < customViews.Count; i++)
{
    // 中断チェック
    if (progressForm.StopRequested)
    {
        cancelCount = customViews.Count - i;
        break;
    }
    
    // 進捗更新
    progressForm.UpdateProgress(customViews[i].Name, i + 1, customViews.Count);
    Application.DoEvents(); // UIを応答させる
    
    // ビュー処理...
}
```

---

### 改修3: NXウィンドウの最小化・復帰

処理中はNXの画面を最小化し、進捗ダイアログだけを表示することで、ユーザーが他のアプリケーションで作業できるようにします。

#### 実装方法

**Win32 API の宣言：**
```csharp
using System.Runtime.InteropServices;
using System.Diagnostics;

[DllImport("user32.dll")]
private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

[DllImport("user32.dll")]
private static extern bool LockSetForegroundWindow(uint uLockCode);

private const int SW_MINIMIZE = 6;
private const int SW_RESTORE = 9;
private const uint LSFW_LOCK = 1;
private const uint LSFW_UNLOCK = 2;
```

**NXウィンドウハンドルの取得：**
```csharp
private static IntPtr GetNXWindowHandle()
{
    return Process.GetCurrentProcess().MainWindowHandle;
}
```

#### 処理フロー

```
1. フォルダ選択ダイアログ表示 → ユーザーが出力先を選択
2. NXウィンドウを最小化（ShowWindow(hWnd, SW_MINIMIZE)）
3. フォーカスロック（LockSetForegroundWindow(LSFW_LOCK)）
4. 画面更新を抑制（SetDisplay SUPPRESS）
5. 進捗ダイアログ（ProgressForm）を表示（TopMost=true）
6. カスタムビューのループ処理
   - 各ビュー処理前に中断チェック
   - 各ビュー処理前に進捗更新 + Application.DoEvents()
7. 進捗ダイアログを閉じる
8. 画面更新を復帰（SetDisplay UNSUPPRESS）
9. 完了ダイアログ（CompletionForm）を表示
10. 完了ダイアログのOKボタン押下 → NXウィンドウを復帰（SW_RESTORE）
    または 中断時 → 完了ダイアログ表示 → OK押下 → NXウィンドウを復帰（SW_RESTORE）
11. フォーカスロック解除（LockSetForegroundWindow(LSFW_UNLOCK)）
```

**重要: NXウィンドウの復帰タイミング**
- 完了ダイアログ（CompletionForm）のOKボタンが押されたときにNXウィンドウを復帰する
- 進捗ダイアログの停止ボタンが押された場合も、完了ダイアログ表示→OK押下の後にNXウィンドウを復帰
- つまり、NXウィンドウの復帰は常に「完了ダイアログのOK押下後」に行う

#### 完了ダイアログ（CompletionForm）の仕様

**正常完了時：**
```
┌──────────────────────────────────────────┐
│ DXFエクスポート完了                 [×]  │
├──────────────────────────────────────────┤
│                                          │
│  ✓  DXFエクスポートが完了しました        │  ← 緑アイコン
│                                          │
│  成功: 9 ファイル                        │  ← 緑文字
│  失敗: 0 ファイル                        │
│  出力先: C:\DXF_Output\                  │
│                                          │
│              [ OK ]                      │
│                                          │
└──────────────────────────────────────────┘
```

**一部失敗・中断時：**
```
┌──────────────────────────────────────────┐
│ DXFエクスポート完了（一部エラー）   [×]  │
├──────────────────────────────────────────┤
│                                          │
│  ⚠  DXFエクスポートが完了しました       │  ← オレンジアイコン
│     （一部エラー）                       │
│                                          │
│  成功: 7 ファイル                        │  ← 緑文字
│  失敗: 1 ファイル                        │  ← 赤文字
│  中断: 1 ファイル（ユーザー停止）        │  ← オレンジ文字
│                                          │
│  エラー詳細:                             │
│  05_DWG_SEC: ビューの投影範囲が          │  ← 赤文字
│  取得できませんでした                    │
│                                          │
│              [ OK ]                      │
│                                          │
└──────────────────────────────────────────┘
```

**CompletionFormの仕様：**
- `FormBorderStyle = FixedDialog`
- `StartPosition = CenterScreen`
- `TopMost = true`
- `ShowInTaskbar = true`
- OKボタン押下時に `DialogResult = DialogResult.OK` を返す
- OKボタン押下のタイミングで、NXウィンドウの復帰処理を実行する

```csharp
// CompletionFormのOKボタンハンドラ
private void btnOk_Click(object sender, EventArgs e)
{
    this.DialogResult = DialogResult.OK;
    this.Close();
}

// メイン処理側
IntPtr nxHandle = GetNXWindowHandle();
try
{
    ShowWindow(nxHandle, SW_MINIMIZE);
    LockSetForegroundWindow(LSFW_LOCK);
    
    // ... 処理ループ ...
    
    // 完了ダイアログ表示
    using (var completionForm = new CompletionForm(successCount, failCount, cancelCount, errors, outputPath))
    {
        completionForm.ShowDialog();
    }
}
finally
{
    // NXウィンドウ復帰（完了ダイアログが閉じた後）
    ShowWindow(nxHandle, SW_RESTORE);
    LockSetForegroundWindow(LSFW_UNLOCK);
}
```

---

## try-finally による保護

全ての処理をtry-finallyで囲み、エラー時も確実に以下を復帰すること：
- 画面更新の復帰（SetDisplay UNSUPPRESS）
- フォーカスロック解除
- NXウィンドウの復帰
- UndoMarkの復帰
- 進捗ダイアログが開いていたら閉じる

```csharp
IntPtr nxHandle = GetNXWindowHandle();
ProgressForm progressForm = null;

try
{
    // NX最小化 + フォーカスロック
    ShowWindow(nxHandle, SW_MINIMIZE);
    LockSetForegroundWindow(LSFW_LOCK);
    
    // 画面更新抑制
    theUfSession.Disp.SetDisplay(UF_DISP_SUPPRESS_DISPLAY);
    
    // 進捗ダイアログ表示
    progressForm = new ProgressForm(totalViews);
    progressForm.Show();
    
    // ループ処理...
    
    // 進捗ダイアログを閉じる
    progressForm.Close();
    progressForm.Dispose();
    progressForm = null;
    
    // 画面更新復帰
    theUfSession.Disp.SetDisplay(UF_DISP_UNSUPPRESS_DISPLAY);
    
    // 完了ダイアログ表示
    using (var completionForm = new CompletionForm(...))
    {
        completionForm.ShowDialog();
    }
}
catch (Exception ex)
{
    // エラーハンドリング
}
finally
{
    // 確実に復帰
    if (progressForm != null)
    {
        progressForm.Close();
        progressForm.Dispose();
    }
    
    try { theUfSession.Disp.SetDisplay(UF_DISP_UNSUPPRESS_DISPLAY); } catch { }
    
    ShowWindow(nxHandle, SW_RESTORE);
    LockSetForegroundWindow(LSFW_UNLOCK);
}
```

---

## 注意事項

1. **既存の処理ロジックは変更しない** — カスタムビュー取得、図面シート作成、DXFエクスポート、dxfdwg.defの埋め込み等はそのまま維持
2. **ProgressFormとCompletionFormは同じ.csファイル内にインナークラスとして定義** — ファイルを分割しない（DLL単体完結）
3. **Application.DoEvents()はループ内でのみ使用** — 他の場所では使わない
4. **TopMost = trueは進捗・完了ダイアログのみ** — NXウィンドウには設定しない
5. **フォントは`Yu Gothic UI`を優先し、フォールバックとして`Segoe UI`を使用** — 日英両方で表示が崩れないように
6. **ビルドまで完了させてください**
7. **完了したらGitにコミット＆プッシュしてください** — コミットメッセージ: `v1.3: 日英自動切替、進捗ダイアログUI、NXウィンドウ制御を追加`

## 参照ファイル
- **claude_code_prompt.md**: 既存のプロンプト（基本仕様・DXFエクスポート設定等はこちらを参照）
- **dxf_export_journal.cs**: NXで手動DXFエクスポートを行った際のジャーナル記録
- **NXCustomViewDxfExporter/CustomViewDxfExporter.cs**: 現在のソースコード（v1.2）
