# 修正指示: NX内部「作業進行中」ダイアログのフォーカス奪取防止

## 問題
DXFエクスポート処理中（DxfdwgCreator.Commit()実行時）に、NXの内部ダイアログ「作業進行中」が前面に表示され、NXウィンドウにフォーカスが移ってしまう。
ユーザーが他のアプリケーションで作業中に割り込まれる。

## 修正方針
バックグラウンドスレッドで定期的にNXの「作業進行中」ダイアログを監視し、検出したら即座に非表示にする。
同時に、自分のProgressFormが常に前面に表示され続けるようにする。

## 実装方法

### 1. Win32 API宣言の追加（既存のDllImportに追加）
以下のWin32 APIが未宣言であれば追加すること:
```csharp
[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

[DllImport("user32.dll", SetLastError = true)]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

[DllImport("user32.dll", SetLastError = true)]
private static extern IntPtr GetForegroundWindow();

[DllImport("user32.dll", SetLastError = true)]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool SetForegroundWindow(IntPtr hWnd);

[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

[DllImport("user32.dll")]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

[DllImport("user32.dll")]
private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

[DllImport("user32.dll")]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool IsWindowVisible(IntPtr hWnd);

private const int SW_HIDE = 0;
private const int SW_MINIMIZE = 6;
```

### 2. NXダイアログ監視用のSystem.Threading.Timerを実装
ビューのループ処理開始前にタイマーを起動し、ループ終了後に停止する。

```csharp
private static System.Threading.Timer _nxDialogSuppressor;

/// <summary>
/// NXの「作業進行中」ダイアログを監視・非表示にするタイマーを開始
/// </summary>
private static void StartNxDialogSuppressor()
{
    // 現在のプロセスIDを取得（NXのプロセス）
    uint currentProcessId = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;

    _nxDialogSuppressor = new System.Threading.Timer(_ =>
    {
        try
        {
            SuppressNxDialogs(currentProcessId);
        }
        catch
        {
            // タイマー内の例外は無視（安全性優先）
        }
    }, null, 0, 200); // 200msごとに監視
}

/// <summary>
/// NXダイアログ監視タイマーを停止
/// </summary>
private static void StopNxDialogSuppressor()
{
    if (_nxDialogSuppressor != null)
    {
        _nxDialogSuppressor.Dispose();
        _nxDialogSuppressor = null;
    }
}

/// <summary>
/// NXプロセスが所有する「作業進行中」ダイアログを検出して非表示にする
/// </summary>
private static void SuppressNxDialogs(uint targetProcessId)
{
    // 抑制対象のダイアログタイトル（日本語・英語両対応）
    string[] suppressTitles = new string[]
    {
        "作業進行中",
        "Work in Progress",
        "Operation in Progress"
    };

    EnumWindows((hWnd, lParam) =>
    {
        if (!IsWindowVisible(hWnd))
            return true; // 非表示ウィンドウはスキップ

        // このウィンドウがNXプロセスのものか確認
        uint processId;
        GetWindowThreadProcessId(hWnd, out processId);
        if (processId != targetProcessId)
            return true; // 別プロセスはスキップ

        // ウィンドウタイトルを取得
        var sb = new System.Text.StringBuilder(256);
        GetWindowText(hWnd, sb, sb.Capacity);
        string title = sb.ToString();

        // 抑制対象のタイトルと一致するか確認
        foreach (string suppressTitle in suppressTitles)
        {
            if (title.Contains(suppressTitle))
            {
                // ダイアログを非表示にする
                ShowWindow(hWnd, SW_HIDE);
                break;
            }
        }

        return true; // 列挙続行
    }, IntPtr.Zero);
}
```

### 3. ループ処理への組み込み
ビューごとのループ処理の前後にタイマーの開始・停止を追加する。

```csharp
// ビューループの直前に追加
StartNxDialogSuppressor();

try
{
    // === 既存のビューごとのループ処理 ===
    foreach (...)
    {
        // ... 図面シート作成 → ベースビュー配置 → DXFエクスポート ...
    }
}
finally
{
    // ループ終了後（正常・異常問わず）に必ず停止
    StopNxDialogSuppressor();
}
```

### 4. 配置場所
- StartNxDialogSuppressor(): NXウィンドウ最小化の直後、ビューループの直前
- StopNxDialogSuppressor(): ビューループ終了後、SafeUndoAndCleanupの前

## 注意事項
- System.Threading.Timer を使用すること（System.Windows.Forms.Timer ではない）。UIスレッドに依存しないため安全。
- タイマーの間隔は200ms。短すぎるとCPU負荷が上がり、長すぎるとダイアログが一瞬見える。200msがバランス良い。
- SW_HIDE（完全非表示）を使用。SW_MINIMIZE だとタスクバーに残り、ユーザーが混乱する可能性がある。
- 自プロセスIDでフィルタしているため、他のアプリケーションのウィンドウには影響しない。
- ProgressFormのTopMost=true設定は維持すること（既存の実装を変更しない）。
- using System.Text; と using System.Threading; と using System.Diagnostics; が必要であれば追加すること。

## ビルド
修正完了後、ビルドまで実行してください。
