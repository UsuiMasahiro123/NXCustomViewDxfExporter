## 根本原因の再分析と完全に異なるアプローチでの修正

これまでの修正で効果がなかった理由を分析し、根本的に異なるアプローチで修正します。
まず現在のソースコード全体を読んで、全体の処理構造を把握してから修正に取りかかってください。

`cat CustomViewDxfExporter.cs` で全文を確認してください。

---

## 不具合1: c0000005 クラッシュ — 根本原因の再分析

### ログの時系列を精密に分析した結果

```
21:33:55 - View 2: 図面シート作成・ベースビュー配置・DxfdwgCreator.Commit()
           → tmp_ugto2d ファイルが保存・読込される
           → UNDO_UG_delete_mark: Mark 776 not found
           → 図面シートの削除・クリーンアップが実行される

12秒後...

21:34:07 - "DXF/DWG Export: Export translation job submitted."
           → 直後に *** ERROR: c0000005
```

**重要な発見: DxfdwgCreator.Commit() は2段階で動作する。**
1. 第1段階（同期）: 一時パートファイルの作成・保存。Commit()はここで戻る
2. 第2段階（非同期）: 実際のDXF変換。バックグラウンドスレッドで後から実行される

つまり、ファイル待機やUndoMark削除の問題ではなく、**図面シートを削除した後に、DxfdwgCreatorの非同期スレッドがその削除済みシートにアクセスしてクラッシュしている**。

### 解決策: 図面シートの削除を全てのエクスポート完了後に一括で行う

**ループ中は図面シートを絶対に削除しない。** 全ビューのエクスポートが完了し、全てのDXFファイルが出力されたことを確認してから、最後にまとめて削除する。

### 具体的な実装

```csharp
// ★ ループ外で、後で削除するシートのリストを用意
List<NXOpen.Drawings.DrawingSheet> sheetsToCleanup = new List<NXOpen.Drawings.DrawingSheet>();
List<string> expectedOutputFiles = new List<string>();

foreach (var viewName in selectedViews)
{
    NXOpen.Drawings.DrawingSheet sheet = null;
    try
    {
        // Step 1: 図面シート作成（既存処理）
        sheet = /* 既存の図面シート作成処理 */;

        // Step 2: ベースビュー配置（既存処理）
        var baseView = /* 既存のベースビュー配置処理 */;

        // Step 3: DXFエクスポート（既存処理）
        DxfdwgCreator dxfdwgCreator = /* 既存のCreator作成・設定処理 */;
        
        NXOpen.NXObject result = dxfdwgCreator.Commit();
        dxfdwgCreator.Destroy();
        dxfdwgCreator = null;

        // ★ 出力予定ファイルをリストに追加
        expectedOutputFiles.Add(outputPath);

        // ★★★ ここでシートを削除しない ★★★
        // ★★★ クリーンアップリストに追加するだけ ★★★
        if (sheet != null)
        {
            sheetsToCleanup.Add(sheet);
        }

        successCount++;
    }
    catch (Exception ex)
    {
        failCount++;
        // エラーでもシートがあればクリーンアップリストに入れる
        if (sheet != null)
        {
            sheetsToCleanup.Add(sheet);
        }
    }

    // 進捗更新
    // ...
}

// ★★★ 全ビューのエクスポート完了後、全DXFファイルの出力を待つ ★★★
foreach (string expectedFile in expectedOutputFiles)
{
    int waitMs = 0;
    while (!File.Exists(expectedFile) && waitMs < 60000)
    {
        System.Threading.Thread.Sleep(1000);
        waitMs += 1000;
    }
}
// さらに安全マージンとして追加待機（バックグラウンドスレッドの完全終了を待つ）
System.Threading.Thread.Sleep(5000);

// ★★★ 全DXFファイル出力完了後に、図面シートを一括削除 ★★★
foreach (var sheet in sheetsToCleanup)
{
    try
    {
        workPart.DrawingSheets.DeleteSheets(
            new NXOpen.Drawings.DrawingSheet[] { sheet });
    }
    catch { } // 削除に失敗しても無視（パートを保存しないので問題ない）
}

// モデリング環境に戻す
try { theSession.ApplicationSwitchImmediate("UG_APP_MODELING"); } catch { }
```

### UndoMark について

UndoMark関連のコードが残っている場合は**全て削除**してください。
grep で確認: `grep -n "UndoMark\|SetUndoMark\|UndoToMark\|DeleteUndoMark\|undoMark" CustomViewDxfExporter.cs`

結果が0件になるまで削除してください。パートは保存しない設計なので、UndoMarkがなくても安全です。

---

## 不具合2: 進捗ダイアログがNXの背面に隠れる — 根本原因の再分析

### なぜ TopMost = true だけでは不十分か

NXはDxfdwgCreator.Commit()実行時に内部的にウィンドウを最前面に持ってくる。
NXは特殊なウィンドウ管理を行っており、WinFormsの TopMost = true を上書きしてしまう場合がある。

### 解決策: System.Windows.Forms.Timer で定期的にダイアログを最前面に強制する

進捗ダイアログに Timer を追加し、500ms間隔でダイアログを最前面に強制的に持ってくる。

```csharp
// 進捗ダイアログクラス内に以下を追加

private System.Windows.Forms.Timer foregroundTimer;

// コンストラクタまたは初期化メソッド内で:
this.TopMost = true;

foregroundTimer = new System.Windows.Forms.Timer();
foregroundTimer.Interval = 500; // 500ms間隔
foregroundTimer.Tick += (s, e) =>
{
    if (this.Visible && !this.IsDisposed)
    {
        // Win32 APIで確実に最前面に持ってくる
        SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    }
};
foregroundTimer.Start();

// Win32 API宣言（クラス内に追加）
[DllImport("user32.dll")]
private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
    int X, int Y, int cx, int cy, uint uFlags);

private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
private const uint SWP_NOMOVE = 0x0002;
private const uint SWP_NOSIZE = 0x0001;
private const uint SWP_NOACTIVATE = 0x0010;

// ダイアログを閉じる時に Timer を停止
// Close または Dispose のタイミングで:
if (foregroundTimer != null)
{
    foregroundTimer.Stop();
    foregroundTimer.Dispose();
    foregroundTimer = null;
}
```

`SWP_NOACTIVATE` フラグにより、ダイアログにフォーカスを奪わずに最前面表示を維持する。
これによりNXの操作を妨げずに、進捗ダイアログが常に見える状態を保つ。

---

## 修正のまとめ

| 項目 | これまでのアプローチ（失敗） | 今回のアプローチ |
|------|---------------------------|----------------|
| クラッシュ | ループ内で毎回シート削除 + 待機 | **ループ内でシートを削除しない。全エクスポート完了後に一括削除** |
| UndoMark | UndoToMarkで一括復元 | **UndoMark完全廃止。手動一括削除** |
| ダイアログ背面 | TopMost = true のみ | **Timer + Win32 SetWindowPos で定期的に最前面強制** |

## 修正後の確認

```bash
echo "=== UndoMark残存チェック（0件であること） ==="
grep -c "UndoMark\|SetUndoMark\|UndoToMark\|DeleteUndoMark\|undoMark" CustomViewDxfExporter.cs

echo "=== シート削除がループ外にあること ==="
grep -n "DeleteSheets" CustomViewDxfExporter.cs

echo "=== Timer実装チェック ==="
grep -n "foregroundTimer\|SetWindowPos\|HWND_TOPMOST" CustomViewDxfExporter.cs

echo "=== TopMost設定 ==="
grep -n "TopMost" CustomViewDxfExporter.cs
```

## 変更しないこと
- DxfdwgCreator の設定内容（SettingsFile、InputFile、OutputFile等）
- dxfdwg.def の埋め込み方式
- ビュー選択ダイアログの機能
- ファイル名の命名規則（DB_PART_NO等）
- 多言語対応
- 出力フォルダ（Pictures）
