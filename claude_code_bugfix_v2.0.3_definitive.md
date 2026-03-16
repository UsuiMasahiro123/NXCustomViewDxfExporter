## 緊急修正: 2つの不具合が未修正です

まず現在のコードの状態を確認してください。以下の3つを出力してください：

1. `grep -n "UndoMark\|SetUndoMark\|UndoToMark\|DeleteUndoMark\|UndoToLastVisibleMark" CustomViewDxfExporter.cs`
2. `grep -n "TopMost" CustomViewDxfExporter.cs`
3. `grep -n "Thread.Sleep" CustomViewDxfExporter.cs`

上記の結果を確認した上で、以下の修正を**全て**実施してください。

---

## 不具合1: c0000005 クラッシュ（最優先）

### ログ分析結果（クラッシュの100%再現パターン）

```
DXF/DWG Export: Export translation job submitted.*** ERROR: Unhandled exception: 0xc0000005
*** Exception was not in main thread
+++ Invalid read from 0000000000000000
UNDO_UG_delete_mark: Mark 22 (DXF Export) not found
```

### 原因
DxfdwgCreator.Commit() は「Export translation job submitted」というメッセージとともに、バックグラウンドスレッドで非同期にDXF変換を実行する。Commit()が返った直後にコードが図面シートやビューを削除・Undoすると、バックグラウンドスレッドが参照中のオブジェクトが消え、NULLポインタアクセス(c0000005)でNXが強制終了する。

また、`UNDO_UG_delete_mark: Mark 22 (DXF Export) not found` は、コード内にまだUndoMark関連の処理が残っていることを示す。DxfdwgCreator内部がUndoMarkを消費するため、外部からのUndoToMarkは常に失敗する。

### 修正手順

#### Step 1: UndoMark関連のコードを全て削除

以下のパターンを含む行を**全て削除**してください：
- `SetUndoMark`
- `UndoToMark`
- `DeleteUndoMark`
- `UndoToLastVisibleMark`
- `undoMark` という変数の宣言・代入・参照
- `undoMarkValid` などのフラグ変数

**1行も残さないでください。** `UNDO_UG_delete_mark: Mark 22 (DXF Export) not found` がログに出る限り、UndoMarkのコードが残っています。

#### Step 2: DXFエクスポート後の待機を追加

DxfdwgCreator.Commit() の後、Destroy() の後に、出力ファイルの存在を確認するループで待機してください。Thread.Sleep(2000)のような固定待機ではなく、ファイル完成を確認する方式にしてください：

```csharp
// Commit実行
NXOpen.NXObject result = dxfdwgCreator.Commit();
dxfdwgCreator.Destroy();
dxfdwgCreator = null;

// DXFファイルの出力完了を待機（バックグラウンド変換の完了待ち）
// outputPath は今回のDXF出力先ファイルパス
int waitMs = 0;
int maxWaitMs = 30000; // 最大30秒
while (!File.Exists(outputPath) && waitMs < maxWaitMs)
{
    System.Threading.Thread.Sleep(500);
    waitMs += 500;
}
// ファイルが書き込み中でないか確認（ロックが解放されるまで待つ）
if (File.Exists(outputPath))
{
    for (int retry = 0; retry < 20; retry++)
    {
        try
        {
            using (var fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                break; // ファイルのロックが解放された = 書き込み完了
            }
        }
        catch (IOException)
        {
            System.Threading.Thread.Sleep(500);
        }
    }
}
```

#### Step 3: 図面シート・ビューの手動クリーンアップ

UndoMarkを使わずに、図面シートとビューを手動で削除してください。
**DXFファイル出力完了の確認後に**実行すること：

```csharp
// ★ DXFファイル出力完了を確認した後に実行 ★

// ベースビューを削除
try
{
    if (baseView != null)
    {
        theSession.UpdateManager.ClearErrorList();
        theSession.UpdateManager.AddToDeleteList(baseView);
        int errsFromDelete = 0;
        theSession.UpdateManager.DoUpdate(out errsFromDelete);
        baseView = null;
    }
}
catch { baseView = null; }

// 図面シートを削除
try
{
    if (sheet != null)
    {
        // まずシートを非アクティブにする
        // DeleteSheets で削除
        workPart.DrawingSheets.DeleteSheets(new NXOpen.Drawings.DrawingSheet[] { sheet });
        sheet = null;
    }
}
catch { sheet = null; }
```

#### Step 4: 全ビュー処理完了後にモデリング環境に戻す

```csharp
// 全ビューの処理ループ終了後
try
{
    theSession.ApplicationSwitchImmediate("UG_APP_MODELING");
}
catch { }
```

#### Step 5: finally ブロック

```csharp
finally
{
    // ★ UndoToMarkは一切呼ばない ★

    // 画面更新の抑制を解除
    if (displaySuppressed)
    {
        try
        {
            theUfSession.Disp.SetDisplay(NXOpen.UF.UFConstants.UF_DISP_UNSUPPRESS_DISPLAY);
        }
        catch { }
    }

    // モデリング環境に戻す（まだ戻っていない場合）
    try { theSession.ApplicationSwitchImmediate("UG_APP_MODELING"); } catch { }

    // 一時ファイル削除
    try
    {
        if (!string.IsNullOrEmpty(tempDefPath) && File.Exists(tempDefPath))
            File.Delete(tempDefPath);
    }
    catch { }

    // 完了メッセージ
    System.Threading.Thread.Sleep(500);
    try
    {
        MessageBox.Show(completionMessage, "DXF Export",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    catch { }
}
```

---

## 不具合2: 進捗ダイアログがNXウィンドウの背面に隠れる

進捗ダイアログ（ProgressForm等）の `TopMost` プロパティを確認し、`true` に設定してください。

```csharp
// 進捗ダイアログのプロパティ設定
progressForm.TopMost = true;  // ← 必ず true にする
```

`grep -n "TopMost" CustomViewDxfExporter.cs` の結果を確認し、`TopMost = false` や `TopMost` 未設定の箇所があれば全て `TopMost = true` に修正してください。

---

## 修正完了後の検証

修正後、以下のgrepで確認してください：

```bash
echo "=== UndoMark残存チェック（0件であること） ==="
grep -c "UndoMark\|SetUndoMark\|UndoToMark\|DeleteUndoMark\|undoMark\|UndoToLastVisibleMark" CustomViewDxfExporter.cs

echo "=== TopMost設定チェック（全てtrueであること） ==="
grep -n "TopMost" CustomViewDxfExporter.cs

echo "=== ファイル待機チェック（存在すること） ==="
grep -n "File.Exists\|FileStream\|FileAccess" CustomViewDxfExporter.cs
```

UndoMark関連が0件、TopMostが全てtrue、ファイル存在チェックが実装されていることを確認してください。
