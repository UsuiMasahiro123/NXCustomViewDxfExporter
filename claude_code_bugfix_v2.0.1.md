以下の2つの不具合を修正してください。

## 不具合1: 進捗ダイアログがNXウィンドウの背面に隠れる

### 現象
処理中の進捗ダイアログがNXウィンドウの背面に回ってしまい、ユーザーから見えなくなる。

### 修正
進捗ダイアログの `TopMost = true` に戻してください。前回の改修で `TopMost = false` に変更しましたが、NXが画面更新時にフォーカスを奪うため、TopMost が必要です。

```csharp
progressForm.TopMost = true;
```

---

## 不具合2: 処理完了後に「Unhandled operating system exception: c0000005」が発生しNXが強制終了する

### 現象
DXFエクスポート処理が完了した後、以下のエラーダイアログが表示される：
```
フェータルエラーが検出されました - 継続することはできません。
Unhandled operating system exception: c0000005
```
OKを押すとNXが強制終了される。c0000005 は ACCESS_VIOLATION（メモリアクセス違反）。

### 原因の可能性と修正

#### 原因1（最も可能性が高い）: UndoMark復元後に無効なオブジェクトにアクセスしている

UndoToMark で状態を巻き戻した後、既に無効になったNXオブジェクト（図面シート、ビュー等）にアクセスしている可能性がある。

**修正:**
- UndoToMark の呼び出し後は、NXオブジェクトへの参照を一切使用しない
- UndoToMark の直前に、全てのNXオブジェクト参照を null にする
- UndoToMark 後の処理は、完了メッセージの表示のみにする

```csharp
// UndoToMark の前にオブジェクト参照をクリア
drawingSheet = null;
baseView = null;
// その他のNXオブジェクト参照も全て null に

// UndoMark復元
theSession.UndoToMark(undoMark, null);

// ★ この後はNXオブジェクトに一切アクセスしない
// 完了メッセージの表示のみ
```

#### 原因2: UndoToMark自体が失敗している

DxfdwgCreator の Commit がUndoMarkを消費してしまい、UndoToMark 時に無効なマークを参照している可能性がある。

**修正:**
- UndoToMark を try-catch で囲み、失敗した場合は UndoToLastVisibleMark にフォールバック

```csharp
try
{
    theSession.UndoToMark(undoMark, null);
}
catch
{
    try
    {
        // フォールバック: 最後の有効なマークまで戻す
        theSession.UndoToLastVisibleMark();
    }
    catch
    {
        // それも失敗した場合は何もしない（パートを保存しない設計なので安全）
    }
}
```

#### 原因3: 画面更新抑制の解除タイミング

UF_DISP_UNSUPPRESS_DISPLAY の解除後にUndoToMarkを呼ぶと、NXが無効な描画対象を再描画しようとしてクラッシュする可能性がある。

**修正:**
- UndoToMark を**先に**実行してから、画面更新の抑制を解除する

```csharp
finally
{
    // 1. まずUndoMark復元（画面更新抑制中に行う）
    try { theSession.UndoToMark(undoMark, null); } catch { }
    
    // 2. その後で画面更新の抑制を解除
    try { theUfSession.Disp.SetDisplay(UF_DISP_UNSUPPRESS_DISPLAY); } catch { }
    
    // 3. 一時ファイルの削除
    try { if (File.Exists(tempDefPath)) File.Delete(tempDefPath); } catch { }
}
```

#### 原因4: 完了メッセージのMessageBox表示タイミング

NXの内部状態が不安定なタイミングでMessageBoxを表示すると、NXのイベントループと干渉してクラッシュする場合がある。

**修正:**
- MessageBoxの代わりに NXOpen の ListingWindow を使用する
- または、MessageBox表示前に少し待機する

```csharp
// ListingWindow を使う方法（より安全）
NXOpen.ListingWindow lw = theSession.ListingWindow;
lw.Open();
lw.WriteLine(completionMessage);
lw.Close();

// MessageBoxを使う場合は、少し待機してから表示
System.Threading.Thread.Sleep(500);
MessageBox.Show(completionMessage, "DXF Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
```

### 推奨される finally ブロックの実装順序

```csharp
finally
{
    // Step 1: NXオブジェクト参照をクリア
    drawingSheet = null;
    baseView = null;
    
    // Step 2: UndoMark復元（画面更新抑制中に実行）
    if (undoMarkValid)
    {
        try
        {
            theSession.UndoToMark(undoMark, null);
        }
        catch
        {
            try { theSession.UndoToLastVisibleMark(); } catch { }
        }
    }
    
    // Step 3: 画面更新の抑制を解除
    if (displaySuppressed)
    {
        try
        {
            theUfSession.Disp.SetDisplay(NXOpen.UF.UFConstants.UF_DISP_UNSUPPRESS_DISPLAY);
        }
        catch { }
    }
    
    // Step 4: 一時ファイル削除
    try { if (!string.IsNullOrEmpty(tempDefPath) && File.Exists(tempDefPath)) File.Delete(tempDefPath); } catch { }
    
    // Step 5: 完了メッセージ（全クリーンアップ後に表示）
    try
    {
        System.Threading.Thread.Sleep(300);
        MessageBox.Show(completionMessage, "DXF Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    catch { }
}
```

### 重要
- 上記の原因1〜4を**全て**反映してください。複合的な問題の可能性が高いです
- 不要になったコードはコメントアウトで残さず完全に削除すること
- 既存のDXFエクスポート処理のコアロジックは変更しないこと
