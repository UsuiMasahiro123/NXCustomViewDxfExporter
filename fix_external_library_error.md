# バグ修正: 処理終了後に「外部ライブラリ中のエラーです」が表示される

## 現象
DXFエクスポート処理の終了後（正常完了時・中断時ともに）、NXが以下のエラーダイアログを表示する：
```
Error
外部ライブラリ中のエラーです。詳細はシステムログをご覧ください
ファイル名: C:\Users\tcadmin\source\repos\NXOpenTest\NXCustomViewDxfExporter\bin\x64\Debug\NXCustomViewDxfExporter.dll
関数名:Main
```

## 原因の可能性（上から順に調査・修正すること）

### 1. WinFormsのダイアログがNXのUIスレッドと競合している
NX Open はシングルスレッドで動作しており、WinFormsの `Application.DoEvents()` や `Form.ShowDialog()` がNXの内部メッセージループと干渉する可能性がある。

**対策:**
- `Application.DoEvents()` の呼び出し箇所を最小限にする（ループ内の進捗更新時のみ）
- 完了ダイアログの表示前に、**NXの画面更新を必ず復帰（UNSUPPRESS）してから**表示する
- ダイアログの `Dispose()` を確実に行う
- `Form.ShowDialog()` ではなく `Form.ShowDialog(NXWindowOwner)` のようにオーナーを指定しない（NXウィンドウをオーナーにするとクラッシュする場合がある）

### 2. SetDisplay の SUPPRESS/UNSUPPRESS の不整合
画面更新の抑制を解除しないまま処理を終了すると、NXが内部エラーを発生させる。

**対策:**
- `try-finally` で必ず `SetDisplay(UF_DISP_UNSUPPRESS_DISPLAY)` を呼ぶ
- SUPPRESSとUNSUPPRESSの呼び出し回数が必ず一致すること
- UNSUPPRESS後に少し待つ: `System.Threading.Thread.Sleep(100)` を入れてみる

### 3. UndoMark の復元タイミング
UndoToMark の呼び出しがダイアログ表示と競合している可能性がある。

**対策:**
- 完了ダイアログを表示する**前に** UndoToMark を実行し、状態復元を完了させる
- UndoToMark の後に `DeleteUndoMark` も呼び、マークを確実にクリーンアップする

### 4. NXウィンドウの ShowWindow/Restore のタイミング
`ShowWindow(hWnd, SW_RESTORE)` がNXの内部状態と矛盾する可能性がある。

**対策:**
- `SW_RESTORE` の呼び出し前にダイアログを完全に閉じて Dispose する
- `SW_RESTORE` の後に `System.Threading.Thread.Sleep(200)` を入れてNXの内部処理を待つ

### 5. 例外がcatchされずにNXに伝播している
Main関数内で発生した例外がNXに到達している可能性がある。

**対策:**
- `Main` 関数の最外部を `try-catch(Exception ex)` で囲み、全ての例外をキャッチする
- catch内でNXのListingWindowにエラーを出力し、例外をNXに伝播させない

## 推奨する修正後の処理順序

```csharp
public static void Main(string[] args)
{
    Session theSession = null;
    UFSession theUfSession = null;
    NXOpen.UF.UFPart.LoadStatus loadStatus;
    IntPtr nxHandle = IntPtr.Zero;
    bool displaySuppressed = false;
    NXOpen.Session.UndoMarkId undoMark = default;
    bool undoMarkCreated = false;
    ProgressForm progressForm = null;

    try
    {
        theSession = Session.GetSession();
        theUfSession = UFSession.GetUFSession();
        
        // ... 初期処理（パート取得、ビュー取得、フォルダ選択）...
        
        // UndoMark作成
        undoMark = theSession.SetUndoMark(Session.MarkVisibility.Invisible, "DXF Export");
        undoMarkCreated = true;
        
        // NXウィンドウ最小化
        nxHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
        ShowWindow(nxHandle, SW_MINIMIZE);
        LockSetForegroundWindow(LSFW_LOCK);
        
        // 画面更新抑制
        theUfSession.Disp.SetDisplay(UF_DISP_SUPPRESS_DISPLAY);
        displaySuppressed = true;
        
        // 進捗ダイアログ表示
        progressForm = new ProgressForm(...);
        progressForm.Show();
        Application.DoEvents();
        
        // === ループ処理 ===
        for (int i = 0; i < customViews.Count; i++)
        {
            if (progressForm.StopRequested) { /* 中断処理 */ break; }
            progressForm.UpdateProgress(...);
            Application.DoEvents();
            // ... ビュー処理 ...
        }
        
        // === ループ終了後のクリーンアップ（順序が重要！） ===
        
        // Step 1: 進捗ダイアログを閉じる
        if (progressForm != null)
        {
            progressForm.Close();
            progressForm.Dispose();
            progressForm = null;
        }
        
        // Step 2: 画面更新を復帰（完了ダイアログ表示前に必ず）
        if (displaySuppressed)
        {
            theUfSession.Disp.SetDisplay(UF_DISP_UNSUPPRESS_DISPLAY);
            displaySuppressed = false;
        }
        
        // Step 3: UndoMarkで状態を復元（完了ダイアログ表示前に）
        if (undoMarkCreated)
        {
            theSession.UndoToMark(undoMark, "DXF Export");
            theSession.DeleteUndoMark(undoMark, "DXF Export");
            undoMarkCreated = false;
        }
        
        // Step 4: モデリングアプリケーションに戻る
        // ※ 必要に応じて
        
        // Step 5: 完了ダイアログ表示
        using (var completionForm = new CompletionForm(...))
        {
            completionForm.ShowDialog();
        }
        // completionFormはここでDispose済み
        
    }
    catch (Exception ex)
    {
        // 全ての例外をキャッチ（NXに伝播させない）
        try
        {
            // 進捗ダイアログが開いていたら閉じる
            if (progressForm != null)
            {
                progressForm.Close();
                progressForm.Dispose();
                progressForm = null;
            }
            
            // エラーメッセージ表示
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        catch { }
    }
    finally
    {
        // 確実に復帰（二重呼び出し防止のフラグ付き）
        try
        {
            if (displaySuppressed)
            {
                theUfSession.Disp.SetDisplay(UF_DISP_UNSUPPRESS_DISPLAY);
            }
        }
        catch { }
        
        try
        {
            if (undoMarkCreated)
            {
                theSession.UndoToMark(undoMark, "DXF Export");
                theSession.DeleteUndoMark(undoMark, "DXF Export");
            }
        }
        catch { }
        
        // NXウィンドウ復帰
        try
        {
            if (nxHandle != IntPtr.Zero)
            {
                ShowWindow(nxHandle, SW_RESTORE);
            }
            LockSetForegroundWindow(LSFW_UNLOCK);
        }
        catch { }
    }
}
```

## 重要ポイントまとめ

1. **順序を厳守**: 進捗ダイアログ閉じる → 画面更新復帰 → UndoMark復元 → 完了ダイアログ表示 → NXウィンドウ復帰
2. **二重呼び出し防止**: `displaySuppressed`, `undoMarkCreated` 等のフラグで、finallyブロックでの二重実行を防ぐ
3. **例外の非伝播**: Main関数の最外部で全例外をcatchし、NXに例外を伝播させない
4. **Dispose の徹底**: 全てのFormを確実にDisposeする
5. **ビルドして動作確認まで実施すること**
