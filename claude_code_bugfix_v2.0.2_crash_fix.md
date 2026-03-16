## c0000005 クラッシュの根本原因分析と修正指示

NXのログを詳細に分析した結果、根本原因を特定しました。以下の修正を実施してください。

---

### ログから判明した事実

1. `UNDO_UG_delete_mark: Mark 155 (Unspecified) not found` — 1回目のDXFエクスポート後、UndoMarkがDxfdwgCreatorのCommitによって消費・削除されている
2. `UNDO_UG_delete_mark: Mark 328 (Unspecified) not found` — 2回目も同様
3. `*** Exception was not in main thread` — クラッシュはメインスレッドではなくバックグラウンドスレッドで発生
4. `+++ Invalid read from 0000000000000000` — NULLポインタ参照（既に存在しないオブジェクトへのアクセス）
5. `DXF/DWG Export: Export translation job submitted.` の直後にクラッシュ — DxfdwgCreator.Commit()は非同期でジョブを投入し、バックグラウンドスレッドで変換処理を実行する
6. `UNDO_UG_delete_mark: Mark 22 (DXF Export) not found` — 最終的なUndoToMarkも失敗している（マークは既に消費済み）

### 根本原因

**DxfdwgCreator.Commit() は非同期でDXF変換ジョブをバックグラウンドスレッドに投入する。** Commit()が戻った直後に図面シートの削除やUndoToMarkを実行すると、バックグラウンドの変換スレッドがまだ動作中であり、変換スレッドがアクセスしようとしたNXオブジェクト（図面シート、ビュー）が既に削除・Undo済みでNULLポインタ参照が発生する。

さらに、DxfdwgCreator.Commit() はNX内部のUndoMarkを消費するため、処理開始時に設定したUndoMarkが最後のUndoToMark時には既に無効になっている。

---

### 修正方針: UndoMark方式を廃止し、手動クリーンアップ方式に変更

#### 理由
- DxfdwgCreator.Commit() がUndoMarkを消費するため、グローバルなUndoToMarkは信頼できない
- 非同期変換ジョブの完了を待つ必要がある

#### 修正内容

**1. グローバルなUndoMarkの設定・復元を廃止する**

```csharp
// ★ 削除: 以下のUndoMark関連コードを全て削除
// NXOpen.Session.UndoMarkId undoMark = theSession.SetUndoMark(...);
// theSession.UndoToMark(undoMark, null);
// theSession.DeleteUndoMark(undoMark, null);
```

**2. 各ビューの処理後、手動でクリーンアップする（Undo不使用）**

各ビューのDXFエクスポート後に、図面シートとベースビューを手動で削除する。
ただし、DxfdwgCreatorのバックグラウンド処理完了を待ってから行うこと。

```csharp
foreach (var viewName in selectedViews)
{
    NXOpen.Drawings.DrawingSheet sheet = null;
    try
    {
        // Step 1: 図面シート作成
        sheet = CreateDrawingSheet(...);
        
        // Step 2: ベースビュー配置
        var baseView = PlaceBaseView(sheet, viewName, ...);
        
        // Step 3: DXFエクスポート
        DxfdwgCreator dxfdwgCreator = ...;
        // ... 設定 ...
        NXOpen.NXObject result = dxfdwgCreator.Commit();
        
        // ★ 重要: Commit後にCreatorを破棄し、バックグラウンド処理の完了を待つ
        dxfdwgCreator.Destroy();
        
        // ★ 重要: バックグラウンドの変換ジョブ完了を待機
        // DxfdwgCreator.Commit()は非同期なので、直後にシート削除するとクラッシュする
        System.Threading.Thread.Sleep(2000);  // 2秒待機
        
        // Step 4: 手動クリーンアップ（図面シートとビューを削除）
        try
        {
            // ベースビューを削除
            if (baseView != null)
            {
                theSession.UpdateManager.ClearErrorList();
                theSession.UpdateManager.AddToDeleteList(baseView);
                // ビュー削除は失敗しても続行
            }
        }
        catch { }
        
        try
        {
            // 図面シートを削除
            if (sheet != null)
            {
                NXOpen.Drawings.DrawingSheet[] sheetsToDelete = new NXOpen.Drawings.DrawingSheet[] { sheet };
                workPart.DrawingSheets.DeleteSheets(sheetsToDelete);
                sheet = null;
            }
        }
        catch { }
        
        successCount++;
    }
    catch (Exception ex)
    {
        failCount++;
        // エラーログに追記
    }
}
```

**3. DxfdwgCreator.Destroy() を Commit() の直後に必ず呼ぶ**

Destroy() を呼ばないと、DxfdwgCreator が内部リソースを保持し続け、後続の処理と干渉する。

```csharp
NXOpen.NXObject result = dxfdwgCreator.Commit();
dxfdwgCreator.Destroy();
dxfdwgCreator = null;
```

**4. 図面シート削除後にモデリングアプリケーションに戻る**

ログに `APP_ENTERED_APPLICATION_ENVIRONMENT_CALLBACK_ID UG_APP_DRAFTING` が出ており、DXFエクスポート後にDraftingアプリケーション環境に入っている。これがクリーンアップ時に問題を起こす可能性がある。

```csharp
// 全ビュー処理完了後、モデリング環境に戻す
try
{
    theSession.ApplicationSwitchImmediate("UG_APP_MODELING");
}
catch { }
```

**5. finally ブロックの修正**

```csharp
finally
{
    // ★ UndoToMarkは使わない

    // 画面更新の抑制を解除
    if (displaySuppressed)
    {
        try
        {
            theUfSession.Disp.SetDisplay(NXOpen.UF.UFConstants.UF_DISP_UNSUPPRESS_DISPLAY);
            theUfSession.Disp.RegenerateDisplay();
        }
        catch { }
    }
    
    // 一時ファイルの削除
    try { if (!string.IsNullOrEmpty(tempDefPath) && File.Exists(tempDefPath)) File.Delete(tempDefPath); } catch { }
    
    // モデリング環境に戻す
    try { theSession.ApplicationSwitchImmediate("UG_APP_MODELING"); } catch { }
    
    // 少し待機してから完了メッセージ表示
    System.Threading.Thread.Sleep(500);
    
    // 完了メッセージ
    MessageBox.Show(completionMessage, "DXF Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
}
```

---

### まとめ: 変更点チェックリスト

- [ ] グローバルUndoMark（SetUndoMark / UndoToMark / DeleteUndoMark）を全て削除
- [ ] DxfdwgCreator.Commit() の直後に必ず dxfdwgCreator.Destroy() を呼ぶ
- [ ] Commit() + Destroy() の後に Thread.Sleep(2000) で待機
- [ ] 各ビュー処理後に図面シートとベースビューを手動で削除（try-catchで囲む）
- [ ] 全ビュー処理完了後に theSession.ApplicationSwitchImmediate("UG_APP_MODELING") でモデリング環境に戻す
- [ ] finally ブロックで UndoToMark を使わず、画面更新解除とファイル削除のみ行う
- [ ] 完了メッセージ表示前に Thread.Sleep(500) を入れる

### 変更しないこと
- DxfdwgCreator の設定内容（SettingsFile、InputFile、OutputFile等）
- dxfdwg.def の埋め込み方式
- ビュー選択ダイアログ
- ファイル名の命名規則
- 多言語対応
- 進捗ダイアログ（TopMost = true を維持）
