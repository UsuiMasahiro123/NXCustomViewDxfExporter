# バグ修正: 「アンドゥマークがありません」エラーの修正

## 現象
DXFエクスポート処理において、以下の**両方**のケースでエラーが表示される：
- 進捗ダイアログの停止ボタンを押して中断した時
- 正常に全ビュー処理が完了した時

```
エラー
予期しないエラーが発生しました:
アンドゥマークがありません
```

## ログ分析結果（重要）

NXのシステムログに、**各DXFエクスポートのたびに**以下が記録されている：
```
UNDO_UG_delete_mark: Mark 37140 (Unspecified) not found
UNDO_UG_delete_mark: Mark 37291 (Unspecified) not found
（各ビューのエクスポートごとに異なるMark番号で繰り返し出現）
```

これはNXのDXFエクスポート処理が**内部的にUndoMarkを操作**しており、
その結果、コード内で `SetUndoMark` で作成したマークが**無効化**されている。
したがって、ループ後に `UndoToMark` / `DeleteUndoMark` を呼ぶと
「アンドゥマークがありません」エラーになる。

## 根本原因（2つの問題の組み合わせ）

### 問題1: NX内部によるUndoMark無効化
NXのDXFエクスポートAPI（DxfdwgCreator.Commit等）が内部でUndoMarkを操作するため、
ループ中に私たちのUndoMarkが無効化される。

### 問題2: UndoMark操作の二重実行
仮にマークが有効であっても、正常フロー内で `UndoToMark` + `DeleteUndoMark` を実行した後、
`undoMarkCreated` フラグが `true` のまま残り、`finally` ブロックで再度実行される。

## 修正方法

### Step 1: SafeUndoAndCleanup ヘルパーメソッドを追加する

コード内に以下のヘルパーメソッドを追加する（Mainメソッドの近くにstaticメソッドとして）：

```csharp
/// <summary>
/// UndoMarkの復元と削除を安全に1回だけ実行するヘルパー
/// NX内部のDXFエクスポートがUndoMarkを無効化する場合があるため、
/// try-catchで保護し、二重実行を防止するためフラグをリセットする
/// </summary>
private static void SafeUndoAndCleanup(Session theSession, ref NXOpen.Session.UndoMarkId undoMark, ref bool undoMarkCreated)
{
    if (!undoMarkCreated) return;

    try
    {
        theSession.UndoToMark(undoMark, "DXF Export");
    }
    catch (Exception) { /* NX内部でマークが無効化された場合は無視 */ }

    try
    {
        theSession.DeleteUndoMark(undoMark, "DXF Export");
    }
    catch (Exception) { /* NX内部でマークが無効化された場合は無視 */ }

    undoMarkCreated = false; // ★重要: フラグをリセットして二重実行を完全に防止
}
```

### Step 2: コード内の全ての UndoToMark / DeleteUndoMark 呼び出しを置換する

コード全体を検索して、`UndoToMark` と `DeleteUndoMark` を直接呼んでいる箇所を**全て** `SafeUndoAndCleanup` に置換する。

**検索対象キーワード:**
- `UndoToMark`
- `DeleteUndoMark`

以下の場所に存在するはず（全て置換する）：
1. **正常フロー**（ループ終了後のクリーンアップ部分）
2. **中断フロー**（StopRequested時の処理部分）
3. **finallyブロック**（最終クリーンアップ部分）
4. **catchブロック**（エラー時のクリーンアップ部分）

**置換ルール:**
```csharp
// 修正前（どこに書いてあっても全て置換）
theSession.UndoToMark(undoMark, "DXF Export");
theSession.DeleteUndoMark(undoMark, "DXF Export");

// 修正後
SafeUndoAndCleanup(theSession, ref undoMark, ref undoMarkCreated);
```

### Step 3: エラーメッセージの表示部分を確認

catchブロック内でエラーメッセージを表示している箇所を確認し、
`UndoToMark` / `DeleteUndoMark` のエラーがユーザーに見える形で表示されないようにする。

具体的には、Main関数の最外部のcatchで全例外をキャッチしているか確認し、
NXに例外を伝播させないようにする。

```csharp
// Main関数の最外部
catch (Exception ex)
{
    // 全ての例外をキャッチ（NXに伝播させない）
    try
    {
        MessageBox.Show(...);
    }
    catch { }
}
```

### Step 4: 修正後のMain関数の構造（参考）

```csharp
public static void Main(string[] args)
{
    Session theSession = null;
    UFSession theUfSession = null;
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

        // NXウィンドウ最小化、画面抑制
        // ...

        // === ループ処理 ===
        for (int i = 0; i < customViews.Count; i++)
        {
            if (progressForm.StopRequested)
            {
                break;
            }
            // ... ビュー処理・DXFエクスポート ...
        }

        // === ループ終了後のクリーンアップ（正常完了・中断共通）===

        // 1. 進捗ダイアログを閉じる
        if (progressForm != null)
        {
            progressForm.Close();
            progressForm.Dispose();
            progressForm = null;
        }

        // 2. 画面更新を復帰
        if (displaySuppressed)
        {
            theUfSession.Disp.SetDisplay(UF_DISP_UNSUPPRESS_DISPLAY);
            displaySuppressed = false;
            System.Threading.Thread.Sleep(100);
        }

        // 3. UndoMarkで状態復元（★SafeUndoAndCleanupを使う）
        SafeUndoAndCleanup(theSession, ref undoMark, ref undoMarkCreated);

        // 4. NXウィンドウ復帰
        if (nxHandle != IntPtr.Zero)
        {
            ShowWindow(nxHandle, SW_RESTORE);
            System.Threading.Thread.Sleep(200);
        }
        LockSetForegroundWindow(LSFW_UNLOCK);

        // 5. 完了ダイアログ表示
        using (var completionForm = new CompletionForm(...))
        {
            completionForm.ShowDialog();
        }
    }
    catch (Exception ex)
    {
        // 全ての例外をキャッチ（NXに伝播させない）
        try
        {
            if (progressForm != null)
            {
                progressForm.Close();
                progressForm.Dispose();
                progressForm = null;
            }
        }
        catch { }

        try
        {
            MessageBox.Show(
                String.Format(GetString("MSG_UNEXPECTED_ERROR"), ex.Message),
                GetString("MSG_ERROR_TITLE"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        // ★ SafeUndoAndCleanup はフラグで二重実行を防止するので安全に呼べる
        try
        {
            SafeUndoAndCleanup(theSession, ref undoMark, ref undoMarkCreated);
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

## 修正のポイントまとめ

1. **SafeUndoAndCleanup メソッド**を新規追加 - try-catchで保護し、フラグで二重実行防止
2. コード内の**全ての** `UndoToMark` / `DeleteUndoMark` 直接呼び出しを `SafeUndoAndCleanup` に置換
3. NX内部のDXFエクスポートがUndoMarkを無効化するため、try-catchでの保護が**必須**
4. Main関数の最外部catchで全例外をキャッチし、NXに伝播させない

## 確認方法
修正後、以下の両方でエラーが出ないことを確認：
1. 全ビュー正常処理完了時
2. 途中で停止ボタンを押して中断した時

## ビルドまでお願いします
