はい、その通りです。完全な仕様を以下に記載します。

## 変更1: DXFファイル名の変更

### ファイル名フォーマット
```
{図面番号}_{累進}_{ビュー名}.dxf
```

- **図面番号**: パート属性 `DB_PART_NO` の値（NXのパート属性から取得）
- **累進**: パート属性 `DB_PART_REV` の値（NXのパート属性から取得）
- **ビュー名**: カスタムビューの Name プロパティ（既存と同じ）

### 属性の取得方法

```csharp
string partNo = "";
string partRev = "";

// 方法1: GetUserAttributes
try
{
    NXObject.AttributeInformation[] attrs = workPart.GetUserAttributes();
    foreach (var attr in attrs)
    {
        if (attr.Title.Equals("DB_PART_NO", StringComparison.OrdinalIgnoreCase))
            partNo = attr.StringValue?.Trim() ?? "";
        if (attr.Title.Equals("DB_PART_REV", StringComparison.OrdinalIgnoreCase))
            partRev = attr.StringValue?.Trim() ?? "";
    }
}
catch
{
    // 属性取得失敗時は空文字のまま
}

// 方法2（方法1が使えない場合の代替）:
// try { partNo = workPart.GetStringAttribute("DB_PART_NO"); } catch (NXException) { partNo = ""; }
// try { partRev = workPart.GetStringAttribute("DB_PART_REV"); } catch (NXException) { partRev = ""; }

// DB_PART_NO が空の場合、ItemID を代用
if (string.IsNullOrEmpty(partNo))
{
    try { partNo = workPart.Leaf; }
    catch { partNo = Path.GetFileNameWithoutExtension(workPart.FullPath); }
}
```

### 重複ファイル名の処理
出力フォルダに同名ファイルが既に存在する場合**のみ**、末尾に連番を付与：

```
{図面番号}_{累進}_{ビュー名}.dxf          ← 初回（番号なし）
{図面番号}_{累進}_{ビュー名}_(1).dxf      ← 同名ファイルが既に存在する場合
{図面番号}_{累進}_{ビュー名}_(2).dxf      ← (1)も存在する場合
```

```csharp
string GetUniqueFilePath(string folder, string baseName, string extension)
{
    string filePath = Path.Combine(folder, baseName + extension);
    if (!File.Exists(filePath))
        return filePath;
    int counter = 1;
    while (true)
    {
        filePath = Path.Combine(folder, $"{baseName}_({counter}){extension}");
        if (!File.Exists(filePath))
            return filePath;
        counter++;
    }
}

// 使用例
string baseName = $"{partNo}_{partRev}_{viewName}";
baseName = string.Join("_", baseName.Split(Path.GetInvalidFileNameChars()));
string outputPath = GetUniqueFilePath(outputFolder, baseName, ".dxf");
```

### ファイル名のサニタイズ
- DB_PART_NO, DB_PART_REV, ビュー名に含まれるファイル名禁止文字（`\ / : * ? " < > |`）はアンダースコア `_` に置換
- 連続するアンダースコアは1つにまとめる（例: `A__B` → `A_B`）
- 先頭・末尾のアンダースコアは除去

---

## 変更2: 出力フォルダの固定化

FolderBrowserDialog を廃止し、Picturesフォルダに固定：

```csharp
string outputFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
// 結果例: "C:\Users\tcadmin\Pictures"
Directory.CreateDirectory(outputFolder); // 存在しない場合は作成
```

- FolderBrowserDialog の生成・表示・結果取得の処理を全て削除
- フォルダ選択ダイアログ関連の多言語メッセージも辞書から削除
- 完了メッセージに出力先フォルダのパスを表示すること

---

## 変更3: NXウィンドウ最小化の削除

- ShowWindow による SW_MINIMIZE / SW_RESTORE 呼び出しを全て削除
- 最小化目的の Win32 API 呼び出し（LockSetForegroundWindow, RestoreFocus等）を削除
- 画面更新の抑制（UF_DISP_SUPPRESS_DISPLAY / UF_DISP_UNSUPPRESS_DISPLAY）は**そのまま残す**
- 進捗ダイアログの TopMost = true は TopMost = false に変更（NXの上に被さる必要がなくなったため）

---

## 変更4: ビュー選択ダイアログの追加

カスタムビュー取得後、ループ処理開始前にダイアログを表示：

### ダイアログ仕様
- 最上部に「ALL（すべて選択）」チェックボックスを配置、**デフォルトでON**
- ALL がONの場合、個別ビューチェックボックスは**全てON＋グレーアウト（無効化）**
- ALL のチェックを外すと個別ビューが有効化され、ユーザーが個別に選択/解除可能
- 個別選択で全てにチェックを入れると ALL が自動的にON
- 1つもチェックがない状態では OKボタンを無効化
- Cancel またはダイアログを閉じた場合は処理全体をキャンセル

### 多言語対応（既存のUGII_LANG判定を使用）
| UI要素 | 日本語 | 英語 |
|--------|--------|------|
| タイトル | ビュー選択 | Select Views |
| ALL | すべて選択 | Select All |
| Cancel | キャンセル | Cancel |

### 実装
- System.Windows.Forms.Form でカスタムダイアログを作成
- CheckedListBox または CheckBox のリストで個別ビューを表示

---

## 変更しないこと
- DXFエクスポート処理のコアロジック（DxfdwgCreator設定、図面シート作成、ベースビュー配置、エクスポート、シート削除）
- dxfdwg.def の埋め込み方式
- UndoMarkによる状態復元
- エラーハンドリングの基本構造
- 多言語対応の基本構造
- 用紙サイズ自動選択ロジック
- 検索パス自動設定
- ラベル欠落対策（InputFile/PartName設定）

不要になったコードは完全に削除し、コメントアウトで残さないでください。
