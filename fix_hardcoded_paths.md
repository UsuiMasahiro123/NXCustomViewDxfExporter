# 修正指示: パートファイルパスの汎用化

## 概要
現在のCustomViewDxfExporter.csには、ジャーナル記録時の特定パートファイルのフルパスがハードコードされている可能性がある。
どのパートファイルを開いていても動作するよう、すべてのパス参照を動的取得に変更する。

## 修正方針
コード内のすべてのファイルパス参照を確認し、ハードコードされたパスを `theSession.Parts.Work` による動的取得に置き換える。

## 確認・修正すべき箇所

### 1. パート取得部分
以下のように、セッションから動的に作業パートを取得していることを確認：
```csharp
Session theSession = Session.GetSession();
Part workPart = theSession.Parts.Work;
```
特定のパート名（例: `2P684813-1_A` 等）が文字列で記述されている場合は削除し、上記に置き換えること。

### 2. BaseViewBuilder の PartName 設定
```csharp
// ★ 動的に現在のパートのフルパスを使用すること
baseViewBuilder.Style.ViewStyleBase.PartName = workPart.FullPath;
```
ここに固定パス文字列（例: `@"C:\path\to\specific\file.prt"` ）が入っていないことを確認。

### 3. DxfdwgCreator の InputFile 設定
```csharp
// ★ 動的に現在のパートのフルパスを使用すること
dxfdwgCreator.InputFile = workPart.FullPath;
```
ここに固定パス文字列が入っていないことを確認。

### 4. 出力ファイル名の構築
出力DXFファイル名は、パート名を動的に取得して構築すること：
```csharp
// パートのファイル名（拡張子なし）を取得
string partFileName = System.IO.Path.GetFileNameWithoutExtension(workPart.FullPath);

// 出力ファイル名: <パート名>_<ビュー名>.dxf
string outputFileName = partFileName + "_" + viewName + ".dxf";
string outputFilePath = System.IO.Path.Combine(selectedFolder, outputFileName);
```
出力ファイル名にも固定のパート名が埋め込まれていないことを確認。

### 5. 検索パス（アセンブリパート対応）
アセンブリパートの場合、コンポーネントの参照先を解決するために検索パスを設定する場合がある。
この検索パスもハードコードされていないことを確認：
```csharp
// パートファイルのあるフォルダを検索パスに追加
string partDirectory = System.IO.Path.GetDirectoryName(workPart.FullPath);
```

### 6. パートが開かれていない場合のガード
処理開始時に、パートが開かれていることを確認するガードが存在すること：
```csharp
Part workPart = theSession.Parts.Work;
if (workPart == null)
{
    // 日本語/英語メッセージで警告表示して終了
    return;
}
```

## 修正手順

1. `CustomViewDxfExporter.cs` を開く
2. ファイル全体で以下のパターンを検索し、ハードコードされたパスがあれば置き換える：
   - `@"C:\` や `@"D:\` などの絶対パス文字列
   - `.prt"` を含む文字列リテラル
   - 特定のパート名（例: `2P684813` など）を含む文字列
3. すべてのパス参照が `workPart.FullPath` や `Path.GetFileNameWithoutExtension(workPart.FullPath)` 等の動的取得になっていることを確認
4. ビルドして成功することを確認

## 変更しない箇所
以下はハードコードで正しい（変更不要）：
- `dxfdwg.def` の埋め込み文字列定数（`EmbeddedDefContent`）→ エクスポート設定なのでパス非依存
- 標準ビュー除外リスト（`Top`, `Front`, `上面`, `正面` 等）→ ビュー名定数なのでOK
- `DrawingList = @"""Sheet 1"""`  → シート名は固定でOK
- `LayerMask = "1-256"` → レイヤ設定は固定でOK

## テスト方法
修正後、異なるパートファイルを2つ以上開いてそれぞれDLL実行し、以下を確認：
- 正しいパート名でDXFファイルが生成されること
- DXFファイルの内容（ラベル・寸法含む）が正しいこと
- エラーが発生しないこと

## ビルド
修正完了後、ビルドまで実行してください。
