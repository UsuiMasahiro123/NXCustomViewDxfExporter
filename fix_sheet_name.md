## 問題
図面シートが毎回「Sheet 1」という名前で作成されるため、
DrawingListが常に最初のシートを参照してしまう。

## 修正
sheetBuilderのCommit前に、シート名をビュー名で明示的に設定する。

ExportViewAsDxf関数内のsheetBuilder.Commit()の直前に以下を追加：
    sheetBuilder.Name = SanitizeFileName(view.Name);

またDEBUGログは不要になったので削除してください。

ビルドまで実施してください。