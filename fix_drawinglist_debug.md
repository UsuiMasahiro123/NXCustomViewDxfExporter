## 目的
DXFファイルが「入力ファイルが無効」エラーで開けない問題を調査する。

## 修正内容
ExportViewAsDxf関数内の以下の行の直前：
    dxfCreator.DrawingList = string.Format(@"""{0}""", sheet.Name);

以下のデバッグログを追加する：
    lw.WriteLine("[DEBUG] sheet.Name = " + sheet.Name);
    lw.WriteLine("[DEBUG] DrawingList = " + string.Format(@"""{0}""", sheet.Name));

ビルドまで実施してください。