# NX Open アドオン開発指示書：カスタムビュー別DXFエクスポーター

## 概要
NXで開いているパートファイルに対して、カスタムビューごとに図面を作成し、DXFファイルとしてエクスポートするNX Openアドオン（C# .NET Framework 4.8 クラスライブラリ）を作成してください。

## 開発環境
- **言語**: C# (.NET Framework 4.8)
- **プロジェクト種別**: クラスライブラリ (.NET Framework)
- **プラットフォーム**: x64
- **NXバージョン**: NX 2406
- **参照DLL** (C:\Program Files\Siemens\NX2406\NXBIN\managed\):
  - NXOpen.dll
  - NXOpen.UF.dll
  - NXOpen.Utilities.dll
  - NXOpenUI.dll

## 機能要件

### 処理フロー
1. **初期処理**
   - NXセッションと作業パートを取得
   - 作業パートが開かれていない場合はエラーメッセージを表示して終了
   - パートファイル名（拡張子なし）を取得（例: `2P684813-1_A`）

2. **出力フォルダの選択**
   - Windows標準のフォルダ選択ダイアログ（FolderBrowserDialog）を表示
   - ユーザーが出力先フォルダを指定
   - キャンセルされた場合は処理を終了

3. **カスタムビューの取得**
   - パートのモデルビュー（ModelingView）を全て取得
   - 以下の標準ビューを除外し、カスタムビューのみを抽出:
     - 上面 (Top)
     - 正面 (Front)
     - 右側面 (Right)
     - 背面 (Back)
     - 下面 (Bottom)
     - 左側面 (Left)
     - 等角投影 (Isometric / Trimetric)
     - 不等角投影
   - ※ 標準ビュー名はNXの言語設定により日本語または英語の場合がある。両方に対応すること
   - カスタムビューが0件の場合はメッセージを表示して終了

4. **カスタムビューごとのループ処理**（以下をカスタムビューの数だけ繰り返す）

   a. **図面シートの作成**
      - 製図（Drafting）アプリケーションに切り替え
      - 新規図面シートを作成
      - シートサイズ: ビューの投影サイズに基づいて自動選択（A4→A3→A2→A1→A0の中から、ビューが収まる最小サイズを選択）
      - シート名: "Sheet 1"（デフォルト）
      - スケール: 1:1（ビューがシートに収まらない場合は縮小）

   b. **ベースビューの配置**
      - カスタムビューを使用してベースビュー（Base View）を作成
      - 「使用するモデルビュー」にカスタムビュー名を指定
      - シートの中央に配置
      - スケールは自動推定（ビュースケールを推定）

   c. **DXFエクスポート**
      - ファイル → エクスポート → AutoCAD DXF/DWG でエクスポート
      - エクスポート設定:
        - エクスポート元: 表示パート
        - 出力タイプ: DXF
        - エクスポートフォーマット: 2D
        - 出力先: レイアウト
        - エクスポートするデータ → 図面 → エクスポート: 選択された図面（現在のシート）
      - 出力ファイル名: `<パート名>_<ビュー名>.dxf`
        - 例: `2P684813-1_A_01_DWG_FRONT.dxf`
      - 出力先: ユーザーが選択したフォルダ

   d. **シートのクリーンアップ**
      - 配置したビューを削除
      - 図面シートを削除
      - ※ 次のビュー処理のためにクリーンな状態にする

5. **終了処理**
   - モデリングアプリケーションに戻る
   - 処理結果をメッセージボックスまたは情報ウィンドウに表示
     - 例: 「DXFエクスポート完了: 9ファイル出力しました」
   - パートファイルは保存しない（図面シートは一時的なもの）

## DXFエクスポート設定の詳細（NX Open API）
NXのDXF/DWGエクスポートは `NXOpen.DxfdwgCreator` クラスを使用します。主要な設定:
- `OutputFileType` = DXF
- `ExportFrom` = ExportFromDisplayPart
- `ProcessHoldFlag` = true
- `FileSaveFlag` = false
- `ExportAs` = ExportAs2d
- `OutputTo` = OutputToLayout
- 出力ファイルパスを設定

## エントリポイント
```csharp
public static void Main(string[] args)
{
    // メイン処理
}

public static int GetUnloadOption(string dummy)
{
    return (int)NXOpen.Session.LibraryUnloadOption.Immediately;
}
```

## エラーハンドリング
- 各ステップでtry-catchを使用
- エラー発生時はNXMessageBoxでエラー内容を表示
- エラーが発生しても、可能な限り次のビューの処理を継続（スキップして続行）
- 最終的に処理結果サマリーを表示（成功数/失敗数）

## 注意事項
- NXの内部処理（インターナルモード）として動作するため、Session.GetSession()でセッションを取得
- UndoMark を使用して、処理前の状態に確実に戻せるようにすること
- 図面シートの作成・削除はNXOpen.Drawings名前空間のAPIを使用
- DXFエクスポートはNXOpen.DxfdwgCreatorを使用
- ビューの配置位置はシートサイズの中央（Width/2, Height/2）
- 標準ビューの判定は、ビュー名の完全一致ではなく、既知の標準ビュー名リストとの比較で行う

## 出力ファイル
- プロジェクト名: `NXCustomViewDxfExporter`
- メインクラスファイル: `CustomViewDxfExporter.cs`
- ビルド出力: `NXCustomViewDxfExporter.dll`

## テスト方法
NXでパートファイルを開いた状態で:
1. Alt + F8（ジャーナルの実行）
2. ファイルの種類をDLLに変更
3. ビルドしたDLLを選択して実行
