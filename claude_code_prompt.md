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
   - **NXウィンドウのフォーカス奪取を抑制する設定を行う**（後述の「バックグラウンド処理」参照）

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
      - ※ シート削除時に「CGMを保存」ダイアログが出る場合があるので、事前に抑制すること

5. **終了処理**
   - 全ての一時的に作成した図面シートが削除されていることを確認
   - モデリングアプリケーションに戻る
   - UndoMarkを使って処理前の状態に完全に戻す（パートに変更を残さない）
   - 処理結果をメッセージボックスまたは情報ウィンドウに表示
     - 例: 「DXFエクスポート完了: 9ファイル出力しました」
   - パートファイルは**絶対に保存しない**（図面シートは一時的なもの。保存するとCGMダイアログが発生する原因になる）

## DXFエクスポート設定の詳細（NX Open API）
NXのDXF/DWGエクスポートは `NXOpen.DxfdwgCreator` クラスを使用します。

### 基本設定
- `OutputFileType` = DXF
- `ExportFrom` = ExportFromDisplayPart
- `ProcessHoldFlag` = true
- `FileSaveFlag` = false
- `ExportAs` = ExportAs2d
- `OutputTo` = OutputToLayout
- **`AutoCADRevision` = AC2018（DXF/DWGリビジョン: 2018-2024）** ※古いリビジョンだと寸法や線種が正しく出力されないため必須
- 出力ファイルパスを設定

### 寸法の出力設定（重要）
DXF変換時に寸法が欠落・破線化する問題を防ぐため、以下を必ず設定すること:
- **寸法はAutoCAD寸法オブジェクトとしてエクスポートする**（分解しない）
  - `ExportDimensionsAs` = AutoCAD寸法オブジェクト（DXF側で編集可能にするため）
- NX寸法スタイルをAutoCAD寸法スタイルに正しくマッピングすること

### 線スタイルの設定（重要）
NXの線種とAutoCADの線種が正しくマッピングされないと、実線が破線になる問題が発生する:
- NX線フォントをAutoCAD線種に正しくマッピングすること
- 特にNXの実線（Solid）がAutoCADのContinuousに対応していることを確認
- `LineFontMappingFile` は使用しない（デフォルトマッピングを使用）

### フォントの設定
- NXフォントをAutoCADフォントにマッピング
- 日本語テキストが含まれる場合、対応するAutoCADフォント（bigfont等）を指定

### レイヤー設定
- 全レイヤー（1-256）をエクスポート対象にすること
- 寸法レイヤーが除外されないよう注意

### ビューのフラット化
- ビューをフラット化（Flatten Views）してエクスポートする
- これにより、ビューポート内のジオメトリが直接レイアウト空間に展開される

### エクスポートデータの選択
- エクスポート: 選択された図面（現在のシート）
- 全てのオブジェクトタイプを含める（寸法、注記、ジオメトリ等）

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

## CGM保存ダイアログの抑制（重要）
処理中に「CGMを保存」ダイアログ（「表示されていないシートを含むパートは、それらと一緒にCGMを保存せず、既存のCGMは削除されます」）が表示され、処理が中断する問題がある。以下の対策を**必ず**実装すること:

### 対策1: UIダイアログの自動応答（推奨）
処理開始前にNXのUIメッセージを抑制する:
```csharp
// UFSessionを使ってUIダイアログを抑制
NXOpen.UF.UFSession theUfSession = NXOpen.UF.UFSession.GetUFSession();

// 処理開始前: UIメッセージフィルタを設定し、ダイアログを自動で閉じる
// NXOpen.Session.GetSession().SetUndoMarkVisibility() 等を活用

// 環境変数でCGM保存を無効化
System.Environment.SetEnvironmentVariable("UGII_CGM_FITS_FILE_SAVE", "0");
```

### 対策2: パートのCGM設定を無効化
```csharp
// パート保存時にCGMデータを含めない設定
// theSession.Parts.Work.SaveOptions でCGM関連オプションを無効化
```

### 対策3: UserInterfaceプレファレンスの変更
```csharp
// NXの設定でCGM保存の確認ダイアログを非表示にする
// Preferences.UserInterface で該当する設定を変更
```

上記のいずれか、または組み合わせで、ダイアログが一切表示されずに処理が完了するようにすること。もしダイアログの抑制が困難な場合は、図面シートの削除順序やタイミングを工夫して、CGMダイアログが発生しない処理フローにすること（例: シートを削除せずに内容をクリアする、全ビュー処理後にまとめて削除する等）。

## 注意事項
- NXの内部処理（インターナルモード）として動作するため、Session.GetSession()でセッションを取得
- UndoMark を使用して、処理前の状態に確実に戻せるようにすること
- 図面シートの作成・削除はNXOpen.Drawings名前空間のAPIを使用
- DXFエクスポートはNXOpen.DxfdwgCreatorを使用
- ビューの配置位置はシートサイズの中央（Width/2, Height/2）
- 標準ビューの判定は、ビュー名の完全一致ではなく、既知の標準ビュー名リストとの比較で行う
- **処理中にNXのダイアログ（CGM保存、確認メッセージ等）が表示されて処理が中断しないよう、全てのダイアログを事前に抑制すること**
- パートファイルは処理完了後に保存しないこと（UndoMarkで元に戻す）。保存しないことでCGMダイアログの発生も防げる可能性がある

## バックグラウンド処理（重要）
変換処理中もユーザーが同じPC上で他の作業（メモ帳でのテキスト入力等）を継続できるようにすること。現状、NXウィンドウがたびたびアクティブ（前面）になり、他のアプリケーションでの作業が中断される問題がある。

### 必須対策
以下の全てを実装すること:

1. **画面更新の抑制**
   - ループ処理の開始前に `theUfSession.Disp.SetDisplay(NXOpen.UF.UFConstants.UF_DISP_SUPPRESS_DISPLAY)` で画面更新を停止
   - ループ処理の完了後に `theUfSession.Disp.SetDisplay(NXOpen.UF.UFConstants.UF_DISP_UNSUPPRESS_DISPLAY)` で画面更新を復帰
   - これにより、シート作成・ビュー配置・削除のたびにNXウィンドウが再描画されるのを防ぐ

2. **NXウィンドウがフォーカスを奪わないようにする**
   - NXウィンドウの自動的な前面化（アクティブ化）を防止する
   - **NXウィンドウの最小化・最大化はしない**（ユーザーが自由に操作できるようにする）
   - Win32 API を使用してフォーカス奪取のみを制御:
   ```csharp
   [DllImport("user32.dll")]
   static extern bool LockSetForegroundWindow(uint uLockCode);
   const uint LSFW_LOCK = 1;
   const uint LSFW_UNLOCK = 2;
   ```
   - 処理開始前: `LockSetForegroundWindow(LSFW_LOCK)` でNXが前面に来るのを防止
   - 処理完了後: `LockSetForegroundWindow(LSFW_UNLOCK)` でロック解除
   - ユーザーは処理中も自由にNXウィンドウを最小化・最大化・移動できる

3. **ビューの更新を抑制**
   - `theSession.SetUndoMarkVisibility()` でアンドゥマークの可視性を制御
   - ステータスバーへの進捗表示のみ行い、画面描画は抑制

### 処理フローへの適用
```
フォルダ選択ダイアログ表示 → ユーザー操作完了
↓
画面更新を抑制（SetDisplay SUPPRESS）
フォーカスロック（LockSetForegroundWindow LOCK）
↓
カスタムビューのループ処理（NXが前面に来ない状態で実行）
↓
画面更新を復帰（SetDisplay UNSUPPRESS）
フォーカスロック解除（LockSetForegroundWindow UNLOCK）
結果メッセージを表示
```

### 注意
- 画面更新の抑制・復帰は必ずtry-finallyで囲み、エラー時も確実に復帰すること
- フォーカスロック・解除も同様にtry-finallyで保護すること
- NXウィンドウの最小化・最大化・移動はユーザーに委ねること（プログラムから操作しない）
- 処理の進捗はNXのステータスバーまたはログ出力で確認可能にする

## 出力ファイル
- プロジェクト名: `NXCustomViewDxfExporter`
- メインクラスファイル: `CustomViewDxfExporter.cs`
- ビルド出力: `NXCustomViewDxfExporter.dll`

## テスト方法
NXでパートファイルを開いた状態で:
1. Alt + F8（ジャーナルの実行）
2. ファイルの種類をDLLに変更
3. ビルドしたDLLを選択して実行
