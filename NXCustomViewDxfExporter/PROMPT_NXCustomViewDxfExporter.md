# NXCustomViewDxfExporter 改修版開発プロンプト

> NX 2312 ジャーナル（2件）確認済み / 作成日：2026-03-13 / FPT Software Japan

---

## タスク

NXCustomViewDxfExporter アドオンの改修版を、NX 2312 で録画したジャーナル（2件）に基づいて実装する。  
既存ソースを参考にしつつ、以下の全仕様を満たすこと。

---

## 環境・前提

- 言語：C# / .NET Framework 4.8
- NX バージョン：NX 2312.8701（Daikin 本番環境と一致）
- Teamcenter 連携あり（読み取り専用パーツへの書き込み試行は警告のみ、処理継続で OK）
- NXOpen 参照 DLL はすべて `<Private>false</Private>` にすること
- `dxfdwg.def` は EmbeddedResource として DLL に埋め込み、実行時に一時ファイルに書き出して `SettingsFile` に設定する
- DLL 署名：`C:\Program Files\Siemens\NX2312\NXBIN\SignDotNet.exe`

---

## 既存ソースの確認

まず以下を実行して既存実装を把握すること：

```bash
find . -name '*.cs' | head -30
cat <メイン実装ファイル>
cat NXCustomViewDxfExporter.csproj
```

---

## 実装仕様

### 1. メイン処理フロー（1ビューにつき）

1. 製図アプリに切り替え（`ApplicationSwitchImmediate` + `EnterDraftingApplication`）
2. 対象 ModelingView の PMI から文字サイズを取得（§6 参照）
3. 仮シート（10000×10000mm）を `CustomSize` で作成
4. BaseView をスケール 1:1・Freeze/Unfreeze で配置
5. `DraftingView.GetBoundingBox()` でビューサイズを取得
6. 最適用紙サイズを決定してシートを最終サイズに再設定
7. DXF 出力（Freeze/Unfreeze で囲む）
8. 一時シートをクローズし UndoMark を安全に削除して元の状態に戻す

---

### 2. DXF 出力（案B：DrawingsFreeze によるモデル更新抑制）

`dxfdwgCreator.Commit()` の直前に `DrawingsFreezeOutOfDateComputation()` を呼び、  
直後に `DrawingsUnfreezeOutOfDateComputation()` を呼ぶこと。  
これにより DXF 変換時の製図シート再計算とモデル更新を抑制し、処理速度を改善する。

DXF 出力設定の確定値（ジャーナルより）：

```csharp
dxfdwgCreator.ExportData       = NXOpen.DxfdwgCreator.ExportDataOption.Drawing;
dxfdwgCreator.OutputTo         = NXOpen.DxfdwgCreator.OutputToOption.Drafting;
dxfdwgCreator.OutputFileType   = NXOpen.DxfdwgCreator.OutputFileTypeOption.Dxf;
dxfdwgCreator.AutoCADRevision  = NXOpen.DxfdwgCreator.AutoCADRevisionOptions.R2018;
dxfdwgCreator.FlattenAssembly  = false;
dxfdwgCreator.ViewEditMode     = true;
dxfdwgCreator.ExportScaleValue = 1.0;
dxfdwgCreator.LayerMask        = "1-256";
dxfdwgCreator.DrawingList      = "_ALL_";
dxfdwgCreator.ProcessHoldFlag  = true;
dxfdwgCreator.WidthFactorMode  = NXOpen.DxfdwgCreator.WidthfactorMethodOptions.AutomaticCalculation;
dxfdwgCreator.ObjectTypes.Curves      = true;
dxfdwgCreator.ObjectTypes.Annotations = true;
dxfdwgCreator.ObjectTypes.Structures  = true;
dxfdwgCreator.InputFile    = workPart.FullPath;
dxfdwgCreator.OutputFile   = outputDxfPath;
dxfdwgCreator.SettingsFile = defPath; // 埋め込み dxfdwg.def を一時展開したパス

// Freeze で囲む
workPart.DraftingManager.DrawingsFreezeOutOfDateComputation();
dxfdwgCreator.Commit();
workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();
dxfdwgCreator.Destroy();
```

---

### 3. DXF 出力後の環境切替を削除

ジャーナルには DXF 出力後に以下の 3 行が残っているが、**アドオンコードには含めないこと**。  
含めるとアプリ環境コールバックが 2 重発火してレスポンスが著しく低下する。

```csharp
// ★ 以下はアドオンに含めない
// theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING");
// workPart.Drafting.EnterDraftingApplication();
// workPart.Views.WorkView.UpdateCustomSymbols();
```

---

### 4. View 参照エラー対策（Expression 削除の Freeze）

投影ビュービルダーをキャンセルする場合（Commit しない場合）、  
Expression の削除は必ず `DrawingsFreeze/Unfreeze` で囲むこと。  
囲まないと破棄済みビューへの参照エラー（`GFX_change_view_attr: View N does not exist`）が発生して DXF 出力が失敗する。

```csharp
workPart.DraftingManager.DrawingsFreezeOutOfDateComputation();
projectedViewBuilder.Destroy();
workPart.MeasureManager.SetPartTransientModification();
workPart.Expressions.Delete(expression1);
workPart.MeasureManager.ClearPartTransientModification();
workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();
```

---

### 5. UndoMark の安全な管理

`DeleteUndoMark` 前に `markId != NXOpen.Session.NullUndoMarkId` をチェックし、  
削除後は `NullUndoMarkId` で無効化すること。連続実行時の二重削除によるセッション汚染を防ぐ。

```csharp
void SafeDeleteMark(ref NXOpen.Session.UndoMarkId id)
{
    if (id != NXOpen.Session.NullUndoMarkId)
    {
        theSession.DeleteUndoMark(id, null);
        id = NXOpen.Session.NullUndoMarkId;
    }
}
```

---

### 6. PMI 文字サイズの取得と製図ビューへの適用（ジャーナル確認済み）

取得するプロパティ名（NX 2312 ジャーナルより確定）：

| NX 設定画面 | プロパティ |
|---|---|
| テキスト → 寸法テキスト → 高さ | `LetteringStyle.DimensionTextSize` |
| テキスト → 公差テキスト → 高さ（二段） | `LetteringStyle.TwoLineToleranceTextSize` |
| テキスト → 一般テキスト → 高さ | `LetteringStyle.GeneralTextSize` |

実装方針：
- `PmiManager.PmiAttributes` から寸法系（`PmiDimension` の派生）を `FirstOrDefault()` で **1件だけ**取得
  - → `DimensionTextSize` と `TwoLineToleranceTextSize` を読む
- 寸法系が存在しない場合は `PmiNote` を `FirstOrDefault()` で 1件取得
  - → `GeneralTextSize` を読む
- どちらも存在しない場合はデフォルト値 3.5mm を使用
- 取得後は即 `esb.Destroy()` を呼ぶこと（**全 PMI 走査は処理速度低下のため禁止**）
- 取得した値を `BaseViewBuilder.Style.ViewStyleAnnotation` の対応プロパティに設定する

```csharp
double dimSize = 3.5, tolSize = 2.5, genSize = 3.5;

var dimPmi = workPart.PmiManager.PmiAttributes
                 .OfType<NXOpen.Annotations.PmiDimension>().FirstOrDefault();
if (dimPmi != null)
{
    var esb = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(
                  new[] { (DisplayableObject)dimPmi });
    workPart.SettingsManager.ProcessForMultipleObjectsSettings(new[] { esb });
    dimSize = esb.AnnotationStyle.LetteringStyle.DimensionTextSize;
    tolSize = esb.AnnotationStyle.LetteringStyle.TwoLineToleranceTextSize;
    esb.Destroy();
}
else
{
    var notePmi = workPart.PmiManager.PmiAttributes
                      .OfType<NXOpen.Annotations.PmiNote>().FirstOrDefault();
    if (notePmi != null)
    {
        var esb = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(
                      new[] { (DisplayableObject)notePmi });
        workPart.SettingsManager.ProcessForMultipleObjectsSettings(new[] { esb });
        genSize = esb.AnnotationStyle.LetteringStyle.GeneralTextSize;
        dimSize = genSize;
        tolSize = genSize * 0.63;
        esb.Destroy();
    }
}

// BaseViewBuilder へ適用
bbBuilder.Style.ViewStyleAnnotation.GeneralTextSize          = genSize;
bbBuilder.Style.ViewStyleAnnotation.DimensionTextSize        = dimSize;
bbBuilder.Style.ViewStyleAnnotation.TwoLineToleranceTextSize = tolSize;
```

---

### 7. 用紙サイズの自動選択

1. BaseView を仮シート（10000×10000）に配置後、`DraftingView.GetBoundingBox()` でビュー幅・高さ（mm）を取得
2. 余白 50mm を加算したサイズと標準用紙を比較
3. A4(297×210) → A3(420×297) → A2(594×420) → A1(841×594) → A0(1189×841) の順で収まるか判定
4. A0 でも収まらない場合は `CustomSize` でビューサイズ＋余白 100mm をカスタム設定
5. 決定したサイズでシートを再作成し BaseView を再配置する
6. スケールは常に 1:1（`ScaleNumerator=1.0, ScaleDenominator=1.0`）

アレンジメントはアセンブリ構成によって null の場合がある（ジャーナルで確認）：

```csharp
NXOpen.Assemblies.Arrangement[] arrs;
workPart.GetArrangements(out arrs);
var arr = (arrs != null && arrs.Length > 0) ? arrs[0] : null;
bbBuilder.Style.ViewStyleBase.Arrangement.SelectedArrangement = arr;
bbBuilder.Style.ViewStyleBase.Arrangement.InheritArrangementFromParent = false;
```

---

### 8. DXF 出力先フォルダのルール

- 基底パス：`Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)`
- フォルダ名：`{図面番号}_{YYMMDD}` 例：`P00AE6571_260313`
- 同名フォルダが既存の場合のみ `_(n)` を付加（n は 2 から）
- DXF ファイルとログファイルを同フォルダに出力
- 図面番号は既存実装の取得方法を踏襲

---

### 9. ビュー選択ダイアログのソート順

1. 先頭文字が数字のもの（0-9 昇順）
2. 先頭文字が ASCII アルファベット（A-Z, a-z 昇順・大文字小文字無視）
3. それ以外（日本語等、Unicode コードポイント昇順）
4. 同一グループ内は `StringComparer.OrdinalIgnoreCase` でソート

---

### 10. 情報ウィンドウの非表示

処理前・中・後を通じて NX 情報ウィンドウ（`ListingWindow` / `LogFile`）への出力を一切行わないこと。  
既存コードに `ListingWindow.Open()` や `WriteLine()` が含まれている場合は削除する。

---

### 11. NX テンプレートパスのバージョン非依存化

```csharp
string nxBase = Environment.GetEnvironmentVariable("UGII_BASE_DIR")
             ?? @"C:\Program Files\Siemens\NX2312";
string tplPath = Path.Combine(nxBase, "DRAFTING", "templates",
                     "Drawing-A0-Size2D-template.prt");
```

---

### 12. ビルドと署名

- 実装完了後に `dotnet build` または `msbuild` でビルドを実行すること
- ビルド成功後に `SignDotNet.exe` で署名すること（PostBuildEvent に記述）
- ビルドエラーが発生した場合は必ず修正してから次に進むこと

---

### 13. 実装完了後のチェックリスト

- [ ] DXF `Commit()` の前後に `DrawingsFreeze/Unfreeze` が記述されている
- [ ] DXF 出力後の `ApplicationSwitchImmediate` + `EnterDraftingApplication` が存在しない
- [ ] `DeleteUndoMark` の前に `NullUndoMarkId` チェックが存在する
- [ ] Expression 削除が `DrawingsFreeze/Unfreeze` で囲まれている
- [ ] PMI 文字サイズを `DimensionTextSize` / `TwoLineToleranceTextSize` / `GeneralTextSize` で取得している
- [ ] PMI 取得が `FirstOrDefault()` で 1件のみ（全 PMI 走査をしていない）
- [ ] 用紙サイズが A4/A3/A2/A1/A0/カスタムから自動選択されている
- [ ] アレンジメント取得が null フォールバック対応している
- [ ] 出力先フォルダが `Pictures/{図面番号}_{YYMMDD}[_(n)]` になっている
- [ ] ビュー一覧が 数字→英字→日本語 の順でソートされている
- [ ] `ListingWindow` への出力が一切含まれていない
- [ ] NXOpen 参照 DLL に `<Private>false</Private>` が設定されている
- [ ] `SettingsFile` に dxfdwg.def（一時展開パス）が設定されている
- [ ] ビルドが成功し `SignDotNet.exe` による署名が完了している
