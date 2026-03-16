# NXCustomViewDxfExporter 改修指示書
# PMIフォントサイズ継承 + シートサイズ/ビューサイズ修正 + バグ修正

## 概要

NXCustomViewDxfExporter（C# .NET Framework 4.8 DLL）に以下の改修を行ってください。
まず `cat CustomViewDxfExporter.cs` で現在のソースコード全体を確認してから改修に着手すること。

**改修項目一覧:**

| No | 種別 | 内容 |
|----|------|------|
| 改修1 | 新機能 | PMIフォントサイズ継承機能の追加 |
| 改修2 | 不具合修正 | シートサイズが形状全体を収められない問題の修正（2パス方式） |
| 改修3 | 不具合修正 | 完了ダイアログの成功/失敗判定の誤り |
| 改修4 | 不具合修正 | GFX_change_view_attr エラーによるレスポンス低下 |
| 改修5 | 不具合修正 | UNDO_UG_delete_mark エラーの累積 |
| 改修6 | 不具合修正 | 処理後にシートが残る問題 |
| 改修7 | 不具合修正 | フォルダ名重複時の連番が_(2)から始まる（_(1)から始めるべき） |
| 改修8 | 不具合修正 | DXF上で寸法が分解されている（寸法エンティティとして出力すべき） |
| 改修9 | 不具合修正 | ループが1ビュー目で停止する問題 |

---

## 改修1: PMIフォントサイズ継承機能の追加

### 1-1. 目的

NXのPMI継承（Inherit PMI）機能では、3Dモデル側のPMIを製図ビューに取り込んだ際、**製図側のアノテーションプリファレンスで一律にフォントサイズが上書きされる**（NXの仕様）。モデル側でPMIごとに個別に設定されたフォント高さが失われる。

この機能は、モデル側のPMIフォント高さを事前に読み取り、製図側に継承された後に同じ値を再設定する。

### 1-2. 処理の挿入位置

```
カスタムビューごとのループ:
  1. ★【新規】モデリングアプリケーションでPMIのフォントサイズを読み取り・保存
  2. 製図アプリケーションに切替
  3. 図面シート作成
  4. ベースビュー配置（PMI継承有効）
  5. ★【新規】継承PMIのフォントサイズをモデル側の値に合わせて更新
  6. DXFエクスポート
  7. シート削除・クリーンアップ
```

### 1-3. ステップ1: モデル側PMIのフォントサイズ読み取り

製図アプリケーションに切り替える**前に**（ループの前で1回だけ）、モデリングアプリケーション上でパート内の全PMIオブジェクトのフォントサイズを読み取り、辞書に保存する。

```csharp
// PMIフォントサイズを保存するデータ構造
public class PmiFontInfo
{
    public double DimensionTextSize { get; set; }     // 寸法テキスト高さ
    public double ToleranceTextSize { get; set; }     // 公差テキスト高さ
    public double GeneralTextSize { get; set; }       // 一般テキスト高さ（注記用）
    public double AppendedTextSize { get; set; }      // 付加テキスト高さ
}

// PMIオブジェクトのTag（NXObject.Tag）をキーとして保存
Dictionary<NXOpen.Tag, PmiFontInfo> pmiFontMap = new Dictionary<NXOpen.Tag, PmiFontInfo>();
```

**読み取り方法（ジャーナルから判明したAPI）:**

```csharp
private static PmiFontInfo ReadPmiFontSize(Part workPart, NXOpen.Annotations.Annotation pmiObj)
{
    PmiFontInfo info = new PmiFontInfo();
    NXOpen.DisplayableObject[] objects = new NXOpen.DisplayableObject[] { (NXOpen.DisplayableObject)pmiObj };
    NXOpen.Annotations.EditSettingsBuilder editSettingsBuilder = null;
    try
    {
        editSettingsBuilder = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects);
        var letteringStyle = editSettingsBuilder.AnnotationStyle.LetteringStyle;

        if (pmiObj is NXOpen.Annotations.Dimension)
        {
            info.DimensionTextSize = letteringStyle.DimensionTextSize;
            info.ToleranceTextSize = letteringStyle.ToleranceTextSize;
            info.AppendedTextSize = letteringStyle.AppendedTextSize;
        }
        info.GeneralTextSize = letteringStyle.GeneralTextSize;
    }
    finally
    {
        // ★重要: Commitせずにキャンセル（読み取りのみ、変更しない）
        if (editSettingsBuilder != null)
            editSettingsBuilder.Destroy();
    }
    return info;
}
```

**対象PMIオブジェクトの列挙:**

```csharp
// パート内の全PMI寸法を列挙
foreach (NXOpen.Annotations.Dimension dim in workPart.Dimensions)
{
    if (dim.GetType().Name.StartsWith("Pmi"))
    {
        PmiFontInfo info = ReadPmiFontSize(workPart, dim);
        pmiFontMap[dim.Tag] = info;
    }
}

// パート内の全PMI注記を列挙
foreach (NXOpen.Annotations.Note note in workPart.Notes)
{
    if (note is NXOpen.Annotations.PmiNote)
    {
        PmiFontInfo info = ReadPmiFontSize(workPart, note);
        pmiFontMap[note.Tag] = info;
    }
}

// PMIラベルも同様
foreach (NXOpen.Annotations.Label label in workPart.Labels)
{
    if (label is NXOpen.Annotations.PmiLabel)
    {
        PmiFontInfo info = ReadPmiFontSize(workPart, label);
        pmiFontMap[label.Tag] = info;
    }
}
```

### 1-4. ステップ2: 製図側の継承PMIにフォントサイズを反映

ベースビュー配置後、製図ビュー内の継承PMIオブジェクトに対応するモデル側PMIのフォントサイズを設定する。

```csharp
private static void ApplyFontSizeToDimension(
    Session theSession, Part workPart,
    NXOpen.Annotations.Dimension dim, PmiFontInfo fontInfo)
{
    NXOpen.DisplayableObject[] objects = new NXOpen.DisplayableObject[] { dim };
    NXOpen.Annotations.EditSettingsBuilder editSettingsBuilder = null;
    try
    {
        editSettingsBuilder = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects);

        // ProcessForMultipleObjectsSettingsの呼出し（ジャーナル準拠: Commit前に必須）
        NXOpen.Drafting.BaseEditSettingsBuilder[] builders =
            new NXOpen.Drafting.BaseEditSettingsBuilder[] { editSettingsBuilder };
        workPart.SettingsManager.ProcessForMultipleObjectsSettings(builders);

        var letteringStyle = editSettingsBuilder.AnnotationStyle.LetteringStyle;
        if (fontInfo.DimensionTextSize > 0)
            letteringStyle.DimensionTextSize = fontInfo.DimensionTextSize;
        if (fontInfo.ToleranceTextSize > 0)
            letteringStyle.ToleranceTextSize = fontInfo.ToleranceTextSize;
        if (fontInfo.AppendedTextSize > 0)
            letteringStyle.AppendedTextSize = fontInfo.AppendedTextSize;

        editSettingsBuilder.Commit();
    }
    catch (Exception ex)
    {
        // フォントサイズ設定失敗は致命的ではないのでログのみ
    }
    finally
    {
        if (editSettingsBuilder != null)
            editSettingsBuilder.Destroy();
    }
}

private static void ApplyFontSizeToNote(
    Session theSession, Part workPart,
    NXOpen.Annotations.Note note, PmiFontInfo fontInfo)
{
    NXOpen.DisplayableObject[] objects = new NXOpen.DisplayableObject[] { (NXOpen.DisplayableObject)note };
    NXOpen.Annotations.EditSettingsBuilder editSettingsBuilder = null;
    try
    {
        editSettingsBuilder = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects);

        NXOpen.Drafting.BaseEditSettingsBuilder[] builders =
            new NXOpen.Drafting.BaseEditSettingsBuilder[] { editSettingsBuilder };
        workPart.SettingsManager.ProcessForMultipleObjectsSettings(builders);

        if (fontInfo.GeneralTextSize > 0)
            editSettingsBuilder.AnnotationStyle.LetteringStyle.GeneralTextSize = fontInfo.GeneralTextSize;

        editSettingsBuilder.Commit();
    }
    catch (Exception) { }
    finally
    {
        if (editSettingsBuilder != null)
            editSettingsBuilder.Destroy();
    }
}
```

### 1-5. PMIマッチング戦略

モデル側PMIと製図側継承PMIの対応付け:

1. PMI継承後、製図ビュー内のPMIオブジェクトを列挙
2. 各継承PMIのTagでモデル側のDictionaryを検索
3. Tagマッチングが効かない場合は、PMIのテキスト内容+PMIタイプの組み合わせでマッチング
4. マッチしたらフォントサイズを設定

**重要:**
- `EditSettingsBuilder.Destroy()` は必ず呼ぶこと（リソースリーク防止）
- `Commit()` を呼ばなければ変更は適用されない
- 読み取りだけの場合は `Destroy()` のみ（Commitしない）

### 1-6. ジャーナルで確認されたAPI

```csharp
// 寸法テキスト高さの設定
editSettingsBuilder.AnnotationStyle.LetteringStyle.DimensionTextSize = 9.4315357344889001;

// 公差テキスト高さの設定
editSettingsBuilder.AnnotationStyle.LetteringStyle.ToleranceTextSize = 5.6910735630474001;

// 注記テキスト高さの設定
editSettingsBuilder.AnnotationStyle.LetteringStyle.GeneralTextSize = 9.4315357344889001;

// PMI継承有効化
projectedViewBuilder1.Style.ViewStyleInheritPmi.Pmi = NXOpen.Preferences.PmiOption.InDrawingPlaneFromView;
```

---

## 改修2: シートサイズが形状全体を収められない問題の修正

### 2-1. 現象

現在のアドオンで出力されたDXFでは、形状全体がシートやビュー内に収まっていない。PMIの寸法線や注記を含めた全体がシートからはみ出している。

### 2-2. 設計方針（★重要 — 3つの絶対ルール）

1. **スケールは常に1:1固定。絶対に縮小しない。** A0を超えるモデルでも1:1を維持する
2. **シートサイズをモデルに合わせる。** A判定形に収まればA判を使い、収まらなければカスタムサイズのシートを作成する
3. **処理速度を犠牲にしない。** A4→A3→…と順番に配置して試す方式は禁止。2パス方式（1回目で測定、2回目で本配置）で最大2回の配置に抑える

### 2-3. 処理フロー: 効率的な2パス方式

```
パス1（サイズ測定 — 十分大きな仮シートで1回だけ仮配置）:
  1. カスタムサイズの巨大シート（5000×5000mm）を作成
     ※ どんなモデルでも確実に収まるサイズにする
  2. ベースビューをスケール1:1で仮配置（シート中央）
  3. UFDraw.AskViewBorders() でPMI含む実際の描画範囲を取得
  4. 仮ビュー・仮シートを削除

パス2（本配置 — 測定結果に基づき1回で正しく配置）:
  1. 測定した描画範囲 + マージンから必要シートサイズを算出
  2. A判定形（A4〜A0）に収まればそれを使用、収まらなければカスタムサイズ
  3. 正しいサイズのシートを作成
  4. ベースビューをスケール1:1でシート中央に配置
  5. DXFエクスポート
```

### 2-4. パス1: 仮配置によるサイズ測定

```csharp
// ★ 仮シートは十分大きなカスタムサイズで作成
// どんなモデルでもはみ出さないよう、5000×5000mmとする
const double MeasureSheetSize = 5000.0;

// 仮シート作成
var measureSheetBuilder = workPart.DraftingDrawingSheets
    .CreateDraftingDrawingSheetBuilder(null);
measureSheetBuilder.AutoStartViewCreation = false;
measureSheetBuilder.Option = DrawingSheetBuilder.SheetOption.CustomSize;
measureSheetBuilder.Height = MeasureSheetSize;
measureSheetBuilder.Length = MeasureSheetSize;
measureSheetBuilder.ScaleNumerator = 1.0;
measureSheetBuilder.ScaleDenominator = 1.0;
measureSheetBuilder.Units = DrawingSheetBuilder.SheetUnits.Metric;
measureSheetBuilder.ProjectionAngle = DrawingSheetBuilder.SheetProjectionAngle.Third;

NXObject measureSheetObj = measureSheetBuilder.Commit();
DrawingSheet measureSheet = (DrawingSheet)measureSheetObj;
measureSheetBuilder.Destroy();
workPart.Drafting.SetTemplateInstantiationIsComplete(true);
measureSheet.Open();

// 仮ベースビュー配置（スケール1:1）
workPart.DraftingManager.DrawingsFreezeOutOfDateComputation();
var measureViewBuilder = workPart.DraftingViews.CreateBaseViewBuilder(null);
measureViewBuilder.Placement.Associative = true;
measureViewBuilder.SelectModelView.SelectedView = view;
measureViewBuilder.Style.ViewStyleBase.Part = workPart;
measureViewBuilder.Style.ViewStyleBase.PartName = workPart.FullPath;
measureViewBuilder.SelectModelView.SelectedView = view;
measureViewBuilder.Scale.Numerator = 1.0;
measureViewBuilder.Scale.Denominator = 1.0;

// PMI継承を有効化（PMIのサイズも測定に含めるため）
measureViewBuilder.Style.ViewStyleInheritPmi.Pmi =
    NXOpen.Preferences.PmiOption.InDrawingPlaneFromView;

Point3d measureOrigin = new Point3d(MeasureSheetSize / 2.0, MeasureSheetSize / 2.0, 0.0);
measureViewBuilder.Placement.Placement.SetValue(null, workPart.Views.WorkView, measureOrigin);

NXObject measureViewObj = measureViewBuilder.Commit();
workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();
measureViewBuilder.Destroy();

// ★ ビューの実際の描画範囲を取得（PMIアノテーション含む）
NXOpen.UF.UFSession theUfSession = NXOpen.UF.UFSession.GetUFSession();
double[] viewBorders = new double[4]; // [minX, maxX, minY, maxY]
theUfSession.Draw.AskViewBorders(((NXOpen.Drawings.BaseView)measureViewObj).Tag, viewBorders);

double actualViewWidth = viewBorders[1] - viewBorders[0];   // maxX - minX
double actualViewHeight = viewBorders[3] - viewBorders[2];  // maxY - minY

// 仮ビュー・仮シートを削除
workPart.DraftingManager.DrawingsFreezeOutOfDateComputation();
try
{
    theSession.UpdateManager.AddToDeleteList((NXOpen.Drawings.BaseView)measureViewObj);
    theSession.UpdateManager.DoUpdate(null);
}
catch { }
finally
{
    workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();
}
try { measureSheet.Delete(); } catch { }
```

### 2-5. パス2: 最適シートサイズの決定とカスタムサイズ対応

```csharp
const double SheetMargin = 10.0; // mm（上下左右）

// 必要なシートサイズ = 実測ビュー範囲 + マージン
double requiredWidth = actualViewWidth + (SheetMargin * 2);
double requiredHeight = actualViewHeight + (SheetMargin * 2);

// A判用紙サイズテーブル（横向き: 幅 > 高さ）
// 小さい順に並べる — 最初にフィットしたものを採用
double[][] standardPaperSizes = new double[][] {
    new double[] { 297.0, 210.0 },   // A4横
    new double[] { 420.0, 297.0 },   // A3横
    new double[] { 594.0, 420.0 },   // A2横
    new double[] { 841.0, 594.0 },   // A1横
    new double[] { 1189.0, 841.0 },  // A0横
};

double sheetWidth = 0;
double sheetHeight = 0;
bool useCustomSize = false;

// 標準A判サイズから最小フィットを検索
foreach (var size in standardPaperSizes)
{
    if (requiredWidth <= size[0] && requiredHeight <= size[1])
    {
        sheetWidth = size[0];
        sheetHeight = size[1];
        break;
    }
}

// ★ A0でも収まらない場合: カスタムサイズを使用（スケールは変えない）
if (sheetWidth == 0)
{
    useCustomSize = true;
    // 必要サイズをそのまま用紙サイズにする（切り上げて余裕を持たせる）
    // 50mm単位に切り上げて見栄えの良いサイズにする
    sheetWidth = Math.Ceiling(requiredWidth / 50.0) * 50.0;
    sheetHeight = Math.Ceiling(requiredHeight / 50.0) * 50.0;
    // 最低でもA0以上にする
    sheetWidth = Math.Max(sheetWidth, 1189.0);
    sheetHeight = Math.Max(sheetHeight, 841.0);
}
```

### 2-6. シートの作成（標準サイズ/カスタムサイズ共通）

```csharp
var sheetBuilder = workPart.DraftingDrawingSheets
    .CreateDraftingDrawingSheetBuilder(null);
sheetBuilder.AutoStartViewCreation = false;

// ★ 常にCustomSizeを使用（標準サイズでもCustomSizeで問題ない）
sheetBuilder.Option = DrawingSheetBuilder.SheetOption.CustomSize;
sheetBuilder.Height = sheetHeight;
sheetBuilder.Length = sheetWidth;

// ★ スケールは常に1:1（絶対に変更しない）
sheetBuilder.ScaleNumerator = 1.0;
sheetBuilder.ScaleDenominator = 1.0;

sheetBuilder.Units = DrawingSheetBuilder.SheetUnits.Metric;
sheetBuilder.ProjectionAngle = DrawingSheetBuilder.SheetProjectionAngle.Third;

NXObject sheetObj = sheetBuilder.Commit();
DrawingSheet sheet = (DrawingSheet)sheetObj;
sheetBuilder.Destroy();
```

### 2-7. ベースビューの配置（スケール1:1固定）

```csharp
// ★ スケールは常に1:1
baseViewBuilder.Scale.Numerator = 1.0;
baseViewBuilder.Scale.Denominator = 1.0;

// シート中央に配置
Point3d viewOrigin = new Point3d(sheetWidth / 2.0, sheetHeight / 2.0, 0.0);
baseViewBuilder.Placement.Placement.SetValue(null, workPart.Views.WorkView, viewOrigin);
```

### 2-8. 処理速度に関する注意

この2パス方式は各ビューにつき最大2回のベースビュー配置（仮配置+本配置）を行う。
仮配置時のDXFエクスポートは行わないため、追加の処理時間は主にベースビュー配置と削除のオーバーヘッドのみ。
A4→A3→…と5回試行する方式（最悪5回配置）より大幅に高速。

**仮配置のシートサイズが5000mmと大きいが、NXは内部的にビューの投影計算をモデル座標系で行うため、シートサイズの大小は計算速度にほぼ影響しない。**

### 2-9. 既存のシートサイズ自動選択コードの削除

現在のコードにあるビュー範囲推定やスケール計算のロジックは、上記の2パス方式に完全に置き換える。特に以下を削除/置換:
- 3Dバウンディングボックスからの2D投影サイズ推定コード（不正確なため）
- スケール縮小のコード（1:1固定のため不要）
- A4〜A0の判定で `if-else if` が連なるコード（テーブルルックアップに置換）

---

## 改修3: 完了ダイアログの成功/失敗判定の誤り（バグB）

### 3-1. 現象

完了ダイアログに「成功0件・失敗10件」と表示されるが、実際にはDXFファイルは正常に生成されている。

### 3-2. 原因

エクスポートの成否判定がUndoMarkの存在やCommitの戻り値など、不正確な基準で行われていると推定。

### 3-3. 修正方針

DXFファイルの実在を確認して成否を判定する:

```csharp
// DxfdwgCreator.Commit() の後、ファイル完成待機の後に:
bool exportSuccess = false;

// DXFファイルの出力完了を待機
int waitMs = 0;
const int maxWaitMs = 30000;
while (!File.Exists(outputDxfPath) && waitMs < maxWaitMs)
{
    System.Threading.Thread.Sleep(500);
    waitMs += 500;
}

// ファイルが書き込み中でないか確認
if (File.Exists(outputDxfPath))
{
    for (int retry = 0; retry < 20; retry++)
    {
        try
        {
            using (var fs = new FileStream(outputDxfPath, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                exportSuccess = fs.Length > 0;
                break;
            }
        }
        catch (IOException)
        {
            System.Threading.Thread.Sleep(500);
        }
    }
}

// ★ 成否判定はファイルの実在で行う
if (exportSuccess)
{
    successCount++;
}
else
{
    failCount++;
    failedViews.Add(viewName + ": DXF file not created");
}
```

**既存の成否判定ロジックを全てこのFile.Exists方式に置き換えること。** UndoMarkやCommit戻り値に依存する判定は削除する。

---

## 改修4: GFX_change_view_attr エラーによるレスポンス低下（バグC-1）

### 4-1. 現象

syslogに `GFX_change_view_attr: View N does not exist` エラーが頻発し、DXF Commit中にも発生。1ビューあたり最大9秒の遅延を引き起こしている。

### 4-2. 原因

- 投影ビューDestroy → Expression削除の際のFreeze/Unfreeze順序の誤り
- DXF `Commit()` が `DrawingsFreezeOutOfDateComputation()` で囲まれていない可能性

### 4-3. 修正方針

**DxfdwgCreatorのCommit()をFreeze/Unfreezeで囲む:**

```csharp
// DXFエクスポート前にFreezeして不要なビュー再計算を抑制
workPart.DraftingManager.DrawingsFreezeOutOfDateComputation();

try
{
    NXObject result = dxfdwgCreator.Commit();
    dxfdwgCreator.Destroy();
    dxfdwgCreator = null;
}
finally
{
    workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();
}
```

**投影ビューやExpressionの削除もFreeze内で行う:**

```csharp
// シート削除・クリーンアップ時もFreezeで囲む
workPart.DraftingManager.DrawingsFreezeOutOfDateComputation();
try
{
    // ビュー削除
    // Expression削除
    // シート削除
}
finally
{
    workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();
}
```

---

## 改修5: UNDO_UG_delete_mark エラーの累積（バグC-2）

### 5-1. 現象

`UNDO_UG_delete_mark: Mark XXXX not found` がビューごとに累積している。

### 5-2. 原因

ループ変数の `markId` が前ビューの値を引き継いでいる、またはNullUndoMarkIdチェックが抜けている。

### 5-3. 修正方針

```csharp
// ループ先頭で初期化
NXOpen.Session.UndoMarkId viewMarkId = NXOpen.Session.NullUndoMarkId;

// UndoMark操作前に有効性チェック
if (viewMarkId != NXOpen.Session.NullUndoMarkId)
{
    try
    {
        theSession.DeleteUndoMark(viewMarkId, null);
    }
    catch (Exception) { /* マークが既に無い場合は無視 */ }
    viewMarkId = NXOpen.Session.NullUndoMarkId; // リセット
}
```

**全てのUndoMark操作（SetUndoMark / UndoToMark / DeleteUndoMark）をtry-catchで囲み、失敗しても処理を継続させる。**

コード全体でUndoMark関連の変数を検索し、以下を確認:

```bash
grep -n "UndoMark\|SetUndoMark\|UndoToMark\|DeleteUndoMark" CustomViewDxfExporter.cs
```

- 全てのSetUndoMarkの直前で変数をNullUndoMarkIdに初期化
- 全てのDeleteUndoMark / UndoToMarkをtry-catchで囲む
- ループの各反復の先頭で前回のmarkIdをリセット

---

## 改修6: 処理後にシートが残る問題（バグD）

### 6-1. 現象

アドオン処理完了後、作成した一時製図シートがパートに残ったままになっている。

### 6-2. 修正方針

各ビューのDXF出力完了後に確実にシートを削除する。例外発生時もfinallyでクリーンアップを行う。

```csharp
NXOpen.Drawings.DrawingSheet createdSheet = null;
try
{
    // シート作成
    createdSheet = (NXOpen.Drawings.DrawingSheet)sheetBuilder.Commit();
    // ... ベースビュー配置、PMIフォントサイズ反映、DXFエクスポート ...
}
finally
{
    // ★ シートの確実な削除（正常時・例外時の両方）
    if (createdSheet != null)
    {
        try
        {
            // ビューの削除（シート削除前に必要）
            workPart.DraftingManager.DrawingsFreezeOutOfDateComputation();
            try
            {
                // シート上の全ビューを削除
                NXOpen.Drawings.DraftingView[] viewsOnSheet = createdSheet.GetDraftingViews();
                foreach (var dv in viewsOnSheet)
                {
                    try { theSession.UpdateManager.AddToDeleteList(dv); } catch { }
                }
                theSession.UpdateManager.DoUpdate(null);
            }
            finally
            {
                workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();
            }

            // シートの削除
            createdSheet.Delete();
        }
        catch (Exception)
        {
            // シート削除失敗はログのみ（次のビュー処理には影響させない）
        }
    }
}
```

**全ビュー処理完了後の最終チェック:**

```csharp
// ループ終了後、念のため残存シートを確認・削除
// （処理途中のクラッシュリカバリとして）
try
{
    foreach (NXOpen.Drawings.DrawingSheet sheet in workPart.DrawingSheets)
    {
        // アドオンが作成したシートのみ削除（既存シートは削除しない）
        // 命名規則やタイムスタンプで判別、または作成したシートのリストと照合
    }
}
catch (Exception) { }
```

---

## 改修7: フォルダ名重複時の連番が_(2)から始まる問題

### 7-1. 現象

出力ファイル名が重複した場合、`filename_(2).dxf` から始まる。`filename_(1).dxf` から始まるべき。

### 7-2. 修正方針

ファイル名重複チェックのカウンター初期値を確認し、`1` から始まるように修正する。

```csharp
// 修正前（推定）
int counter = 2; // ← これが原因

// 修正後
int counter = 1;
```

コード内で `GetUniqueFilePath` や重複チェックに関する箇所を検索:

```bash
grep -n "counter\|suffix\|_(.*)\|GetUnique\|Exists.*dxf" CustomViewDxfExporter.cs
```

---

## 改修8: DXF上で寸法が分解されている問題

### 8-1. 現象

出力されたDXFファイルをAutoCADで開くと、寸法が「寸法エンティティ」ではなく、線分・テキスト・矢印などのバラバラの要素に分解されている。寸法値の変更や寸法スタイルの適用ができない。

### 8-2. 原因

DLL内に埋め込まれている `dxfdwg.def`（`EmbeddedDefContent`文字列）の `EXPORT_DIMENSIONS_AS` パラメータが `EXPLODED`（分解）になっている。

### 8-3. 修正方針

`EmbeddedDefContent` 文字列内の設定を修正する:

```
修正前: EXPORT_DIMENSIONS_AS = EXPLODED  （または未設定）
修正後: EXPORT_DIMENSIONS_AS = REAL
```

**「REAL」は寸法をAutoCADネイティブの寸法エンティティとして出力する設定。**

コード内で該当箇所を検索:

```bash
grep -n "EXPORT_DIMENSIONS_AS\|EmbeddedDefContent\|dxfdwg.def" CustomViewDxfExporter.cs
```

`EmbeddedDefContent` の文字列定数の中に `EXPORT_DIMENSIONS_AS` がなければ追加し、ある場合は `REAL` に変更する。

また、以下の関連設定も確認すること:

```
EXPORT_TOLERANCES_AS = REAL
EXPORT_APPENDED_TEXT_AS = REAL
EXPORT_CROSSHATCH_AS = REAL
```

---

## 改修9: ループが1ビュー目で停止する問題

### 9-1. 現象

syslogで `DXF/DWG Export: Export translation job submitted.` が1回しか出現しない。カスタムビューが複数あるにも関わらず、1ビュー目の処理後にループが停止している。

### 9-2. 推定原因

以下のいずれかが原因:

- DXFエクスポート後の例外（`UNDO_UG_delete_mark` エラー等）がキャッチされずにループを抜けている
- シート削除やUndoMarkの処理で未処理例外が発生し、ループ全体のtry-catchに到達している
- 「読取り専用オブジェクトが修正されています」エラーでNXがモーダルダイアログを表示し、処理がブロックされている
- DxfdwgCreatorのバックグラウンドスレッドが完了する前に次のビュー処理に入り、競合が発生

### 9-3. 修正方針

**A. 各ビュー処理を個別のtry-catchで確実に囲む:**

```csharp
for (int i = 0; i < customViews.Count; i++)
{
    try
    {
        // ビュー処理全体
    }
    catch (Exception ex)
    {
        // エラーログに記録してスキップ
        failCount++;
        failedViews.Add(customViews[i].Name + ": " + ex.Message);
        continue; // ★ 次のビューに進む
    }
}
```

**B. DXFエクスポート後のファイル完成待機を確実に実装:**

```csharp
// DxfdwgCreator.Commit() の後
dxfdwgCreator.Destroy();
dxfdwgCreator = null;

// ファイル完成まで待機（次のビュー処理に入る前に必須）
int waitMs = 0;
const int maxWaitMs = 30000;
while (!File.Exists(outputDxfPath) && waitMs < maxWaitMs)
{
    System.Threading.Thread.Sleep(500);
    waitMs += 500;
}
if (File.Exists(outputDxfPath))
{
    for (int retry = 0; retry < 20; retry++)
    {
        try
        {
            using (var fs = new FileStream(outputDxfPath, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                break; // ファイルロック解放済み
            }
        }
        catch (IOException) { System.Threading.Thread.Sleep(500); }
    }
}
```

**C. syslogの「読取り専用オブジェクトが修正されています」への対策:**

このメッセージはNXがモーダルダイアログを表示する可能性がある。Win32 APIの `EnumWindows` でこのダイアログを検出し、自動的に閉じる処理を追加するか、または `theSession.SetUndoMark` の前に一時パートへの書き込みを避ける設計にする。

---

## 統合処理フロー（修正後の完全なフロー）

```
Main()
  |
  +-- 初期処理（セッション取得、パート確認、フォルダ選択等）
  |
  +-- ★ PMIフォントサイズ読み取り（モデリングアプリケーション上で1回だけ実行）
  |     全PMI寸法・注記・ラベルのフォントサイズをDictionaryに保存
  |
  +-- カスタムビュー取得
  |
  +-- カスタムビューごとのループ:
        |
        +-- ★ ループ先頭でUndoMark変数をNullUndoMarkIdに初期化
        |
        +-- 製図アプリケーションに切替
        |
        +-- ★ 2パス方式（スケールは常に1:1、絶対に縮小しない）:
        |     パス1: 5000mm巨大仮シートでベースビュー仮配置 → AskViewBordersで範囲測定 → 削除
        |     パス2: A判定形に収まればA判、収まらなければカスタムサイズシートで本配置
        |
        +-- ★ 継承PMIフォントサイズ反映（Dictionary参照 → EditSettingsBuilder）
        |
        +-- ★ DXFエクスポート（Freeze/Unfreezeで囲む）
        |
        +-- ★ DXFファイル完成待機（File.Exists + ファイルロック解放確認）
        |
        +-- ★ 成否判定: File.Exists(outputDxfPath) で判定
        |
        +-- ★ シート削除（finallyで確実に実行、Freeze/Unfreezeで囲む）
        |
        +-- ★ 例外発生時は failCount++ して continue（次のビューへ）
  |
  +-- ★ 残存シートの最終チェック・クリーンアップ
  |
  +-- 完了ダイアログ表示
```

---

## 変更しないこと

- DxfdwgCreator の設定内容（SettingsFile、InputFile、OutputFile等）— ただし改修8でEmbeddedDefContent内のEXPORT_DIMENSIONS_ASは変更する
- ファイル名の命名規則
- 多言語対応（日英切替）
- 出力フォルダの選択ダイアログ
- フォーカス制御（Win32 API）
- 進捗表示ダイアログ（ProgressForm）
- エラーハンドリングの基本方針（ビュー単位でtry-catch、失敗時はスキップ続行）

---

## ビルド確認

```bash
msbuild /t:Rebuild /p:Configuration=Debug /p:Platform=x64 NXCustomViewDxfExporter.sln
```

---

## 検証ポイント

```bash
# PMIフォントサイズ関連
grep -n "DimensionTextSize\|GeneralTextSize\|ToleranceTextSize\|PmiFontInfo\|ReadPmiFontSize" CustomViewDxfExporter.cs

# シートサイズ自動選択（A4〜A0 + カスタムサイズ対応）
grep -n "297\|420\|594\|841\|1189\|5000\|paperSizes\|requiredWidth\|useCustomSize\|MeasureSheetSize" CustomViewDxfExporter.cs

# スケール1:1固定（Numerator=1, Denominator=1 のみ存在し、縮小コードがないこと）
grep -n "Scale.Numerator\|Scale.Denominator" CustomViewDxfExporter.cs
# → Scale.Numerator に 1.0 以外の値が設定されている箇所がないことを確認

# AskViewBorders による実測（仮配置でPMI含む範囲を取得していること）
grep -n "AskViewBorders\|viewBorders\|actualViewWidth\|actualViewHeight" CustomViewDxfExporter.cs

# 成否判定がFile.Existsベースであること
grep -n "File.Exists.*dxf\|File.Exists.*output\|exportSuccess" CustomViewDxfExporter.cs

# Freeze/Unfreezeの対称性
grep -c "DrawingsFreezeOutOfDateComputation" CustomViewDxfExporter.cs
grep -c "DrawingsUnfreezeOutOfDateComputation" CustomViewDxfExporter.cs
# → 上記2つのカウントが一致すること

# UndoMark初期化
grep -n "NullUndoMarkId" CustomViewDxfExporter.cs

# シート削除
grep -n "\.Delete()\|createdSheet\|DeleteSheets\|sheet.*Delete" CustomViewDxfExporter.cs

# 改修7: フォルダ連番（カウンタ初期値が1であること）
grep -n "counter\|suffix\|GetUnique" CustomViewDxfExporter.cs

# 改修8: 寸法エクスポート設定（REALであること）
grep -n "EXPORT_DIMENSIONS_AS" CustomViewDxfExporter.cs
# → REAL が設定されていること。EXPLODEDがないこと

# 改修9: ループ内のtry-catch（continueがあること）
grep -n "continue\|catch.*Exception" CustomViewDxfExporter.cs

# DXFファイル完成待機
grep -n "File.Exists\|FileStream\|waitMs\|maxWaitMs" CustomViewDxfExporter.cs
```
