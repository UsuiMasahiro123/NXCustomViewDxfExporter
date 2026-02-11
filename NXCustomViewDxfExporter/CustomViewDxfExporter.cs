using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using NXOpen;
using NXOpen.Drawings;
using NXOpen.UF;

namespace NXCustomViewDxfExporter
{
    public class CustomViewDxfExporter
    {
        private static Session theSession;
        private static UFSession theUFSession;
        private static UI theUI;
        private static Part workPart;

        // 標準ビュー名リスト（英語・日本語）
        private static readonly HashSet<string> StandardViewNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // 英語
            "Top", "Front", "Right", "Back", "Bottom", "Left",
            "Isometric", "Trimetric",
            // 日本語
            "上", "正面", "右", "背面", "下", "左",
            "等角投影", "不等角投影"
        };

        // シートサイズ定義 (幅mm, 高さmm)
        private static readonly double[][] SheetSizes =
        {
            new double[] { 210.0, 297.0 },   // A4
            new double[] { 420.0, 297.0 },   // A3
            new double[] { 594.0, 420.0 },   // A2
            new double[] { 841.0, 594.0 },   // A1
            new double[] { 1189.0, 841.0 }   // A0
        };

        private const double SheetMargin = 10.0; // 各辺のマージン(mm)

        public static void Main(string[] args)
        {
            theSession = Session.GetSession();
            theUFSession = UFSession.GetUFSession();
            theUI = UI.GetUI();

            Session.UndoMarkId undoMark = theSession.SetUndoMark(Session.MarkVisibility.Visible, "DXF Export");

            int successCount = 0;
            int failCount = 0;

            try
            {
                // 1. 作業パートの確認
                workPart = theSession.Parts.Work;
                if (workPart == null)
                {
                    theUI.NXMessageBox.Show("エラー", NXMessageBox.DialogType.Error,
                        "パートファイルが開かれていません。");
                    return;
                }

                // パート名（拡張子なし）を取得
                string partName = Path.GetFileNameWithoutExtension(workPart.FullPath);

                // 2. 出力フォルダの選択
                string outputFolder = SelectOutputFolder();
                if (outputFolder == null)
                {
                    return; // キャンセル
                }

                // 3. カスタムビューの取得
                List<ModelingView> customViews = GetCustomViews();
                if (customViews.Count == 0)
                {
                    theUI.NXMessageBox.Show("情報", NXMessageBox.DialogType.Information,
                        "カスタムビューが見つかりません。");
                    return;
                }

                // 4. カスタムビューごとのループ処理
                foreach (ModelingView view in customViews)
                {
                    try
                    {
                        ExportViewAsDxf(view, partName, outputFolder);
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        theSession.LogFile.WriteLine(
                            string.Format("ビュー '{0}' のエクスポートに失敗: {1}", view.Name, ex.Message));
                    }
                }

                // 5. モデリングアプリケーションに戻る
                try
                {
                    theSession.ApplicationSwitchImmediate("UG_APP_MODELING");
                }
                catch
                {
                    // 既にモデリングの場合は無視
                }

                // 結果表示
                string message = string.Format("DXFエクスポート完了: {0}ファイル出力しました", successCount);
                if (failCount > 0)
                {
                    message += string.Format("\n失敗: {0}件", failCount);
                }
                theUI.NXMessageBox.Show("完了", NXMessageBox.DialogType.Information, message);
            }
            catch (Exception ex)
            {
                theUI.NXMessageBox.Show("エラー", NXMessageBox.DialogType.Error,
                    string.Format("予期しないエラーが発生しました:\n{0}", ex.Message));
            }
            finally
            {
                // UndoMarkまで戻す（図面シート等の変更を取り消す）
                theSession.UndoToMark(undoMark, null);
            }
        }

        private static string SelectOutputFolder()
        {
            string selectedPath = null;
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "DXFファイルの出力先フォルダを選択してください";
                dialog.ShowNewFolderButton = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    selectedPath = dialog.SelectedPath;
                }
            }
            return selectedPath;
        }

        private static List<ModelingView> GetCustomViews()
        {
            List<ModelingView> customViews = new List<ModelingView>();
            ModelingViewCollection views = workPart.ModelingViews;

            foreach (ModelingView view in views)
            {
                if (!StandardViewNames.Contains(view.Name))
                {
                    customViews.Add(view);
                }
            }

            return customViews;
        }

        private static void ExportViewAsDxf(ModelingView view, string partName, string outputFolder)
        {
            DrawingSheet sheet = null;
            NXObject baseViewObj = null;
            Session.UndoMarkId viewUndoMark = theSession.SetUndoMark(Session.MarkVisibility.Invisible, "ExportView");

            try
            {
                // a. 製図アプリケーションに切り替え
                theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING");

                // ビューのバウンディングボックスからシートサイズを決定
                double viewWidth;
                double viewHeight;
                GetViewBounds(out viewWidth, out viewHeight);

                double sheetWidth;
                double sheetHeight;
                SelectSheetSize(viewWidth, viewHeight, out sheetWidth, out sheetHeight);

                // b. 図面シートの作成 (DrawingSheetBuilder使用)
                DrawingSheetBuilder sheetBuilder = workPart.DrawingSheets.DrawingSheetBuilder(null);
                sheetBuilder.Option = DrawingSheetBuilder.SheetOption.CustomSize;
                sheetBuilder.AutoStartViewCreation = false;
                sheetBuilder.Height = sheetHeight;
                sheetBuilder.Length = sheetWidth;
                sheetBuilder.ScaleNumerator = 1.0;
                sheetBuilder.ScaleDenominator = 1.0;
                sheetBuilder.Units = DrawingSheetBuilder.SheetUnits.Metric;
                sheetBuilder.ProjectionAngle = DrawingSheetBuilder.SheetProjectionAngle.Third;

                NXObject sheetObj = sheetBuilder.Commit();
                sheet = (DrawingSheet)sheetObj;
                sheetBuilder.Destroy();

                // シートを開く
                sheet.Open();

                // c. ベースビューの配置
                Point3d viewOrigin = new Point3d(sheetWidth / 2.0, sheetHeight / 2.0, 0.0);

                BaseViewBuilder baseViewBuilder = workPart.DraftingViews.CreateBaseViewBuilder(null);
                baseViewBuilder.SelectModelView.SelectedView = view;
                baseViewBuilder.Placement.Placement.SetValue(null, workPart.Views.WorkView, viewOrigin);
                baseViewBuilder.Scale.Denominator = 1.0;
                baseViewBuilder.Scale.Numerator = 1.0;

                // ビューがシートに収まるようスケール調整
                double availableWidth = sheetWidth - (SheetMargin * 2);
                double availableHeight = sheetHeight - (SheetMargin * 2);
                if (viewWidth > 0 && viewHeight > 0)
                {
                    double scaleX = availableWidth / viewWidth;
                    double scaleY = availableHeight / viewHeight;
                    double scale = Math.Min(scaleX, scaleY);
                    if (scale < 1.0)
                    {
                        baseViewBuilder.Scale.Denominator = 1.0;
                        baseViewBuilder.Scale.Numerator = scale;
                    }
                }

                baseViewObj = baseViewBuilder.Commit();
                baseViewBuilder.Destroy();

                // d. DXFエクスポート
                string dxfFileName = string.Format("{0}_{1}.dxf", partName, view.Name);
                string dxfFilePath = Path.Combine(outputFolder, dxfFileName);

                DxfdwgCreator dxfCreator = theSession.DexManager.CreateDxfdwgCreator();
                dxfCreator.ExportFrom = DxfdwgCreator.ExportFromOption.DisplayPart;
                dxfCreator.OutputFileType = DxfdwgCreator.OutputFileTypeOption.Dxf;
                dxfCreator.ExportAs = DxfdwgCreator.ExportAsOption.TwoD;
                dxfCreator.ProcessHoldFlag = true;
                dxfCreator.FileSaveFlag = false;
                dxfCreator.OutputTo = DxfdwgCreator.OutputToOption.Drafting;
                dxfCreator.ExportData = DxfdwgCreator.ExportDataOption.Drawing;
                dxfCreator.OutputFile = dxfFilePath;

                // 図面シートを選択
                dxfCreator.ExportSelectionBlock.SelectionScope = ObjectSelector.Scope.SelectedObjects;
                dxfCreator.ExportSelectionBlock.SelectionComp.Add(sheet);

                dxfCreator.Commit();
                dxfCreator.Destroy();

                // e. クリーンアップ：ビューとシートを削除
                CleanupViewAndSheet(baseViewObj, sheet, viewUndoMark);
                sheet = null;
            }
            catch
            {
                // クリーンアップ試行
                CleanupViewAndSheet(baseViewObj, sheet, viewUndoMark);
                sheet = null;
                throw;
            }
        }

        private static void CleanupViewAndSheet(NXObject baseViewObj, DrawingSheet sheet, Session.UndoMarkId undoMark)
        {
            // ビューの削除
            if (baseViewObj != null)
            {
                try
                {
                    theSession.UpdateManager.AddObjectsToDeleteList(new NXObject[] { baseViewObj });
                    theSession.UpdateManager.DoUpdate(undoMark);
                }
                catch
                {
                    // ビュー削除失敗は無視
                }
            }

            // シートの削除
            if (sheet != null)
            {
                try
                {
                    theUFSession.Draw.DeleteDrawing(sheet.Tag);
                }
                catch
                {
                    // シート削除失敗は無視
                }
            }
        }

        private static void GetViewBounds(out double width, out double height)
        {
            // パートのボディからバウンディングボックスを推定
            try
            {
                double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
                double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
                bool found = false;

                foreach (Body body in workPart.Bodies)
                {
                    double[] minCorner = new double[3];
                    double[,] directions = new double[3, 3];
                    double[] distances = new double[3];
                    theUFSession.Modl.AskBoundingBoxExact(body.Tag,
                        Tag.Null, minCorner, directions, distances);

                    // min_corner + directions * distances でバウンディングボックスの頂点を計算
                    double bodyMaxX = minCorner[0] + distances[0];
                    double bodyMaxY = minCorner[1] + distances[1];
                    double bodyMaxZ = minCorner[2] + distances[2];

                    if (minCorner[0] < minX) minX = minCorner[0];
                    if (minCorner[1] < minY) minY = minCorner[1];
                    if (minCorner[2] < minZ) minZ = minCorner[2];
                    if (bodyMaxX > maxX) maxX = bodyMaxX;
                    if (bodyMaxY > maxY) maxY = bodyMaxY;
                    if (bodyMaxZ > maxZ) maxZ = bodyMaxZ;
                    found = true;
                }

                if (found)
                {
                    double dx = maxX - minX;
                    double dy = maxY - minY;
                    double dz = maxZ - minZ;

                    // 最大の2軸を幅・高さとする
                    double[] dims = new double[] { dx, dy, dz };
                    Array.Sort(dims);
                    width = dims[2];   // 最大
                    height = dims[1];  // 2番目

                    if (width <= 0) width = 200.0;
                    if (height <= 0) height = 200.0;
                }
                else
                {
                    width = 200.0;
                    height = 200.0;
                }
            }
            catch
            {
                // デフォルトサイズ
                width = 200.0;
                height = 200.0;
            }
        }

        private static void SelectSheetSize(double viewWidth, double viewHeight, out double sheetWidth, out double sheetHeight)
        {
            double requiredWidth = viewWidth + (SheetMargin * 2);
            double requiredHeight = viewHeight + (SheetMargin * 2);

            foreach (double[] size in SheetSizes)
            {
                double w = size[0];
                double h = size[1];

                // 横向き・縦向きの両方を試す
                if ((w >= requiredWidth && h >= requiredHeight) ||
                    (h >= requiredWidth && w >= requiredHeight))
                {
                    if (w >= requiredWidth && h >= requiredHeight)
                    {
                        sheetWidth = w;
                        sheetHeight = h;
                    }
                    else
                    {
                        sheetWidth = h;
                        sheetHeight = w;
                    }
                    return;
                }
            }

            // A0でも収まらない場合はA0を使用（スケール縮小で対応）
            sheetWidth = SheetSizes[SheetSizes.Length - 1][0];
            sheetHeight = SheetSizes[SheetSizes.Length - 1][1];
        }

        public static int GetUnloadOption(string dummy)
        {
            return (int)Session.LibraryUnloadOption.Immediately;
        }
    }
}
