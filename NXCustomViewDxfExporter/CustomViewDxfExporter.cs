using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NXOpen;
using NXOpen.Drawings;
using NXOpen.UF;

namespace NXCustomViewDxfExporter
{
    public class CustomViewDxfExporter
    {
        // Win32 API: フォーカス奪取を制御
        [DllImport("user32.dll")]
        private static extern bool LockSetForegroundWindow(uint uLockCode);
        private const uint LSFW_LOCK = 1;
        private const uint LSFW_UNLOCK = 2;

        // Win32 API: フォーカスの保存・復元
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private static Session theSession;
        private static UFSession theUFSession;
        private static UI theUI;
        private static Part workPart;
        private static ListingWindow lw;

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
            lw = theSession.ListingWindow;
            lw.Open();

            Stopwatch totalSw = Stopwatch.StartNew();

            // CGM環境変数の元の値を退避
            string originalCgmEnv = Environment.GetEnvironmentVariable("UGII_CGM_FITS_FILE_SAVE");

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

                // CGMダイアログ抑制（処理開始前に設定）
                SuppressCgmDialog();

                // パート名（拡張子なし）を取得
                string partName = Path.GetFileNameWithoutExtension(workPart.FullPath);
                lw.WriteLine("========================================");
                lw.WriteLine("  カスタムビュー DXF エクスポーター");
                lw.WriteLine("========================================");
                lw.WriteLine("パート名: " + partName);

                // 2. 出力フォルダの選択
                string outputFolder = SelectOutputFolder();
                if (outputFolder == null)
                {
                    lw.WriteLine("フォルダ選択がキャンセルされました。");
                    return;
                }
                lw.WriteLine("出力先:   " + outputFolder);

                // 3. カスタムビューの取得
                List<ModelingView> customViews = GetCustomViews();
                if (customViews.Count == 0)
                {
                    lw.WriteLine("カスタムビューが見つかりません。");
                    theUI.NXMessageBox.Show("情報", NXMessageBox.DialogType.Information,
                        "カスタムビューが見つかりません。");
                    return;
                }

                lw.WriteLine("----------------------------------------");
                lw.WriteLine("変換対象ビュー: " + customViews.Count + " 件");
                for (int i = 0; i < customViews.Count; i++)
                {
                    lw.WriteLine(string.Format("  {0}. {1}", i + 1, customViews[i].Name));
                }
                lw.WriteLine("----------------------------------------");

                // 4. 製図アプリケーションに切り替え（ループ外で1回だけ実行）
                //    ループ内で毎回切り替えるとNXウィンドウがアクティブになるため
                theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING");

                // 5. バックグラウンド処理の開始
                //    ユーザーの作業中ウィンドウを保存し、フォーカスロックで
                //    NXが前面に来るのを防止する。
                IntPtr userWindow = GetForegroundWindow();

                try
                {
                    LockSetForegroundWindow(LSFW_LOCK);
                }
                catch { }

                try
                {
                    // 6. カスタムビューごとのループ処理
                    for (int i = 0; i < customViews.Count; i++)
                    {
                        ModelingView view = customViews[i];
                        lw.WriteLine(string.Format("\n[{0}/{1}] 処理中: {2}", i + 1, customViews.Count, view.Name));

                        Stopwatch viewSw = Stopwatch.StartNew();
                        try
                        {
                            ExportViewAsDxf(view, partName, outputFolder);
                            viewSw.Stop();
                            successCount++;
                            lw.WriteLine(string.Format("  → 成功 ({0:F1}秒)", viewSw.Elapsed.TotalSeconds));
                        }
                        catch (Exception ex)
                        {
                            viewSw.Stop();
                            failCount++;
                            lw.WriteLine(string.Format("  → 失敗 ({0:F1}秒): {1}", viewSw.Elapsed.TotalSeconds, ex.Message));
                            theSession.LogFile.WriteLine(
                                string.Format("ビュー '{0}' のエクスポートに失敗: {1}", view.Name, ex.Message));
                        }

                        // 各ビュー処理後にユーザーのウィンドウを前面に復帰
                        try
                        {
                            if (userWindow != IntPtr.Zero)
                                SetForegroundWindow(userWindow);
                        }
                        catch { }
                    }
                }
                finally
                {
                    // フォーカスロック解除（エラー時も確実に解除）
                    try
                    {
                        LockSetForegroundWindow(LSFW_UNLOCK);
                    }
                    catch { }

                    // ユーザーのウィンドウを前面に復帰
                    try
                    {
                        if (userWindow != IntPtr.Zero)
                            SetForegroundWindow(userWindow);
                    }
                    catch { }
                }

                // 7. モデリングアプリケーションに戻る
                try
                {
                    theSession.ApplicationSwitchImmediate("UG_APP_MODELING");
                }
                catch
                {
                    // 既にモデリングの場合は無視
                }

                // 結果表示
                totalSw.Stop();
                lw.WriteLine("\n========================================");
                lw.WriteLine("  処理結果");
                lw.WriteLine("========================================");
                lw.WriteLine(string.Format("成功: {0} 件", successCount));
                if (failCount > 0)
                    lw.WriteLine(string.Format("失敗: {0} 件", failCount));
                lw.WriteLine(string.Format("合計処理時間: {0:F1}秒", totalSw.Elapsed.TotalSeconds));
                lw.WriteLine("========================================");

                string message = string.Format("DXFエクスポート完了: {0}ファイル出力しました", successCount);
                if (failCount > 0)
                {
                    message += string.Format("\n失敗: {0}件", failCount);
                }
                theUI.NXMessageBox.Show("完了", NXMessageBox.DialogType.Information, message);
            }
            catch (Exception ex)
            {
                lw.WriteLine("致命的エラー: " + ex.Message);
                theUI.NXMessageBox.Show("エラー", NXMessageBox.DialogType.Error,
                    string.Format("予期しないエラーが発生しました:\n{0}", ex.Message));
            }
            finally
            {
                // UndoMarkまで戻す（図面シート等の変更を取り消す）
                theSession.UndoToMark(undoMark, null);

                // CGM環境変数を元に戻す
                try
                {
                    Environment.SetEnvironmentVariable("UGII_CGM_FITS_FILE_SAVE", originalCgmEnv);
                }
                catch { }

                // パートは絶対に保存しない（保存するとCGMダイアログの原因になる）
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

                // DXF/DWGリビジョン: 2018-2024（寸法・線種の正確な出力に必須）
                dxfCreator.AutoCADRevision = DxfdwgCreator.AutoCADRevisionOptions.R2018;

                // 寸法・注記を含める（寸法線の欠落防止）
                dxfCreator.ObjectTypes.Annotations = true;
                dxfCreator.ObjectTypes.Curves = true;
                dxfCreator.ObjectTypes.Solids = true;

                // 全レイヤーをエクスポート対象にする（寸法レイヤーの除外を防止）
                dxfCreator.LayerMask = "1-256";

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

        /// <summary>
        /// ビューとシートをUndoで巻き戻してクリーンアップ。
        /// 明示的なシート削除（Draw.DeleteDrawing）はCGMダイアログの原因になるため、
        /// UndoToMarkで巻き戻すことでダイアログの発生条件自体を回避する。
        /// </summary>
        private static void CleanupViewAndSheet(NXObject baseViewObj, DrawingSheet sheet, Session.UndoMarkId undoMark)
        {
            try
            {
                // UndoToMarkでシート作成・ビュー配置を巻き戻す
                // → 明示的な削除が不要になり、CGMダイアログを回避
                theSession.UndoToMark(undoMark, null);
            }
            catch
            {
                // Undo失敗時はフォールバックとして明示的削除を試行
                if (baseViewObj != null)
                {
                    try
                    {
                        theSession.UpdateManager.AddObjectsToDeleteList(new NXObject[] { baseViewObj });
                        theSession.UpdateManager.DoUpdate(undoMark);
                    }
                    catch { }
                }

                if (sheet != null)
                {
                    try
                    {
                        theUFSession.Draw.DeleteDrawing(sheet.Tag);
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// CGM保存ダイアログを抑制するための多層的な対策。
        /// ダイアログ: 「表示されていないシートを含むパートは、それらと一緒にCGMを
        /// 保存せず、既存のCGMは削除されます」
        ///
        /// 対策1: 環境変数 UGII_CGM_FITS_FILE_SAVE=0
        /// 対策2: SaveOptions.DrawingCgmData=false
        /// 対策3: 処理フローの工夫 — シートの明示的削除をせずUndoToMarkで巻き戻す
        ///        （CleanupViewAndSheet内で実装）
        /// 対策4: パートを保存しない（UndoMarkで元に戻すため保存不要）
        /// </summary>
        private static void SuppressCgmDialog()
        {
            // 対策1: 環境変数でCGM保存を無効化
            Environment.SetEnvironmentVariable("UGII_CGM_FITS_FILE_SAVE", "0");

            // 対策2: パートのSaveOptionsで図面CGMデータ保存を無効化
            try
            {
                workPart.SaveOptions.DrawingCgmData = false;
            }
            catch
            {
                // NXバージョンによってはこのAPIが利用できない場合がある
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
