using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
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
        private static Part workPart;

        private static string tempDefPath;

        // デバッグモードフラグ（trueのとき製図シートを残してDXFをスキップ）
        private static bool DebugDraftingCheck = false;

        private static bool isJapanese;
        private static Dictionary<string, string> messagesJa;
        private static Dictionary<string, string> messagesEn;

        // Win32 API
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;

        // 標準ビュー名（フィルタ対象）
        private static readonly HashSet<string> StandardViewNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Top", "Front", "Right", "Back", "Bottom", "Left", "Isometric", "Trimetric",
            "上", "正面", "右", "背面", "下", "左", "等角投影", "不等角投影"
        };

        // 標準用紙サイズ (幅mm, 高さmm)
        private static readonly double[][] StandardPapers =
        {
            new double[] { 297, 210 },   // A4
            new double[] { 420, 297 },   // A3
            new double[] { 594, 420 },   // A2
            new double[] { 841, 594 },   // A1
            new double[] { 1189, 841 },  // A0
        };
        private static readonly string[] PaperNames = { "A4", "A3", "A2", "A1", "A0" };

        // PMIフォントサイズ情報
        public class PmiFontInfo
        {
            public double DimensionTextSize { get; set; }
            public double ToleranceTextSize { get; set; }
            public double GeneralTextSize { get; set; }
            public double AppendedTextSize { get; set; }
            public bool IsDimension { get; set; }
            public string PmiTypeName { get; set; }
            public string TextContent { get; set; }
        }

        // dxfdwg.def 埋め込み内容
        private static readonly string EmbeddedDefContent =
            "ACAD_LAYOUTS_TO_IMPORT =\r\n" +
            "ACAD_VERSION = 2018\r\n" +
            "ALTERNATE_SYMBOL_FONT_FOR_EXPORT =Arial Unicode MS\r\n" +
            "ASPECT_RATIO_CALCULATION_ON_IMPORT =AUTOMATIC_CALCULATION\r\n" +
            "ASSEMBLY_MAP = ON\r\n" +
            "AVOID_NX_TEMPLATE_PART_LAYERS =YES\r\n" +
            "BASE_PART_IN =dwgnull_in.prt\r\n" +
            "BASE_PART_METER =dwgnull_meter.prt\r\n" +
            "BASE_PART_MICRON =dwgnull_micron.prt\r\n" +
            "BASE_PART_MM =dwgnull_mm.prt\r\n" +
            "BSPLINE_TO_PLINE_CONV_TOL =0.08\r\n" +
            "CHOOSE_DIRECTION = UG TO DXF\r\n" +
            "DATA_REDUCTION =no\r\n" +
            "DSP_MASK =VOLUMINOUS\r\n" +
            "EXPORT_CURVE_ATTRIBUTES =NO\r\n" +
            "EXPORT_APPENDED_TEXT_AS =REAL\r\n" +
            "EXPORT_CROSSHATCH_AS =REAL\r\n" +
            "EXPORT_DIMENSIONS_AS =REAL\r\n" +
            "EXPORT_DRAWING_USING_CGM =NO\r\n" +
            "EXPORT_TOLERANCES_AS =REAL\r\n" +
            "EXPORT_SCALE =1.0\r\n" +
            "FILL_MODE =OFF\r\n" +
            "HEAL_GEOMETRY_ON_IMPORT =YES\r\n" +
            "IMPORT_ACAD_BLOCK_AS =GROUP\r\n" +
            "IMPORT_ACAD_CURVES_ON_SKETCH =NO\r\n" +
            "IMPORT_ACAD_LAYOUTS =YES\r\n" +
            "IMPORT_ACAD_LAYOUTS_TO =IMPORTED_VIEW\r\n" +
            "IMPORT_ACAD_MODEL_DATA =YES\r\n" +
            "IMPORT_ACAD_MODEL_DATA_TO =MODELING\r\n" +
            "IMPORT_ALL_ACAD_LAYOUTS =YES\r\n" +
            "IMPORT_OBJECTS_FROM_FROZEN_LAYER =NO\r\n" +
            "IMPORT_OBJECTS_FROM_INVISIBLE_LAYER =NO\r\n" +
            "IMPORT_UNSELECTED_ACAD_LAYERS =NO\r\n" +
            "IMPORT_UNSELECTED_ACAD_LAYERS_TO =256\r\n" +
            "LOG_FILE =\r\n" +
            "MAX_SPLINE_DEGREE =3\r\n" +
            "MINIMUM_OBJECT_SIZE_TO_EXCLUDE_ON_IMPORT =0.0\r\n" +
            "MSG_MASK =VOLUMINOUS\r\n" +
            "NON_NUMERIC_LAYER_SORTING_CRITERIA =ALPHABETICAL\r\n" +
            "OPTIMIZE_GEOMETRY =NO\r\n" +
            "SET_NX_LAYER_NUMBER_FROM_PREFIX =NO\r\n" +
            "SIMPLIFY_GEOMETRY =NO\r\n" +
            "SKIP_UNREFERENCED_ACAD_LAYERS =YES\r\n" +
            "SUPPORT_MTEXT_FORMATTING_ON_IMPORT =NO\r\n" +
            "SURFU =8\r\n" +
            "SURFV =8\r\n" +
            "UGI_ANNOT_MASK =Dimensions,Notes,Labels,ID Symbols,Tolerances,Centerlines,Crosshatching,Draft Aid by Parts,Stand Alone Symbols,Symbol Fonts\r\n" +
            "UGI_COMP_FAIL = Continue if Load Fails\r\n" +
            "UGI_COMP_SUB = Do not Allow Substitution\r\n" +
            "UGI_CURVE_MASK =Points,Lines,Arcs,Conics,B-Curves,Silhouette Curves,Solid Edges on Drawings\r\n" +
            "UGI_DIM_IMPORT_FLAG =GROUP\r\n" +
            "UGI_DRAWING_NAMES =\r\n" +
            "UGI_GDT_EXPORT_AS_BLOCK =YES\r\n" +
            "UGI_LAYER_MASK =1-256\r\n" +
            "UGI_LOAD_COMP = Load Components\r\n" +
            "UGI_LOAD_OPTION = Load From Assem Dir\r\n" +
            "UGI_LOAD_VER = Load Exact Version\r\n" +
            "UGI_PROC_ASSEM = Overwrite load_options.def values\r\n" +
            "UGI_SEARCH_DIRS ={PART_DIR}\r\n" +
            "UGI_SOLID_EXPORT = FACET\r\n" +
            "UGI_SOLID_MASK = \r\n" +
            "UGI_SPLINE_EXPORT = SPLINE\r\n" +
            "UGI_STRUCT_MASK =Groups,Views,Drawings,Components,Reference Sets\r\n" +
            "UGI_SURF_MASK = \r\n" +
            "UGI_USER_DEFINED_VIEWS =\r\n" +
            "UNITS =Metric\r\n" +
            "UNSELECTED_ACAD_LAYER_LIST =\r\n" +
            "VIEW_MODERASE_MODE = YES\r\n" +
            "WIDTHFACTOR_CALCULATION_ON_EXPORT = AUTOMATIC_CALCULATION\r\n";

        // ================================================================
        // エントリーポイント
        // ================================================================

        public static void Main(string[] args)
        {
            theSession = Session.GetSession();
            theUFSession = UFSession.GetUFSession();
            workPart = theSession.Parts.Work;

            InitializeLanguage();

            if (workPart == null)
            {
                ShowError(GetMsg("ErrorNoPartOpen"));
                return;
            }

            try
            {
                RunExport();
            }
            catch (Exception ex)
            {
                ShowError(string.Format(GetMsg("UnexpectedError"), ex.Message));
            }
            finally
            {
                CleanupTempFiles();
            }
        }

        public static int GetUnloadOption(string dummy)
        {
            return (int)Session.LibraryUnloadOption.Immediately;
        }

        // ================================================================
        // メインエクスポート処理
        // ================================================================

        private static void RunExport()
        {
            // 1. パート属性
            string partNo, partRev;
            GetPartAttributes(out partNo, out partRev);

            // 2. ★【改修1】PMIフォントサイズ読み取り（モデリングアプリ上で1回だけ）
            Dictionary<Tag, PmiFontInfo> pmiFontMap = ReadAllPmiFontSizes();

            // 3. カスタムビュー取得・ソート
            List<ModelingView> customViews = GetCustomViews();
            if (customViews.Count == 0)
            {
                ShowError(GetMsg("ErrorNoCustomView"));
                return;
            }
            customViews.Sort((a, b) => CompareViewNames(a.Name, b.Name));

            // 4. ビュー選択ダイアログ
            List<ModelingView> selectedViews;
            using (var form = new ViewSelectionForm(customViews))
            {
                form.TopMost = true;
                if (form.ShowDialog() != DialogResult.OK || form.SelectedIndices.Count == 0)
                    return;

                selectedViews = new List<ModelingView>();
                foreach (int idx in form.SelectedIndices)
                    selectedViews.Add(customViews[idx]);
            }

            // 5. 出力フォルダ作成
            string outputFolder = CreateOutputFolder(partNo);
            Directory.CreateDirectory(outputFolder);

            // 6. DEFファイル展開
            tempDefPath = ExtractDefFile();

            // 7. 製図アプリに切替
            theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING");
            workPart.Drafting.EnterDraftingApplication();

            // ★【改修6】処理前の既存シートTagを記録（残存シートクリーンアップ用）
            HashSet<Tag> originalSheetTags = new HashSet<Tag>();
            try
            {
                foreach (DrawingSheet s in workPart.DrawingSheets)
                    originalSheetTags.Add(s.Tag);
            }
            catch { }

            // 8. 進捗ダイアログ（STAスレッド）
            ProgressForm progressForm = null;
            ManualResetEvent formReady = new ManualResetEvent(false);
            Thread progressThread = new Thread(() =>
            {
                progressForm = new ProgressForm(selectedViews.Count);
                formReady.Set();
                Application.Run(progressForm);
            });
            progressThread.SetApartmentState(ApartmentState.STA);
            progressThread.IsBackground = true;
            progressThread.Start();
            formReady.WaitOne(5000);

            // 9. エクスポートループ
            int successCount = 0, failCount = 0, cancelCount = 0;
            List<string> errorDetails = new List<string>();
            List<string> logEntries = new List<string>();
            Stopwatch totalSw = Stopwatch.StartNew();

            try
            {
                for (int i = 0; i < selectedViews.Count; i++)
                {
                    if (progressForm != null && progressForm.StopRequested)
                    {
                        cancelCount = selectedViews.Count - i;
                        break;
                    }

                    ModelingView view = selectedViews[i];
                    if (progressForm != null)
                        progressForm.UpdateProgress(view.Name, i + 1, selectedViews.Count);

                    Stopwatch viewSw = Stopwatch.StartNew();

                    // ★【改修3】DXFパスを先に決定（例外発生時もFile.Existsで成否判定するため）
                    string baseName = SanitizeFileName(string.Format("{0}_{1}", partNo, view.Name));
                    string dxfPath = GetUniqueFilePath(outputFolder, baseName, ".dxf");

                    try
                    {
                        ExportViewAsDxf(view, dxfPath, pmiFontMap);
                    }
                    catch (Exception)
                    {
                        // 例外発生時もFile.Existsで成否判定を続行
                    }

                    viewSw.Stop();

                    // ★【改修3】成否判定: File.Existsで判定（UndoMark/Commit戻り値に依存しない）
                    bool exportSuccess = false;
                    if (dxfPath != null)
                    {
                        int waitMs = 0;
                        while (!File.Exists(dxfPath) && waitMs < 30000)
                        {
                            Thread.Sleep(500);
                            waitMs += 500;
                        }
                        if (File.Exists(dxfPath))
                        {
                            for (int retry = 0; retry < 20; retry++)
                            {
                                try
                                {
                                    using (var fs = new FileStream(dxfPath, FileMode.Open,
                                        FileAccess.Read, FileShare.None))
                                    {
                                        exportSuccess = fs.Length > 0;
                                        break;
                                    }
                                }
                                catch (IOException) { Thread.Sleep(500); }
                            }
                        }
                    }

                    if (exportSuccess)
                    {
                        successCount++;
                        logEntries.Add(string.Format("[OK] {0} ({1:F1}s)",
                            view.Name, viewSw.Elapsed.TotalSeconds));
                    }
                    else
                    {
                        failCount++;
                        string errorMsg = view.Name + ": DXF file not created";
                        errorDetails.Add(errorMsg);
                        logEntries.Add(string.Format("[FAIL] {0}: DXF file not created", view.Name));
                    }
                }
            }
            finally
            {
                // デバッグモード中はシートを残すためクリーンアップをスキップ
                if (!DebugDraftingCheck)
                {
                    // ★【改修6】残存シートの最終チェック・クリーンアップ
                    CleanupResidualSheets(originalSheetTags);

                    // モデリングに復帰
                    try { theSession.ApplicationSwitchImmediate("UG_APP_MODELING"); } catch { }
                }

                // 進捗ダイアログを閉じる
                if (progressForm != null)
                {
                    try { progressForm.ForceClose(); } catch { }
                }

                totalSw.Stop();
            }

            // 11. ログファイル出力
            WriteLogFile(outputFolder, partNo, logEntries, successCount, failCount, cancelCount,
                totalSw.Elapsed.TotalSeconds);

            // 12. 完了ダイアログ
            using (var completionForm = new CompletionForm(
                successCount, failCount, cancelCount, errorDetails, outputFolder))
            {
                completionForm.TopMost = true;
                completionForm.ShowDialog();
            }
        }

        // ================================================================
        // 個別ビューエクスポート（BoundingBox方式: 仮配置不要）
        // ================================================================

        private static void ExportViewAsDxf(ModelingView view, string dxfPath,
            Dictionary<Tag, PmiFontInfo> pmiFontMap)
        {
            const double SheetMargin = 10.0;
            const double PmiMarginFactor = 0.3; // PMI用30%マージン

            // === 3Dバウンディングボックスからビューサイズを推定 ===
            double bbMinX = double.MaxValue, bbMinY = double.MaxValue, bbMinZ = double.MaxValue;
            double bbMaxX = double.MinValue, bbMaxY = double.MinValue, bbMaxZ = double.MinValue;
            bool hasBodies = false;

            foreach (Body body in workPart.Bodies)
            {
                try
                {
                    double[] bbox = new double[6];
                    theUFSession.Modl.AskBoundingBox(body.Tag, bbox);
                    bbMinX = Math.Min(bbMinX, bbox[0]);
                    bbMinY = Math.Min(bbMinY, bbox[1]);
                    bbMinZ = Math.Min(bbMinZ, bbox[2]);
                    bbMaxX = Math.Max(bbMaxX, bbox[3]);
                    bbMaxY = Math.Max(bbMaxY, bbox[4]);
                    bbMaxZ = Math.Max(bbMaxZ, bbox[5]);
                    hasBodies = true;
                }
                catch { }
            }

            if (!hasBodies)
            {
                bbMinX = bbMinY = bbMinZ = -100;
                bbMaxX = bbMaxY = bbMaxZ = 100;
            }

            double modelSizeX = bbMaxX - bbMinX;
            double modelSizeY = bbMaxY - bbMinY;
            double modelSizeZ = bbMaxZ - bbMinZ;

            // ビューの向きに基づいて2D投影サイズを計算
            Matrix3x3 mtx = view.Matrix;
            double[][] corners = new double[][] {
                new double[] { bbMinX, bbMinY, bbMinZ }, new double[] { bbMaxX, bbMinY, bbMinZ },
                new double[] { bbMinX, bbMaxY, bbMinZ }, new double[] { bbMaxX, bbMaxY, bbMinZ },
                new double[] { bbMinX, bbMinY, bbMaxZ }, new double[] { bbMaxX, bbMinY, bbMaxZ },
                new double[] { bbMinX, bbMaxY, bbMaxZ }, new double[] { bbMaxX, bbMaxY, bbMaxZ },
            };

            double pMinX = double.MaxValue, pMaxX = double.MinValue;
            double pMinY = double.MaxValue, pMaxY = double.MinValue;

            foreach (double[] c in corners)
            {
                double px = c[0] * mtx.Xx + c[1] * mtx.Xy + c[2] * mtx.Xz;
                double py = c[0] * mtx.Yx + c[1] * mtx.Yy + c[2] * mtx.Yz;
                pMinX = Math.Min(pMinX, px);
                pMaxX = Math.Max(pMaxX, px);
                pMinY = Math.Min(pMinY, py);
                pMaxY = Math.Max(pMaxY, py);
            }

            double viewWidth = pMaxX - pMinX;
            double viewHeight = pMaxY - pMinY;

            // PMIマージン（30%）
            double pmiMargin = Math.Max(viewWidth, viewHeight) * PmiMarginFactor;
            double requiredWidth = viewWidth + pmiMargin + (SheetMargin * 2);
            double requiredHeight = viewHeight + pmiMargin + (SheetMargin * 2);

            // デバッグログ
            theSession.ListingWindow.Open();
            theSession.ListingWindow.WriteLine("=== DEBUG: BoundingBox方式 ===");
            theSession.ListingWindow.WriteLine("modelSize=" + modelSizeX + " x " + modelSizeY + " x " + modelSizeZ);
            theSession.ListingWindow.WriteLine("viewWidth=" + viewWidth + " viewHeight=" + viewHeight);
            theSession.ListingWindow.WriteLine("pmiMargin=" + pmiMargin);
            theSession.ListingWindow.WriteLine("requiredWidth=" + requiredWidth + " requiredHeight=" + requiredHeight);

            // === 用紙サイズ決定 ===
            double sheetW, sheetH;
            string paperName;
            GetOptimalPaperSize(requiredWidth, requiredHeight, out sheetW, out sheetH, out paperName);

            theSession.ListingWindow.WriteLine("sheetW=" + sheetW + " sheetH=" + sheetH + " paper=" + paperName);

            // === シート作成・ベースビュー配置・DXFエクスポート ===
            DrawingSheet exportSheet = null;
            try
            {
                exportSheet = CreateSheet(sheetW, sheetH);
                exportSheet.Open();
                workPart.Drafting.SetTemplateInstantiationIsComplete(true);

                // スナップショット（PlaceBaseView前）：既存Tagを記録
                HashSet<Tag> snapshotBeforeView = SnapshotDraftingEntityTags();

                DraftingView exportView = PlaceBaseView(view, sheetW / 2.0, sheetH / 2.0);

                // PmiNoteフォントサイズ適用（スナップショット差分方式・PmiLabel等）
                ApplyPmiFontSizes(pmiFontMap, snapshotBeforeView, exportView);

                // ★デバッグモード：無効（DXFエクスポートを実行）
                // DXFエクスポート
                PerformDxfExport(dxfPath);

                // DXF後処理：フォントサイズをPMIから反映
                PostProcessDxfFontSizes(dxfPath, pmiFontMap);
            }
            finally
            {
                DeleteSheetAndViews(exportSheet);
            }
        }

        // ================================================================
        // 製図エンティティTagのスナップショット取得（Bug①対策用）
        // ================================================================

        private static HashSet<Tag> SnapshotDraftingEntityTags()
        {
            var snapshot = new HashSet<Tag>();
            try
            {
                Tag objTag = Tag.Null;
                theUFSession.Obj.CycleObjsInPart(workPart.Tag,
                    NXOpen.UF.UFConstants.UF_drafting_entity_type, ref objTag);
                while (objTag != Tag.Null)
                {
                    snapshot.Add(objTag);
                    theUFSession.Obj.CycleObjsInPart(workPart.Tag,
                        NXOpen.UF.UFConstants.UF_drafting_entity_type, ref objTag);
                }
            }
            catch { }

            // ★type=0（全タイプ）でも追加スキャン
            // 製図側PmiNoteインスタンスはUF_drafting_entity_typeに含まれない場合があるため
            try
            {
                Tag objTag2 = Tag.Null;
                theUFSession.Obj.CycleObjsInPart(workPart.Tag, 0, ref objTag2);
                while (objTag2 != Tag.Null)
                {
                    snapshot.Add(objTag2);
                    theUFSession.Obj.CycleObjsInPart(workPart.Tag, 0, ref objTag2);
                }
            }
            catch { }

            theSession.ListingWindow.WriteLine("  Snapshot: " + snapshot.Count + " existing tags (all types)");
            return snapshot;
        }

        // ================================================================
        // 図面シート作成
        // ================================================================

        private static DrawingSheet CreateSheet(double width, double height)
        {
            DraftingDrawingSheet nullSheet = null;
            var builder = workPart.DraftingDrawingSheets.CreateDraftingDrawingSheetBuilder(nullSheet);
            builder.AutoStartViewCreation = false;
            builder.Option = DrawingSheetBuilder.SheetOption.CustomSize;
            builder.Length = width;
            builder.Height = height;
            builder.ScaleNumerator = 1.0;
            builder.ScaleDenominator = 1.0;
            builder.Units = DrawingSheetBuilder.SheetUnits.Metric;
            builder.ProjectionAngle = DrawingSheetBuilder.SheetProjectionAngle.Third;

            NXObject sheetObj = builder.Commit();
            builder.Destroy();
            return (DrawingSheet)sheetObj;
        }

        // ================================================================
        // BaseView配置（Freeze/Unfreezeで囲む）
        // ================================================================

        private static DraftingView PlaceBaseView(ModelingView view,
            double centerX, double centerY)
        {
            workPart.DraftingManager.DrawingsFreezeOutOfDateComputation();

            BaseView nullBaseView = null;
            BaseViewBuilder builder = workPart.DraftingViews.CreateBaseViewBuilder(nullBaseView);

            try
            {
                builder.Placement.Associative = true;
                builder.SelectModelView.SelectedView = view;

                builder.SecondaryComponents.ObjectType =
                    DraftingComponentSelectionBuilder.Geometry.PrimaryGeometry;

                // パート参照の設定
                builder.Style.ViewStyleBase.Part = workPart;
                builder.Style.ViewStyleBase.PartName = workPart.FullPath;
                builder.SelectModelView.SelectedView = view;

                // スケール 1:1（常に固定、縮小しない）
                builder.Scale.Numerator = 1.0;
                builder.Scale.Denominator = 1.0;

                // アレンジメント（null対応）
                try
                {
                    NXOpen.Assemblies.Arrangement[] arrs;
                    workPart.GetArrangements(out arrs);
                    if (arrs != null && arrs.Length > 0)
                    {
                        builder.Style.ViewStyleBase.Arrangement.SelectedArrangement = arrs[0];
                        builder.Style.ViewStyleBase.Arrangement.InheritArrangementFromParent = false;
                    }
                }
                catch { }

                // ★【改修1/2】PMI継承設定: InDrawingPlaneFromView（ジャーナル確認済み）
                try
                {
                    builder.Style.ViewStyleInheritPmi.Pmi =
                        NXOpen.Preferences.PmiOption.InDrawingPlaneFromView;
                    builder.Style.ViewStyleInheritPmi.PmiToDrawing = true;
                }
                catch { }

                // GDT継承設定（InDrawingPlane = 2D図面平面に投影）
                try
                {
                    builder.Style.ViewStyleInheritPmi.Gdt =
                        NXOpen.Preferences.GdtOption.InDrawingPlane;
                }
                catch { }

                // 配置位置（シート中央）
                Point3d origin = new Point3d(centerX, centerY, 0.0);
                builder.Placement.Placement.SetValue(null, workPart.Views.WorkView, origin);

                NXObject viewObj = builder.Commit();
                workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();

                return (DraftingView)viewObj;
            }
            catch
            {
                workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();
                builder.Destroy();
                throw;
            }
            finally
            {
                try { builder.Destroy(); } catch { }
            }
        }

        // ================================================================
        // DXF後処理：PMIフォントサイズをDXFファイルに反映
        // NX Open APIでは製図ビュー内の寸法・注記のフォントサイズを
        // 変更できないため、出力後のDXFファイルを直接書き換える
        // ================================================================

        private static void PostProcessDxfFontSizes(string dxfPath,
            Dictionary<Tag, PmiFontInfo> pmiFontMap)
        {
            if (pmiFontMap == null || pmiFontMap.Count == 0) return;
            if (string.IsNullOrEmpty(dxfPath) || !File.Exists(dxfPath)) return;

            try
            {
                // テキスト値 → DimensionTextSize のマップを構築
                // 同じテキスト値で複数のサイズがある場合は最大サイズを採用
                var textToDimSize = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                // GenSize（注記・ラベル用）マップ
                var textToGenSize = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

                foreach (var kv in pmiFontMap.Values)
                {
                    if (kv.IsDimension && kv.DimensionTextSize > 0
                        && !string.IsNullOrEmpty(kv.TextContent))
                    {
                        double existing;
                        if (!textToDimSize.TryGetValue(kv.TextContent, out existing)
                            || kv.DimensionTextSize > existing)
                            textToDimSize[kv.TextContent] = kv.DimensionTextSize;
                    }
                    if (kv.GeneralTextSize > 3.5 && !string.IsNullOrEmpty(kv.TextContent))
                    {
                        double existing;
                        if (!textToGenSize.TryGetValue(kv.TextContent, out existing)
                            || kv.GeneralTextSize > existing)
                            textToGenSize[kv.TextContent] = kv.GeneralTextSize;
                    }
                }

                string[] lines = File.ReadAllLines(dxfPath, System.Text.Encoding.UTF8);
                var result = new System.Text.StringBuilder();

                // DXFをエンティティ単位でパースしてDIMENSION/TEXTのフォントサイズを修正
                int i = 0;
                int dimApplied = 0;
                int textApplied = 0;

                while (i < lines.Length)
                {
                    string codeLine = lines[i].Trim();
                    int groupCode;
                    if (!int.TryParse(codeLine, out groupCode))
                    {
                        result.AppendLine(lines[i]);
                        i++;
                        continue;
                    }

                    string valueLine = (i + 1 < lines.Length) ? lines[i + 1] : "";

                    // DIMENSION エンティティを検出してバッファに収集
                    if (groupCode == 0 && valueLine.Trim() == "DIMENSION")
                    {
                        // エンティティ全体を収集
                        var entity = new List<KeyValuePair<int, string>>();
                        entity.Add(new KeyValuePair<int, string>(0, valueLine.Trim()));
                        i += 2;

                        while (i < lines.Length)
                        {
                            string c = lines[i].Trim();
                            int gc;
                            if (!int.TryParse(c, out gc)) { break; }
                            string v = (i + 1 < lines.Length) ? lines[i + 1].Trim() : "";
                            // 次のエンティティ開始（group code 0）で終了
                            if (gc == 0) break;
                            entity.Add(new KeyValuePair<int, string>(gc, v));
                            i += 2;
                        }

                        // テキスト値を取得（group code 1）
                        string dimText = "";
                        foreach (var kv in entity)
                            if (kv.Key == 1) { dimText = kv.Value.Trim(); break; }

                        // マッチングしてDIMTXT（140）を設定
                        double newDimSize = 0;
                        if (!string.IsNullOrEmpty(dimText) && textToDimSize.ContainsKey(dimText))
                            newDimSize = textToDimSize[dimText];

                        // エンティティを書き出し（140を挿入または上書き）
                        bool has140 = false;
                        foreach (var kv in entity)
                        {
                            if (kv.Key == 140) has140 = true;
                        }

                        bool inserted140 = false;
                        foreach (var kv in entity)
                        {
                            if (kv.Key == 140 && newDimSize > 0)
                            {
                                // 既存の140を上書き
                                result.AppendLine("140");
                                result.AppendLine(newDimSize.ToString("F6",
                                    System.Globalization.CultureInfo.InvariantCulture));
                                dimApplied++;
                                inserted140 = true;
                                continue;
                            }
                            // group code 3（スタイル名）の直後に140を挿入
                            result.AppendLine(kv.Key.ToString());
                            result.AppendLine(kv.Value);
                            if (!has140 && !inserted140 && kv.Key == 3 && newDimSize > 0)
                            {
                                result.AppendLine("140");
                                result.AppendLine(newDimSize.ToString("F6",
                                    System.Globalization.CultureInfo.InvariantCulture));
                                dimApplied++;
                                inserted140 = true;
                            }
                        }
                        // 次のエンティティはループ先頭で処理
                        continue;
                    }

                    // MTEXT / TEXT エンティティのフォントサイズ修正（注記・ラベル用）
                    if (groupCode == 0 &&
                        (valueLine.Trim() == "TEXT" || valueLine.Trim() == "MTEXT"))
                    {
                        var entity = new List<KeyValuePair<int, string>>();
                        entity.Add(new KeyValuePair<int, string>(0, valueLine.Trim()));
                        i += 2;

                        while (i < lines.Length)
                        {
                            string c = lines[i].Trim();
                            int gc;
                            if (!int.TryParse(c, out gc)) { break; }
                            string v = (i + 1 < lines.Length) ? lines[i + 1].Trim() : "";
                            if (gc == 0) break;
                            entity.Add(new KeyValuePair<int, string>(gc, v));
                            i += 2;
                        }

                        // テキスト内容（code 1）を取得
                        string textContent = "";
                        foreach (var kv in entity)
                            if (kv.Key == 1) { textContent = kv.Value.Trim(); break; }

                        // MTEXTはMTEXTフォーマット制御文字を除去
                        string cleanText = System.Text.RegularExpressions.Regex
                            .Replace(textContent, @"\\[A-Za-z][^;]*;|[{}]", "").Trim();

                        double newGenSize = 0;
                        if (!string.IsNullOrEmpty(cleanText))
                        {
                            if (textToGenSize.ContainsKey(cleanText))
                                newGenSize = textToGenSize[cleanText];
                            else if (textToGenSize.ContainsKey(textContent))
                                newGenSize = textToGenSize[textContent];
                        }

                        foreach (var kv in entity)
                        {
                            // code 40 = テキスト高さ
                            if (kv.Key == 40 && newGenSize > 0)
                            {
                                result.AppendLine("40");
                                result.AppendLine(newGenSize.ToString("F6",
                                    System.Globalization.CultureInfo.InvariantCulture));
                                textApplied++;
                                continue;
                            }
                            result.AppendLine(kv.Key.ToString());
                            result.AppendLine(kv.Value);
                        }
                        continue;
                    }

                    // 通常行はそのまま出力
                    result.AppendLine(lines[i]);
                    if (i + 1 < lines.Length)
                        result.AppendLine(lines[i + 1]);
                    i += 2;
                }

                File.WriteAllText(dxfPath, result.ToString(), System.Text.Encoding.UTF8);

                theSession.ListingWindow.WriteLine(
                    "=== DXF後処理完了: dimensions=" + dimApplied +
                    " texts=" + textApplied + " ===");
            }
            catch (Exception ex)
            {
                theSession.ListingWindow.WriteLine("=== DXF後処理エラー: " + ex.Message + " ===");
            }
        }

        // ================================================================
        // DXFエクスポート（★【改修4】Freeze/Unfreezeで確実に囲む）
        // ================================================================

        private static void PerformDxfExport(string dxfPath)
        {
            DxfdwgCreator creator = theSession.DexManager.CreateDxfdwgCreator();

            try
            {
                creator.ExportData = DxfdwgCreator.ExportDataOption.Drawing;
                creator.OutputTo = DxfdwgCreator.OutputToOption.Drafting;
                creator.OutputFileType = DxfdwgCreator.OutputFileTypeOption.Dxf;
                creator.AutoCADRevision = DxfdwgCreator.AutoCADRevisionOptions.R2018;
                creator.FlattenAssembly = false;
                creator.ViewEditMode = false;
                creator.ExportScaleValue = 1.0;
                creator.LayerMask = "1-256";
                creator.DrawingList = "_ALL_";
                creator.ProcessHoldFlag = true;
                creator.WidthFactorMode = DxfdwgCreator.WidthfactorMethodOptions.AutomaticCalculation;

                creator.ObjectTypes.Curves = true;
                creator.ObjectTypes.Annotations = true;
                creator.ObjectTypes.Structures = true;

                creator.InputFile = workPart.FullPath;
                creator.OutputFile = dxfPath;
                creator.SettingsFile = tempDefPath;

                // ★【改修4】Freeze/Unfreezeで囲んで不要なビュー再計算を抑制
                workPart.DraftingManager.DrawingsFreezeOutOfDateComputation();
                try
                {
                    creator.Commit();
                }
                finally
                {
                    workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();
                }
            }
            finally
            {
                try { creator.Destroy(); } catch { }
            }
        }

        // ================================================================
        // ★【改修1】PMIフォントサイズ読み取り（モデル側、1回だけ）
        // ================================================================

        private static PmiFontInfo ReadPmiFontSize(NXOpen.Annotations.Annotation pmiObj)
        {
            PmiFontInfo info = new PmiFontInfo();
            NXOpen.Annotations.EditSettingsBuilder esb = null;
            // ★読み取り専用: CommitもProcessForMultipleObjectsSettingsも呼ばない
            // CreateAnnotationEditSettingsBuilder + 値読み取り + Destroy のみ
            // ProcessForMultipleObjectsSettingsはモデルを書き換えるリスクがあるため使用禁止
            try
            {
                DisplayableObject[] objects = new DisplayableObject[] { (DisplayableObject)pmiObj };
                esb = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects);

                var letteringStyle = esb.AnnotationStyle.LetteringStyle;
                if (pmiObj is NXOpen.Annotations.Dimension)
                {
                    info.IsDimension = true;
                    try { info.DimensionTextSize = letteringStyle.DimensionTextSize; } catch { }
                    try { info.ToleranceTextSize = letteringStyle.ToleranceTextSize; } catch { }
                    try { info.AppendedTextSize = letteringStyle.AppendedTextSize; } catch { }
                }
                try { info.GeneralTextSize = letteringStyle.GeneralTextSize; } catch { }
            }
            finally
            {
                // Commitなし・DestroyのみでモデルへのSide effectを排除
                if (esb != null) try { esb.Destroy(); } catch { }
            }
            return info;
        }

        private static void ReadAndStorePmiNote(NXOpen.Annotations.Note pmiNote,
            Dictionary<Tag, PmiFontInfo> pmiFontMap)
        {
            var objects = new DisplayableObject[] { (DisplayableObject)pmiNote };
            NXOpen.Annotations.EditSettingsBuilder esb = null;
            // ★読み取り専用: CommitもProcessForMultipleObjectsSettingsも呼ばない
            try
            {
                esb = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects);

                string noteText = "";
                try
                {
                    string[] t = pmiNote.GetText();
                    if (t != null && t.Length > 0) noteText = string.Join("\n", t);
                }
                catch { }
                if (string.IsNullOrEmpty(noteText)) noteText = pmiNote.ToString();

                double genSize = 0;
                try { genSize = esb.AnnotationStyle.LetteringStyle.GeneralTextSize; } catch { }

                var info = new PmiFontInfo
                {
                    DimensionTextSize = 0,
                    ToleranceTextSize = 0,
                    AppendedTextSize = 0,
                    GeneralTextSize = genSize,
                    TextContent = noteText,
                    PmiTypeName = "PmiNote",
                    IsDimension = false
                };
                pmiFontMap[pmiNote.Tag] = info;

                theSession.ListingWindow.WriteLine("  PMI Tag=" + pmiNote.Tag +
                    " type=PmiNote text=\"" + noteText + "\"" +
                    " dimSize=0 genSize=" + info.GeneralTextSize);
            }
            finally
            {
                if (esb != null) try { esb.Destroy(); } catch { }
            }
        }

        // PmiNote読み取り（Annotation基底型版）
        // ★PmiNoteはNXOpen.Annotations.Noteを継承しない可能性があるためAnnotation型で受け取る
        private static void ReadAndStorePmiAnnotation(NXOpen.Annotations.Annotation pmiAnn,
            Dictionary<Tag, PmiFontInfo> pmiFontMap)
        {
            NXOpen.Annotations.EditSettingsBuilder esb = null;
            try
            {
                var objects = new DisplayableObject[] { (DisplayableObject)pmiAnn };
                esb = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects);

                // テキスト取得（Note型にキャスト試行 → 失敗時はToString）
                string noteText = "";
                try
                {
                    if (pmiAnn is NXOpen.Annotations.Note)
                    {
                        string[] t = ((NXOpen.Annotations.Note)pmiAnn).GetText();
                        if (t != null && t.Length > 0) noteText = string.Join("\n", t);
                    }
                }
                catch { }
                if (string.IsNullOrEmpty(noteText))
                    noteText = pmiAnn.ToString();

                double genSize = 0;
                try { genSize = esb.AnnotationStyle.LetteringStyle.GeneralTextSize; } catch { }

                var info = new PmiFontInfo
                {
                    DimensionTextSize = 0,
                    ToleranceTextSize = 0,
                    AppendedTextSize = 0,
                    GeneralTextSize = genSize,
                    TextContent = noteText,
                    PmiTypeName = "PmiNote",
                    IsDimension = false
                };
                pmiFontMap[pmiAnn.Tag] = info;

                theSession.ListingWindow.WriteLine("  PMI Tag=" + pmiAnn.Tag +
                    " type=PmiNote(Ann) text=\"" + noteText + "\"" +
                    " genSize=" + genSize);
            }
            finally
            {
                if (esb != null) try { esb.Destroy(); } catch { }
            }
        }

        private static Dictionary<Tag, PmiFontInfo> ReadAllPmiFontSizes()
        {
            var pmiFontMap = new Dictionary<Tag, PmiFontInfo>();

            try
            {
                // パート内の全PMI寸法を列挙
                foreach (NXOpen.Annotations.Dimension dim in workPart.Dimensions)
                {
                    try
                    {
                        if (dim.GetType().Name.StartsWith("Pmi"))
                        {
                            PmiFontInfo info = ReadPmiFontSize(dim);
                            info.PmiTypeName = dim.GetType().Name;
                            info.TextContent = GetPmiText(dim);
                            pmiFontMap[dim.Tag] = info;
                        }
                    }
                    catch { }
                }
            }
            catch { }

            // === PmiNote検索：UF_drafting_entity_type スキャン ===
            // ★根本修正: PmiNoteはNXOpen.Annotations.Noteを継承していない可能性がある。
            // is NoteではなくGetType().Name == "PmiNote" で判定する。
            // type=0(全タイプ)はPMIオブジェクトを返さない場合があるため、
            // UF_drafting_entity_typeを使う。
            int noteReadCount = 0;
            try
            {
                Tag scanTag = Tag.Null;
                theUFSession.Obj.CycleObjsInPart(workPart.Tag,
                    NXOpen.UF.UFConstants.UF_drafting_entity_type, ref scanTag);
                while (scanTag != Tag.Null)
                {
                    try
                    {
                        if (!pmiFontMap.ContainsKey(scanTag))
                        {
                            NXOpen.TaggedObject tObj = NXOpen.Utilities.NXObjectManager.Get(scanTag);
                            // ★is Annotation でチェック（is Note は不要、PmiNoteはNoteを継承しない）
                            if (tObj is NXOpen.Annotations.Annotation
                                && tObj.GetType().Name == "PmiNote")
                            {
                                ReadAndStorePmiAnnotation(
                                    (NXOpen.Annotations.Annotation)tObj, pmiFontMap);
                                noteReadCount++;
                            }
                        }
                    }
                    catch { }
                    theUFSession.Obj.CycleObjsInPart(workPart.Tag,
                        NXOpen.UF.UFConstants.UF_drafting_entity_type, ref scanTag);
                }
            }
            catch { }

            // === フォールバック：type=0 全スキャン（DraftingEntityTypeに入らない場合）===
            if (noteReadCount == 0)
            {
                try
                {
                    Tag scanTag = Tag.Null;
                    theUFSession.Obj.CycleObjsInPart(workPart.Tag, 0, ref scanTag);
                    while (scanTag != Tag.Null)
                    {
                        try
                        {
                            if (!pmiFontMap.ContainsKey(scanTag))
                            {
                                NXOpen.TaggedObject tObj = NXOpen.Utilities.NXObjectManager.Get(scanTag);
                                if (tObj is NXOpen.Annotations.Annotation
                                    && tObj.GetType().Name == "PmiNote")
                                {
                                    ReadAndStorePmiAnnotation(
                                        (NXOpen.Annotations.Annotation)tObj, pmiFontMap);
                                    noteReadCount++;
                                }
                            }
                        }
                        catch { }
                        theUFSession.Obj.CycleObjsInPart(workPart.Tag, 0, ref scanTag);
                    }
                }
                catch { }
            }

            // === フォールバック：workPart.Notes ===
            if (noteReadCount == 0)
            {
                try
                {
                    foreach (NXOpen.Annotations.Note note in workPart.Notes)
                    {
                        if (note.GetType().Name != "PmiNote") continue;
                        try
                        {
                            ReadAndStorePmiAnnotation(note, pmiFontMap);
                            noteReadCount++;
                        }
                        catch { }
                    }
                }
                catch { }
            }

            // === フォールバック：type=158 ===
            if (noteReadCount == 0)
            {
                try
                {
                    Tag pmiTag = Tag.Null;
                    theUFSession.Obj.CycleObjsInPart(workPart.Tag, 158, ref pmiTag);
                    while (pmiTag != Tag.Null)
                    {
                        try
                        {
                            if (!pmiFontMap.ContainsKey(pmiTag))
                            {
                                NXOpen.TaggedObject tObj = NXOpen.Utilities.NXObjectManager.Get(pmiTag);
                                if (tObj is NXOpen.Annotations.Annotation)
                                {
                                    ReadAndStorePmiAnnotation(
                                        (NXOpen.Annotations.Annotation)tObj, pmiFontMap);
                                    noteReadCount++;
                                }
                            }
                        }
                        catch { }
                        theUFSession.Obj.CycleObjsInPart(workPart.Tag, 158, ref pmiTag);
                    }
                }
                catch { }
            }

            theSession.ListingWindow.WriteLine("  PmiNote search: found " + noteReadCount + " notes");

            try
            {
                // パート内の全PMIラベルを列挙
                foreach (NXOpen.Annotations.Label label in workPart.Labels)
                {
                    try
                    {
                        if (label.GetType().Name.Contains("Pmi"))
                        {
                            PmiFontInfo info = ReadPmiFontSize(label);
                            info.PmiTypeName = label.GetType().Name;
                            info.TextContent = GetPmiText(label);
                            pmiFontMap[label.Tag] = info;
                        }
                    }
                    catch { }
                }
            }
            catch { }

            theSession.ListingWindow.Open();
            theSession.ListingWindow.WriteLine("=== DEBUG PMI: ReadAllPmiFontSizes count=" + pmiFontMap.Count + " (notes=" + noteReadCount + ") ===");
            foreach (var kvp in pmiFontMap)
            {
                theSession.ListingWindow.WriteLine("  PMI Tag=" + kvp.Key + " type=" + kvp.Value.PmiTypeName +
                    " text=\"" + kvp.Value.TextContent + "\" dimSize=" + kvp.Value.DimensionTextSize +
                    " genSize=" + kvp.Value.GeneralTextSize);
            }

            return pmiFontMap;
        }

        private static string GetPmiText(NXOpen.Annotations.Annotation pmi)
        {
            try
            {
                if (pmi is NXOpen.Annotations.Dimension)
                {
                    var dim = (NXOpen.Annotations.Dimension)pmi;
                    try
                    {
                        string[] mainText;
                        string[] dualText;
                        dim.GetDimensionText(out mainText, out dualText);
                        if (mainText != null && mainText.Length > 0)
                            return string.Join(" ", mainText);
                    }
                    catch { }
                    return dim.ToString();
                }
                if (pmi is NXOpen.Annotations.Note)
                {
                    var note = (NXOpen.Annotations.Note)pmi;
                    try
                    {
                        string[] lines = note.GetText();
                        if (lines != null && lines.Length > 0)
                            return string.Join(" ", lines);
                    }
                    catch { }
                    return note.ToString();
                }
                if (pmi is NXOpen.Annotations.Label)
                {
                    var label = (NXOpen.Annotations.Label)pmi;
                    try
                    {
                        string[] lines = label.GetText();
                        if (lines != null && lines.Length > 0)
                            return string.Join(" ", lines);
                    }
                    catch { }
                    return label.ToString();
                }
            }
            catch { }
            return pmi.ToString();
        }

        // 製図側アノテーションのテキスト取得（NXOpen API → UF API フォールバック）
        private static string GetAnnotationText(NXOpen.NXObject ann)
        {
            // まず NXOpen API で試行
            try
            {
                if (ann is NXOpen.Annotations.Note)
                {
                    string[] t = ((NXOpen.Annotations.Note)ann).GetText();
                    if (t != null && t.Length > 0)
                    {
                        string result = string.Join(" ", t);
                        if (!string.IsNullOrEmpty(result) && !result.StartsWith("PmiNote"))
                            return result;
                    }
                }
                if (ann is NXOpen.Annotations.Label)
                {
                    string[] t = ((NXOpen.Annotations.Label)ann).GetText();
                    if (t != null && t.Length > 0)
                        return string.Join(" ", t);
                }
            }
            catch { }

            // NXOpen API失敗時、Dimension のテキスト取得を試行
            try
            {
                if (ann is NXOpen.Annotations.Dimension)
                {
                    var dim = (NXOpen.Annotations.Dimension)ann;
                    string[] mainText;
                    string[] dualText;
                    dim.GetDimensionText(out mainText, out dualText);
                    if (mainText != null && mainText.Length > 0)
                        return string.Join(" ", mainText);
                }
            }
            catch { }

            return ann.ToString();
        }

        // ================================================================
        // 継承PMIへのフォントサイズ反映（スナップショット差分方式）
        // ================================================================

        private static void ApplyPmiFontSizes(Dictionary<Tag, PmiFontInfo> pmiFontMap,
            HashSet<Tag> snapshotBeforeView, DraftingView exportView)
        {
            if (pmiFontMap == null || pmiFontMap.Count == 0)
            {
                theSession.ListingWindow.WriteLine("=== DEBUG PMI: pmiFontMap is null or empty ===");
                return;
            }

            theSession.ListingWindow.WriteLine("=== DEBUG PMI: ApplyPmiFontSizes start, map count=" + pmiFontMap.Count + " ===");

            int noteCount = 0;

            // ─────────────────────────────────────────────────────────────
            // ★Dimension は PmiToDrawing で DimensionTextSize が正しく継承される。
            // Commit() を呼ぶと制図アプリのスケール変換が二重適用されて縮小するため
            // Dimension には一切触れない。
            // ─────────────────────────────────────────────────────────────

            // PmiNoteエントリの一覧（新規Tagの照合用）
            var pmiNotesByTag = new Dictionary<Tag, PmiFontInfo>();
            foreach (var kv in pmiFontMap)
                if (kv.Value.PmiTypeName == "PmiNote")
                    pmiNotesByTag[kv.Key] = kv.Value;

            // PmiNoteの全GenSize一覧（フォールバック用）
            var pmiNoteSizeList = new List<double>();
            foreach (var kv in pmiNotesByTag)
                if (kv.Value.GeneralTextSize > 0)
                    pmiNoteSizeList.Add(kv.Value.GeneralTextSize);

            bool allPmiNotesSameSize = (pmiNoteSizeList.Count > 0);
            if (allPmiNotesSameSize)
            {
                for (int k = 1; k < pmiNoteSizeList.Count; k++)
                    if (Math.Abs(pmiNoteSizeList[k] - pmiNoteSizeList[0]) > 0.001)
                    { allPmiNotesSameSize = false; break; }
            }
            double uniformPmiNoteSize = allPmiNotesSameSize && pmiNoteSizeList.Count > 0
                ? pmiNoteSizeList[0] : 0;

            theSession.ListingWindow.WriteLine(
                "  PmiNote entries in map: " + pmiNotesByTag.Count +
                " uniformSize=" + (allPmiNotesSameSize ? uniformPmiNoteSize.ToString() : "mixed"));

            try
            {
                // ★type=0（全タイプ）でスキャン
                // 製図側PmiNoteインスタンスはUF_drafting_entity_typeに含まれない場合があるため
                var newDraftingObjects = new List<Tag>();

                Tag objTag = Tag.Null;
                theUFSession.Obj.CycleObjsInPart(workPart.Tag,
                    NXOpen.UF.UFConstants.UF_drafting_entity_type, ref objTag);
                while (objTag != Tag.Null)
                {
                    if (!snapshotBeforeView.Contains(objTag))
                        newDraftingObjects.Add(objTag);
                    theUFSession.Obj.CycleObjsInPart(workPart.Tag,
                        NXOpen.UF.UFConstants.UF_drafting_entity_type, ref objTag);
                }

                // type=0 でも追加スキャン（PmiNote描画インスタンス検出）
                Tag objTag2 = Tag.Null;
                theUFSession.Obj.CycleObjsInPart(workPart.Tag, 0, ref objTag2);
                while (objTag2 != Tag.Null)
                {
                    if (!snapshotBeforeView.Contains(objTag2)
                        && !newDraftingObjects.Contains(objTag2))
                        newDraftingObjects.Add(objTag2);
                    theUFSession.Obj.CycleObjsInPart(workPart.Tag, 0, ref objTag2);
                }

                theSession.ListingWindow.WriteLine(
                    "  New Objects (post-PlaceBaseView, all types): " + newDraftingObjects.Count);

                foreach (Tag dTag in newDraftingObjects)
                {
                    try
                    {
                        NXOpen.TaggedObject tObj = NXOpen.Utilities.NXObjectManager.Get(dTag);
                        if (tObj == null) continue;

                        // ─── Note処理（PmiNote含む全Note型）───
                        bool isNote = tObj is NXOpen.Annotations.Note;
                        bool isPmiNote = isNote || tObj.GetType().Name == "PmiNote"
                            || tObj.GetType().Name.Contains("PmiNote");

                        if (isPmiNote)
                        {
                            NXOpen.Annotations.Annotation draftNoteAnn =
                                tObj as NXOpen.Annotations.Annotation;
                            if (draftNoteAnn == null) continue;
                            NXOpen.Annotations.Note draftNote =
                                draftNoteAnn as NXOpen.Annotations.Note;
                            PmiFontInfo noteInfo = null;

                            // 優先①: 同一TagがpmiFontMap（PmiNote）に存在 → Tag直接照合
                            if (pmiNotesByTag.ContainsKey(dTag))
                            {
                                noteInfo = pmiNotesByTag[dTag];
                                theSession.ListingWindow.WriteLine(
                                    "  Note[TagMatch] tag=" + dTag + " -> GenSize=" + noteInfo.GeneralTextSize);
                            }

                            // 優先②: テキスト照合（完全一致）
                            if (noteInfo == null)
                            {
                                string noteText = draftNote != null
                                    ? GetAnnotationText(draftNote)
                                    : draftNoteAnn.ToString();
                                bool textUnavailable = string.IsNullOrEmpty(noteText)
                                    || noteText.StartsWith("PmiNote ") || noteText.StartsWith("Note ");

                                if (!textUnavailable)
                                {
                                    foreach (var entry in pmiNotesByTag.Values)
                                    {
                                        if (!string.IsNullOrEmpty(entry.TextContent) &&
                                            entry.TextContent == noteText &&
                                            entry.GeneralTextSize > 0)
                                        { noteInfo = entry; break; }
                                    }
                                    // 優先③: テキスト部分一致（3文字以上）
                                    if (noteInfo == null && noteText.Length >= 3)
                                    {
                                        foreach (var entry in pmiNotesByTag.Values)
                                        {
                                            if (!string.IsNullOrEmpty(entry.TextContent) &&
                                                entry.TextContent.Length >= 3 &&
                                                (entry.TextContent.Contains(noteText)
                                                 || noteText.Contains(entry.TextContent)) &&
                                                entry.GeneralTextSize > 0)
                                            { noteInfo = entry; break; }
                                        }
                                    }
                                    if (noteInfo != null)
                                        theSession.ListingWindow.WriteLine(
                                            "  Note[TextMatch] text=\"" + noteText + "\" -> GenSize=" + noteInfo.GeneralTextSize);
                                    else
                                        theSession.ListingWindow.WriteLine(
                                            "  Note[NoMatch] text=\"" + noteText + "\"");
                                }
                                else
                                {
                                    // ★決定的な手がかり:
                                    // 製図側PmiNoteのGetText()/ToString()は "PmiNote {モデルTag}" を返す
                                    // 例: "PmiNote 65428" → モデルTag=65428 → pmiFontMapから正しいGenSizeを取得
                                    string rawText = draftNoteAnn.ToString();
                                    Tag modelTag = Tag.Null;
                                    if (rawText.StartsWith("PmiNote "))
                                    {
                                        string tagStr = rawText.Substring("PmiNote ".Length).Trim();
                                        ulong tagVal;
                                        if (ulong.TryParse(tagStr, out tagVal))
                                            modelTag = (Tag)tagVal;
                                    }

                                    if (modelTag != Tag.Null && pmiNotesByTag.ContainsKey(modelTag))
                                    {
                                        noteInfo = pmiNotesByTag[modelTag];
                                        theSession.ListingWindow.WriteLine(
                                            "  Note[TagFromText] drawing=" + dTag +
                                            " modelTag=" + modelTag +
                                            " -> GenSize=" + noteInfo.GeneralTextSize);
                                    }
                                    else
                                    {
                                        theSession.ListingWindow.WriteLine(
                                            "  Note[Skip] tag=" + dTag +
                                            " rawText=\"" + rawText + "\"" +
                                            " modelTag=" + modelTag);
                                    }
                                }
                            }

                            if (noteInfo != null && noteInfo.GeneralTextSize > 0)
                            {
                                var dstObjects = new DisplayableObject[] {
                                    (DisplayableObject)draftNoteAnn };
                                NXOpen.Annotations.EditSettingsBuilder dstEsb = null;
                                try
                                {
                                    dstEsb = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(dstObjects);
                                    var builders = new NXOpen.Drafting.BaseEditSettingsBuilder[] { dstEsb };
                                    workPart.SettingsManager.ProcessForMultipleObjectsSettings(builders);
                                    dstEsb.AnnotationStyle.LetteringStyle.GeneralTextSize = noteInfo.GeneralTextSize;
                                    dstEsb.Commit();
                                    noteCount++;
                                    theSession.ListingWindow.WriteLine(
                                        "  NewNote[Apply] type=" + tObj.GetType().Name +
                                        " tag=" + dTag + " -> GenSize=" + noteInfo.GeneralTextSize);
                                }
                                finally
                                {
                                    if (dstEsb != null) try { dstEsb.Destroy(); } catch { }
                                }
                            }
                            continue;
                        }

                        // ─── Label・その他Annotation処理 ───
                        NXOpen.Annotations.Annotation ann = tObj as NXOpen.Annotations.Annotation;
                        if (ann == null) continue;

                        string annText = GetAnnotationText(ann);
                        if (string.IsNullOrEmpty(annText)) continue;
                        if (annText.StartsWith("PmiNote ") || annText.StartsWith("PmiLabel ")
                            || annText.StartsWith("Note ") || annText.StartsWith("Label ")) continue;

                        // 完全一致
                        PmiFontInfo match = null;
                        foreach (var entry in pmiFontMap.Values)
                        {
                            if (!string.IsNullOrEmpty(entry.TextContent) &&
                                entry.TextContent == annText && entry.GeneralTextSize > 0)
                            { match = entry; break; }
                        }
                        // 部分一致（3文字以上）
                        if (match == null && annText.Length >= 3)
                        {
                            foreach (var entry in pmiFontMap.Values)
                            {
                                if (!string.IsNullOrEmpty(entry.TextContent) &&
                                    entry.TextContent.Length >= 3 &&
                                    (entry.TextContent.Contains(annText) || annText.Contains(entry.TextContent)) &&
                                    entry.GeneralTextSize > 0)
                                { match = entry; break; }
                            }
                        }

                        if (match != null && match.GeneralTextSize > 0)
                        {
                            var objects = new DisplayableObject[] { (DisplayableObject)ann };
                            NXOpen.Annotations.EditSettingsBuilder esb = null;
                            try
                            {
                                esb = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects);
                                var builders = new NXOpen.Drafting.BaseEditSettingsBuilder[] { esb };
                                workPart.SettingsManager.ProcessForMultipleObjectsSettings(builders);
                                esb.AnnotationStyle.LetteringStyle.GeneralTextSize = match.GeneralTextSize;
                                esb.Commit();
                                noteCount++;
                                theSession.ListingWindow.WriteLine("  DraftObj: " + ann.GetType().Name +
                                    " text=\"" + annText + "\" -> GenSize=" + match.GeneralTextSize);
                            }
                            finally
                            {
                                if (esb != null) try { esb.Destroy(); } catch { }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                theSession.ListingWindow.WriteLine("  DraftingObj enumeration error: " + ex.Message);
            }

            // ─── DraftingView内のアノテーションをビューAPIで直接取得（寸法・公差・PmiNote対応）───
            if (exportView != null)
            {
                try
                {
                    // AskVisibleObjects でビュー内のオブジェクトを取得し、Tagリストに変換
                    Tag[] viewAnnTags = null;
                    try
                    {
                        DisplayableObject[] visObjs = exportView.AskVisibleObjects();
                        if (visObjs != null && visObjs.Length > 0)
                        {
                            var tagList = new List<Tag>();
                            foreach (var vo in visObjs)
                                tagList.Add(vo.Tag);
                            viewAnnTags = tagList.ToArray();
                        }
                    }
                    catch { viewAnnTags = null; }

                    if (viewAnnTags != null && viewAnnTags.Length > 0)
                    {
                        theSession.ListingWindow.WriteLine(
                            "  ViewAnn scan: " + viewAnnTags.Length + " objects in view");

                        foreach (Tag vTag in viewAnnTags)
                        {
                            try
                            {
                                NXOpen.TaggedObject tObj = NXOpen.Utilities.NXObjectManager.Get(vTag);
                                if (tObj == null) continue;

                                var dispObj = tObj as DisplayableObject;
                                if (dispObj == null) continue;

                                bool isDim = tObj is NXOpen.Annotations.Dimension;

                                // pmiFontMapからTagで直接照合
                                PmiFontInfo info = null;
                                if (pmiFontMap.ContainsKey(vTag))
                                {
                                    info = pmiFontMap[vTag];
                                }

                                // Tag照合失敗時：Dimensionはテキスト値で照合
                                if (info == null && isDim)
                                {
                                    var dimAnn = tObj as NXOpen.Annotations.Annotation;
                                    if (dimAnn != null)
                                    {
                                        string dimText = GetAnnotationText(dimAnn);
                                        if (!string.IsNullOrEmpty(dimText))
                                        {
                                            foreach (var kv in pmiFontMap)
                                            {
                                                if (kv.Value.IsDimension &&
                                                    !string.IsNullOrEmpty(kv.Value.TextContent) &&
                                                    kv.Value.TextContent == dimText &&
                                                    kv.Value.DimensionTextSize > 0)
                                                { info = kv.Value; break; }
                                            }
                                        }
                                    }
                                }

                                // Tag照合失敗時：PmiNoteはrawText "PmiNote {modelTag}" で照合
                                if (info == null && !isDim)
                                {
                                    string rawText = tObj.ToString();
                                    if (rawText != null && rawText.StartsWith("PmiNote "))
                                    {
                                        string tagStr = rawText.Substring("PmiNote ".Length).Trim();
                                        ulong tagVal;
                                        if (ulong.TryParse(tagStr, out tagVal))
                                        {
                                            Tag modelTag = (Tag)tagVal;
                                            if (pmiNotesByTag.ContainsKey(modelTag))
                                                info = pmiNotesByTag[modelTag];
                                        }
                                    }
                                }

                                if (info == null) continue;

                                NXOpen.Annotations.EditSettingsBuilder esb = null;
                                try
                                {
                                    esb = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(
                                        new DisplayableObject[] { dispObj });
                                    var builders = new NXOpen.Drafting.BaseEditSettingsBuilder[] { esb };
                                    workPart.SettingsManager.ProcessForMultipleObjectsSettings(builders);

                                    bool changed = false;
                                    if (isDim && info.DimensionTextSize > 0)
                                    {
                                        esb.AnnotationStyle.LetteringStyle.DimensionTextSize = info.DimensionTextSize;
                                        changed = true;
                                    }
                                    if (info.GeneralTextSize > 0)
                                    {
                                        esb.AnnotationStyle.LetteringStyle.GeneralTextSize = info.GeneralTextSize;
                                        changed = true;
                                    }

                                    if (changed)
                                    {
                                        esb.Commit();
                                        noteCount++;
                                        theSession.ListingWindow.WriteLine(
                                            "  ViewAnn[Apply] type=" + tObj.GetType().Name +
                                            " tag=" + vTag +
                                            (isDim ? " dimSize=" + info.DimensionTextSize : "") +
                                            " genSize=" + info.GeneralTextSize);
                                    }
                                }
                                finally
                                {
                                    if (esb != null) try { esb.Destroy(); } catch { }
                                }
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        theSession.ListingWindow.WriteLine("  ViewAnn scan: 0 annotations (AskAnnotationsOfView returned empty)");
                    }
                }
                catch (Exception ex)
                {
                    theSession.ListingWindow.WriteLine("  ViewAnn scan error: " + ex.Message);
                }
            }

            theSession.ListingWindow.WriteLine(
                "=== DEBUG PMI: ApplyPmiFontSizes end, applied=" + noteCount + " ===");
        }

        private static PmiFontInfo FindBestMatch(
            Dictionary<Tag, PmiFontInfo> pmiFontMap, string inheritedType, string inheritedText)
        {
            // 第1優先: 型名＋テキスト内容の完全一致
            foreach (var entry in pmiFontMap)
            {
                if (entry.Value.PmiTypeName == inheritedType &&
                    entry.Value.TextContent == inheritedText)
                {
                    return entry.Value;
                }
            }

            // 第2優先: 型名＋テキスト内容の部分一致（Contains）
            if (!string.IsNullOrEmpty(inheritedText))
            {
                foreach (var entry in pmiFontMap)
                {
                    if (entry.Value.PmiTypeName == inheritedType &&
                        !string.IsNullOrEmpty(entry.Value.TextContent) &&
                        (entry.Value.TextContent.Contains(inheritedText) ||
                         inheritedText.Contains(entry.Value.TextContent)))
                    {
                        return entry.Value;
                    }
                }
            }

            // 第3優先: 同じ型の最初のエントリ（フォールバック）
            foreach (var entry in pmiFontMap)
            {
                if (entry.Value.PmiTypeName == inheritedType)
                {
                    return entry.Value;
                }
            }

            // 第4優先: 同じカテゴリ（Dimension/Note）のフォールバック
            bool isDim = inheritedType.Contains("Dimension") || inheritedType.Contains("dim");
            foreach (var entry in pmiFontMap)
            {
                if (entry.Value.IsDimension == isDim)
                {
                    return entry.Value;
                }
            }

            return null;
        }

        private static void ApplyFontSizeToDimension(
            NXOpen.Annotations.Dimension dim, PmiFontInfo fontInfo)
        {
            NXOpen.Annotations.EditSettingsBuilder esb = null;
            try
            {
                DisplayableObject[] objects = new DisplayableObject[] { dim };
                esb = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects);

                NXOpen.Drafting.BaseEditSettingsBuilder[] builders =
                    new NXOpen.Drafting.BaseEditSettingsBuilder[] { esb };
                workPart.SettingsManager.ProcessForMultipleObjectsSettings(builders);

                var letteringStyle = esb.AnnotationStyle.LetteringStyle;
                if (fontInfo.DimensionTextSize > 0)
                    letteringStyle.DimensionTextSize = fontInfo.DimensionTextSize;
                if (fontInfo.ToleranceTextSize > 0)
                    letteringStyle.ToleranceTextSize = fontInfo.ToleranceTextSize;
                if (fontInfo.AppendedTextSize > 0)
                    letteringStyle.AppendedTextSize = fontInfo.AppendedTextSize;

                esb.Commit();
            }
            catch { }
            finally
            {
                if (esb != null)
                    try { esb.Destroy(); } catch { }
            }
        }

        private static void ApplyFontSizeToNote(
            NXOpen.Annotations.Note note, PmiFontInfo fontInfo)
        {
            NXOpen.Annotations.EditSettingsBuilder esb = null;
            try
            {
                DisplayableObject[] objects = new DisplayableObject[] { (DisplayableObject)note };
                esb = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects);

                NXOpen.Drafting.BaseEditSettingsBuilder[] builders =
                    new NXOpen.Drafting.BaseEditSettingsBuilder[] { esb };
                workPart.SettingsManager.ProcessForMultipleObjectsSettings(builders);

                if (fontInfo.GeneralTextSize > 0)
                    esb.AnnotationStyle.LetteringStyle.GeneralTextSize = fontInfo.GeneralTextSize;

                esb.Commit();
            }
            catch { }
            finally
            {
                if (esb != null)
                    try { esb.Destroy(); } catch { }
            }
        }

        // ================================================================
        // ★【改修6】シート・ビューの確実な削除
        // ================================================================

        private static void DeleteSheetAndViews(DrawingSheet sheet)
        {
            if (sheet == null) return;

            try
            {
                // ★【改修4】ビュー・シート削除もFreeze/Unfreezeで囲む
                workPart.DraftingManager.DrawingsFreezeOutOfDateComputation();
                try
                {
                    // シート上の全ビューを削除してからシートも削除
                    Session.UndoMarkId delMark = default(Session.UndoMarkId);
                    try
                    {
                        delMark = theSession.SetUndoMark(
                            Session.MarkVisibility.Invisible, "DeleteSheet");

                        DraftingView[] views = sheet.GetDraftingViews();
                        if (views != null)
                        {
                            foreach (var dv in views)
                            {
                                try { theSession.UpdateManager.AddToDeleteList(dv); } catch { }
                            }
                        }

                        // シートも削除リストに追加
                        try { theSession.UpdateManager.AddToDeleteList(sheet); } catch { }

                        theSession.UpdateManager.DoUpdate(delMark);
                    }
                    catch { }
                    finally
                    {
                        SafeDeleteMark(ref delMark);
                    }
                }
                finally
                {
                    workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();
                }
            }
            catch { }
        }

        // ★【改修6】残存シートの最終チェック・クリーンアップ
        private static void CleanupResidualSheets(HashSet<Tag> originalSheetTags)
        {
            try
            {
                List<DrawingSheet> toDelete = new List<DrawingSheet>();
                foreach (DrawingSheet s in workPart.DrawingSheets)
                {
                    if (!originalSheetTags.Contains(s.Tag))
                        toDelete.Add(s);
                }
                foreach (DrawingSheet s in toDelete)
                {
                    DeleteSheetAndViews(s);
                }
            }
            catch { }
        }

        // ================================================================
        // ★【改修2】用紙サイズ自動選択（必要サイズからフィット検索）
        // ================================================================

        private static void GetOptimalPaperSize(double requiredWidth, double requiredHeight,
            out double sheetWidth, out double sheetHeight, out string paperName)
        {
            // A判定形から最小フィットを検索（横置きのみ）
            for (int i = 0; i < StandardPapers.Length; i++)
            {
                double pw = StandardPapers[i][0];
                double ph = StandardPapers[i][1];

                if (requiredWidth <= pw && requiredHeight <= ph)
                {
                    sheetWidth = pw;
                    sheetHeight = ph;
                    paperName = PaperNames[i];
                    return;
                }
            }

            // ★ A0でも収まらない場合: カスタムサイズ（スケールは変えない）
            // 50mm単位に切り上げて見栄えの良いサイズにする
            sheetWidth = Math.Ceiling(requiredWidth / 50.0) * 50.0;
            sheetHeight = Math.Ceiling(requiredHeight / 50.0) * 50.0;
            // 最低でもA0以上にする
            sheetWidth = Math.Max(sheetWidth, 1189.0);
            sheetHeight = Math.Max(sheetHeight, 841.0);
            paperName = "Custom";
        }

        // ================================================================
        // ★【改修5】UndoMark安全管理
        // ================================================================

        private static void SafeDeleteMark(ref Session.UndoMarkId id)
        {
            try { theSession.DeleteUndoMark(id, null); } catch { }
            id = default(Session.UndoMarkId);
        }

        // ================================================================
        // パート属性取得
        // ================================================================

        private static void GetPartAttributes(out string partNo, out string partRev)
        {
            partNo = "";
            partRev = "";

            try
            {
                NXObject.AttributeInformation[] attrs = workPart.GetUserAttributes();
                foreach (var attr in attrs)
                {
                    if (attr.Title.Equals("DB_PART_NO", StringComparison.OrdinalIgnoreCase))
                        partNo = attr.StringValue != null ? attr.StringValue.Trim() : "";
                    if (attr.Title.Equals("DB_PART_REV", StringComparison.OrdinalIgnoreCase))
                        partRev = attr.StringValue != null ? attr.StringValue.Trim() : "";
                }
            }
            catch { }

            if (string.IsNullOrEmpty(partNo))
            {
                try { partNo = workPart.Leaf; }
                catch { partNo = Path.GetFileNameWithoutExtension(workPart.FullPath); }
            }
        }

        // ================================================================
        // カスタムビュー取得
        // ================================================================

        private static List<ModelingView> GetCustomViews()
        {
            List<ModelingView> customViews = new List<ModelingView>();
            foreach (ModelingView view in workPart.ModelingViews)
            {
                if (!StandardViewNames.Contains(view.Name))
                    customViews.Add(view);
            }
            return customViews;
        }

        // ================================================================
        // ビュー名ソート（数字→英字→その他）
        // ================================================================

        private static int CompareViewNames(string a, string b)
        {
            int ga = GetSortGroup(a);
            int gb = GetSortGroup(b);
            if (ga != gb) return ga.CompareTo(gb);
            return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
        }

        private static int GetSortGroup(string name)
        {
            if (string.IsNullOrEmpty(name)) return 3;
            char c = name[0];
            if (c >= '0' && c <= '9') return 0;
            if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')) return 1;
            return 2;
        }

        // ================================================================
        // 出力フォルダ作成
        // ================================================================

        private static string CreateOutputFolder(string partNo)
        {
            string basePath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            string dateSuffix = DateTime.Now.ToString("yyMMdd");
            string baseName = string.Format("{0}_{1}", partNo, dateSuffix);
            string folder = Path.Combine(basePath, baseName);

            if (!Directory.Exists(folder))
                return folder;

            int n = 1;
            while (Directory.Exists(Path.Combine(basePath, string.Format("{0}_({1})", baseName, n))))
                n++;
            return Path.Combine(basePath, string.Format("{0}_({1})", baseName, n));
        }

        // ================================================================
        // DEFファイル展開
        // ================================================================

        private static string ExtractDefFile()
        {
            string path = Path.Combine(Path.GetTempPath(),
                "nxcvde_" + Guid.NewGuid().ToString("N") + ".def");
            File.WriteAllText(path, EmbeddedDefContent);
            return path;
        }

        // ================================================================
        // ログファイル出力
        // ================================================================

        private static void WriteLogFile(string outputFolder, string partNo,
            List<string> entries, int success, int fail, int cancel, double totalSeconds)
        {
            try
            {
                string logPath = Path.Combine(outputFolder,
                    string.Format("{0}_export_log.txt", partNo));
                var lines = new List<string>();
                lines.Add("NXCustomViewDxfExporter - Export Log");
                lines.Add(string.Format("Date: {0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                lines.Add(string.Format("Part: {0}", partNo));
                lines.Add(string.Format("Output: {0}", outputFolder));
                lines.Add("");
                lines.Add("Views:");
                lines.AddRange(entries);
                lines.Add("");
                lines.Add(string.Format("Summary: {0} success, {1} failed, {2} cancelled ({3:F1}s)",
                    success, fail, cancel, totalSeconds));
                File.WriteAllLines(logPath, lines.ToArray());
            }
            catch { }
        }

        // ================================================================
        // ファイルユーティリティ
        // ================================================================

        private static string SanitizeFileName(string name)
        {
            char[] invalid = Path.GetInvalidFileNameChars();
            string result = string.Join("_", name.Split(invalid));
            while (result.Contains("__"))
                result = result.Replace("__", "_");
            return result.Trim('_');
        }

        private static string GetUniqueFilePath(string folder, string baseName, string extension)
        {
            string path = Path.Combine(folder, baseName + extension);
            if (!File.Exists(path)) return path;

            int counter = 1;
            while (true)
            {
                path = Path.Combine(folder, string.Format("{0}_({1}){2}", baseName, counter, extension));
                if (!File.Exists(path)) return path;
                counter++;
            }
        }

        private static void CleanupTempFiles()
        {
            if (!string.IsNullOrEmpty(tempDefPath))
            {
                try { if (File.Exists(tempDefPath)) File.Delete(tempDefPath); } catch { }
            }
        }

        // ================================================================
        // 言語・メッセージ
        // ================================================================

        private static void InitializeLanguage()
        {
            string lang = Environment.GetEnvironmentVariable("UGII_LANG");
            if (!string.IsNullOrEmpty(lang))
                isJapanese = lang.Equals("japanese", StringComparison.OrdinalIgnoreCase);
            else
                isJapanese = System.Globalization.CultureInfo.CurrentUICulture.Name.StartsWith("ja");

            messagesJa = MessageStrings.CreateJaDictionary();
            messagesEn = MessageStrings.CreateEnDictionary();
        }

        private static string GetMsg(string key)
        {
            var dict = isJapanese ? messagesJa : messagesEn;
            return dict.ContainsKey(key) ? dict[key] : key;
        }

        private static string GetMsg(string key, params object[] args)
        {
            return string.Format(GetMsg(key), args);
        }

        private static Font GetUIFont(float size, FontStyle style = FontStyle.Regular)
        {
            if (isJapanese)
            {
                string[] jaFonts = { "Meiryo UI", "Yu Gothic UI", "MS UI Gothic" };
                foreach (string name in jaFonts)
                {
                    try
                    {
                        Font font = new Font(name, size, style);
                        if (font.Name == name) return font;
                        font.Dispose();
                    }
                    catch { }
                }
            }
            return new Font("Segoe UI", size, style);
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(message, GetMsg("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // ================================================================
        // 進捗ダイアログ
        // ================================================================

        private class ProgressForm : Form
        {
            private System.Windows.Forms.Label labelProgress;
            private ProgressBar progressBar;
            private System.Windows.Forms.Label labelPercent;
            private Button btnStop;
            private System.Windows.Forms.Timer foregroundTimer;
            private volatile bool _stopRequested;
            private bool _allowClose;

            public bool StopRequested { get { return _stopRequested; } }

            public ProgressForm(int totalViews)
            {
                _stopRequested = false;
                _allowClose = false;

                this.Text = GetMsg("ProgressTitle");
                this.Width = 400;
                this.Height = 200;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.TopMost = true;
                this.ShowInTaskbar = true;
                this.Font = GetUIFont(9f);

                this.FormClosing += (s, e) =>
                {
                    if (!_allowClose && e.CloseReason == CloseReason.UserClosing)
                        e.Cancel = true;
                };

                labelProgress = new System.Windows.Forms.Label();
                labelProgress.Text = GetMsg("ProgressMessage", "...", 0, totalViews);
                labelProgress.Location = new System.Drawing.Point(20, 20);
                labelProgress.Size = new System.Drawing.Size(350, 25);
                labelProgress.Font = GetUIFont(10f);
                this.Controls.Add(labelProgress);

                progressBar = new ProgressBar();
                progressBar.Location = new System.Drawing.Point(20, 55);
                progressBar.Size = new System.Drawing.Size(280, 25);
                progressBar.Minimum = 0;
                progressBar.Maximum = totalViews;
                this.Controls.Add(progressBar);

                labelPercent = new System.Windows.Forms.Label();
                labelPercent.Text = "0%";
                labelPercent.Location = new System.Drawing.Point(310, 58);
                labelPercent.Size = new System.Drawing.Size(50, 25);
                labelPercent.Font = GetUIFont(10f);
                this.Controls.Add(labelPercent);

                btnStop = new Button();
                btnStop.Text = GetMsg("StopButton");
                btnStop.Location = new System.Drawing.Point(140, 100);
                btnStop.Size = new System.Drawing.Size(120, 40);
                btnStop.BackColor = Color.FromArgb(220, 53, 69);
                btnStop.ForeColor = Color.White;
                btnStop.FlatStyle = FlatStyle.Flat;
                btnStop.Font = GetUIFont(10f, FontStyle.Bold);
                btnStop.Click += (s, e) =>
                {
                    _stopRequested = true;
                    btnStop.Enabled = false;
                    btnStop.Text = GetMsg("StopRequesting");
                };
                this.Controls.Add(btnStop);

                foregroundTimer = new System.Windows.Forms.Timer();
                foregroundTimer.Interval = 200;
                foregroundTimer.Tick += (s, e) =>
                {
                    if (this.Visible && !this.IsDisposed)
                    {
                        SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                    }
                };
                foregroundTimer.Start();
            }

            public void UpdateProgress(string viewName, int current, int total)
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(() => UpdateProgress(viewName, current, total)));
                    return;
                }
                labelProgress.Text = GetMsg("ProgressMessage", viewName, current, total);
                progressBar.Maximum = total;
                progressBar.Value = current;
                labelPercent.Text = string.Format("{0}%", (int)((double)current / total * 100));
            }

            public void ForceClose()
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(ForceClose));
                    return;
                }
                if (foregroundTimer != null)
                {
                    foregroundTimer.Stop();
                    foregroundTimer.Dispose();
                    foregroundTimer = null;
                }
                _allowClose = true;
                Close();
            }
        }

        // ================================================================
        // 完了ダイアログ
        // ================================================================

        private class CompletionForm : Form
        {
            public CompletionForm(int successCount, int failCount, int cancelCount,
                List<string> errors, string outputPath)
            {
                bool hasErrors = failCount > 0;

                this.Text = hasErrors ? GetMsg("CompleteTitleError") : GetMsg("CompleteTitleSuccess");
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.TopMost = true;
                this.ShowInTaskbar = true;
                this.Font = GetUIFont(9f);

                int y = 20;

                System.Windows.Forms.Label labelStatus = new System.Windows.Forms.Label();
                labelStatus.Text = (hasErrors ? "\u26A0 " : "\u2713 ") +
                    (hasErrors ? GetMsg("CompleteMessageError") : GetMsg("CompleteMessage"));
                labelStatus.Location = new System.Drawing.Point(20, y);
                labelStatus.Size = new System.Drawing.Size(400, 30);
                labelStatus.Font = GetUIFont(11f, FontStyle.Bold);
                this.Controls.Add(labelStatus);
                y += 40;

                System.Windows.Forms.Label labelSuccess = new System.Windows.Forms.Label();
                labelSuccess.Text = GetMsg("SuccessCount", successCount);
                labelSuccess.Location = new System.Drawing.Point(30, y);
                labelSuccess.Size = new System.Drawing.Size(380, 22);
                labelSuccess.Font = GetUIFont(10f);
                this.Controls.Add(labelSuccess);
                y += 25;

                if (failCount > 0)
                {
                    System.Windows.Forms.Label lbl = new System.Windows.Forms.Label();
                    lbl.Text = GetMsg("FailureCount", failCount);
                    lbl.Location = new System.Drawing.Point(30, y);
                    lbl.Size = new System.Drawing.Size(380, 22);
                    lbl.Font = GetUIFont(10f);
                    this.Controls.Add(lbl);
                    y += 25;
                }

                if (cancelCount > 0)
                {
                    System.Windows.Forms.Label lbl = new System.Windows.Forms.Label();
                    lbl.Text = GetMsg("CancelCount", cancelCount);
                    lbl.Location = new System.Drawing.Point(30, y);
                    lbl.Size = new System.Drawing.Size(380, 22);
                    lbl.Font = GetUIFont(10f);
                    this.Controls.Add(lbl);
                    y += 25;
                }

                System.Windows.Forms.Label labelOutput = new System.Windows.Forms.Label();
                labelOutput.Text = GetMsg("OutputPath", outputPath);
                labelOutput.Location = new System.Drawing.Point(30, y);
                labelOutput.MaximumSize = new System.Drawing.Size(380, 0);
                labelOutput.AutoSize = true;
                labelOutput.Font = GetUIFont(9f);
                this.Controls.Add(labelOutput);
                y += Math.Max(labelOutput.PreferredHeight, 22) + 8;

                if (errors != null && errors.Count > 0)
                {
                    System.Windows.Forms.Label hdr = new System.Windows.Forms.Label();
                    hdr.Text = GetMsg("ErrorDetails");
                    hdr.Location = new System.Drawing.Point(30, y);
                    hdr.Size = new System.Drawing.Size(380, 22);
                    hdr.Font = GetUIFont(9f, FontStyle.Bold);
                    this.Controls.Add(hdr);
                    y += 25;

                    foreach (string error in errors)
                    {
                        System.Windows.Forms.Label errLbl = new System.Windows.Forms.Label();
                        errLbl.Text = "  " + error;
                        errLbl.Location = new System.Drawing.Point(30, y);
                        errLbl.Size = new System.Drawing.Size(380, 40);
                        errLbl.Font = GetUIFont(9f);
                        this.Controls.Add(errLbl);
                        y += 40;
                    }
                }

                y += 10;

                Button btnOk = new Button();
                btnOk.Text = GetMsg("OkButton");
                btnOk.Location = new System.Drawing.Point(170, y);
                btnOk.Size = new System.Drawing.Size(100, 35);
                btnOk.Font = GetUIFont(10f);
                btnOk.Click += (s, e) => { this.DialogResult = DialogResult.OK; this.Close(); };
                this.Controls.Add(btnOk);
                this.AcceptButton = btnOk;

                y += 55;
                this.ClientSize = new System.Drawing.Size(430, y);
            }
        }

        // ================================================================
        // ビュー選択ダイアログ
        // ================================================================

        private class ViewSelectionForm : Form
        {
            private CheckBox chkAll;
            private CheckedListBox checkedListBox;
            private Button btnOk;
            private Button btnCancel;
            private bool _updatingAll;

            public List<int> SelectedIndices { get; private set; }

            public ViewSelectionForm(List<ModelingView> views)
            {
                SelectedIndices = new List<int>();
                _updatingAll = false;

                this.Text = GetMsg("ViewSelectTitle");
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.ShowInTaskbar = true;

                Font viewFont;
                Font viewFontBold;
                var testFont = new Font("Meiryo UI", 10f);
                if (testFont.Name == "Meiryo UI")
                {
                    viewFont = testFont;
                    viewFontBold = new Font("Meiryo UI", 10f, FontStyle.Bold);
                }
                else
                {
                    testFont.Dispose();
                    viewFont = new Font("MS UI Gothic", 10f);
                    viewFontBold = new Font("MS UI Gothic", 10f, FontStyle.Bold);
                }
                this.Font = viewFont;

                int y = 15;

                chkAll = new CheckBox();
                chkAll.Text = GetMsg("ViewSelectAll");
                chkAll.Location = new System.Drawing.Point(20, y);
                chkAll.Size = new System.Drawing.Size(350, 25);
                chkAll.Font = viewFontBold;
                chkAll.Checked = true;
                chkAll.CheckedChanged += ChkAll_CheckedChanged;
                this.Controls.Add(chkAll);
                y += 30;

                checkedListBox = new CheckedListBox();
                checkedListBox.Location = new System.Drawing.Point(20, y);
                checkedListBox.Font = new Font("Meiryo UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
                checkedListBox.CheckOnClick = true;
                checkedListBox.DrawMode = DrawMode.OwnerDrawFixed;
                checkedListBox.ItemHeight = 20;
                checkedListBox.DrawItem += CheckedListBox_DrawItem;

                for (int i = 0; i < views.Count; i++)
                    checkedListBox.Items.Add(views[i].Name, true);

                int listHeight = Math.Min(views.Count * 22 + 5, 300);
                checkedListBox.Size = new System.Drawing.Size(350, listHeight);
                checkedListBox.Enabled = false;
                checkedListBox.ItemCheck += CheckedListBox_ItemCheck;
                this.Controls.Add(checkedListBox);
                y += listHeight + 15;

                btnOk = new Button();
                btnOk.Text = GetMsg("OkButton");
                btnOk.Location = new System.Drawing.Point(120, y);
                btnOk.Size = new System.Drawing.Size(90, 35);
                btnOk.Font = viewFont;
                btnOk.Click += BtnOk_Click;
                this.Controls.Add(btnOk);
                this.AcceptButton = btnOk;

                btnCancel = new Button();
                btnCancel.Text = GetMsg("CancelButton");
                btnCancel.Location = new System.Drawing.Point(220, y);
                btnCancel.Size = new System.Drawing.Size(90, 35);
                btnCancel.Font = viewFont;
                btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
                this.Controls.Add(btnCancel);
                this.CancelButton = btnCancel;

                y += 55;
                this.ClientSize = new System.Drawing.Size(390, y);
            }

            private void BtnOk_Click(object sender, EventArgs e)
            {
                SelectedIndices.Clear();
                if (chkAll.Checked)
                {
                    for (int i = 0; i < checkedListBox.Items.Count; i++)
                        SelectedIndices.Add(i);
                }
                else
                {
                    for (int i = 0; i < checkedListBox.Items.Count; i++)
                    {
                        if (checkedListBox.GetItemChecked(i))
                            SelectedIndices.Add(i);
                    }
                }
                this.DialogResult = DialogResult.OK;
                this.Close();
            }

            private void ChkAll_CheckedChanged(object sender, EventArgs e)
            {
                _updatingAll = true;
                if (chkAll.Checked)
                {
                    checkedListBox.Enabled = false;
                    for (int i = 0; i < checkedListBox.Items.Count; i++)
                        checkedListBox.SetItemChecked(i, true);
                }
                else
                {
                    checkedListBox.Enabled = true;
                }
                UpdateOkButton();
                _updatingAll = false;
            }

            private void CheckedListBox_ItemCheck(object sender, ItemCheckEventArgs e)
            {
                if (_updatingAll) return;

                int checkedCount = checkedListBox.CheckedItems.Count;
                if (e.NewValue == CheckState.Checked) checkedCount++;
                else if (e.NewValue == CheckState.Unchecked) checkedCount--;

                if (checkedCount == checkedListBox.Items.Count)
                {
                    this.BeginInvoke(new Action(() =>
                    {
                        if (!chkAll.Checked) chkAll.Checked = true;
                    }));
                }

                btnOk.Enabled = checkedCount > 0;
            }

            private void UpdateOkButton()
            {
                if (chkAll.Checked) { btnOk.Enabled = true; return; }
                bool any = false;
                for (int i = 0; i < checkedListBox.Items.Count; i++)
                {
                    if (checkedListBox.GetItemChecked(i)) { any = true; break; }
                }
                btnOk.Enabled = any;
            }

            private void CheckedListBox_DrawItem(object sender, DrawItemEventArgs e)
            {
                if (e.Index < 0) return;
                var clb = (CheckedListBox)sender;

                e.DrawBackground();
                e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

                int checkWidth = 16;
                var textBounds = new RectangleF(
                    e.Bounds.X + checkWidth, e.Bounds.Y,
                    e.Bounds.Width - checkWidth, e.Bounds.Height);

                bool isChecked = clb.GetItemChecked(e.Index);
                CheckBoxRenderer.DrawCheckBox(e.Graphics,
                    new System.Drawing.Point(e.Bounds.X + 1, e.Bounds.Y + (e.Bounds.Height - 13) / 2),
                    isChecked
                        ? System.Windows.Forms.VisualStyles.CheckBoxState.CheckedNormal
                        : System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal);

                Color textColor = (e.State & DrawItemState.Selected) != 0
                    ? SystemColors.HighlightText : clb.ForeColor;

                using (var brush = new SolidBrush(textColor))
                {
                    e.Graphics.DrawString(clb.Items[e.Index].ToString(), clb.Font, brush, textBounds);
                }

                e.DrawFocusRectangle();
            }
        }
    }

    // ================================================================
    // メッセージ文字列定数クラス（日英対応）
    // ================================================================

    internal static class MessageStrings
    {
        public static Dictionary<string, string> CreateJaDictionary()
        {
            return new Dictionary<string, string>
            {
                { "Error", "エラー" },
                { "ErrorNoPartOpen", "パートファイルが開かれていません。" },
                { "ErrorNoCustomView", "カスタムビューが見つかりません。" },
                { "UnexpectedError", "予期しないエラーが発生しました:\n{0}" },
                { "ProgressTitle", "DXFエクスポート中..." },
                { "ProgressMessage", "処理中: {0}（{1}/{2}）" },
                { "StopButton", "停止" },
                { "StopRequesting", "停止中..." },
                { "ViewSelectTitle", "ビュー選択" },
                { "ViewSelectAll", "すべて選択" },
                { "CancelButton", "キャンセル" },
                { "OkButton", "OK" },
                { "CompleteTitleSuccess", "DXFエクスポート完了" },
                { "CompleteTitleError", "DXFエクスポート完了（一部エラー）" },
                { "CompleteMessage", "DXFエクスポートが完了しました" },
                { "CompleteMessageError", "DXFエクスポートが完了しました（一部エラー）" },
                { "SuccessCount", "成功: {0} ファイル" },
                { "FailureCount", "失敗: {0} ファイル" },
                { "CancelCount", "中断: {0} ファイル（ユーザー停止）" },
                { "OutputPath", "出力先: {0}" },
                { "ErrorDetails", "エラー詳細:" },
            };
        }

        public static Dictionary<string, string> CreateEnDictionary()
        {
            return new Dictionary<string, string>
            {
                { "Error", "Error" },
                { "ErrorNoPartOpen", "No part file is open." },
                { "ErrorNoCustomView", "No custom views found." },
                { "UnexpectedError", "An unexpected error occurred:\n{0}" },
                { "ProgressTitle", "Exporting DXF..." },
                { "ProgressMessage", "Processing: {0} ({1}/{2})" },
                { "StopButton", "Stop" },
                { "StopRequesting", "Stopping..." },
                { "ViewSelectTitle", "Select Views" },
                { "ViewSelectAll", "Select All" },
                { "CancelButton", "Cancel" },
                { "OkButton", "OK" },
                { "CompleteTitleSuccess", "DXF Export Complete" },
                { "CompleteTitleError", "DXF Export Complete (with errors)" },
                { "CompleteMessage", "DXF export completed successfully" },
                { "CompleteMessageError", "DXF export completed with some errors" },
                { "SuccessCount", "Success: {0} files" },
                { "FailureCount", "Failed: {0} files" },
                { "CancelCount", "Cancelled: {0} files (user stopped)" },
                { "OutputPath", "Output: {0}" },
                { "ErrorDetails", "Error details:" },
            };
        }
    }
}
