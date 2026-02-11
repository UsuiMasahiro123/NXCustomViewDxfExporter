using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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

        // 一時defファイルのパス（処理全体で共有し、終了時に削除）
        private static string tempDefPath;

        // Tempにコピーしたパートファイルのパス（終了時に削除）
        private static readonly List<string> tempPartCopies = new List<string>();

        // 停止ボタン用フラグ
        private static volatile bool cancelRequested;
        private static Form cancelForm;

        // dxfdwg.def 埋め込み内容（DLL経由エクスポート時の設定ズレを防止）
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
            "EXPORT_DIMENSIONS_AS =REAL\r\n" +
            "EXPORT_DRAWING_USING_CGM =NO\r\n" +
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
            "UGI_LOAD_OPTION = Load From Search Dirs\r\n" +
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

                // パートが保存済み（ディスク上に存在）であることを確認
                if (string.IsNullOrEmpty(workPart.FullPath))
                {
                    theUI.NXMessageBox.Show("エラー", NXMessageBox.DialogType.Error,
                        "パートファイルが保存されていません。\n先にパートを保存してから実行してください。");
                    return;
                }

                // CGMダイアログ抑制（処理開始前に設定）
                SuppressCgmDialog();

                // パート名（拡張子なし）とディレクトリを取得
                string partName = Path.GetFileNameWithoutExtension(workPart.FullPath);
                string partDir = Path.GetDirectoryName(workPart.FullPath);
                lw.WriteLine("========================================");
                lw.WriteLine("  カスタムビュー DXF エクスポーター");
                lw.WriteLine("========================================");
                lw.WriteLine("パート名: " + partName);
                lw.WriteLine("パートDir: " + partDir);

                // 2. パートファイル検索パスの設定（Failed to retrieve file 対策）
                SetupPartSearchPaths(partDir);

                // 3. dxfdwg.def 一時ファイルの生成（{PART_DIR}をパートディレクトリに置換）
                string defContent = EmbeddedDefContent.Replace("{PART_DIR}", partDir);
                tempDefPath = Path.Combine(Path.GetTempPath(), "nxdxf_export_settings.def");
                File.WriteAllText(tempDefPath, defContent);
                lw.WriteLine("DEF設定ファイル: " + tempDefPath);

                // 4. パートファイルをTempディレクトリにコピー（トランスレータが確実に参照可能にする）
                CopyPartFilesToTemp(workPart.FullPath);

                // 5. 出力フォルダの選択
                string outputFolder = SelectOutputFolder();
                if (outputFolder == null)
                {
                    lw.WriteLine("フォルダ選択がキャンセルされました。");
                    return;
                }
                lw.WriteLine("出力先:   " + outputFolder);

                // 6. カスタムビューの取得
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

                // 7. 製図アプリケーションに切り替え（ループ外で1回だけ実行）
                theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING");

                // 8. バックグラウンド処理の開始
                IntPtr userWindow = GetForegroundWindow();

                try
                {
                    theUFSession.Disp.SetDisplay(NXOpen.UF.UFConstants.UF_DISP_SUPPRESS_DISPLAY);
                    lw.WriteLine("[バックグラウンド] 画面更新を抑制しました");
                }
                catch (Exception ex)
                {
                    lw.WriteLine("[バックグラウンド] 画面更新抑制失敗（続行）: " + ex.Message);
                }

                try
                {
                    LockSetForegroundWindow(LSFW_LOCK);
                    lw.WriteLine("[バックグラウンド] フォーカスロックを設定しました");
                }
                catch { }

                // 9. 停止ボタンを表示
                ShowCancelForm(customViews.Count);

                try
                {
                    // 10. カスタムビューごとのループ処理
                    for (int i = 0; i < customViews.Count; i++)
                    {
                        // 停止ボタンが押されたかチェック
                        if (cancelRequested)
                        {
                            lw.WriteLine("\n★ ユーザーにより処理が中止されました。");
                            break;
                        }

                        ModelingView view = customViews[i];
                        lw.WriteLine(string.Format("\n[{0}/{1}] 処理中: {2}", i + 1, customViews.Count, view.Name));

                        // 停止フォームの進捗を更新
                        UpdateCancelFormProgress(i + 1, customViews.Count, view.Name);

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
                    // 停止フォームを閉じる
                    CloseCancelForm();

                    // 画面更新を復帰（エラー時も確実に復帰）
                    try
                    {
                        theUFSession.Disp.SetDisplay(NXOpen.UF.UFConstants.UF_DISP_UNSUPPRESS_DISPLAY);
                        lw.WriteLine("[バックグラウンド] 画面更新を復帰しました");
                    }
                    catch (Exception ex)
                    {
                        lw.WriteLine("[バックグラウンド] 画面更新復帰失敗: " + ex.Message);
                    }

                    // フォーカスロック解除（エラー時も確実に解除）
                    try
                    {
                        LockSetForegroundWindow(LSFW_UNLOCK);
                        lw.WriteLine("[バックグラウンド] フォーカスロックを解除しました");
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

                // 11. モデリングアプリケーションに戻る
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
                if (cancelRequested)
                    lw.WriteLine("※ ユーザーにより中止");
                lw.WriteLine(string.Format("合計処理時間: {0:F1}秒", totalSw.Elapsed.TotalSeconds));
                lw.WriteLine("========================================");

                string message = string.Format("DXFエクスポート完了: {0}ファイル出力しました", successCount);
                if (failCount > 0)
                {
                    message += string.Format("\n失敗: {0}件", failCount);
                }
                if (cancelRequested)
                {
                    message += "\n（ユーザーにより中止）";
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
                // 停止フォームが残っていたら閉じる
                CloseCancelForm();

                // UndoMarkまで戻す（図面シート等の変更を取り消す）
                theSession.UndoToMark(undoMark, null);

                // CGM環境変数を元に戻す
                try
                {
                    Environment.SetEnvironmentVariable("UGII_CGM_FITS_FILE_SAVE", originalCgmEnv);
                }
                catch { }

                // 一時defファイルを削除
                try
                {
                    if (tempDefPath != null && File.Exists(tempDefPath))
                        File.Delete(tempDefPath);
                }
                catch { }

                // Tempにコピーしたパートファイルを削除
                CleanupTempPartCopies();

                // パートは絶対に保存しない（保存するとCGMダイアログの原因になる）
            }
        }

        // ================================================================
        // パートファイル検索パスの設定
        // ================================================================

        private static void SetupPartSearchPaths(string partDir)
        {
            // 対策1: 既存の検索ディレクトリにパートDirを追加（上書きではなく追加）
            try
            {
                string[] existingDirs;
                bool[] existingFlags;
                theSession.Parts.LoadOptions.GetSearchDirectories(out existingDirs, out existingFlags);

                List<string> allDirs = new List<string>();
                List<bool> allFlags = new List<bool>();

                // パートDirを先頭に追加（優先度最高）
                allDirs.Add(partDir);
                allFlags.Add(true); // サブディレクトリも検索

                // 既存のディレクトリを追加
                if (existingDirs != null)
                {
                    for (int i = 0; i < existingDirs.Length; i++)
                    {
                        if (!string.Equals(existingDirs[i], partDir, StringComparison.OrdinalIgnoreCase))
                        {
                            allDirs.Add(existingDirs[i]);
                            allFlags.Add(existingFlags[i]);
                        }
                    }
                }

                theSession.Parts.LoadOptions.SetSearchDirectories(
                    allDirs.ToArray(), allFlags.ToArray());
                lw.WriteLine("[検索パス] LoadOptions.SetSearchDirectories 設定済み (" + allDirs.Count + "件)");
            }
            catch (Exception ex)
            {
                lw.WriteLine("[検索パス] LoadOptions設定失敗（続行）: " + ex.Message);
            }

            // 対策2: UF_ASSEMレベルで検索ディレクトリを設定
            try
            {
                theUFSession.Assem.SetSearchDirectories(
                    1,
                    new string[] { partDir },
                    new bool[] { true });
                lw.WriteLine("[検索パス] UF_ASSEM.SetSearchDirectories 設定済み");
            }
            catch (Exception ex)
            {
                lw.WriteLine("[検索パス] UF_ASSEM設定失敗（続行）: " + ex.Message);
            }
        }

        /// <summary>
        /// パートファイルとその関連ファイルをTempディレクトリにコピー。
        /// DXFトランスレータは%TEMP%と%TEMP%\2\に一時パートを作成し、
        /// 元パートをファイル名のみで参照する。defの検索パス設定が効かない場合の
        /// フォールバックとして、参照先にパートファイルを直接配置する。
        /// </summary>
        private static void CopyPartFilesToTemp(string partFullPath)
        {
            string partFileName = Path.GetFileName(partFullPath);
            string tempDir = Path.GetTempPath();

            // %TEMP% にコピー
            CopySinglePartToDir(partFullPath, tempDir, partFileName);

            // %TEMP%\2\ にコピー（DXFトランスレータが使うサブディレクトリ）
            string tempDir2 = Path.Combine(tempDir, "2");
            if (Directory.Exists(tempDir2))
            {
                CopySinglePartToDir(partFullPath, tempDir2, partFileName);
            }
        }

        private static void CopySinglePartToDir(string srcPath, string destDir, string fileName)
        {
            try
            {
                string destPath = Path.Combine(destDir, fileName);
                File.Copy(srcPath, destPath, true);
                tempPartCopies.Add(destPath);
                lw.WriteLine("[検索パス] パートファイルをコピー: " + destPath);
            }
            catch (Exception ex)
            {
                lw.WriteLine("[検索パス] パートコピー失敗（続行）: " + destDir + " - " + ex.Message);
            }
        }

        private static void CleanupTempPartCopies()
        {
            foreach (string path in tempPartCopies)
            {
                try
                {
                    if (File.Exists(path))
                        File.Delete(path);
                }
                catch { }
            }
            tempPartCopies.Clear();
        }

        // ================================================================
        // 停止ボタン（別スレッドのモードレスフォーム）
        // ================================================================

        private static void ShowCancelForm(int totalViews)
        {
            cancelRequested = false;
            cancelForm = null;

            Thread uiThread = new Thread(() =>
            {
                Form form = new Form();
                form.Text = "DXFエクスポート";
                form.Width = 340;
                form.Height = 130;
                form.FormBorderStyle = FormBorderStyle.FixedToolWindow;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.TopMost = true;

                Label label = new Label();
                label.Text = string.Format("0/{0} 処理中...", totalViews);
                label.Dock = DockStyle.Top;
                label.Height = 40;
                label.Padding = new Padding(10, 10, 10, 0);
                form.Controls.Add(label);

                Button btn = new Button();
                btn.Text = "処理を停止";
                btn.Dock = DockStyle.Bottom;
                btn.Height = 40;
                btn.Click += (s, e) =>
                {
                    cancelRequested = true;
                    btn.Enabled = false;
                    btn.Text = "停止中...";
                    label.Text = "現在のビュー処理完了後に停止します...";
                };
                form.Controls.Add(btn);

                cancelForm = form;
                Application.Run(form);
            });
            uiThread.SetApartmentState(ApartmentState.STA);
            uiThread.IsBackground = true;
            uiThread.Start();

            // フォームが初期化されるまで少し待つ
            Thread.Sleep(300);
        }

        private static void UpdateCancelFormProgress(int current, int total, string viewName)
        {
            try
            {
                Form form = cancelForm;
                if (form != null && form.IsHandleCreated && !form.IsDisposed)
                {
                    form.Invoke(new Action(() =>
                    {
                        if (form.Controls.Count > 0)
                        {
                            // Labelはindex 0（最初に追加したコントロール）
                            form.Controls[0].Text = string.Format(
                                "[{0}/{1}] {2}", current, total, viewName);
                        }
                    }));
                }
            }
            catch { }
        }

        private static void CloseCancelForm()
        {
            try
            {
                Form form = cancelForm;
                if (form != null && form.IsHandleCreated && !form.IsDisposed)
                {
                    form.Invoke(new Action(() => form.Close()));
                }
                cancelForm = null;
            }
            catch { }
        }

        // ================================================================
        // フォルダ選択
        // ================================================================

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

        // ================================================================
        // カスタムビューのフィルタリング
        // ================================================================

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

        // ================================================================
        // 個別ビューのエクスポート処理
        // ================================================================

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

                // ★ 最重要: 埋め込みdef設定を最初に読み込む（他のプロパティ設定より前）
                dxfCreator.SettingsFile = tempDefPath;

                // ★ 入力ファイルを明示的に指定（トランスレータがパートを見つけられるようにする）
                dxfCreator.InputFile = workPart.FullPath;

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
        /// </summary>
        private static void CleanupViewAndSheet(NXObject baseViewObj, DrawingSheet sheet, Session.UndoMarkId undoMark)
        {
            try
            {
                theSession.UndoToMark(undoMark, null);
            }
            catch
            {
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

        // ================================================================
        // CGMダイアログ抑制
        // ================================================================

        private static void SuppressCgmDialog()
        {
            Environment.SetEnvironmentVariable("UGII_CGM_FITS_FILE_SAVE", "0");

            try
            {
                workPart.SaveOptions.DrawingCgmData = false;
            }
            catch { }
        }

        // ================================================================
        // ビューサイズ・シートサイズ
        // ================================================================

        private static void GetViewBounds(out double width, out double height)
        {
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

                    double[] dims = new double[] { dx, dy, dz };
                    Array.Sort(dims);
                    width = dims[2];
                    height = dims[1];

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

            sheetWidth = SheetSizes[SheetSizes.Length - 1][0];
            sheetHeight = SheetSizes[SheetSizes.Length - 1][1];
        }

        public static int GetUnloadOption(string dummy)
        {
            return (int)Session.LibraryUnloadOption.Immediately;
        }
    }
}
