using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using NXOpen;
using NXOpen.Drawings;
using NXOpen.UF;

namespace NXCustomViewDxfExporter
{
    public class CustomViewDxfExporter
    {
        // Win32 API: フォーカス制御・ウィンドウ制御
        [DllImport("user32.dll")]
        private static extern bool LockSetForegroundWindow(uint uLockCode);
        private const uint LSFW_LOCK = 1;
        private const uint LSFW_UNLOCK = 2;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_MINIMIZE = 6;
        private const int SW_RESTORE = 9;

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        private const int SW_HIDE = 0;
        private const int SW_SHOWNOACTIVATE = 8;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int X, Y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct WINDOWPLACEMENT
        {
            public uint length;
            public uint flags;
            public uint showCmd;
            public POINT ptMinPosition;
            public POINT ptMaxPosition;
            public RECT rcNormalPosition;
        }

        private static Session theSession;
        private static UFSession theUFSession;
        private static UI theUI;
        private static Part workPart;
        private static ListingWindow lw;

        // 一時defファイルのパス（処理全体で共有し、終了時に削除）
        private static string tempDefPath;

        // Tempにコピーしたパートファイルのパス（終了時に削除）
        private static readonly List<string> tempPartCopies = new List<string>();

        // NXプロセスのウィンドウハンドル一覧と元の配置（保存・復元用）
        private static readonly List<IntPtr> nxWindows = new List<IntPtr>();
        private static readonly Dictionary<IntPtr, WINDOWPLACEMENT> _savedPlacements = new Dictionary<IntPtr, WINDOWPLACEMENT>();

        // NXダイアログ監視タイマー
        private static System.Threading.Timer _nxDialogSuppressor;

        // 進捗ダイアログのウィンドウハンドル（監視対象から除外用）
        private static IntPtr _progressFormHandle = IntPtr.Zero;

        // フォーカス奪取防止: ユーザーが操作中のウィンドウを追跡
        private static IntPtr _userForegroundWindow = IntPtr.Zero;

        // 処理中に非表示にしたNXダイアログの追跡（クリーンアップ時に閉じるため）
        private static readonly HashSet<IntPtr> _suppressedWindows = new HashSet<IntPtr>();

        // 言語設定
        private static bool isJapanese;
        private static Dictionary<string, string> messagesJa;
        private static Dictionary<string, string> messagesEn;

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

        // ================================================================
        // 言語初期化・メッセージ取得
        // ================================================================

        private static void InitializeLanguage()
        {
            string lang = Environment.GetEnvironmentVariable("UGII_LANG");
            if (!string.IsNullOrEmpty(lang))
            {
                isJapanese = lang.Equals("japanese", StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                isJapanese = System.Globalization.CultureInfo.CurrentUICulture.Name.StartsWith("ja");
            }

            messagesJa = new Dictionary<string, string>
            {
                // ダイアログタイトル
                { "Error", "エラー" },
                { "Info", "情報" },
                // エラーメッセージ
                { "ErrorNoPartOpen", "パートファイルが開かれていません。" },
                { "ErrorPartNotSaved", "パートファイルが保存されていません。\n先にパートを保存してから実行してください。" },
                { "ErrorNoCustomView", "カスタムビューが見つかりません。" },
                // リスティングウィンドウ
                { "HeaderTitle", "カスタムビュー DXF エクスポーター" },
                { "LwPartName", "パート名: {0}" },
                { "LwPartDir", "パートDir: {0}" },
                { "LwDefFile", "DEF設定ファイル: {0}" },
                { "LwSearchPathSet", "[検索パス] LoadOptions.SetSearchDirectories 設定済み ({0}件)" },
                { "LwSearchPathFailed", "[検索パス] LoadOptions設定失敗（続行）: {0}" },
                { "LwSearchPathAssemSet", "[検索パス] UF_ASSEM.SetSearchDirectories 設定済み" },
                { "LwSearchPathAssemFailed", "[検索パス] UF_ASSEM設定失敗（続行）: {0}" },
                { "LwPartCopied", "[検索パス] パートファイルをコピー: {0}" },
                { "LwPartCopyFailed", "[検索パス] パートコピー失敗（続行）: {0} - {1}" },
                { "FolderSelectTitle", "DXF出力先フォルダを選択" },
                { "FolderCancelled", "フォルダ選択がキャンセルされました。" },
                { "LwOutput", "出力先: {0}" },
                { "LwTargetViews", "変換対象ビュー: {0} 件" },
                { "LwDisplaySuppressed", "[バックグラウンド] 画面更新を抑制しました" },
                { "LwDisplaySuppressFailed", "[バックグラウンド] 画面更新抑制失敗（続行）: {0}" },
                { "UserStopped", "★ ユーザーにより処理が中止されました。" },
                { "LwProcessing", "[{0}/{1}] 処理中: {2}" },
                { "LwSuccess", "  → 成功 ({0:F1}秒)" },
                { "LwFailed", "  → 失敗 ({0:F1}秒): {1}" },
                { "LwViewExportFailed", "ビュー '{0}' のエクスポートに失敗: {1}" },
                { "LwDisplayRestored", "[バックグラウンド] 画面更新を復帰しました" },
                { "LwDisplayRestoreFailed", "[バックグラウンド] 画面更新復帰失敗: {0}" },
                { "ResultTitle", "処理結果" },
                { "LwTotalTime", "合計処理時間: {0:F1}秒" },
                { "FatalError", "致命的エラー: {0}" },
                { "UnexpectedError", "予期しないエラーが発生しました:\n{0}" },
                // 進捗ダイアログ
                { "ProgressTitle", "DXFエクスポート中..." },
                { "ProgressMessage", "処理中: {0}（{1}/{2}）" },
                { "StopButton", "停止" },
                { "StopRequesting", "停止中..." },
                // 完了ダイアログ
                { "CompleteTitleSuccess", "DXFエクスポート完了" },
                { "CompleteTitleError", "DXFエクスポート完了（一部エラー）" },
                { "CompleteMessage", "DXFエクスポートが完了しました" },
                { "CompleteMessageError", "DXFエクスポートが完了しました（一部エラー）" },
                { "SuccessCount", "成功: {0} ファイル" },
                { "FailureCount", "失敗: {0} ファイル" },
                { "CancelCount", "中断: {0} ファイル（ユーザー停止）" },
                { "OutputPath", "出力先: {0}" },
                { "ErrorDetails", "エラー詳細:" },
                { "OkButton", "OK" },
            };

            messagesEn = new Dictionary<string, string>
            {
                // Dialog titles
                { "Error", "Error" },
                { "Info", "Information" },
                // Error messages
                { "ErrorNoPartOpen", "No part file is open." },
                { "ErrorPartNotSaved", "Part file is not saved.\nPlease save the part before running." },
                { "ErrorNoCustomView", "No custom views found." },
                // Listing window
                { "HeaderTitle", "Custom View DXF Exporter" },
                { "LwPartName", "Part: {0}" },
                { "LwPartDir", "Part Dir: {0}" },
                { "LwDefFile", "DEF Settings: {0}" },
                { "LwSearchPathSet", "[Search Path] LoadOptions.SetSearchDirectories configured ({0} entries)" },
                { "LwSearchPathFailed", "[Search Path] LoadOptions configuration failed (continuing): {0}" },
                { "LwSearchPathAssemSet", "[Search Path] UF_ASSEM.SetSearchDirectories configured" },
                { "LwSearchPathAssemFailed", "[Search Path] UF_ASSEM configuration failed (continuing): {0}" },
                { "LwPartCopied", "[Search Path] Part file copied: {0}" },
                { "LwPartCopyFailed", "[Search Path] Part copy failed (continuing): {0} - {1}" },
                { "FolderSelectTitle", "Select DXF Output Folder" },
                { "FolderCancelled", "Folder selection cancelled." },
                { "LwOutput", "Output: {0}" },
                { "LwTargetViews", "Target views: {0}" },
                { "LwDisplaySuppressed", "[Background] Display updates suppressed" },
                { "LwDisplaySuppressFailed", "[Background] Display suppression failed (continuing): {0}" },
                { "UserStopped", "* Processing stopped by user." },
                { "LwProcessing", "[{0}/{1}] Processing: {2}" },
                { "LwSuccess", "  -> Success ({0:F1}s)" },
                { "LwFailed", "  -> Failed ({0:F1}s): {1}" },
                { "LwViewExportFailed", "View '{0}' export failed: {1}" },
                { "LwDisplayRestored", "[Background] Display updates restored" },
                { "LwDisplayRestoreFailed", "[Background] Display restoration failed: {0}" },
                { "ResultTitle", "Results" },
                { "LwTotalTime", "Total time: {0:F1}s" },
                { "FatalError", "Fatal error: {0}" },
                { "UnexpectedError", "An unexpected error occurred:\n{0}" },
                // Progress dialog
                { "ProgressTitle", "Exporting DXF..." },
                { "ProgressMessage", "Processing: {0} ({1}/{2})" },
                { "StopButton", "Stop" },
                { "StopRequesting", "Stopping..." },
                // Completion dialog
                { "CompleteTitleSuccess", "DXF Export Complete" },
                { "CompleteTitleError", "DXF Export Complete (with errors)" },
                { "CompleteMessage", "DXF export completed successfully" },
                { "CompleteMessageError", "DXF export completed with some errors" },
                { "SuccessCount", "Success: {0} files" },
                { "FailureCount", "Failed: {0} files" },
                { "CancelCount", "Cancelled: {0} files (user stopped)" },
                { "OutputPath", "Output: {0}" },
                { "ErrorDetails", "Error details:" },
                { "OkButton", "OK" },
            };
        }

        private static string GetMessage(string key)
        {
            var dict = isJapanese ? messagesJa : messagesEn;
            return dict.ContainsKey(key) ? dict[key] : key;
        }

        private static string GetMessage(string key, params object[] args)
        {
            return string.Format(GetMessage(key), args);
        }

        private static Font GetUIFont(float size, FontStyle style = FontStyle.Regular)
        {
            try
            {
                Font font = new Font("Yu Gothic UI", size, style);
                if (font.Name == "Yu Gothic UI") return font;
                font.Dispose();
            }
            catch { }
            return new Font("Segoe UI", size, style);
        }

        // ================================================================
        // メインエントリポイント
        // ================================================================

        public static void Main(string[] args)
        {
            theSession = Session.GetSession();
            theUFSession = UFSession.GetUFSession();
            theUI = UI.GetUI();
            lw = theSession.ListingWindow;
            lw.Open();

            InitializeLanguage();

            Stopwatch totalSw = Stopwatch.StartNew();

            // CGM環境変数の元の値を退避
            string originalCgmEnv = Environment.GetEnvironmentVariable("UGII_CGM_FITS_FILE_SAVE");

            // 二重呼び出し防止フラグ
            bool displaySuppressed = false;
            bool undoMarkCreated = false;
            Session.UndoMarkId undoMark = default(Session.UndoMarkId);

            int successCount = 0;
            int failCount = 0;
            int cancelCount = 0;
            List<string> errorDetails = new List<string>();
            string outputFolder = null;

            CollectNxWindows();
            ProgressForm progressForm = null;
            Thread progressThread = null;

            try
            {
                // 1. 作業パートの確認
                workPart = theSession.Parts.Work;
                if (workPart == null)
                {
                    theUI.NXMessageBox.Show(GetMessage("Error"), NXMessageBox.DialogType.Error,
                        GetMessage("ErrorNoPartOpen"));
                    return;
                }

                // パートが保存済み（ディスク上に存在）であることを確認
                if (string.IsNullOrEmpty(workPart.FullPath))
                {
                    theUI.NXMessageBox.Show(GetMessage("Error"), NXMessageBox.DialogType.Error,
                        GetMessage("ErrorPartNotSaved"));
                    return;
                }

                // CGMダイアログ抑制（処理開始前に設定）
                SuppressCgmDialog();

                // パート名（拡張子なし）とディレクトリを取得
                string partName = Path.GetFileNameWithoutExtension(workPart.FullPath);
                string partDir = Path.GetDirectoryName(workPart.FullPath);
                lw.WriteLine("========================================");
                lw.WriteLine("  " + GetMessage("HeaderTitle"));
                lw.WriteLine("========================================");
                lw.WriteLine(GetMessage("LwPartName", partName));
                lw.WriteLine(GetMessage("LwPartDir", partDir));

                // 2. パートファイル検索パスの設定（Failed to retrieve file 対策）
                SetupPartSearchPaths(partDir);

                // 3. dxfdwg.def 一時ファイルの生成（{PART_DIR}をパートディレクトリに置換）
                string defContent = EmbeddedDefContent.Replace("{PART_DIR}", partDir);
                tempDefPath = Path.Combine(Path.GetTempPath(), "nxdxf_export_settings.def");
                File.WriteAllText(tempDefPath, defContent);
                lw.WriteLine(GetMessage("LwDefFile", tempDefPath));

                // 4. パートファイルをTempディレクトリにコピー（トランスレータが確実に参照可能にする）
                CopyPartFilesToTemp(workPart.FullPath);

                // 5. 出力フォルダの選択（NX最小化の前に実行）
                outputFolder = SelectOutputFolder();
                if (outputFolder == null)
                {
                    lw.WriteLine(GetMessage("FolderCancelled"));
                    return;
                }
                lw.WriteLine(GetMessage("LwOutput", outputFolder));

                // 6. カスタムビューの取得
                List<ModelingView> customViews = GetCustomViews();
                if (customViews.Count == 0)
                {
                    lw.WriteLine(GetMessage("ErrorNoCustomView"));
                    theUI.NXMessageBox.Show(GetMessage("Info"), NXMessageBox.DialogType.Information,
                        GetMessage("ErrorNoCustomView"));
                    return;
                }

                lw.WriteLine("----------------------------------------");
                lw.WriteLine(GetMessage("LwTargetViews", customViews.Count));
                for (int i = 0; i < customViews.Count; i++)
                {
                    lw.WriteLine(string.Format("  {0}. {1}", i + 1, customViews[i].Name));
                }
                lw.WriteLine("----------------------------------------");

                // 7. 製図アプリケーションに切り替え（ループ外で1回だけ実行）
                theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING");

                // 8. UndoMark作成（製図変更開始直前、早期returnでは作成しない）
                undoMark = theSession.SetUndoMark(Session.MarkVisibility.Invisible, "DXF Export");
                undoMarkCreated = true;

                // 9. NXの全ウィンドウを最小化（メイン＋情報ウィンドウ等）
                lw.Close();
                MinimizeAllNxWindows();

                // 10. フォーカスロック
                try
                {
                    LockSetForegroundWindow(LSFW_LOCK);
                }
                catch { }

                // 11. 画面更新を抑制
                theUFSession.Disp.SetDisplay(NXOpen.UF.UFConstants.UF_DISP_SUPPRESS_DISPLAY);
                displaySuppressed = true;

                // 12. 進捗ダイアログを専用STAスレッドで表示（NXOpen APIブロック中も操作可能）
                ManualResetEvent formReady = new ManualResetEvent(false);
                progressThread = new Thread(() =>
                {
                    progressForm = new ProgressForm(customViews.Count);
                    progressForm.HandleCreated += (s, e) =>
                    {
                        _progressFormHandle = progressForm.Handle;
                        formReady.Set();
                    };
                    Application.Run(progressForm);
                    // Application.Run終了 = ForceCloseされた
                    progressForm.Dispose();
                    progressForm = null;
                });
                progressThread.SetApartmentState(ApartmentState.STA);
                progressThread.IsBackground = true;
                progressThread.Start();
                formReady.WaitOne(5000);

                // 13. NXダイアログ監視タイマー開始 + カスタムビューごとのループ処理
                _userForegroundWindow = GetForegroundWindow();
                StartNxDialogSuppressor();
                try
                {
                    for (int i = 0; i < customViews.Count; i++)
                    {
                        // 中断チェック
                        if (progressForm != null && progressForm.StopRequested)
                        {
                            cancelCount = customViews.Count - i;
                            break;
                        }

                        ModelingView view = customViews[i];

                        // 進捗更新
                        if (progressForm != null)
                            progressForm.UpdateProgress(view.Name, i + 1, customViews.Count);

                        Stopwatch viewSw = Stopwatch.StartNew();
                        try
                        {
                            ExportViewAsDxf(view, partName, outputFolder);
                            viewSw.Stop();
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            viewSw.Stop();
                            failCount++;
                            string errorMsg = string.Format("{0}: {1}", view.Name, ex.Message);
                            errorDetails.Add(errorMsg);
                        }
                    }
                }
                finally
                {
                    // 注: ここではSuppressorを停止しない。
                    // クリーンアップ中もNXダイアログを抑制し続ける必要がある。
                }

                // === ループ終了後のクリーンアップ（順序が重要！） ===

                bool userStopped = (cancelCount > 0);

                // Step 1: 進捗ダイアログを閉じる（STAスレッドのApplication.Runを終了）
                _progressFormHandle = IntPtr.Zero;
                if (progressForm != null)
                {
                    try { progressForm.ForceClose(); } catch { }
                    progressForm = null;
                }
                // スレッド完了を待つ（DLLアンロード時のクラッシュ防止）
                if (progressThread != null && progressThread.IsAlive)
                {
                    progressThread.Join(2000);
                }

                if (userStopped)
                {
                    // === ユーザー停止: NXバックグラウンドDXF翻訳との競合を回避 ===

                    // 最優先: NXウィンドウを即座に復元（応答性向上）
                    StopNxDialogSuppressor();
                    // 非表示ダイアログは再表示しない（DXF翻訳中のダイアログ操作はクラッシュ原因）
                    lock (_suppressedWindows) { _suppressedWindows.Clear(); }
                    RestoreAllNxWindows();
                    try { LockSetForegroundWindow(LSFW_UNLOCK); } catch { }

                    // 画面更新を復帰
                    if (displaySuppressed)
                    {
                        try { theUFSession.Disp.SetDisplay(NXOpen.UF.UFConstants.UF_DISP_UNSUPPRESS_DISPLAY); }
                        catch { }
                        displaySuppressed = false;
                    }

                    // UndoToMarkをスキップ（バックグラウンドDXF翻訳中のUndoはNXクラッシュの原因）
                    // 各ビューのCleanupViewAndSheetで個別変更は既に巻き戻し済み
                    undoMarkCreated = false;
                }
                else
                {
                    // === 正常完了: フルクリーンアップ ===

                    // 画面更新を復帰
                    if (displaySuppressed)
                    {
                        theUFSession.Disp.SetDisplay(NXOpen.UF.UFConstants.UF_DISP_UNSUPPRESS_DISPLAY);
                        displaySuppressed = false;
                        Thread.Sleep(50);
                    }

                    // UndoMarkで状態を復元
                    SafeUndoAndCleanup(theSession, ref undoMark, ref undoMarkCreated);

                    // モデリングアプリケーションに戻る
                    try { theSession.ApplicationSwitchImmediate("UG_APP_MODELING"); }
                    catch { }

                    // NXダイアログ監視を停止し、NXウィンドウを復元
                    StopNxDialogSuppressor();
                    DismissSuppressedDialogs();
                    RestoreAllNxWindows();
                    try { LockSetForegroundWindow(LSFW_UNLOCK); } catch { }
                }

                // リスティングウィンドウに結果を出力
                totalSw.Stop();
                try
                {
                    lw.Open();
                    lw.WriteLine("\n========================================");
                    lw.WriteLine("  " + GetMessage("ResultTitle"));
                    lw.WriteLine("========================================");
                    lw.WriteLine(GetMessage("SuccessCount", successCount));
                    if (failCount > 0)
                        lw.WriteLine(GetMessage("FailureCount", failCount));
                    if (cancelCount > 0)
                        lw.WriteLine(GetMessage("CancelCount", cancelCount));
                    lw.WriteLine(GetMessage("LwTotalTime", totalSw.Elapsed.TotalSeconds));
                    lw.WriteLine("========================================");
                }
                catch { }

                // 完了ダイアログを表示
                using (var completionForm = new CompletionForm(
                    successCount, failCount, cancelCount, errorDetails, outputFolder))
                {
                    completionForm.ShowDialog();
                }

                // 停止時: 完了ダイアログ後にモデリング復帰を試行
                // （ユーザーがダイアログを閲覧中にDXF翻訳が完了している可能性が高い）
                if (userStopped)
                {
                    try { theSession.ApplicationSwitchImmediate("UG_APP_MODELING"); }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                // 全ての例外をキャッチ（NXに伝播させない）
                try
                {
                    _progressFormHandle = IntPtr.Zero;
                    if (progressForm != null)
                    {
                        progressForm.ForceClose();
                        progressForm = null;
                    }
                    if (progressThread != null && progressThread.IsAlive)
                        progressThread.Join(2000);
                }
                catch { }

                try
                {
                    MessageBox.Show(
                        GetMessage("UnexpectedError", ex.Message),
                        GetMessage("Error"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                catch { }
            }
            finally
            {
                // 確実に復帰（二重呼び出し防止のフラグ付き）

                try
                {
                    _progressFormHandle = IntPtr.Zero;
                    if (progressForm != null)
                    {
                        progressForm.ForceClose();
                        progressForm = null;
                    }
                    if (progressThread != null && progressThread.IsAlive)
                        progressThread.Join(2000);
                }
                catch { }

                try
                {
                    StopNxDialogSuppressor();
                    lock (_suppressedWindows) { _suppressedWindows.Clear(); }
                }
                catch { }

                try
                {
                    RestoreAllNxWindows();
                    LockSetForegroundWindow(LSFW_UNLOCK);
                }
                catch { }

                try
                {
                    if (displaySuppressed)
                    {
                        theUFSession.Disp.SetDisplay(NXOpen.UF.UFConstants.UF_DISP_UNSUPPRESS_DISPLAY);
                    }
                }
                catch { }

                try
                {
                    SafeUndoAndCleanup(theSession, ref undoMark, ref undoMarkCreated);
                }
                catch { }

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
            }
        }

        // ================================================================
        // UndoMark安全復元ヘルパー
        // ================================================================

        /// <summary>
        /// UndoMarkの復元と削除を安全に1回だけ実行するヘルパー。
        /// NX内部のDXFエクスポートがUndoMarkを無効化する場合があるため、
        /// try-catchで保護し、二重実行を防止するためフラグをリセットする。
        /// </summary>
        private static void SafeUndoAndCleanup(Session session, ref Session.UndoMarkId undoMark, ref bool undoMarkCreated)
        {
            if (!undoMarkCreated) return;

            try
            {
                session.UndoToMark(undoMark, "DXF Export");
            }
            catch (Exception) { /* NX内部でマークが無効化された場合は無視 */ }

            try
            {
                session.DeleteUndoMark(undoMark, "DXF Export");
            }
            catch (Exception) { /* NX内部でマークが無効化された場合は無視 */ }

            undoMarkCreated = false;
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
                lw.WriteLine(GetMessage("LwSearchPathSet", allDirs.Count));
            }
            catch (Exception ex)
            {
                lw.WriteLine(GetMessage("LwSearchPathFailed", ex.Message));
            }

            // 対策2: UF_ASSEMレベルで検索ディレクトリを設定
            try
            {
                theUFSession.Assem.SetSearchDirectories(
                    1,
                    new string[] { partDir },
                    new bool[] { true });
                lw.WriteLine(GetMessage("LwSearchPathAssemSet"));
            }
            catch (Exception ex)
            {
                lw.WriteLine(GetMessage("LwSearchPathAssemFailed", ex.Message));
            }
        }

        /// <summary>
        /// パートファイルとその関連ファイルをTempディレクトリにコピー。
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
                lw.WriteLine(GetMessage("LwPartCopied", destPath));
            }
            catch (Exception ex)
            {
                lw.WriteLine(GetMessage("LwPartCopyFailed", destDir, ex.Message));
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
        // フォルダ選択
        // ================================================================

        private static string SelectOutputFolder()
        {
            string selectedPath = null;
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = GetMessage("FolderSelectTitle");
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

                // b. 図面シートの作成 (DraftingDrawingSheetBuilder使用 - ジャーナル準拠)
                NXOpen.Drawings.DraftingDrawingSheet nullDraftingSheet = null;
                var sheetBuilder = workPart.DraftingDrawingSheets.CreateDraftingDrawingSheetBuilder(nullDraftingSheet);
                sheetBuilder.AutoStartViewCreation = false;
                sheetBuilder.Option = DrawingSheetBuilder.SheetOption.CustomSize;
                sheetBuilder.Height = sheetHeight;
                sheetBuilder.Length = sheetWidth;
                sheetBuilder.ScaleNumerator = 1.0;
                sheetBuilder.ScaleDenominator = 1.0;
                sheetBuilder.Units = DrawingSheetBuilder.SheetUnits.Metric;
                sheetBuilder.ProjectionAngle = DrawingSheetBuilder.SheetProjectionAngle.Third;

                NXObject sheetObj = sheetBuilder.Commit();
                sheet = (DrawingSheet)sheetObj;
                sheetBuilder.Destroy();

                // ジャーナル準拠: テンプレートインスタンス化完了を通知
                workPart.Drafting.SetTemplateInstantiationIsComplete(true);

                // シートを開く
                sheet.Open();

                // c. ベースビューの配置（ジャーナル準拠）
                workPart.DraftingManager.DrawingsFreezeOutOfDateComputation();

                NXOpen.Drawings.BaseView nullBaseView = null;
                BaseViewBuilder baseViewBuilder = workPart.DraftingViews.CreateBaseViewBuilder(nullBaseView);

                baseViewBuilder.Placement.Associative = true;

                // カスタムビューを選択
                baseViewBuilder.SelectModelView.SelectedView = view;

                baseViewBuilder.SecondaryComponents.ObjectType =
                    NXOpen.Drawings.DraftingComponentSelectionBuilder.Geometry.PrimaryGeometry;

                // ★重要: PartNameに元パートのフルパスを明示的に設定（参照解決に必要）
                baseViewBuilder.Style.ViewStyleBase.Part = workPart;
                baseViewBuilder.Style.ViewStyleBase.PartName = workPart.FullPath;

                // ジャーナル準拠: PartName設定後にビューを再選択
                baseViewBuilder.SelectModelView.SelectedView = view;

                // ビューがシートに収まるようスケール調整
                double availableWidth = sheetWidth - (SheetMargin * 2);
                double availableHeight = sheetHeight - (SheetMargin * 2);
                baseViewBuilder.Scale.Denominator = 1.0;
                baseViewBuilder.Scale.Numerator = 1.0;
                if (viewWidth > 0 && viewHeight > 0)
                {
                    double scaleX = availableWidth / viewWidth;
                    double scaleY = availableHeight / viewHeight;
                    double scale = Math.Min(scaleX, scaleY);
                    if (scale < 1.0)
                    {
                        baseViewBuilder.Scale.Numerator = scale;
                    }
                }

                // 配置位置（シート中央）
                Point3d viewOrigin = new Point3d(sheetWidth / 2.0, sheetHeight / 2.0, 0.0);
                baseViewBuilder.Placement.Placement.SetValue(null, workPart.Views.WorkView, viewOrigin);

                baseViewObj = baseViewBuilder.Commit();

                workPart.DraftingManager.DrawingsUnfreezeOutOfDateComputation();
                baseViewBuilder.Destroy();

                // d. DXFエクスポート（ジャーナル準拠の順序）
                string dxfFileName = string.Format("{0}_{1}.dxf", partName, view.Name);
                string dxfFilePath = Path.Combine(outputFolder, dxfFileName);

                DxfdwgCreator dxfCreator = theSession.DexManager.CreateDxfdwgCreator();

                // --- SettingsFileより前の初期設定（ジャーナル順序）---
                dxfCreator.ExportData = DxfdwgCreator.ExportDataOption.Drawing;
                dxfCreator.AutoCADRevision = DxfdwgCreator.AutoCADRevisionOptions.R2004;
                dxfCreator.ViewEditMode = true;
                dxfCreator.FlattenAssembly = true;
                dxfCreator.ExportScaleValue = 1.0;

                // --- dxfdwg.def設定ファイルの読み込み ---
                dxfCreator.SettingsFile = tempDefPath;

                // --- SettingsFile読み込み後の設定（defの値を上書き）---
                dxfCreator.OutputTo = DxfdwgCreator.OutputToOption.Drafting;
                dxfCreator.ObjectTypes.Curves = true;
                dxfCreator.ObjectTypes.Annotations = true;
                dxfCreator.ObjectTypes.Structures = true;
                dxfCreator.AutoCADRevision = DxfdwgCreator.AutoCADRevisionOptions.R2018;
                dxfCreator.FlattenAssembly = false;

                // ★重要: InputFileに元パートのフルパスを設定（参照解決に必要）
                dxfCreator.InputFile = workPart.FullPath;

                dxfCreator.OutputFile = dxfFilePath;

                // --- Commit直前の最終設定 ---
                dxfCreator.WidthFactorMode = DxfdwgCreator.WidthfactorMethodOptions.AutomaticCalculation;
                dxfCreator.LayerMask = "1-256";
                dxfCreator.DrawingList = @"""Sheet 1""";
                dxfCreator.ProcessHoldFlag = true;

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
        // NXメインウィンドウ検索
        // ================================================================

        private static void CollectNxWindows()
        {
            nxWindows.Clear();
            _savedPlacements.Clear();
            uint currentPid = (uint)Process.GetCurrentProcess().Id;

            EnumWindows((hWnd, lParam) =>
            {
                uint windowPid;
                GetWindowThreadProcessId(hWnd, out windowPid);
                if (windowPid == currentPid && IsWindowVisible(hWnd))
                {
                    nxWindows.Add(hWnd);
                    // ウィンドウの配置情報を保存（最大化状態等を復元するため）
                    WINDOWPLACEMENT wp = new WINDOWPLACEMENT();
                    wp.length = (uint)Marshal.SizeOf(typeof(WINDOWPLACEMENT));
                    GetWindowPlacement(hWnd, ref wp);
                    _savedPlacements[hWnd] = wp;
                }
                return true;
            }, IntPtr.Zero);
        }

        private static void MinimizeAllNxWindows()
        {
            foreach (IntPtr hWnd in nxWindows)
            {
                ShowWindow(hWnd, SW_MINIMIZE);
            }
        }

        private static void RestoreAllNxWindows()
        {
            // 保存した配置情報で復元（最大化状態等を正確に再現）
            foreach (var kvp in _savedPlacements)
            {
                WINDOWPLACEMENT wp = kvp.Value;
                SetWindowPlacement(kvp.Key, ref wp);
            }
            Thread.Sleep(100);
        }

        // ================================================================
        // NX「作業進行中」ダイアログ監視・非表示
        // ================================================================

        private static void StartNxDialogSuppressor()
        {
            uint currentProcessId = (uint)Process.GetCurrentProcess().Id;
            lock (_suppressedWindows) { _suppressedWindows.Clear(); }

            _nxDialogSuppressor = new System.Threading.Timer(_ =>
            {
                try
                {
                    SuppressNxDialogs(currentProcessId);
                }
                catch
                {
                    // タイマー内の例外は無視（安全性優先）
                }
            }, null, 0, 50); // 50msごとに監視（ダイアログ出現を即座に検出）
        }

        private static void StopNxDialogSuppressor()
        {
            var timer = _nxDialogSuppressor;
            _nxDialogSuppressor = null;
            if (timer != null)
            {
                // 最後のコールバック完了を待ってから破棄（レースコンディション防止）
                using (var waitHandle = new ManualResetEvent(false))
                {
                    timer.Dispose(waitHandle);
                    waitHandle.WaitOne(200);
                }
            }
        }

        /// <summary>
        /// 処理中に非表示にしたNXダイアログを再表示し、NXが自然にクローズするのを待つ。
        /// WM_CLOSEによる強制クローズはNX内部スレッドのクラッシュ原因になるため使用しない。
        /// </summary>
        private static void DismissSuppressedDialogs()
        {
            List<IntPtr> windows;
            lock (_suppressedWindows)
            {
                windows = new List<IntPtr>(_suppressedWindows);
                _suppressedWindows.Clear();
            }

            if (windows.Count == 0)
                return;

            // 非表示にしたダイアログを再表示（フォーカスは奪わない）
            foreach (IntPtr hWnd in windows)
            {
                ShowWindow(hWnd, SW_SHOWNOACTIVATE);
            }

            // NXがダイアログを自動的に閉じるのを待つ（最大1秒、100msごとにチェック）
            for (int i = 0; i < 10; i++)
            {
                Thread.Sleep(100);
                bool anyVisible = false;
                foreach (IntPtr hWnd in windows)
                {
                    if (IsWindowVisible(hWnd))
                    {
                        anyVisible = true;
                        break;
                    }
                }
                if (!anyVisible)
                    break;
            }
        }

        /// <summary>
        /// 処理中にNXプロセスが新規表示したウィンドウを非表示にし、
        /// NXがフォーカスを奪った場合はユーザーの元のウィンドウに復元する。
        /// </summary>
        private static void SuppressNxDialogs(uint targetProcessId)
        {
            IntPtr progressHandle = _progressFormHandle;

            // 1. NXプロセスが新たに表示したウィンドウを非表示にする
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd))
                    return true;

                // 自分のProgressFormはスキップ
                if (progressHandle != IntPtr.Zero && hWnd == progressHandle)
                    return true;

                // NXプロセスのウィンドウか確認
                uint processId;
                GetWindowThreadProcessId(hWnd, out processId);
                if (processId != targetProcessId)
                    return true;

                if (nxWindows.Contains(hWnd))
                {
                    // 処理開始前に記録済みのNXウィンドウが復元された → 再最小化
                    if (!IsIconic(hWnd))
                        ShowWindow(hWnd, SW_MINIMIZE);
                }
                else
                {
                    // NXプロセスが新たに表示したウィンドウ → 非表示にする
                    ShowWindow(hWnd, SW_HIDE);
                    lock (_suppressedWindows) { _suppressedWindows.Add(hWnd); }
                }

                return true;
            }, IntPtr.Zero);

            // 2. NXがフォーカスを奪っていたら、ユーザーの元のウィンドウに復元
            IntPtr fg = GetForegroundWindow();
            if (fg == IntPtr.Zero)
                return;

            // ProgressFormがフォアグラウンドなら問題なし
            if (progressHandle != IntPtr.Zero && fg == progressHandle)
                return;

            uint fgProcessId;
            uint fgThreadId = GetWindowThreadProcessId(fg, out fgProcessId);
            if (fgProcessId != targetProcessId)
            {
                // ユーザーが別アプリを操作中 → そのウィンドウを記憶
                _userForegroundWindow = fg;
                return;
            }

            // NXがフォーカスを奪っている → ユーザーの元のウィンドウに返す
            IntPtr target = _userForegroundWindow;
            if (target == IntPtr.Zero)
                return;

            uint curThreadId = GetCurrentThreadId();
            AttachThreadInput(curThreadId, fgThreadId, true);
            SetForegroundWindow(target);
            AttachThreadInput(curThreadId, fgThreadId, false);
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

        // ================================================================
        // 進捗ダイアログ（ProgressForm）
        // ================================================================

        private class ProgressForm : Form
        {
            private Label labelProgress;
            private ProgressBar progressBar;
            private Label labelPercent;
            private Button btnStop;
            private volatile bool _stopRequested;
            private bool _allowClose;

            public bool StopRequested => _stopRequested;

            public ProgressForm(int totalViews)
            {
                _stopRequested = false;
                _allowClose = false;

                this.Text = GetMessage("ProgressTitle");
                this.Width = 400;
                this.Height = 200;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.TopMost = true;
                this.ShowInTaskbar = true;
                this.Font = GetUIFont(9f);

                // ×ボタンの無効化（ForceClose以外では閉じさせない）
                this.FormClosing += (s, e) =>
                {
                    if (!_allowClose && e.CloseReason == CloseReason.UserClosing)
                        e.Cancel = true;
                };

                labelProgress = new Label();
                labelProgress.Text = GetMessage("ProgressMessage", "...", 0, totalViews);
                labelProgress.Location = new System.Drawing.Point(20, 20);
                labelProgress.Size = new System.Drawing.Size(350, 25);
                labelProgress.Font = GetUIFont(10f);
                this.Controls.Add(labelProgress);

                progressBar = new ProgressBar();
                progressBar.Location = new System.Drawing.Point(20, 55);
                progressBar.Size = new System.Drawing.Size(280, 25);
                progressBar.Minimum = 0;
                progressBar.Maximum = totalViews;
                progressBar.Value = 0;
                this.Controls.Add(progressBar);

                labelPercent = new Label();
                labelPercent.Text = "0%";
                labelPercent.Location = new System.Drawing.Point(310, 58);
                labelPercent.Size = new System.Drawing.Size(50, 25);
                labelPercent.Font = GetUIFont(10f);
                this.Controls.Add(labelPercent);

                btnStop = new Button();
                btnStop.Text = GetMessage("StopButton");
                btnStop.Location = new System.Drawing.Point(140, 100);
                btnStop.Size = new System.Drawing.Size(120, 40);
                btnStop.BackColor = Color.FromArgb(220, 53, 69);
                btnStop.ForeColor = Color.White;
                btnStop.FlatStyle = FlatStyle.Flat;
                btnStop.Font = GetUIFont(10f, FontStyle.Bold);
                btnStop.Click += btnStop_Click;
                this.Controls.Add(btnStop);
            }

            private void btnStop_Click(object sender, EventArgs e)
            {
                _stopRequested = true;
                btnStop.Enabled = false;
                btnStop.Text = GetMessage("StopRequesting");
            }

            public void UpdateProgress(string viewName, int current, int total)
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(() => UpdateProgress(viewName, current, total)));
                    return;
                }
                labelProgress.Text = GetMessage("ProgressMessage", viewName, current, total);
                progressBar.Maximum = total;
                progressBar.Value = current;
                labelPercent.Text = string.Format("{0}%", (int)((double)current / total * 100));
            }

            /// <summary>
            /// 別スレッドから安全にフォームを閉じる。Application.Runを終了させる。
            /// </summary>
            public void ForceClose()
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(ForceClose));
                    return;
                }
                _allowClose = true;
                Close();
            }
        }

        // ================================================================
        // 完了ダイアログ（CompletionForm）
        // ================================================================

        private class CompletionForm : Form
        {
            public CompletionForm(int successCount, int failCount, int cancelCount,
                List<string> errors, string outputPath)
            {
                bool hasErrors = failCount > 0;

                this.Text = hasErrors ? GetMessage("CompleteTitleError") : GetMessage("CompleteTitleSuccess");
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.TopMost = true;
                this.ShowInTaskbar = true;
                this.Font = GetUIFont(9f);

                int y = 20;

                // ステータスアイコン + メッセージ
                Label labelStatus = new Label();
                if (hasErrors)
                {
                    labelStatus.Text = "\u26A0 " + GetMessage("CompleteMessageError");
                }
                else
                {
                    labelStatus.Text = "\u2713 " + GetMessage("CompleteMessage");
                }
                labelStatus.Location = new System.Drawing.Point(20, y);
                labelStatus.Size = new System.Drawing.Size(400, 30);
                labelStatus.Font = GetUIFont(11f, FontStyle.Bold);
                this.Controls.Add(labelStatus);
                y += 40;

                // 成功数
                Label labelSuccess = new Label();
                labelSuccess.Text = GetMessage("SuccessCount", successCount);
                labelSuccess.Location = new System.Drawing.Point(30, y);
                labelSuccess.Size = new System.Drawing.Size(380, 22);
                labelSuccess.Font = GetUIFont(10f);
                this.Controls.Add(labelSuccess);
                y += 25;

                // 失敗数
                if (failCount > 0)
                {
                    Label labelFail = new Label();
                    labelFail.Text = GetMessage("FailureCount", failCount);
                    labelFail.Location = new System.Drawing.Point(30, y);
                    labelFail.Size = new System.Drawing.Size(380, 22);
                    labelFail.Font = GetUIFont(10f);
                    this.Controls.Add(labelFail);
                    y += 25;
                }

                // 中断数
                if (cancelCount > 0)
                {
                    Label labelCancel = new Label();
                    labelCancel.Text = GetMessage("CancelCount", cancelCount);
                    labelCancel.Location = new System.Drawing.Point(30, y);
                    labelCancel.Size = new System.Drawing.Size(380, 22);
                    labelCancel.Font = GetUIFont(10f);
                    this.Controls.Add(labelCancel);
                    y += 25;
                }

                // 出力先パス
                Label labelOutput = new Label();
                labelOutput.Text = GetMessage("OutputPath", outputPath);
                labelOutput.Location = new System.Drawing.Point(30, y);
                labelOutput.Size = new System.Drawing.Size(380, 22);
                labelOutput.Font = GetUIFont(9f);
                this.Controls.Add(labelOutput);
                y += 30;

                // エラー詳細
                if (errors != null && errors.Count > 0)
                {
                    Label labelErrorHeader = new Label();
                    labelErrorHeader.Text = GetMessage("ErrorDetails");
                    labelErrorHeader.Location = new System.Drawing.Point(30, y);
                    labelErrorHeader.Size = new System.Drawing.Size(380, 22);
                    labelErrorHeader.Font = GetUIFont(9f, FontStyle.Bold);
                    this.Controls.Add(labelErrorHeader);
                    y += 25;

                    foreach (string error in errors)
                    {
                        Label labelError = new Label();
                        labelError.Text = "  " + error;
                        labelError.Location = new System.Drawing.Point(30, y);
                        labelError.Size = new System.Drawing.Size(380, 40);
                        labelError.Font = GetUIFont(9f);
                        this.Controls.Add(labelError);
                        y += 40;
                    }
                }

                y += 10;

                // OKボタン
                Button btnOk = new Button();
                btnOk.Text = GetMessage("OkButton");
                btnOk.Location = new System.Drawing.Point(170, y);
                btnOk.Size = new System.Drawing.Size(100, 35);
                btnOk.Font = GetUIFont(10f);
                btnOk.Click += (s, e) =>
                {
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                };
                this.Controls.Add(btnOk);
                this.AcceptButton = btnOk;

                y += 55;

                this.ClientSize = new System.Drawing.Size(430, y);
            }
        }
    }
}
