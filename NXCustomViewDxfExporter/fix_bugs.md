【緊急修正】以下の2つのバグを修正してください。

まず現在の状態を確認してください：
1. このフォルダにあるCustomViewDxfExporter.csを開く
2. grep -n "GetUnloadOption" -A5 CustomViewDxfExporter.cs
3. grep -n "Immediately\|AtTermination" CustomViewDxfExporter.cs
4. Main()メソッドの最後20行を表示
5. ProgressFormクラスのforegroundTimerの部分を表示
6. UpdateProgressメソッドを表示

---

## 修正1: NXクラッシュ(c0000005)の回避

原因: DxfdwgCreator.Commit()はバックグラウンドスレッドでDXF変換する。Main()終了後もスレッドが実行中のままDLLがアンロードされてクラッシュする。

### 修正1-A: GetUnloadOptionの変更
GetUnloadOptionがImmediatelyを返している場合、AtTerminationに変更する：
```csharp
public static int GetUnloadOption(string dummy)
{
    return (int)Session.LibraryUnloadOption.AtTermination;
}
```

### 修正1-B: Main()終了前に待機を追加
Main()メソッドのforループ（全ビューのエクスポート）が終わった後、progressFormのForceClose()より前に以下を追加：
```csharp
// DxfdwgCreatorのバックグラウンドDXF変換スレッド完了を待機
// これがないとMain()終了後にDLL内のコードが実行中のままアンロードされてクラッシュする
Thread.Sleep(15000);
```

---

## 修正2: 進捗ダイアログが2つ目のビュー処理時にNXウィンドウの後ろに隠れる

原因: NXのDxfdwgCreator操作やシート作成時にNXウィンドウが前面に来る。現在のTopMost+SetWindowPosだけでは不十分。

### 修正2-A: Win32 API追加
ProgressFormクラスの既存のDllImport宣言の近くに以下を追加：
```csharp
[DllImport("user32.dll")]
private static extern bool SetForegroundWindow(IntPtr hWnd);
```

### 修正2-B: foregroundTimerのInterval短縮とTick処理強化
foregroundTimerのIntervalを200に変更し、Tickイベントの中身を以下に変更：
```csharp
foregroundTimer.Interval = 200;
foregroundTimer.Tick += (s2, e2) =>
{
    if (this.Visible && !this.IsDisposed)
    {
        this.TopMost = true;
        SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        SetForegroundWindow(this.Handle);
        this.BringToFront();
    }
};
```

### 修正2-C: UpdateProgressメソッド内の最前面維持を強化
UpdateProgressメソッド内で、TopMostやBringToFrontを設定している箇所を以下に変更：
```csharp
this.TopMost = true;
SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0,
    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
SetForegroundWindow(this.Handle);
this.BringToFront();
```

---

修正後、ソリューションをリビルド：
msbuild /t:Rebuild /p:Configuration=Debug /p:Platform=x64 NXCustomViewDxfExporter.sln
