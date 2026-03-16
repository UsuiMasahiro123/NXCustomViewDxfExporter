【緊急】NXクラッシュ(c0000005)の回避のみ修正してください。他の変更は一切不要です。

まず現在の状態を確認してください：
1. grep -n "GetUnloadOption" -A5 CustomViewDxfExporter.cs
2. grep -n "Immediately\|AtTermination" CustomViewDxfExporter.cs  
3. Main()メソッドの最後20行を表示

修正は以下の2点のみ：

【修正1】GetUnloadOptionがImmediatelyを返している場合、AtTerminationに変更：
```csharp
public static int GetUnloadOption(string dummy)
{
    return (int)Session.LibraryUnloadOption.AtTermination;
}
```

【修正2】Main()メソッドのreturnの直前（progressFormのForceCloseより前）に以下を追加：
```csharp
// DxfdwgCreatorのバックグラウンドDXF変換スレッド完了を待機
// これがないとMain()終了後にDLL内のコードが実行中のままアンロードされてクラッシュする
Thread.Sleep(15000);
```

この2点だけ修正してリビルド：
msbuild /t:Rebuild /p:Configuration=Debug /p:Platform=x64 NXCustomViewDxfExporter.sln
