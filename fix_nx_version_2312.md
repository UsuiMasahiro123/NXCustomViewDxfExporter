# 修正指示: NX Open 参照DLLをNX 2312に変更（バージョン不一致エラー修正）

## エラー内容
NX 2312環境で実行時に以下のエラーが発生:
```
System.EntryPointNotFoundException: DLL 'libjam' の 'ReportAutomationClientExtern' というエントリ ポイントが見つかりません。
```
原因: ビルド出力フォルダにNX 2406版のNXOpen.Utilities.dll（Version=2406.0.0.0）がコピーされており、NX 2312のlibjam.dllと互換性がない。

## 修正手順

### 1. NX 2312のmanaged DLLフォルダを確認
以下のパスが存在することを確認:
```
C:\Program Files\Siemens\NX2312\nxbin\managed\
```
存在しない場合は以下で検索:
```powershell
Get-ChildItem -Path "C:\Program Files\Siemens" -Recurse -Filter "NXOpen.dll" -ErrorAction SilentlyContinue | Select-Object FullName
```

### 2. .csprojファイルの参照パスを変更
`NXCustomViewDxfExporter\NXCustomViewDxfExporter.csproj` を編集する。

NX Open関連の全ての参照について、以下の2点を変更すること:
- **HintPath**: NX 2312のパスに変更
- **Private**: false に設定（= Copy Local を無効化。これが最重要）

変更後の形式:
```xml
<Reference Include="NXOpen">
  <HintPath>C:\Program Files\Siemens\NX2312\nxbin\managed\NXOpen.dll</HintPath>
  <Private>false</Private>
</Reference>
<Reference Include="NXOpen.UF">
  <HintPath>C:\Program Files\Siemens\NX2312\nxbin\managed\NXOpen.UF.dll</HintPath>
  <Private>false</Private>
</Reference>
<Reference Include="NXOpen.Utilities">
  <HintPath>C:\Program Files\Siemens\NX2312\nxbin\managed\NXOpen.Utilities.dll</HintPath>
  <Private>false</Private>
</Reference>
<Reference Include="NXOpenUI">
  <HintPath>C:\Program Files\Siemens\NX2312\nxbin\managed\NXOpenUI.dll</HintPath>
  <Private>false</Private>
</Reference>
```

対象の参照名（上記以外にもNX関連があれば全て同様に変更）:
- NXOpen
- NXOpen.UF
- NXOpen.Utilities
- NXOpenUI

### 3. ビルド出力フォルダから古いDLLを削除
```powershell
Remove-Item "NXCustomViewDxfExporter\bin\x64\Debug\NXOpen*.dll" -ErrorAction SilentlyContinue
Remove-Item "NXCustomViewDxfExporter\bin\x64\Debug\NXOpenUI.dll" -ErrorAction SilentlyContinue
```

### 4. リビルド
クリーンビルドを実行:
```
dotnet clean
dotnet build
```
またはMSBuildの場合:
```
msbuild /t:Clean
msbuild /t:Build
```

### 5. ビルド後の確認
ビルド出力フォルダ `bin\x64\Debug\` に以下のDLLが**存在しないこと**を確認:
- NXOpen.dll
- NXOpen.UF.dll
- NXOpen.Utilities.dll
- NXOpenUI.dll

これらはNX実行時に `C:\Program Files\Siemens\NX2312\nxbin\managed\` から自動ロードされるため、ビルド出力にコピーしてはならない。

## 注意事項
- ソースコード（CustomViewDxfExporter.cs）は変更しないこと
- `<Private>false</Private>` の設定が最も重要。これによりNX Open DLLがビルド出力にコピーされなくなり、実行時にNX側の正しいバージョンが使用される
- System.Windows.Forms等の.NET標準ライブラリの参照は変更不要
