# NX Open 署名（サイン）の組み込み

## 背景
NX Open for .NETで開発したDLLをお客様環境（NX Open開発ライセンスがない端末）で実行するには、DLLにNX署名を組み込む必要がある。署名がないDLLはNXが読み込みを拒否する。

現在の開発環境では開発ライセンスがあるため署名なしでも動作しているが、お客様環境では動作しない。

## 対応内容

### 1. NXSigningResource.res をプロジェクトにコピー

まず、NXインストールディレクトリからリソースファイルを探してプロジェクトディレクトリにコピーする。

```powershell
# NXSigningResource.res の場所を探す
dir "C:\Program Files\Siemens" -Recurse -Filter "NXSigningResource.res" -ErrorAction SilentlyContinue
```

見つかったファイルを以下にコピー：
```powershell
copy "<見つかったパス>\NXSigningResource.res" "C:\Users\tcadmin\source\repos\NXOpenTest\NXCustomViewDxfExporter\NXSigningResource.res"
```

### 2. .csproj にリソースファイルを埋め込みリソースとして追加

`NXCustomViewDxfExporter.csproj` に以下を追加：

```xml
<ItemGroup>
  <EmbeddedResource Include="NXSigningResource.res" />
</ItemGroup>
```

### 3. ビルド後イベントにSignLibraryコマンドを追加

`NXCustomViewDxfExporter.csproj` の `<PropertyGroup>` 内に以下を追加：

```xml
<PostBuildEvent>"$(UGII_ROOT_DIR)\SignLibrary" "$(TargetPath)"</PostBuildEvent>
```

※ `UGII_ROOT_DIR` はNXの環境変数。もし環境変数が設定されていない場合は、SignLibrary.exeのフルパスを直接指定する。

```powershell
# SignLibrary.exe の場所を探す
dir "C:\Program Files\Siemens" -Recurse -Filter "SignLibrary.exe" -ErrorAction SilentlyContinue
```

フルパスが分かったら、PostBuildEventを以下のように書く（例）：
```xml
<PostBuildEvent>"C:\Program Files\Siemens\NX2312\NXBIN\SignLibrary.exe" "$(TargetPath)"</PostBuildEvent>
```

### 4. ビルドして署名を確認

x64 Debugでビルドし、ビルド後イベントのログにSignLibraryの実行結果が表示されることを確認する。

## 注意事項
- NXSigningResource.res は「埋め込まれたリソース（EmbeddedResource）」として追加すること。「コンテンツ」や「なし」ではダメ。
- SignLibrary.exe はNXコマンドプロンプト環境でなくても、フルパス指定で実行可能。
- 既存のコードには一切変更を加えないこと。.csprojの変更のみ。
