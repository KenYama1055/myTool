# 必要なモジュールのインポート
Import-Module ImportExcel

# スクリプトの引数を取得
param (
    [string]$ExcelFilePath  # 読み込むExcelファイルのパス
)

# コマンドライン引数からExcelファイルのパスを取得
if (-not $ExcelFilePath) {
    if ($args.Count -eq 1) {
        $ExcelFilePath = $args[0]
    } else {
        Write-Output "使用法: .\generate_bcp_commands.ps1 <Excelファイルのフルパス>"
        exit
    }
}

# スクリプトのディレクトリを取得
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
# 出力ファイルのパスを設定
$outputFilePath = Join-Path -Path $scriptDirectory -ChildPath "generated_bcp_commands.txt"

# テキストファイルの初期化（内容をクリア）
Out-File -FilePath $outputFilePath -Force

# Excelファイルを読み込む シート名: Parameters
$parameters = Import-Excel -Path $excelFilePath -WorksheetName "Parameters"

# 各行のパラメータからBCPコマンドを生成
foreach ($row in $parameters) {
    # テーブル名
    $tableName = $row.TableName
    # テーブルデータが格納された/テーブルデータを格納するファイルのディレクトリ
    $dataFilePath = $row.DataFilePath
    # エクスポート = "out" / インポート = "in" / フォーマットファイル作成 = "format nul"
    $direction = $row.Direction

    # オプションを生成する
    $options = ""

    if ($row.MaxErrors) { $options += " -m $($row.MaxErrors)" } # [-m 最大エラー数]
    if ($row.FormatFile) { $options += " -f `"$($row.FormatFile)`"" } # [-f フォーマット ファイル]
    if ($row.ErrorFile) { $options += " -e `"$($row.ErrorFile)`"" } # [-e エラー ファイル]
    if ($row.FirstRow) { $options += " -F $($row.FirstRow)" } # [-F 先頭行]
    if ($row.LastRow) { $options += " -L $($row.LastRow)" } # [-L 最終行]
    if ($row.BatchSize) { $options += " -b $($row.BatchSize)" } # [-b バッチ サイズ]
    if ($row.NativeType) { $options += " -n" } # [-n ネイティブ型]
    if ($row.CharacterType) { $options += " -c" } # [-c 文字型]
    if ($row.UnicodeType) { $options += " -w" } # [-w UNICODE 文字型]
    if ($row.PreserveNonText) { $options += " -N" } # [-N text 以外のネイティブ型を保持]
    if ($row.FileVersion) { $options += " -V $($row.FileVersion)" } # [-V ファイル フォーマットのバージョン]
    if ($row.QuotedIdentifier) { $options += " -q" } # [-q 引用符で囲まれた識別子]
    if ($row.CodePage) { $options += " -C $($row.CodePage)" } # [-C コード ページ指定子]
    if ($row.FieldTerminator) { $options += " -t $($row.FieldTerminator)" } # [-t フィールド ターミネータ]
    if ($row.RowTerminator) { $options += " -r $($row.RowTerminator)" } # [-r 行ターミネータ]
    if ($row.InputFile) { $options += " -i `"$($row.InputFile)`"" } # [-i 入力ファイル]
    if ($row.OutputFile) { $options += " -o `"$($row.OutputFile)`"" } # [-o 出力ファイル]
    if ($row.PacketSize) { $options += " -a $($row.PacketSize)" } # [-a パケット サイズ]
    if ($row.ServerNameOption) { $options += " -S $($row.ServerNameOption)" } # [-S サーバー名]
    if ($row.UserName) { $options += " -U $($row.UserName)" } # [-U ユーザー名]
    if ($row.Password) { $options += " -P $($row.Password)" } # [-P パスワード]
    if ($row.TrustedConnection -eq "Yes") { $options += " -T" } # [-T 信頼関係接続]
    if ($row.Version) { $options += " -v" } # [-v バージョン]
    if ($row.RegionalSettings) { $options += " -R" } # [-R 地域別設定有効]
    if ($row.PreserveNulls -eq "Yes") { $options += " -k" } # [-k NULL 値を保持]
    if ($row.KeepIdentity -eq "Yes") { $options += " -E" } # [-E ID 値を保持]
    if ($row.AzureAD -eq "Yes") { $options += " -G" } # [-G Azure Active Directory 認証]
    if ($row.GenerateXmlFormat -eq "Yes") { $options += " -x" } # [-x XML フォーマット ファイルを生成]
    if ($row.DBName) { $options += " -d $($row.DBName)" } # [-d DB 名]
    if ($row.ApplicationIntent) { $options += " -K $($row.ApplicationIntent)" } # [-K アプリケーション インテント]
    if ($row.LoginTimeout) { $options += " -l $($row.LoginTimeout)" } # [-l ログイン タイムアウト]

    # BCPコマンドの生成
    $bcpCommand = "bcp $($tableName) $($direction) `"$($dataFilePath)`" $options"
    
    # 生成されたBCPコマンドの出力（コンソールおよびファイルへ）
    Write-Output "生成されたBCPコマンド: $bcpCommand"
    Add-Content -Path $outputFilePath -Value $bcpCommand
}

Write-Output "BCPコマンドが $outputFilePath に出力されました。"
