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

    # 引数なしのオプションについては"y"(yes)でオプション追加するよう設定。
    # 使用頻度が高めのオプション
    if ($row.ServerNameOption) { $options += " -S $($row.ServerNameOption)" } # [-S サーバー名]引数あり
    if ($row.UserName) { $options += " -U $($row.UserName)" } # [-U ユーザー名]引数あり
    if ($row.Password) { $options += " -P $($row.Password)" } # [-P パスワード]引数あり
    if ($row.TrustedConnection -eq "y") { $options += " -T" } # [-T 信頼関係接続]
    if ($row.DBName) { $options += " -d $($row.DBName)" } # [-d DB 名]引数あり
    if ($row.FieldTerminator) { $options += " -t $($row.FieldTerminator)" } # [-t フィールド ターミネータ]引数あり
    if ($row.RowTerminator) { $options += " -r $($row.RowTerminator)" } # [-r 行ターミネータ]引数あり
    if ($row.PreserveNulls -eq "y") { $options += " -k" } # [-k NULL 値を保持]
    if ($row.FirstRow) { $options += " -F $($row.FirstRow)" } # [-F 先頭行]引数あり
    if ($row.LastRow) { $options += " -L $($row.LastRow)" } # [-L 最終行]引数あり
    if ($row.FormatOption) { $options += " -c $($row.FormatOption)" } # フォーマットオプション [-c 文字型]/[-n ネイティブ型]/[-w UNICODE文字型]/[-N text以外のネイティブ型を保持]
    if ($row.GenerateXmlFormat -eq "y") { $options += " -x" } # [-x XML フォーマット ファイルを生成]
    if ($row.FormatFile) { $options += " -f `"$($row.FormatFile)`"" } # [-f フォーマット ファイル]引数あり
    if ($row.OutputFile) { $options += " -o `"$($row.OutputFile)`"" } # [-o 出力ファイル]引数あり

    # 使用頻度が低めのオプション
    if ($row.MaxErrors) { $options += " -m $($row.MaxErrors)" } # [-m 最大エラー数]引数あり
    if ($row.ErrorFile) { $options += " -e `"$($row.ErrorFile)`"" } # [-e エラー ファイル]引数あり
    if ($row.BatchSize) { $options += " -b $($row.BatchSize)" } # [-b バッチ サイズ]引数あり

    #フォーマットオプション [-c]/[-n]/[-w]/[-N]については併用できないオプションのため、変数名「FormatOption」として取りまとめる。
    # if ($row.CharacterType -eq "y") { $options += " -c" } # [-c 文字型]
    # if ($row.NativeType -eq "y") { $options += " -n" } # [-n ネイティブ型]
    # if ($row.UnicodeType -eq "y") { $options += " -w" } # [-w UNICODE 文字型]
    # if ($row.PreserveNonText -eq "y") { $options += " -N" } # [-N text 以外のネイティブ型を保持]

    if ($row.FileVersion) { $options += " -V $($row.FileVersion)" } # [-V ファイル フォーマットのバージョン]引数あり
    if ($row.QuotedIdentifier -eq "y") { $options += " -q" } # [-q 引用符で囲まれた識別子]
    if ($row.CodePage) { $options += " -C $($row.CodePage)" } # [-C コード ページ指定子]引数あり
    if ($row.InputFile) { $options += " -i `"$($row.InputFile)`"" } # [-i 入力ファイル]引数あり
    if ($row.PacketSize) { $options += " -a $($row.PacketSize)" } # [-a パケット サイズ]引数あり
    # if ($row.Version -eq "y") { $options += " -v" } # [-v バージョン]
    if ($row.RegionalSettings -eq "y") { $options += " -R" } # [-R 地域別設定有効]
    if ($row.KeepIdentity -eq "y") { $options += " -E" } # [-E ID 値を保持]
    if ($row.AzureAD -eq "y") { $options += " -G" } # [-G Azure Active Directory 認証]
    if ($row.LoadHints) { $options += " -h `"$($row.LoadHints)`"" }# [-h "読み込みヒント"]引数あり
    if ($row.ApplicationIntent) { $options += " -K $($row.ApplicationIntent)" } # [-K アプリケーション インテント]引数あり
    if ($row.LoginTimeout) { $options += " -l $($row.LoginTimeout)" } # [-l ログイン タイムアウト]引数あり

    # BCPコマンドの生成
    $bcpCommand = "bcp $($tableName) $($direction) `"$($dataFilePath)`" $options"
    
    # 生成されたBCPコマンドの出力（コンソールおよびファイルへ）
    Write-Output "生成されたBCPコマンド: $bcpCommand"
    Add-Content -Path $outputFilePath -Value $bcpCommand
}

Write-Output "BCPコマンドが $outputFilePath に出力されました。"
