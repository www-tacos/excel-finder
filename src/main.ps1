<#====================================================================================================
Excel検索ツール
====================================================================================================#>
<#--------------------------------------------------
関数定義
--------------------------------------------------#>
# 列番号をR1C1形式からA1形式に変換する関数
# 例) 50 -> AX
function R1C1_to_A1([int] $n) {
  if($n -le 0) { throw "Error: argument should be greater than 0" }
  $rest = $n % 26
  if($rest -eq 0) { $rest = 26; $n -= 1 }
  $over = [Math]::Truncate($n / 26)
  if($over -gt 0) {
    $pre = "$(R1C1_to_A1($over))"
  } else {
    $pre = ''
  }
  $az = [System.Convert]::ToChar($rest+64)
  return "${pre}${az}"
}


<#--------------------------------------------------
事前設定
--------------------------------------------------#>
$ROOT = (Get-Location).Path

$CSV_DIR = Join-Path $ROOT 'results'
if ( ! (Test-Path $CSV_DIR -PathType Container) ) {
  New-Item -ItemType Directory -Path $CSV_DIR | Out-Null
}

$dt = (Get-Date)
$CSV = Join-Path $CSV_DIR "result.$($dt.ToString('yyyyMMdd')).$($dt.ToString('HHmmss')).csv"


<#--------------------------------------------------
ディレクトリ指定
--------------------------------------------------#>
Clear-Host
Write-Host "対象ディレクトリを指定してください"
Write-Host "--------------------------------------------------"
while($True) {
  $dpath = Read-Host
  if( ! (Test-Path $dpath -PathType Container) ) {
    Write-Warning "無効なディレクトリパスです"
    Write-Host ("`n" * 2)
  } else {
    break
  }
}


<#--------------------------------------------------
Excelファイル一覧取得
--------------------------------------------------#>
$files = (Get-ChildItem $dpath -Include "*.xlsx","*.xls","*.xlt","*.xlsm","*.xlm" -Recurse)


<#--------------------------------------------------
キーワード指定
--------------------------------------------------#>
Write-Host ("`n" * 2)
Write-Host "キーワードを指定してください"
Write-Host "--------------------------------------------------"
$keyword = Read-Host


<#--------------------------------------------------
パスワード指定
--------------------------------------------------#>
Write-Host ("`n" * 2)
Write-Host "パスワードがかかっていた場合に試すパスワードを指定してください"
Write-Host "※パスワード未設定のファイルに対しては影響しません"
Write-Host "※パスワードが異なる場合は最後に開けなかったファイルとして表示されます"
Write-Host "--------------------------------------------------"
$password = Read-Host


<#--------------------------------------------------
確認
--------------------------------------------------#>
Write-Host ("`n" * 2)
Write-Host "以下の設定で検索します。よろしいですか？ [y/n]"
Write-Host "--------------------------------------------------"
Write-Host "対象ディレクトリ : ${dpath}"
Write-Host "対象ファイル数   : $( $files.Count )"
Write-Host "検索キーワード   : ${keyword}"
Write-Host "使用パスワード   : ${password}"
Write-Host "--------------------------------------------------"
while ($True) {
  $yn = Read-Host
  if( $yn -notin @('y', 'n') ) {
    Write-Warning "yまたはnを入力してください"
    Write-Host ("`n" * 2)
  } else {
    break
  }
}


<#--------------------------------------------------
検索実行
--------------------------------------------------#>
$results = @()
$errors = @()

$EXCELAPP = New-Object -ComObject Excel.Application
$EXCELAPP.Visible = $False

# 各ファイルに対して繰り返し処理
Clear-Host
$files | % { $cnt = 0 } {
  $file = $_
  $cnt += 1

  try {
    try {
      # 1: ファイルパス
      # 2: 0ならシート内の外部参照を更新しない
      # 3: Trueなら読み取り専用で開く
      # 4: テキストファイルを開く場合の区切り文字、不要なのでMissingでスキップ
      # 5: パスワードがかかっている場合に試すパスワード
      $wb = $EXCELAPP.Workbooks.Open($file, 0, $True, [Type]::Missing, $password)
    } catch {
      $errors += $file.FullName
      throw New-Object System.IO.IOException
    }

    Write-Host "($($cnt)/$($files.Count))"
    Write-Host "Book: $($file.Name)"

    # 各シートに対して繰り返し処理
    $wb.Worksheets | ForEach-Object {
      $ws = $_
      $wsName = $ws.Name
      # 最初の検索結果を覚えておく
      $first = $found = $ws.Cells.Find($keyword)
      Write-Host "  Sheet: ${wsName}"
      while ($null -ne $found) {
        Write-Host "    Cell: $($found.Text)" -BackgroundColor Yellow -ForegroundColor Black
        $result = New-Object PSObject | Select-Object Path, Sheet, Pos, Text
        $result.Path  = $file.FullName
        $result.Sheet = $wsName
        # $result.Pos   = "$($found.Row),$($found.Column)"  # R1C1形式
        $result.Pos   = "$(R1C1_to_A1($found.Column))$($found.Row)"  # A1形式
        $result.Text  = $found.Text
        $results += $result
        $found = $ws.Cells.FindNext($found)
        # 検索結果が1つ目に戻ってきたら終了
        if ($found.Address() -eq $first.Address()) {
          break
        }
      }
    }
    $wb.Close(0)
  } catch [System.IO.IOException] {
    # ファイルを開けない場合はスキップ
  } catch {
    # 強制終了などした場合はエクセルを閉じる
    Write-Error $PSItem.Exception
    $EXCELAPP.Quit()
    # 明示的なGC
    $ws = $null
    $wb = $null
    $EXCELAPP = $null
    # [System.GC]::Collect([System.GC]::MaxGeneration)
    [System.GC]::Collect()
    Write-Host "処理が異常終了しました"
    Write-Host "終了するにはEnterを押してください"
    Read-Host
    exit 1
  }
}
$EXCELAPP.Quit()

# 明示的なGC
$ws = $null
$wb = $null
$EXCELAPP = $null
# [System.GC]::Collect([System.GC]::MaxGeneration)
[System.GC]::Collect()


<#--------------------------------------------------
エラー分の表示
--------------------------------------------------#>
if($errors.Count -gt 0) {
  Write-Host ("`n" * 2)
  Write-Host "以下のファイルは開けませんでした"
  Write-Host "--------------------------------------------------"
  $errors | % { Write-Host $_ -ForegroundColor Red }
}


<#--------------------------------------------------
検索の保存
--------------------------------------------------#>
Write-Host ("`n" * 2)
if ($results.Count -gt 0) {
  $results | Export-Csv -Path $CSV -Encoding UTF8 -NoTypeInformation
  Write-Host "検索結果を保存しました"
  Write-Host "${CSV}"
} else {
  Write-Host "検索結果がありませんでした"
}

Write-Host ("`n" * 2)
Write-Host "処理が終了しました"
Write-Host "終了するにはEnterを押してください"
Read-Host
exit 0
