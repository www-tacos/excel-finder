<#====================================================================================================
Excel�����c�[��
====================================================================================================#>
<#--------------------------------------------------
�֐���`
--------------------------------------------------#>
# ��ԍ���R1C1�`������A1�`���ɕϊ�����֐�
# ��) 50 -> AX
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
���O�ݒ�
--------------------------------------------------#>
$ROOT = (Get-Location).Path

$CSV_DIR = Join-Path $ROOT 'results'
if ( ! (Test-Path $CSV_DIR -PathType Container) ) {
  New-Item -ItemType Directory -Path $CSV_DIR | Out-Null
}

$dt = (Get-Date)
$CSV = Join-Path $CSV_DIR "result.$($dt.ToString('yyyyMMdd')).$($dt.ToString('HHmmss')).csv"


<#--------------------------------------------------
�f�B���N�g���w��
--------------------------------------------------#>
Clear-Host
Write-Host "�Ώۃf�B���N�g�����w�肵�Ă�������"
Write-Host "--------------------------------------------------"
while($True) {
  $dpath = Read-Host
  if( ! (Test-Path $dpath -PathType Container) ) {
    Write-Warning "�����ȃf�B���N�g���p�X�ł�"
    Write-Host ("`n" * 2)
  } else {
    break
  }
}


<#--------------------------------------------------
Excel�t�@�C���ꗗ�擾
--------------------------------------------------#>
$files = (Get-ChildItem $dpath -Include "*.xlsx","*.xls","*.xlt","*.xlsm","*.xlm" -Recurse)


<#--------------------------------------------------
�L�[���[�h�w��
--------------------------------------------------#>
Write-Host ("`n" * 2)
Write-Host "�L�[���[�h���w�肵�Ă�������"
Write-Host "--------------------------------------------------"
$keyword = Read-Host


<#--------------------------------------------------
�p�X���[�h�w��
--------------------------------------------------#>
Write-Host ("`n" * 2)
Write-Host "�p�X���[�h���������Ă����ꍇ�Ɏ����p�X���[�h���w�肵�Ă�������"
Write-Host "���p�X���[�h���ݒ�̃t�@�C���ɑ΂��Ă͉e�����܂���"
Write-Host "���p�X���[�h���قȂ�ꍇ�͍Ō�ɊJ���Ȃ������t�@�C���Ƃ��ĕ\������܂�"
Write-Host "--------------------------------------------------"
$password = Read-Host


<#--------------------------------------------------
�m�F
--------------------------------------------------#>
Write-Host ("`n" * 2)
Write-Host "�ȉ��̐ݒ�Ō������܂��B��낵���ł����H [y/n]"
Write-Host "--------------------------------------------------"
Write-Host "�Ώۃf�B���N�g�� : ${dpath}"
Write-Host "�Ώۃt�@�C����   : $( $files.Count )"
Write-Host "�����L�[���[�h   : ${keyword}"
Write-Host "�g�p�p�X���[�h   : ${password}"
Write-Host "--------------------------------------------------"
while ($True) {
  $yn = Read-Host
  if( $yn -notin @('y', 'n') ) {
    Write-Warning "y�܂���n����͂��Ă�������"
    Write-Host ("`n" * 2)
  } else {
    break
  }
}


<#--------------------------------------------------
�������s
--------------------------------------------------#>
$results = @()
$errors = @()

$EXCELAPP = New-Object -ComObject Excel.Application
$EXCELAPP.Visible = $False

# �e�t�@�C���ɑ΂��ČJ��Ԃ�����
Clear-Host
$files | % { $cnt = 0 } {
  $file = $_
  $cnt += 1

  try {
    try {
      # 1: �t�@�C���p�X
      # 2: 0�Ȃ�V�[�g���̊O���Q�Ƃ��X�V���Ȃ�
      # 3: True�Ȃ�ǂݎ���p�ŊJ��
      # 4: �e�L�X�g�t�@�C�����J���ꍇ�̋�؂蕶���A�s�v�Ȃ̂�Missing�ŃX�L�b�v
      # 5: �p�X���[�h���������Ă���ꍇ�Ɏ����p�X���[�h
      $wb = $EXCELAPP.Workbooks.Open($file, 0, $True, [Type]::Missing, $password)
    } catch {
      $errors += $file.FullName
      throw New-Object System.IO.IOException
    }

    Write-Host "($($cnt)/$($files.Count))"
    Write-Host "Book: $($file.Name)"

    # �e�V�[�g�ɑ΂��ČJ��Ԃ�����
    $wb.Worksheets | ForEach-Object {
      $ws = $_
      $wsName = $ws.Name
      # �ŏ��̌������ʂ��o���Ă���
      $first = $found = $ws.Cells.Find($keyword)
      Write-Host "  Sheet: ${wsName}"
      while ($null -ne $found) {
        Write-Host "    Cell: $($found.Text)" -BackgroundColor Yellow -ForegroundColor Black
        $result = New-Object PSObject | Select-Object Path, Sheet, Pos, Text
        $result.Path  = $file.FullName
        $result.Sheet = $wsName
        # $result.Pos   = "$($found.Row),$($found.Column)"  # R1C1�`��
        $result.Pos   = "$(R1C1_to_A1($found.Column))$($found.Row)"  # A1�`��
        $result.Text  = $found.Text
        $results += $result
        $found = $ws.Cells.FindNext($found)
        # �������ʂ�1�ڂɖ߂��Ă�����I��
        if ($found.Address() -eq $first.Address()) {
          break
        }
      }
    }
    $wb.Close(0)
  } catch [System.IO.IOException] {
    # �t�@�C�����J���Ȃ��ꍇ�̓X�L�b�v
  } catch {
    # �����I���Ȃǂ����ꍇ�̓G�N�Z�������
    Write-Error $PSItem.Exception
    $EXCELAPP.Quit()
    # �����I��GC
    $ws = $null
    $wb = $null
    $EXCELAPP = $null
    # [System.GC]::Collect([System.GC]::MaxGeneration)
    [System.GC]::Collect()
    Write-Host "�������ُ�I�����܂���"
    Write-Host "�I������ɂ�Enter�������Ă�������"
    Read-Host
    exit 1
  }
}
$EXCELAPP.Quit()

# �����I��GC
$ws = $null
$wb = $null
$EXCELAPP = $null
# [System.GC]::Collect([System.GC]::MaxGeneration)
[System.GC]::Collect()


<#--------------------------------------------------
�G���[���̕\��
--------------------------------------------------#>
if($errors.Count -gt 0) {
  Write-Host ("`n" * 2)
  Write-Host "�ȉ��̃t�@�C���͊J���܂���ł���"
  Write-Host "--------------------------------------------------"
  $errors | % { Write-Host $_ -ForegroundColor Red }
}


<#--------------------------------------------------
�����̕ۑ�
--------------------------------------------------#>
Write-Host ("`n" * 2)
if ($results.Count -gt 0) {
  $results | Export-Csv -Path $CSV -Encoding UTF8 -NoTypeInformation
  Write-Host "�������ʂ�ۑ����܂���"
  Write-Host "${CSV}"
} else {
  Write-Host "�������ʂ�����܂���ł���"
}

Write-Host ("`n" * 2)
Write-Host "�������I�����܂���"
Write-Host "�I������ɂ�Enter�������Ă�������"
Read-Host
exit 0
