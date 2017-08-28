'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'####################################################################
'### 設定
'####################################################################
Const DIFF_TRGT_PATH_FILE_NAME = "diff_target_path.tmp"

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "X-Finder 初期化時実行処理"

Dim bIsContinue
bIsContinue = True

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    Dim sTmpPath
    If PRODUCTION_ENVIRONMENT = 0 Then
        sTmpPath = "C:\prg_exe\X-Finder\" & DIFF_TRGT_PATH_FILE_NAME
    Else
        sTmpPath = WScript.Env("X-Finder") & DIFF_TRGT_PATH_FILE_NAME
    End If
Else
    'Do Nothing
End If

'*** diff_target_path.tmp 削除 ***
If bIsContinue = True Then
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists( sTmpPath ) Then
      objFSO.DeleteFile sTmpPath, True
    Else
      'Do Nothing
    End If
Else
    'Do Nothing
End If
