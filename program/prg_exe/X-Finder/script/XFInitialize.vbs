'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################
Const TEMP_FILE_NAME = "xf_diff_target_path.tmp"

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "X-Finder 初期化時実行処理"

Dim bIsContinue
bIsContinue = True

Dim sTmpPath

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Dim objWshShell
        Set objWshShell = WScript.CreateObject("WScript.Shell")
        sTmpPath = objWshShell.SpecialFolders("Templates") & "\" & TEMP_FILE_NAME
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        sTmpPath = WScript.Env("Temp") & "\" & TEMP_FILE_NAME
    Else 'デバッグ実行
        sTmpPath = "C:\prg_exe\X-Finder\" & TEMP_FILE_NAME
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
