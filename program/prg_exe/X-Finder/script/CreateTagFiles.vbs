'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "タグファイル作成"

Dim bIsContinue
bIsContinue = True

If bIsContinue = True Then
    Dim sCTagExePath
    Dim sGTagExePath
    Dim sTrgtDirPath
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "デバッグモードです。"
        sCTagExePath = "C:\prg_exe\Ctags\ctags.exe"
        sGTagExePath = "C:\prg_exe\Gtags\bin\gtags.exe"
        sTrgtDirPath = "C:\codes\c"
    Else
        sCTagExePath = "C:\prg_exe\Ctags\ctags.exe"
        sGTagExePath = "C:\prg_exe\Gtags\bin\gtags.exe"
        sTrgtDirPath = WScript.Env("Current")
    End If
Else
    'Do Nothing
End If

'*** タグファイル作成 ***
If bIsContinue = True Then
    MsgBox "「ctags」と「gtags」にパスが通っていない場合は、パスを通してから実行してください。", vbOk, PROG_NAME
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    objWshShell.Run "cmd /c cd """ & sTrgtDirPath & """ & ctags -R", 0, True
    objWshShell.Run "cmd /c cd """ & sTrgtDirPath & """ & gtags -v", 0, True
    MsgBox "タグファイルの作成が完了しました。", vbOk, PROG_NAME
Else
    'Do Nothing
End If
