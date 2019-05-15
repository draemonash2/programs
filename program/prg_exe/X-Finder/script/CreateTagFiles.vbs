'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "タグファイル作成"

Dim bIsContinue
bIsContinue = True

Dim sCTagExePath
Dim sGTagExePath
Dim sTrgtDirPath

If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Dim sArg
        Dim sDefaultPath
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        For Each sArg In WScript.Arguments
            If sDefaultPath = "" Then
                sDefaultPath = objFSO.GetParentFolderName( sArg )
            End If
        Next
        sCTagExePath = "C:\prg_exe\Ctags\ctags.exe"
        sGTagExePath = "C:\prg_exe\Gtags\bin\gtags.exe"
        sTrgtDirPath = InputBox( "ファイルパスを指定してください", PROG_NAME, sDefaultPath )
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        sCTagExePath = "C:\prg_exe\Ctags\ctags.exe"
        sGTagExePath = "C:\prg_exe\Gtags\bin\gtags.exe"
        sTrgtDirPath = WScript.Env("Current")
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        sCTagExePath = "C:\prg_exe\Ctags\ctags.exe"
        sGTagExePath = "C:\prg_exe\Gtags\bin\gtags.exe"
        sTrgtDirPath = "C:\codes\c"
    End If
Else
    'Do Nothing
End If

'*** タグファイル作成 ***
If bIsContinue = True Then
    MsgBox "「ctags」と「gtags」にパスが通っていない場合は、パスを通してから実行してください。", vbOk, PROG_NAME
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    objWshShell.Run "cmd /c pushd """ & sTrgtDirPath & """ & ctags -R", 0, True
    objWshShell.Run "cmd /c pushd """ & sTrgtDirPath & """ & gtags -v", 0, True
    MsgBox "タグファイルの作成が完了しました。", vbOk, PROG_NAME
Else
    'Do Nothing
End If
