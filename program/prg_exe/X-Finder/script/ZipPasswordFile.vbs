'Option Explicit

'<<7-Zip usage>>
'  7z a <zip_file_path> <target_dir_path> -p<password>

'★TODO★：ZIP 以外も圧縮できるようにする。

Const PROG_NAME = "7-Zip でパスワード圧縮 (zip)"

Dim sExePath
Dim cSelectedPaths

Dim bIsContinue
bIsContinue = True

If bIsContinue = True Then
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "デバッグモードです。"
        sExePath = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\aa"
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\b b"
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\d.txt"
    Else
        sExePath = WScript.Env("X-Finder") & "..\" & WScript.Env("7-Zip")
        Set cSelectedPaths = WScript.Col( WScript.Env("Selected") )
    End If
Else
    'Do Nothing
End If

'*** ファイルパスチェック ***
If bIsContinue = True Then
    If cSelectedPaths.Count = 0 Then
        MsgBox "ファイル/フォルダが選択されていません。", vbOKOnly, PROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'************************
'*** 対象ファイル選定 ***
'************************
'*** ファイル選定 ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If bIsContinue = True Then
    Dim cTrgtPaths
    Set cTrgtPaths = CreateObject("System.Collections.ArrayList")
    Dim sSelectedPath
    For Each sSelectedPath In cSelectedPaths
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( sSelectedPath )
        bFileExists = objFSO.FileExists( sSelectedPath )
        If bFolderExists = False And bFileExists = True Then
            cTrgtPaths.Add sSelectedPath
        ElseIf bFolderExists = True And bFileExists = False Then
            cTrgtPaths.Add sSelectedPath
        Else
            MsgBox "選択されたオブジェクトが存在しません" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, PROG_NAME
            MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
            bIsContinue = False
        End If
    Next
Else
    'Do Nothing
End If

'*** ファイルパスチェック ***
If bIsContinue = True Then
    If cTrgtPaths.Count = 0 Then
        MsgBox "対象となるファイル/フォルダが存在しません。", vbYes, PROG_NAME
        MsgBox "処理を中断します。", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'****************
'*** 実行確認 ***
'****************
If bIsContinue = True Then
    Dim sTrgtPath
    Dim sTrgtPathsStr
    sTrgtPathsStr = ""
    For Each sTrgtPath In cTrgtPaths
        sTrgtPathsStr = sTrgtPathsStr & vbNewLine & sTrgtPath
    Next
    Dim lAnswer
    lAnswer = MsgBox ( _
                    "以下を【パスワード付き圧縮】して、選択ファイルと同じフォルダに格納します。よろしいですか？" & vbNewLine & _
                    sTrgtPathsStr & vbNewLine & _
                    vbNewLine & _
                    "(※) それぞれのファイル/フォルダが圧縮されます！" & vbNewLine & _
                    "     一つの圧縮ファイルになる訳ではありません！", _
                    vbYesNo, _
                    PROG_NAME _
                )
    If lAnswer = vbYes Then
        'Do Nothing
    Else
        MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
        bIsContinue = False
    End If
Else
    'Do Nothing
End If

'**********************
'*** パスワード設定 ***
'**********************
If bIsContinue = True Then
    MsgBox _
        "圧縮ファイルの解凍パスワードを設定します。" & vbNewLine & _
        vbNewLine & _
        "(※) 選択されたファイル/フォルダが全て同じパスワードで圧縮されます。", _
        vbOKOnly, PROG_NAME
    
    Dim sPassword
    Dim sPasswordCheck
    Dim bIsReEnter
    bIsReEnter = False
    Do
        If bIsReEnter = True Then
            sPassword = InputBox( "再度、パスワードを入力してください。", PROG_NAME )
        Else
            sPassword = InputBox( "パスワードを入力してください。", PROG_NAME )
        End If
        If sPassword = "" Then
            MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
            MsgBox "処理を中断します。", vbYes, PROG_NAME
            bIsContinue = False
            bIsReEnter = False
        Else
            sPasswordCheck = InputBox( "確認のため、もう一度パスワードを入力してください。", PROG_NAME )
            If sPasswordCheck = "" Then
                MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
                MsgBox "処理を中断します。", vbYes, PROG_NAME
                bIsContinue = False
                bIsReEnter = False
            Else
                If sPassword = sPasswordCheck Then
                    bIsReEnter = False
                Else
                    MsgBox "パスワードが一致していません。", vbOKOnly, PROG_NAME
                    bIsReEnter = True
                End If
            End If
        End If
    Loop While bIsReEnter = True
Else
    'Do Nothing
End If

'****************
'*** 圧縮実行 ***
'****************
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
If bIsContinue = True Then
    For Each sTrgtPath In cTrgtPaths
        Dim sArchiveFilePath
        sArchiveFilePath = sTrgtPath & ".zip"
        Dim sExecCmd
        sExecCmd = """" & sExePath & """ a """ & sArchiveFilePath & """ """ & sTrgtPath & """ -p" & sPassword
        objWshShell.Run sExecCmd, 1, True
    Next
    MsgBox "圧縮が完了しました。", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

Set objFSO = Nothing
Set objWshShell = Nothing
