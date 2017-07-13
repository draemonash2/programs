'Option Explicit

'<<7-Zip usage>>
'  7z x -o<target_dir_path> <archive_file_path>

'特記事項：解凍対象とする拡張子を増やしたい場合、
'          「対象ファイル選定」内の「Select Case sFileExt」の
'          Case 分岐を増やしてください。

'★TODO★：ZIP 以外も解凍できるようにする。

Const PROG_NAME = "7-Zip で解凍"

Dim sExePath
Dim cSelectedPaths

Dim bIsContinue
bIsContinue = True

If bIsContinue = True Then
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "デバッグモードです。"
        sExePath = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\aa.zip"
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\b b.zip"
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\cc"
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
            Dim sFileExt
            sFileExt = objFSO.GetExtensionName( sSelectedPath )
            
            Dim bIsTrgtFile
            Select Case sFileExt
                Case "zip": bIsTrgtFile = True
                Case Else: bIsTrgtFile = False
            End Select
            If bIsTrgtFile = True Then
                cTrgtPaths.Add sSelectedPath
            Else
                'Do Nothing
            End If
        ElseIf bFolderExists = True And bFileExists = False Then
            'Do Nothing
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
        MsgBox "対象となるファイルが存在しません。", vbYes, PROG_NAME
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
                    "以下を【解凍】して、選択ファイルと同じフォルダに格納します。よろしいですか？" & vbNewLine & _
                    sTrgtPathsStr, _
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

'****************
'*** 解凍実行 ***
'****************
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
If bIsContinue = True Then
    For Each sTrgtPath In cTrgtPaths
        Dim sOutputDirPath
        sOutputDirPath = objFSO.GetParentFolderName( sTrgtPath )
        Dim sExecCmd
        sExecCmd = """" & sExePath & """ x -o""" & sOutputDirPath & """ """ & sTrgtPath & """"
        objWshShell.Run sExecCmd, 1, True
    Next
    MsgBox "解凍が完了しました。", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

Set objFSO = Nothing
Set objWshShell = Nothing
