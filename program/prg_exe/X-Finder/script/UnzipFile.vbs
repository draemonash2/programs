'Option Explicit

'<<7-Zip usage>>
'  7z x -o<target_dir_path> <archive_file_path>

Const PROG_NAME = "7-Zip で解凍"

Dim sExePath
Dim sArchiveFilePaths

Dim cFilePaths
If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "デバッグモードです。"
    sExePath = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"
    Set cFilePaths = CreateObject("System.Collections.ArrayList")
    cFilePaths.Add "C:\Users\draem_000\Desktop\aa.zip"
    cFilePaths.Add "C:\Users\draem_000\Desktop\b b.zip"
Else
    sExePath = WScript.Env("X-Finder") & "..\" & WScript.Env("7-Zip")
    Set cFilePaths = WScript.Col( WScript.Env("Selected") )
End If

'*** ファイルパスチェック ***
If cFilePaths.Count = 0 Then
    MsgBox "ファイルが選択されていません。", vbYes, PROG_NAME
    MsgBox "処理を中断します", vbYes, PROG_NAME
    WScript.Quit
Else
    'Do Nothing
End If

'****************
'*** 解凍実行 ***
'****************
Dim sArchiveFilePath
For Each sArchiveFilePath In cFilePaths
    Dim oFileSys
    Dim bFolderExists
    Dim bFileExists
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    bFolderExists = oFileSys.FolderExists( sArchiveFilePath )
    bFileExists = oFileSys.FileExists( sArchiveFilePath )
    
    If bFolderExists = False And bFileExists = True Then
        Dim sTrgtDirPath
        sTrgtDirPath = oFileSys.GetParentFolderName( sArchiveFilePath ) & "\" & oFileSys.GetBaseName( sArchiveFilePath )
        Dim lAnswer
        lAnswer = MsgBox ( _
                        "以下を解凍します。よろしいですか？" & vbNewLine & _
                        vbNewLine & _
                        "<<アーカイブファイル名>>" & vbNewLine & _
                        sArchiveFilePath & vbNewLine & _
                        vbNewLine & _
                        "<<解凍フォルダ>>" & vbNewLine & _
                        sTrgtDirPath, _
                        vbYesNo, _
                        PROG_NAME _
                    )
        If lAnswer = vbYes Then
            Dim sExecCmd
            sExecCmd = """" & sExePath & """ x -o""" & sTrgtDirPath & """ """ & sArchiveFilePath & """"
            Dim objWshShell
            Set objWshShell = WScript.CreateObject("WScript.Shell")
            objWshShell.Run sExecCmd, 1, True
        Else
            MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
        End If
    ElseIf bFolderExists = True And bFileExists = False Then
        MsgBox "ファイルを指定してください" & vbNewLine & vbNewLine & sArchiveFilePath, vbOKOnly, PROG_NAME
    Else
        MsgBox "指定したファイルが存在しません" & vbNewLine & vbNewLine & sArchiveFilePath, vbOKOnly, PROG_NAME
    End If
    
    Set oFileSys = Nothing
Next
