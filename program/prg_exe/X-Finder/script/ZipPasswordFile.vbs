'Option Explicit

'<<7-Zip usage>>
'  7z a <zip_file_path> <target_dir_path> -p<password>

Const PROG_NAME = "7-Zip でパスワード圧縮 (zip)"

Dim sExePath
Dim cFilePaths

If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "デバッグモードです。"
    sExePath = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"
    Set cFilePaths = CreateObject("System.Collections.ArrayList")
    cFilePaths.Add "C:\Users\draem_000\Desktop\aa"
    cFilePaths.Add "C:\Users\draem_000\Desktop\b b"
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
'*** 圧縮実行 ***
'****************
Dim sTrgtDirPath
For Each sTrgtDirPath In cFilePaths
    Dim sArchiveFilePath
    sArchiveFilePath = sTrgtDirPath & ".zip"
    
    Dim oFileSys
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    Dim bFolderExists
    Dim bFileExists
    bFolderExists = oFileSys.FolderExists( sTrgtDirPath )
    bFileExists = oFileSys.FileExists( sTrgtDirPath )
    
    If bFolderExists = False And bFileExists = True Then
        MsgBox "フォルダを指定してください" & vbNewLine & vbNewLine & sTrgtDirPath, vbOKOnly, PROG_NAME
    ElseIf bFolderExists = True And bFileExists = False Then
        Dim lAnswer
        lAnswer = MsgBox ( _
                        "以下をパスワード付き ZIP 圧縮します。よろしいですか？" & vbNewLine & _
                        vbNewLine & _
                        "<<対象フォルダ>>" & vbNewLine & _
                        sTrgtDirPath & vbNewLine & _
                        vbNewLine & _
                        "<<Zip ファイル名>>" & vbNewLine & _
                        sArchiveFilePath , _
                        vbYesNo, _
                        PROG_NAME _
                    )
        If lAnswer = vbYes Then
            Dim sPassword
            Dim sPasswordCheck
            Dim sExecCmd
            Dim bIsContinue
            Do
                sPassword = InputBox( "パスワードを入力してください。", PROG_NAME )
                If sPassword = "" Then
                    MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
                    bIsContinue = False
                Else
                    sPasswordCheck = InputBox( "もう一度パスワードを入力してください。", PROG_NAME )
                    If sPassword = sPasswordCheck Then
                        sExecCmd = """" & sExePath & """ a """ & sArchiveFilePath & """ """ & sTrgtDirPath & """ -p" & sPassword
                        WScript.CreateObject("WScript.Shell").Run sExecCmd, 1, True
                        bIsContinue = False
                    Else
                        MsgBox "パスワードが一致しません。", vbOKOnly, PROG_NAME
                        bIsContinue = True
                    End If
                End If
            Loop While bIsContinue = True
        Else
            MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
        End If
    Else
        MsgBox "指定したフォルダが存在しません" & vbNewLine & vbNewLine & sTrgtDirPath, vbOKOnly, PROG_NAME
    End If
Next
Set oFileSys = Nothing
Set objWshShell = Nothing
