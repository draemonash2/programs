'Option Explicit

'<<7-Zip usage>>
'  7z a <zip_file_path> <target_dir_path>

Const PROG_NAME = "7-Zip で圧縮 (zip)"

Dim sExePath
Dim sTrgtDirPaths

If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "デバッグモードです。"
    sExePath = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"
    sTrgtDirPaths = "C:\Users\draem_000\Desktop\aa ""C:\Users\draem_000\Desktop\b b"""
    'sTrgtDirPaths = """C:\Users\draem_000\Desktop\b b"""
    'sTrgtDirPaths = "C:\Users\draem_000\Desktop\aa"
    'sTrgtDirPaths = ""
Else
    sExePath = WScript.Env("X-Finder") & "..\" & WScript.Env("7-Zip")
    sTrgtDirPaths = WScript.Env("Selected")
End If

'****************************
'*** ファイルパス一覧作成 ***
'****************************
Dim asPathList()
If sTrgtDirPaths = "" Then
    ReDim Preserve asPathList(-1)
Else
    ReDim Preserve asPathList(0)
    
    Dim sCurPath
    Dim bIsPathContinue
    sCurPath = ""
    bIsPathContinue = False
    Dim lTrgtStrIdx
    For lTrgtStrIdx = 1 To Len( sTrgtDirPaths )
        Dim sTrgtChar
        sTrgtChar = Mid( sTrgtDirPaths, lTrgtStrIdx, 1 )
        If sTrgtChar = """" Then
            If bIsPathContinue = True Then
                bIsPathContinue = False
            Else
                bIsPathContinue = True
            End If
            'sCurPath = sCurPath & sTrgtChar
        ElseIf sTrgtChar = " " Then
            If bIsPathContinue = True Then
                sCurPath = sCurPath & sTrgtChar
            Else
                asPathList( UBound( asPathList ) ) = sCurPath
                ReDim Preserve asPathList( UBound( asPathList ) + 1 )
                sCurPath = ""
            End If
        Else
            sCurPath = sCurPath & sTrgtChar
        End If
    Next
    asPathList( UBound( asPathList ) ) = sCurPath
End If

'****************
'*** 圧縮実行 ***
'****************
If UBound( asPathList ) = -1 Then
    MsgBox "ファイルが選択されていません。", vbYes, PROG_NAME
Else
    Dim lPathListIdx
    For lPathListIdx = 0 To Ubound( asPathList )
        Dim sTrgtDirPath
        Dim sArchiveFilePath
        sTrgtDirPath = asPathList( lPathListIdx )
        sArchiveFilePath = sTrgtDirPath & ".zip"
        
        Dim oFileSys
        Dim bFolderExists
        Dim bFileExists
        Set oFileSys = CreateObject("Scripting.FileSystemObject")
        bFolderExists = oFileSys.FolderExists( sTrgtDirPath )
        bFileExists = oFileSys.FileExists( sTrgtDirPath )
        Set oFileSys = Nothing
        
        If bFolderExists = False And bFileExists = True Then
            MsgBox "フォルダを指定してください！" & vbNewLine & vbNewLine & sTrgtDirPath, vbOKOnly, PROG_NAME
        ElseIf bFolderExists = True And bFileExists = False Then
            Dim lAnswer
            lAnswer = MsgBox ( _
                            "以下を ZIP 圧縮します。よろしいですか？" & vbNewLine & _
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
                Dim sExecCmd
                sExecCmd = """" & sExePath & """ a """ & sArchiveFilePath & """ """ & sTrgtDirPath & """"
                Dim objWshShell
                Set objWshShell = WScript.CreateObject("WScript.Shell")
                objWshShell.Run sExecCmd, 1, True
            Else
                MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
            End If
        Else
            MsgBox "指定したフォルダが存在しません！" & vbNewLine & vbNewLine & sTrgtDirPath, vbOKOnly, PROG_NAME
        End If
    Next
End If
