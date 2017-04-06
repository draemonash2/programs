'Option Explicit

'<<7-Zip usage>>
'  7z x -o<target_dir_path> <archive_file_path>

Const PROG_NAME = "7-Zip で解凍"

Dim sExePath
Dim sArchiveFilePaths

If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "デバッグモードです。"
    sExePath = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"
    sArchiveFilePaths = "C:\Users\draem_000\Desktop\aa.zip ""C:\Users\draem_000\Desktop\b b.zip"""
    'sArchiveFilePaths = """C:\Users\draem_000\Desktop\b b.zip"""
    'sArchiveFilePaths = "C:\Users\draem_000\Desktop\aa.zip"
    'sArchiveFilePaths = ""
Else
    sExePath = WScript.Env("X-Finder") & "..\" & WScript.Env("7-Zip")
    sArchiveFilePaths = WScript.Env("Selected")
End If

'****************************
'*** ファイルパス一覧作成 ***
'****************************
Dim asPathList()
If sArchiveFilePaths = "" Then
    ReDim Preserve asPathList(-1)
Else
    ReDim Preserve asPathList(0)
    
    Dim sCurPath
    Dim bIsPathContinue
    sCurPath = ""
    bIsPathContinue = False
    Dim lTrgtStrIdx
    For lTrgtStrIdx = 1 To Len( sArchiveFilePaths )
        Dim sTrgtChar
        sTrgtChar = Mid( sArchiveFilePaths, lTrgtStrIdx, 1 )
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
'*** 解凍実行 ***
'****************
If UBound( asPathList ) = -1 Then
    MsgBox "ファイルが選択されていません。", vbYes, PROG_NAME
Else
    Dim lPathListIdx
    For lPathListIdx = 0 To UBound( asPathList )
        Dim sArchiveFilePath
        sArchiveFilePath = asPathList( lPathListIdx )
        
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
            MsgBox "ファイルを指定してください！" & vbNewLine & vbNewLine & sArchiveFilePath, vbOKOnly, PROG_NAME
        Else
            MsgBox "指定したファイルが存在しません！" & vbNewLine & vbNewLine & sArchiveFilePath, vbOKOnly, PROG_NAME
        End If
        
        Set oFileSys = Nothing
    Next
End If
