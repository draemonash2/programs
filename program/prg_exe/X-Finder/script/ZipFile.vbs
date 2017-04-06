'Option Explicit

'<<7-Zip usage>>
'  7z a <zip_file_path> <target_dir_path>

Const PROG_NAME = "7-Zip �ň��k (zip)"

Dim sExePath
Dim sTrgtDirPaths

If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "�f�o�b�O���[�h�ł��B"
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
'*** �t�@�C���p�X�ꗗ�쐬 ***
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
'*** ���k���s ***
'****************
If UBound( asPathList ) = -1 Then
    MsgBox "�t�@�C�����I������Ă��܂���B", vbYes, PROG_NAME
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
            MsgBox "�t�H���_���w�肵�Ă��������I" & vbNewLine & vbNewLine & sTrgtDirPath, vbOKOnly, PROG_NAME
        ElseIf bFolderExists = True And bFileExists = False Then
            Dim lAnswer
            lAnswer = MsgBox ( _
                            "�ȉ��� ZIP ���k���܂��B��낵���ł����H" & vbNewLine & _
                            vbNewLine & _
                            "<<�Ώۃt�H���_>>" & vbNewLine & _
                            sTrgtDirPath & vbNewLine & _
                            vbNewLine & _
                            "<<Zip �t�@�C����>>" & vbNewLine & _
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
                MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
            End If
        Else
            MsgBox "�w�肵���t�H���_�����݂��܂���I" & vbNewLine & vbNewLine & sTrgtDirPath, vbOKOnly, PROG_NAME
        End If
    Next
End If
