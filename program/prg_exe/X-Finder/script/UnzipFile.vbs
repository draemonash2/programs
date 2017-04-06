'Option Explicit

'<<7-Zip usage>>
'  7z x -o<target_dir_path> <archive_file_path>

Const PROG_NAME = "7-Zip �ŉ�"

Dim sExePath
Dim sArchiveFilePaths

If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "�f�o�b�O���[�h�ł��B"
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
'*** �t�@�C���p�X�ꗗ�쐬 ***
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
'*** �𓀎��s ***
'****************
If UBound( asPathList ) = -1 Then
    MsgBox "�t�@�C�����I������Ă��܂���B", vbYes, PROG_NAME
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
                            "�ȉ����𓀂��܂��B��낵���ł����H" & vbNewLine & _
                            vbNewLine & _
                            "<<�A�[�J�C�u�t�@�C����>>" & vbNewLine & _
                            sArchiveFilePath & vbNewLine & _
                            vbNewLine & _
                            "<<�𓀃t�H���_>>" & vbNewLine & _
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
                MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
            End If
        ElseIf bFolderExists = True And bFileExists = False Then
            MsgBox "�t�@�C�����w�肵�Ă��������I" & vbNewLine & vbNewLine & sArchiveFilePath, vbOKOnly, PROG_NAME
        Else
            MsgBox "�w�肵���t�@�C�������݂��܂���I" & vbNewLine & vbNewLine & sArchiveFilePath, vbOKOnly, PROG_NAME
        End If
        
        Set oFileSys = Nothing
    Next
End If
