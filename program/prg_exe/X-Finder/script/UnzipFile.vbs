'Option Explicit

'<<7-Zip usage>>
'  7z x -o<target_dir_path> <archive_file_path>

Const PROG_NAME = "7-Zip �ŉ�"

Dim sExePath
Dim sArchiveFilePaths

Dim cFilePaths
If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "�f�o�b�O���[�h�ł��B"
    sExePath = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"
    Set cFilePaths = CreateObject("System.Collections.ArrayList")
    cFilePaths.Add "C:\Users\draem_000\Desktop\aa.zip"
    cFilePaths.Add "C:\Users\draem_000\Desktop\b b.zip"
Else
    sExePath = WScript.Env("X-Finder") & "..\" & WScript.Env("7-Zip")
    Set cFilePaths = WScript.Col( WScript.Env("Selected") )
End If

'*** �t�@�C���p�X�`�F�b�N ***
If cFilePaths.Count = 0 Then
    MsgBox "�t�@�C�����I������Ă��܂���B", vbYes, PROG_NAME
    MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
    WScript.Quit
Else
    'Do Nothing
End If

'****************
'*** �𓀎��s ***
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
        MsgBox "�t�@�C�����w�肵�Ă�������" & vbNewLine & vbNewLine & sArchiveFilePath, vbOKOnly, PROG_NAME
    Else
        MsgBox "�w�肵���t�@�C�������݂��܂���" & vbNewLine & vbNewLine & sArchiveFilePath, vbOKOnly, PROG_NAME
    End If
    
    Set oFileSys = Nothing
Next
