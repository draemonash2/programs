'Option Explicit

'<<7-Zip usage>>
'  7z x -o<target_dir_path> <archive_file_path>

'���L�����F�𓀑ΏۂƂ���g���q�𑝂₵�����ꍇ�A
'          �u�Ώۃt�@�C���I��v���́uSelect Case sFileExt�v��
'          Case ����𑝂₵�Ă��������B

'��TODO���FZIP �ȊO���𓀂ł���悤�ɂ���B

Const PROG_NAME = "7-Zip �ŉ�"

Dim sExePath
Dim cSelectedPaths

Dim bIsContinue
bIsContinue = True

If bIsContinue = True Then
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "�f�o�b�O���[�h�ł��B"
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

'*** �t�@�C���p�X�`�F�b�N ***
If bIsContinue = True Then
    If cSelectedPaths.Count = 0 Then
        MsgBox "�t�@�C��/�t�H���_���I������Ă��܂���B", vbOKOnly, PROG_NAME
        MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'************************
'*** �Ώۃt�@�C���I�� ***
'************************
'*** �t�@�C���I�� ***
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
            MsgBox "�I�����ꂽ�I�u�W�F�N�g�����݂��܂���" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, PROG_NAME
            MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
            bIsContinue = False
        End If
    Next
Else
    'Do Nothing
End If

'*** �t�@�C���p�X�`�F�b�N ***
If bIsContinue = True Then
    If cTrgtPaths.Count = 0 Then
        MsgBox "�ΏۂƂȂ�t�@�C�������݂��܂���B", vbYes, PROG_NAME
        MsgBox "�����𒆒f���܂��B", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'****************
'*** ���s�m�F ***
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
                    "�ȉ����y�𓀁z���āA�I���t�@�C���Ɠ����t�H���_�Ɋi�[���܂��B��낵���ł����H" & vbNewLine & _
                    sTrgtPathsStr, _
                    vbYesNo, _
                    PROG_NAME _
                )
    If lAnswer = vbYes Then
        'Do Nothing
    Else
        MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
        MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
        bIsContinue = False
    End If
Else
    'Do Nothing
End If

'****************
'*** �𓀎��s ***
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
    MsgBox "�𓀂��������܂����B", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

Set objFSO = Nothing
Set objWshShell = Nothing
