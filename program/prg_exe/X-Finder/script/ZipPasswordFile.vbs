'Option Explicit

'<<7-Zip usage>>
'  7z a <zip_file_path> <target_dir_path> -p<password>

'��TODO���FZIP �ȊO�����k�ł���悤�ɂ���B

Const PROG_NAME = "7-Zip �Ńp�X���[�h���k (zip)"

Dim sExePath
Dim cSelectedPaths

Dim bIsContinue
bIsContinue = True

If bIsContinue = True Then
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "�f�o�b�O���[�h�ł��B"
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
            cTrgtPaths.Add sSelectedPath
        ElseIf bFolderExists = True And bFileExists = False Then
            cTrgtPaths.Add sSelectedPath
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
        MsgBox "�ΏۂƂȂ�t�@�C��/�t�H���_�����݂��܂���B", vbYes, PROG_NAME
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
                    "�ȉ����y�p�X���[�h�t�����k�z���āA�I���t�@�C���Ɠ����t�H���_�Ɋi�[���܂��B��낵���ł����H" & vbNewLine & _
                    sTrgtPathsStr & vbNewLine & _
                    vbNewLine & _
                    "(��) ���ꂼ��̃t�@�C��/�t�H���_�����k����܂��I" & vbNewLine & _
                    "     ��̈��k�t�@�C���ɂȂ��ł͂���܂���I", _
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

'**********************
'*** �p�X���[�h�ݒ� ***
'**********************
If bIsContinue = True Then
    MsgBox _
        "���k�t�@�C���̉𓀃p�X���[�h��ݒ肵�܂��B" & vbNewLine & _
        vbNewLine & _
        "(��) �I�����ꂽ�t�@�C��/�t�H���_���S�ē����p�X���[�h�ň��k����܂��B", _
        vbOKOnly, PROG_NAME
    
    Dim sPassword
    Dim sPasswordCheck
    Dim bIsReEnter
    bIsReEnter = False
    Do
        If bIsReEnter = True Then
            sPassword = InputBox( "�ēx�A�p�X���[�h����͂��Ă��������B", PROG_NAME )
        Else
            sPassword = InputBox( "�p�X���[�h����͂��Ă��������B", PROG_NAME )
        End If
        If sPassword = "" Then
            MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
            MsgBox "�����𒆒f���܂��B", vbYes, PROG_NAME
            bIsContinue = False
            bIsReEnter = False
        Else
            sPasswordCheck = InputBox( "�m�F�̂��߁A������x�p�X���[�h����͂��Ă��������B", PROG_NAME )
            If sPasswordCheck = "" Then
                MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
                MsgBox "�����𒆒f���܂��B", vbYes, PROG_NAME
                bIsContinue = False
                bIsReEnter = False
            Else
                If sPassword = sPasswordCheck Then
                    bIsReEnter = False
                Else
                    MsgBox "�p�X���[�h����v���Ă��܂���B", vbOKOnly, PROG_NAME
                    bIsReEnter = True
                End If
            End If
        End If
    Loop While bIsReEnter = True
Else
    'Do Nothing
End If

'****************
'*** ���k���s ***
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
    MsgBox "���k���������܂����B", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

Set objFSO = Nothing
Set objWshShell = Nothing
