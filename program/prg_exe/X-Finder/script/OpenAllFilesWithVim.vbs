'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################


'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�J�����g�t�H���_�z���̓���t�@�C���� Vim �őS�ĊJ��"

Dim bIsContinue
bIsContinue = True

Dim objFSO
Dim sExePath
Dim sCurDirPath

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
        Dim sArg
        Dim sDefaultPath
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        For Each sArg In WScript.Arguments
            If sDefaultPath = "" Then
                sDefaultPath = objFSO.GetParentFolderName( sArg )
            End If
        Next
        sExePath = "C:\prg_exe\Vim\gvim.exe"
        sCurDirPath = InputBox( "�t�@�C���p�X���w�肵�Ă�������", PROG_NAME, sDefaultPath )
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        sExePath = WScript.Env("Vim")
        sCurDirPath = WScript.Env("Current")
    Else '�f�o�b�O���s
        MsgBox "�f�o�b�O���[�h�ł��B"
        sExePath = "C:\prg_exe\Vim\gvim.exe"
        sCurDirPath = "C:\codes\c"
    End If
Else
    'Do Nothing
End If

'*** �g���q�I�� ***
If bIsContinue = True Then
    Dim sExtNames
    sExtNames = InputBox( _
        "�g���q��I�����Ă��������B" & vbNewLine & _
        "�����̊g���q���w�肷�鎞�̓X�y�[�X�ŋ�؂�܂��B" & vbNewLine & _
        "  ��P�j*.txt *.c" & vbNewLine & _
        "  ��Q�j*.*" & vbNewLine & _
        "" , _
        "title", _
        "*.c *.h" _
    )
    If sExtNames = "" Then
        MsgBox "�g���q���I������Ă��܂���", vbYes, PROG_NAME
        MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** �t�@�C�����X�g�쐬 ***
If bIsContinue = True Then
    '�t�@�C�����X�g�o�̓R�}���h���s
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    
    Dim sTmpFilePath
    Dim sExecCmd
    sTmpFilePath = objWshShell.SpecialFolders("Templates") & "\open_file_list.txt"
    'MsgBox sTmpFilePath '��DEBUG��
    sExecCmd = "cd """ & sCurDirPath & """ & dir " & sExtNames & " /b /s /a:a-d > """ & sTmpFilePath & """"
    'MsgBox sExecCmd '��DEBUG��
    objWshShell.Run "cmd /c" & sExecCmd, 0, True
    
    '�o�͂����t�@�C�����X�g��荞��
    Dim objFile
    Dim sTextAll
    On Error Resume Next
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    Dim asFileList
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
        'MsgBox Err.Number '��DEBUG��
        If Err.Number = 0 Then
            sTextAll = objFile.ReadAll
            sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '�����ɉ��s���t�^����Ă��܂����߁A�폜
            asFileList = Split( sTextAll, vbNewLine )
            objFile.Close
        Else
            MsgBox "�G���[���������܂����B [ErrorNo." & Err.Number & "] " & Err.Description, vbYes, PROG_NAME
            MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
            bIsContinue = False
        End If
        Set objFile = Nothing   '�I�u�W�F�N�g�̔j��
    Else
        MsgBox "�G���[���������܂����B [ErrorNo." & Err.Number & "] " & Err.Description, vbYes, PROG_NAME
        MsgBox "�����𒆒f���܂��B", vbYes, PROG_NAME
        bIsContinue = False
    End If
    objFSO.DeleteFile sTmpFilePath, True
    Set objFSO = Nothing    '�I�u�W�F�N�g�̔j��
    On Error Goto 0
    'MsgBox Ubound(asFileList) '��DEBUG��
Else
    'Do Nothing
End If

'*** �t�@�C���I�[�v�����s ***
If bIsContinue = True Then
    Dim sFilePathList
    sFilePathList = """"
    Dim lIdx
    lIdx = 0
    For Each sFilePath In asFileList
        If lIdx = 0 Then
            sFilePathList = """" & sFilePath & """"
        Else
            sFilePathList = sFilePathList & " """ & sFilePath & """"
        End If
        lIdx = lIdx + 1
    Next
    'MsgBox sFilePathList '��DEBUG��
    
    objWshShell.Run "cmd /c " & sExePath & " " & sFilePathList, 0, False
Else
    'Do Nothing
End If
