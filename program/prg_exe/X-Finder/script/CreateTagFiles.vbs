'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################


'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�^�O�t�@�C���쐬"

Dim bIsContinue
bIsContinue = True

Dim sCTagExePath
Dim sGTagExePath
Dim sTrgtDirPath

If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
        Dim sArg
        Dim sDefaultPath
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        For Each sArg In WScript.Arguments
            If sDefaultPath = "" Then
                sDefaultPath = objFSO.GetParentFolderName( sArg )
            End If
        Next
        sCTagExePath = "C:\prg_exe\Ctags\ctags.exe"
        sGTagExePath = "C:\prg_exe\Gtags\bin\gtags.exe"
        sTrgtDirPath = InputBox( "�t�@�C���p�X���w�肵�Ă�������", PROG_NAME, sDefaultPath )
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        sCTagExePath = "C:\prg_exe\Ctags\ctags.exe"
        sGTagExePath = "C:\prg_exe\Gtags\bin\gtags.exe"
        sTrgtDirPath = WScript.Env("Current")
    Else '�f�o�b�O���s
        MsgBox "�f�o�b�O���[�h�ł��B"
        sCTagExePath = "C:\prg_exe\Ctags\ctags.exe"
        sGTagExePath = "C:\prg_exe\Gtags\bin\gtags.exe"
        sTrgtDirPath = "C:\codes\c"
    End If
Else
    'Do Nothing
End If

'*** �^�O�t�@�C���쐬 ***
If bIsContinue = True Then
    MsgBox "�uctags�v�Ɓugtags�v�Ƀp�X���ʂ��Ă��Ȃ��ꍇ�́A�p�X��ʂ��Ă�����s���Ă��������B", vbOk, PROG_NAME
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    objWshShell.Run "cmd /c pushd """ & sTrgtDirPath & """ & ctags -R", 0, True
    objWshShell.Run "cmd /c pushd """ & sTrgtDirPath & """ & gtags -v", 0, True
    MsgBox "�^�O�t�@�C���̍쐬���������܂����B", vbOk, PROG_NAME
Else
    'Do Nothing
End If
