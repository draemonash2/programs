'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'####################################################################
'### �ݒ�
'####################################################################


'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�^�O�t�@�C���쐬"

Dim bIsContinue
bIsContinue = True

If bIsContinue = True Then
    Dim sCTagExePath
    Dim sGTagExePath
    Dim sTrgtDirPath
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "�f�o�b�O���[�h�ł��B"
        sCTagExePath = "C:\prg_exe\Ctags\ctags.exe"
        sGTagExePath = "C:\prg_exe\Gtags\bin\gtags.exe"
        sTrgtDirPath = "C:\codes\c"
    Else
        sCTagExePath = "C:\prg_exe\Ctags\ctags.exe"
        sGTagExePath = "C:\prg_exe\Gtags\bin\gtags.exe"
        sTrgtDirPath = WScript.Env("Current")
    End If
Else
    'Do Nothing
End If

'*** �^�O�t�@�C���쐬 ***
If bIsContinue = True Then
    MsgBox "�uctags�v�Ɓugtags�v�Ƀp�X���ʂ��Ă��Ȃ��ꍇ�́A�p�X��ʂ��Ă�����s���Ă��������B", vbOk, PROG_NAME
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    objWshShell.Run "cmd /c cd """ & sTrgtDirPath & """ & ctags -R", 0, True
    objWshShell.Run "cmd /c cd """ & sTrgtDirPath & """ & gtags -v", 0, True
    MsgBox "�^�O�t�@�C���̍쐬���������܂����B", vbOk, PROG_NAME
Else
    'Do Nothing
End If
