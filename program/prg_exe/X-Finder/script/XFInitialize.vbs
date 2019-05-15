'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################
Const TEMP_FILE_NAME = "xf_diff_target_path.tmp"

'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "X-Finder �����������s����"

Dim bIsContinue
bIsContinue = True

Dim sTmpPath

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
        Dim objWshShell
        Set objWshShell = WScript.CreateObject("WScript.Shell")
        sTmpPath = objWshShell.SpecialFolders("Templates") & "\" & TEMP_FILE_NAME
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        sTmpPath = WScript.Env("Temp") & "\" & TEMP_FILE_NAME
    Else '�f�o�b�O���s
        sTmpPath = "C:\prg_exe\X-Finder\" & TEMP_FILE_NAME
    End If
Else
    'Do Nothing
End If

'*** diff_target_path.tmp �폜 ***
If bIsContinue = True Then
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists( sTmpPath ) Then
      objFSO.DeleteFile sTmpPath, True
    Else
      'Do Nothing
    End If
Else
    'Do Nothing
End If
