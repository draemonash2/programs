'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################


'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�v���O����No.���R�s�["

Dim bIsContinue
bIsContinue = True

Dim cFilePaths

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        Dim sArg
        For Each sArg In WScript.Arguments
            cFilePaths.add sArg
        Next
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        Set cFilePaths = WScript.Col( WScript.Env("Selected") )
    Else '�f�o�b�O���s
        MsgBox "�f�o�b�O���[�h�ł��B"
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\aabbbbb.txt"
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\b b"
    End If
Else
    'Do Nothing
End If

'*** �t�@�C���p�X�`�F�b�N ***
If bIsContinue = True Then
    If cFilePaths.Count = 0 Then
        MsgBox "�t�@�C�����I������Ă��܂���", vbYes, PROG_NAME
        MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** �N���b�v�{�[�h�փR�s�[ ***
If bIsContinue = True Then
    Dim sOutString
    Dim bFirstStore
    bFirstStore = True
    Dim objTxtFile
    Dim sPrgNo
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sOutString = ""
    For Each oFilePath In cFilePaths
        Set objTxtFile = objFSO.OpenTextFile( oFilePath, 1, False)
        sPrgNo = objTxtFile.ReadLine
        sPrgNo = Replace( sPrgNo, "/* ", "" )
        sPrgNo = Replace( sPrgNo, " */", "" )
        If bFirstStore = True Then
            sOutString = sPrgNo
            bFirstStore = False
        Else
            sOutString = sOutString & vbNewLine & sPrgNo
        End If
        objTxtFile.Close
    Next
    Set objTxtFile = Nothing
    Set objFSO = Nothing
    CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
    'Do Nothing
End If
