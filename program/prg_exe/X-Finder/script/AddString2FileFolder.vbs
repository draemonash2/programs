'Option Explicit

'####################################################################
'### �ݒ�
'####################################################################


'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�t�@�C��/�t�H���_������������t�^"

Dim lAnswer
lAnswer = MsgBox ( _
                "�t�@�C��/�t�H���_���̖����ɕ������t�^���܂��B��낵���ł����H", _
                vbYesNo, _
                PROG_NAME _
            )
If lAnswer = vbYes Then
    'Do Nothing
Else
    MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
    WScript.Quit()
End If

'*******************************************************
'* �t�@�C��/�t�H���_���擾
'*******************************************************
Dim cFilePaths
If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "�f�o�b�O���[�h�ł��B"
    Set cFilePaths = CreateObject("System.Collections.ArrayList")
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    objWshShell.Run "cmd /c echo.> ""C:\Users\draem_000\Desktop\test.txt""", 0, True
    objWshShell.Run "cmd /c mkdir ""C:\Users\draem_000\Desktop\test2""", 0, True
    cFilePaths.Add "C:\Users\draem_000\Desktop\test.txt"
    cFilePaths.Add "C:\Users\draem_000\Desktop\test2"
Else
    Set cFilePaths = WScript.Col( WScript.Env("Selected") )
End If

'*** �t�@�C���p�X�`�F�b�N ***
If cFilePaths.Count = 0 Then
    MsgBox "�t�@�C�����I������Ă��܂���", vbYes, PROG_NAME
    MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
    WScript.Quit
Else
    'Do Nothing
End If

'*******************************************************
'* �ǉ�������擾
'*******************************************************
Dim sSearchPattern
Dim sTargetStr
sSearchPattern = "(\w{4})/(\w{1,2})/(\w{1,2}) (\w{1,2}):(\w{1,2}):(\w{1,2})"
sTargetStr = Now()

Dim oRegExp
Set oRegExp = CreateObject("VBScript.RegExp")
oRegExp.Pattern = sSearchPattern                '�����p�^�[����ݒ�
oRegExp.IgnoreCase = True                       '�啶���Ə���������ʂ��Ȃ�
oRegExp.Global = True                           '������S�̂�����
Dim oMatchResult
Set oMatchResult = oRegExp.Execute(sTargetStr)  '�p�^�[���}�b�`���s
Dim sNowStr
With oMatchResult(0)
    sNowStr = String( 4 - Len( .SubMatches(0) ), "0" ) & .SubMatches(0) & _
              String( 2 - Len( .SubMatches(1) ), "0" ) & .SubMatches(1) & _
              String( 2 - Len( .SubMatches(2) ), "0" ) & .SubMatches(2) & _
              "_" & _
              String( 2 - Len( .SubMatches(3) ), "0" ) & .SubMatches(3) & _
              String( 2 - Len( .SubMatches(4) ), "0" ) & .SubMatches(4) & _
              String( 2 - Len( .SubMatches(5) ), "0" ) & .SubMatches(5)
End With
Set oMatchResult = Nothing
Set oRegExp = Nothing

Dim sAddStr
sAddStr = InputBox( "�����ɕt�^���镶�������͂��Ă�������", PROG_NAME, "_" & sNowStr )

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim oFilePath
For Each oFilePath In cFilePaths
    '*******************************************************
    '* �t�@�C��/�t�H���_������
    '*******************************************************
    
    Dim lFileOrFolder '1:�t�@�C���A2:�t�H���_�A0:�G���[�i���݂��Ȃ��p�X�j
    Dim bFolderExists
    Dim bFileExists
    bFolderExists = objFSO.FolderExists( oFilePath )
    bFileExists = objFSO.FileExists( oFilePath )
    If bFolderExists = False And bFileExists = True Then
        lFileOrFolder = 1 '�t�@�C��
    ElseIf bFolderExists = True And bFileExists = False Then
        lFileOrFolder = 2 '�t�H���_�[
    Else
        lFileOrFolder = 0 '�G���[�i���݂��Ȃ��p�X�j
    End If
    
    '*******************************************************
    '* �t�@�C��/�t�H���_���ύX
    '*******************************************************
    Dim sTrgtDirPath
    Dim sTrgtFileName
    sTrgtDirPath = Mid( oFilePath, 1, InStrRev( oFilePath, "\" ) - 1 )
    sTrgtFileName = Mid( oFilePath, InStrRev( oFilePath, "\" ) + 1, Len( oFilePath ) )
    
    If lFileOrFolder = 1 Then
        If InStr( sTrgtFileName, "." ) > 0 Then
            Dim sTrgtFileBaseName
            Dim sTrgtFileExt
            sTrgtFileExt = Mid( sTrgtFileName, InStrRev( sTrgtFileName, "." ) + 1, Len( sTrgtFileName ) )
            sTrgtFileBaseName = Mid( _
                    sTrgtFileName, _
                    InStrRev( sTrgtFileName, "\" ) + 1, _
                    InStrRev( sTrgtFileName, "." ) - InStrRev( sTrgtFileName, "\" ) - 1 _
                )
            objFSO.MoveFile _
                oFilePath, _
                sTrgtDirPath & "\" & sTrgtFileBaseName & sAddStr & "." & sTrgtFileExt
        Else
            objFSO.MoveFile _
                oFilePath, _
                sTrgtDirPath & "\" & sTrgtFileName & sAddStr
        End If
    ElseIf lFileOrFolder = 2 Then
        objFSO.MoveFolder _
            oFilePath, _
            sTrgtDirPath & "\" & sTrgtFileName & sAddStr
    Else
        MsgBox "�t�@�C��/�t�H���_���s���ł��B", vbOKOnly, PROG_NAME
        WScript.Quit()
    End If
Next

Set objFSO = Nothing
