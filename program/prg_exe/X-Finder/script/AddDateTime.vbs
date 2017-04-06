'Option Explicit

Const PROG_NAME = "���ݎ����t�^"

Dim lAnswer
lAnswer = MsgBox ( _
                "�t�@�C��/�t�H���_���Ɍ��ݎ�����t�^���܂��B��낵���ł����H", _
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
Dim sTrgtFilePath

If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "�f�o�b�O���[�h�ł��B"
    sTrgtFilePath = WScript.Arguments(0)
Else
    sTrgtFilePath = WScript.Env("Focused")
End If

'*******************************************************
'* ���ݓ����擾
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

'*******************************************************
'* �t�@�C��/�t�H���_������
'*******************************************************
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim lFileOrFolder '1:�t�@�C���A2:�t�H���_�A0:�G���[�i���݂��Ȃ��p�X�j
Dim bFolderExists
Dim bFileExists
bFolderExists = objFSO.FolderExists( sTrgtFilePath )
bFileExists = objFSO.FileExists( sTrgtFilePath )
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
sTrgtDirPath = Mid( sTrgtFilePath, 1, InStrRev( sTrgtFilePath, "\" ) - 1 )
sTrgtFileName = Mid( sTrgtFilePath, InStrRev( sTrgtFilePath, "\" ) + 1, Len( sTrgtFilePath ) )

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
            sTrgtFilePath, _
            sTrgtDirPath & "\" & sTrgtFileBaseName & "_" & sNowStr & "." & sTrgtFileExt
    Else
        objFSO.MoveFile _
            sTrgtFilePath, _
            sTrgtDirPath & "\" & sTrgtFileName & "_" & sNowStr
    End If
ElseIf lFileOrFolder = 2 Then
    objFSO.MoveFolder _
        sTrgtFilePath, _
        sTrgtDirPath & "\" & sTrgtFileName & "_" & sNowStr
Else
    MsgBox "�t�@�C��/�t�H���_���s���ł��B", vbOKOnly, PROG_NAME
    WScript.Quit()
End If

Set objFSO = Nothing
