'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'####################################################################
'### �ݒ�
'####################################################################


'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "���ݓ����ǉ������l�[��"

Dim bIsContinue
bIsContinue = True

Dim lAnswer
lAnswer = MsgBox ( _
                "�t�@�C��/�t�H���_���̖����Ɍ��ݓ�����t�^���܂��B��낵���ł����H", _
                vbYesNo, _
                PROG_NAME _
            )
If lAnswer = vbYes Then
    'Do Nothing
Else
    MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
    bIsContinue = False
End If

'*******************************************************
'* �t�@�C��/�t�H���_���擾
'*******************************************************
If bIsContinue = True Then
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
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*******************************************************
'* �ǉ�������擾
'*******************************************************
If bIsContinue = True Then
    Dim sDateRaw
    Dim sDateStr
    Dim sAddStr
    sDateRaw = Now()
    sDateStr = ConvDate2String( sDateRaw )
    sAddStr = "_" & sDateStr
    
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
            bIsContinue = False
        End If
        
        If bIsContinue = True Then
            'Do Nothing
        Else
            Exit For
        End If
    Next
    
    Set objFSO = Nothing
Else
    'Do Nothing
End If

' ==================================================================
' = �T�v    ������������t�@�C��/�t�H���_���ɓK�p�ł���`���ɕϊ�����
' = ����    sDateRaw    String  [in]    �����i��F2017/8/5 12:59:58�j
' = �ߒl                String          �����i��F20170805_125958�j
' = �o��    �Ȃ�
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateRaw _
)
    Dim sSearchPattern
    Dim sTargetStr
    sSearchPattern = "(\w{4})/(\w{1,2})/(\w{1,2}) (\w{1,2}):(\w{1,2}):(\w{1,2})"
    sTargetStr = sDateRaw
    
    Dim oRegExp
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = sSearchPattern                '�����p�^�[����ݒ�
    oRegExp.IgnoreCase = True                       '�啶���Ə���������ʂ��Ȃ�
    oRegExp.Global = True                           '������S�̂�����
    Dim oMatchResult
    Set oMatchResult = oRegExp.Execute(sTargetStr)  '�p�^�[���}�b�`���s
    Dim sDateStr
    With oMatchResult(0)
        sDateStr = String( 4 - Len( .SubMatches(0) ), "0" ) & .SubMatches(0) & _
                   String( 2 - Len( .SubMatches(1) ), "0" ) & .SubMatches(1) & _
                   String( 2 - Len( .SubMatches(2) ), "0" ) & .SubMatches(2) & _
                   "-" & _
                   String( 2 - Len( .SubMatches(3) ), "0" ) & .SubMatches(3) & _
                   String( 2 - Len( .SubMatches(4) ), "0" ) & .SubMatches(4) & _
                   String( 2 - Len( .SubMatches(5) ), "0" ) & .SubMatches(5)
    End With
    Set oMatchResult = Nothing
    Set oRegExp = Nothing
    ConvDate2String = sDateStr
End Function

