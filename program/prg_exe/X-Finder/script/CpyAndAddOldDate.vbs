'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'####################################################################
'### �ݒ�
'####################################################################


'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�O���������ǉ����R�s�["

Dim bIsContinue
bIsContinue = True

Dim lAnswer
lAnswer = MsgBox ( _
                "�t�@�C��/�t�H���_���̖����ɑO����������t�^���܂��B��낵���ł����H", _
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
    sDateRaw = DateAdd("d", -1, sDateRaw)
    sDateRaw = Year( sDateRaw ) & "/" & _
               Month( sDateRaw ) & "/" & _
               Day( sDateRaw ) & " " & _
               "17:45:00"
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
                objFSO.CopyFile _
                    oFilePath, _
                    sTrgtDirPath & "\" & sTrgtFileBaseName & sAddStr & "." & sTrgtFileExt
            Else
                objFSO.CopyFile _
                    oFilePath, _
                    sTrgtDirPath & "\" & sTrgtFileName & sAddStr
            End If
        ElseIf lFileOrFolder = 2 Then
            objFSO.CopyFolder _
                oFilePath, _
                sTrgtDirPath & "\" & sTrgtFileName & sAddStr, _
                True
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
' = �T�v    �����`����ϊ�����B�i��F2017/03/22 18:20:14 �� 20170322-182014�j
' = ����    sDateTime   String  [in]  �����iYYYY/MM/DD HH:MM:SS�j
' = �ߒl                String        �����iYYYYMMDD-HHMMSS�j
' = �o��    ��ɓ������t�@�C������t�H���_���Ɏg�p����ۂɎg�p����B
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateTime _
)
    On Error Resume Next
    Dim sDateStr
    sDateStr = _
        String(4 - Len(Year(sDateTime)),   "0") & Year(sDateTime)   & _
        String(2 - Len(Month(sDateTime)),  "0") & Month(sDateTime)  & _
        String(2 - Len(Day(sDateTime)),    "0") & Day(sDateTime)    & _
        "-" & _
        String(2 - Len(Hour(sDateTime)),   "0") & Hour(sDateTime)   & _
        String(2 - Len(Minute(sDateTime)), "0") & Minute(sDateTime) & _
        String(2 - Len(Second(sDateTime)), "0") & Second(sDateTime)
    If Err.Number = 0 Then
        ConvDate2String = sDateStr
    Else
        ConvDate2String = ""
    End If
    On Error Goto 0
End Function

