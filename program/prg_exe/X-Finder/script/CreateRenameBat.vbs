'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

Const PROG_NAME = "���l�[���p�o�b�`�t�@�C���o��"
Const OUTPUT_BAT_FILE_BASE_NAME = "rename"

Dim bIsContinue
bIsContinue = True

If bIsContinue = True Then
    Dim sOutputBatDirPath
    Dim sExePath
    Dim cFilePaths
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "�f�o�b�O���[�h�ł��B"
        sOutputBatDirPath = "C:\Users\draem_000\Desktop\test"
        sExePath = "C:\prg_exe\Vim\gvim.exe"
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\aabbbbb.txt"
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\b b"
    Else
        sOutputBatDirPath = WScript.Env("Current")
        sExePath = WScript.Env("X-Finder") & "..\" & WScript.Env("Vim")
        Set cFilePaths = WScript.Col( WScript.Env("Selected") )
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

'*** �t�@�C���p�X����t�@�C�������o�� ***
If bIsContinue = True Then
    Dim cFileNames
    Set cFileNames = CreateObject("System.Collections.ArrayList")
    Dim oFilePath
    For Each oFilePath In cFilePaths
        cFileNames.Add Mid( oFilePath, InStrRev( oFilePath, "\" ) + 1, Len( oFilePath ) )
    Next
Else
    'Do Nothing
End If

'*** �ő�̕����񒷂��擾 ***
If bIsContinue = True Then
    Dim lFileNameLenMax
    lFileNameLenMax = 0
    Dim oFileName
    For Each oFileName In cFileNames
        Dim lTrgtFileNameLen
        lTrgtFileNameLen = LenByte( oFileName )
        If lTrgtFileNameLen > lFileNameLenMax Then
            lFileNameLenMax = lTrgtFileNameLen
        Else
            'Do Nothing
        End If
    Next
Else
    'Do Nothing
End If

'*** ���l�[���O�t�@�C�����X�g�o�� ***
If bIsContinue = True Then
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objTxtFile
    Dim sBakFilePath
    sBakFilePath = sOutputBatDirPath & "\" & OUTPUT_BAT_FILE_BASE_NAME & "_bak.txt"
    Set objTxtFile = objFSO.OpenTextFile( sBakFilePath, 2, True)
    For Each oFileName In cFileNames
        objTxtFile.WriteLine oFileName
    Next
    objTxtFile.Close
    Set objTxtFile = Nothing
Else
    'Do Nothing
End If

'*** ���l�[���p�o�b�`�t�@�C���o�� ***
If bIsContinue = True Then
    Dim sBatFilePath
    sBatFilePath = sOutputBatDirPath & "\" & OUTPUT_BAT_FILE_BASE_NAME & ".bat"
    Set objTxtFile = objFSO.OpenTextFile( sBatFilePath, 2, True)
    For Each oFileName In cFileNames
        objTxtFile.WriteLine _
            "rename " & _
            """" & oFileName & """" & _
            String(lFileNameLenMax - LenByte( oFileName ) + 1, " ") & _
            """" & oFileName & """"
    Next
    objTxtFile.WriteLine "pause"
    objTxtFile.Close
    Set objTxtFile = Nothing
    
    MsgBox OUTPUT_BAT_FILE_BASE_NAME & ".bat ���o�͂��܂����B"
Else
    'Do Nothing
End If

' ==================================================================
' = �T�v    �w�肳�ꂽ������̕����񒷁i�o�C�g���j��ԋp����
' = ����    sInStr      String  [in]  ������
' = �ߒl                Long          �����񒷁i�o�C�g���j
' = �o��    �W���ŗp�ӂ���Ă��� LenB() �֐��́AUnicode �ɂ�����
' =         �o�C�g����ԋp���邽�ߔ��p�������Q�����Ƃ��ăJ�E���g����B
' =           �i��FLenB("�t�@�C���T�C�Y ") �� 16�j
' =         ���̂��߁A���p�������P�����Ƃ��ăJ�E���g����{�֐���p�ӁB
' ==================================================================
Public Function LenByte( _
    ByVal sInStr _
)
    Dim lIdx, sChar
    LenByte = 0
    If Trim(sInStr) <> "" Then
        For lIdx = 1 To Len(sInStr)
            sChar = Mid(sInStr, lIdx, 1)
            '�Q�o�C�g�����́{�Q
            If (Asc(sChar) And &HFF00) <> 0 Then
                LenByte = LenByte + 2
            Else
                LenByte = LenByte + 1
            End If
        Next
    End If
End Function
'   Call Test_LenByte()
    Private Sub Test_LenByte()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & LenByte( "aaa" )      ' 3
        Result = Result & vbNewLine & LenByte( "aaa " )     ' 4
        Result = Result & vbNewLine & LenByte( "" )         ' 0
        Result = Result & vbNewLine & LenByte( "������" )   ' 6
        Result = Result & vbNewLine & LenByte( "������ " )  ' 7
        Result = Result & vbNewLine & LenByte( "���� ��" )  ' 7
        Result = Result & vbNewLine & LenByte( Chr(9) )     ' 1
        Result = Result & vbNewLine & LenByte( Chr(10) )    ' 1
        MsgBox Result
    End Sub
