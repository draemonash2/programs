'Option Explicit

Const PROG_NAME = "���l�[���p�o�b�`�t�@�C���o��"
Const OUTPUT_BAT_FILE_BASE_NAME = "rename"

Dim sTrgtPathsRaw
Dim sOutputBatDirPath
If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "�f�o�b�O���[�h�ł��B"
    sOutputBatDirPath = "C:\Users\draem_000\Desktop\test"
'   sTrgtPathsRaw = "C:\Users\draem_000\Desktop\aa ""C:\Users\draem_000\Desktop\b b"""
'   sTrgtPathsRaw = """C:\Users\draem_000\D esktop\aa"" ""C:\Users\draem_000\Desktop\b b"""
    sTrgtPathsRaw = """C:\Users\draem_000\Desktop\test\aabbbbb.txt"" ""C:\Users\draem_000\Desktop\test\b b"""
'   sTrgtPathsRaw = """C:\Users\draem_000\Desktop\b b"""
'   sTrgtPathsRaw = """C:\Users\draem_000\Desktop\b b.txt"""
'   sTrgtPathsRaw = "C:\Users\draem_000\Desktop\aa"
'   sTrgtPathsRaw = ""
Else
    sOutputBatDirPath = "C:\Users\draem_000\Desktop"
    sTrgtPathsRaw = WScript.Env("Selected")
    Dim sExePath
    sExePath = WScript.Env("X-Finder") & "..\" & WScript.Env("Vim")
End If

'*** �p�X�ϊ��i������^�˔z�񕶎���^�j ***
Dim asTrgtPaths()
Dim bRet
bRet = ConvPathsStringToArray( sTrgtPathsRaw, False, asTrgtPaths )
If bRet = True Then
    'Do Nothing
Else
    MsgBox "�t�@�C���p�X���w�肳��Ă��܂���I"
    MsgBox "�����𒆒f���܂�"
    WScript.Quit
End If

'*** �t�@�C���p�X����t�@�C�������o�� ***
Dim asTrgtNames()
ReDim Preserve asTrgtNames( UBound( asTrgtPaths ) )
Dim lIdx
For lIdx = LBound( asTrgtPaths ) To UBound( asTrgtPaths )
    Dim sTrgtPath
    sTrgtPath = asTrgtPaths( lIdx )
    asTrgtNames( lIdx ) = Mid( sTrgtPath, InStrRev( sTrgtPath, "\" ) + 1, Len( sTrgtPath ) )
Next

'*** �ő�̕����񒷂��擾 ***
Dim lFileNameLenMax
lFileNameLenMax = 0
For lIdx = LBound( asTrgtNames ) To UBound( asTrgtNames )
    Dim lTrgtFileNameLen
    lTrgtFileNameLen = Len( asTrgtNames( lIdx ) )
    If lTrgtFileNameLen > lFileNameLenMax Then
        lFileNameLenMax = lTrgtFileNameLen
    Else
        'Do Nothing
    End If
Next

'*** ���l�[���O�t�@�C�����X�g�o�� ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objTxtFile
Dim sBakFilePath
sBakFilePath = sOutputBatDirPath & "\" & OUTPUT_BAT_FILE_BASE_NAME & "_bak.txt"
Set objTxtFile = objFSO.OpenTextFile( sBakFilePath, 2, True)
For lIdx = LBound( asTrgtNames ) To UBound( asTrgtNames )
    objTxtFile.WriteLine asTrgtNames( lIdx )
Next
objTxtFile.Close
Set objTxtFile = Nothing

'*** ���l�[���p�o�b�`�t�@�C���o�� ***
Dim sBatFilePath
sBatFilePath = sOutputBatDirPath & "\" & OUTPUT_BAT_FILE_BASE_NAME & ".bat"
Set objTxtFile = objFSO.OpenTextFile( sBatFilePath, 2, True)
For lIdx = LBound( asTrgtNames ) To UBound( asTrgtNames )
    objTxtFile.WriteLine _
        "rename " & _
        """" & asTrgtNames( lIdx ) & """" & _
        String(lFileNameLenMax - Len( asTrgtNames( lIdx ) ) + 1, " ") & _
        """" & asTrgtNames( lIdx ) & """"
Next
objTxtFile.WriteLine "pause"
objTxtFile.Close
Set objTxtFile = Nothing

MsgBox OUTPUT_BAT_FILE_BASE_NAME & ".bat ���o�͂��܂����B"
'WScript.CreateObject("WScript.Shell").Run "%comspec% /c """ & sExePath & """ """ & sBatFilePath & """", 0, True

'�󕶎�����w�肵���ꍇ�A�߂�l False ��ԋp����
Public Function ConvPathsStringToArray( _
    ByVal sTrgtPaths, _
    ByVal bIncludeDblQuote, _
    ByRef asTrgtPaths() _
)
    
    If sTrgtPaths = "" Then
        ReDim Preserve asTrgtPaths(-1)
        ConvPathsStringToArray = False
    Else
        ReDim Preserve asTrgtPaths(0)
        Dim sCurPathStr
        sCurPathStr = ""
        Dim bIsPathContinue
        bIsPathContinue = False
        Dim lTrgtStrIdx
        For lTrgtStrIdx = 1 To Len( sTrgtPaths )
            Dim sChar
            sChar = Mid( sTrgtPaths, lTrgtStrIdx, 1 )
            If sChar = """" Then
                If bIsPathContinue = True Then
                    bIsPathContinue = False
                Else
                    bIsPathContinue = True
                End If
                If bIncludeDblQuote = True Then
                    sCurPathStr = sCurPathStr & sChar
                Else
                    'Do Nothing
                End If
            ElseIf sChar = " " Then
                If bIsPathContinue = True Then
                    sCurPathStr = sCurPathStr & sChar
                Else
                    asTrgtPaths( UBound( asTrgtPaths ) ) = sCurPathStr
                    ReDim Preserve asTrgtPaths( UBound( asTrgtPaths ) + 1 )
                    sCurPathStr = ""
                End If
            Else
                sCurPathStr = sCurPathStr & sChar
            End If
        Next
        asTrgtPaths( UBound( asTrgtPaths ) ) = sCurPathStr
        ConvPathsStringToArray = True
    End If
End Function
'   Call Test_ConvPathsStringToArray()
    Private Function Test_ConvPathsStringToArray()
        Dim sTrgtPaths
    '   sTrgtPaths = "C:\Users\draem_000\Desktop\mp4 C:\Users\draem_000\Desktop\temp.txt C:\Users\draem_000\Desktop\test.vbs"
    '   sTrgtPaths = "C:\Users\draem_000\Desktop\mp4"
    '   sTrgtPaths = """C:\Users\draem_000\Des ktop\mp4"" C:\Users\draem_000\Desktop\temp.txt ""C:\Use rs\draem_000\Desktop\test.vbs"""
        sTrgtPaths = ""
        Dim asTrgtPaths()
        Dim bRet
        bRet = ConvPathsStringToArray( sTrgtPaths, asTrgtPaths )
        Dim sBuf
        sBuf = bRet
        sBuf = sBuf & vbNewLine & UBound( asTrgtPaths ) + 1
        Dim i
        For i = LBound( asTrgtPaths ) to UBound( asTrgtPaths )
            sBuf = sBuf & vbNewLine & asTrgtPaths( i )
        '   MsgBox asTrgtPaths( i )
        Next
        MsgBox sBuf
    End Function
