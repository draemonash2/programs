'Option Explicit

'�t�@�C���p�X�R�s�[�� Clippath:18�A�t�@�C�����R�s�[�� Clippath:12 �Ŏ��s�ł���B
'�������擪�ɉ��s���܂܂�Ă��܂����߁A�g��Ȃ��B

Const PROG_NAME = "�t�@�C���p�X���R�s�["
Const INCLUDE_DOUBLE_QUOTATION = False

Dim sTrgtPathsRaw
If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "�f�o�b�O���[�h�ł��B"
    sTrgtPathsRaw = "C:\Users\draem_000\Desktop\aa ""C:\Users\draem_000\Desktop\b b"""
'   sTrgtPathsRaw = """C:\Users\draem_000\D esktop\aa"" ""C:\Users\draem_000\Desktop\b b"""
'   sTrgtPathsRaw = """C:\Users\draem_000\Desktop\b b"""
'   sTrgtPathsRaw = """C:\Users\draem_000\Desktop\b b.txt"""
'   sTrgtPathsRaw = "C:\Users\draem_000\Desktop\aa"
'   sTrgtPathsRaw = ""
Else
    sTrgtPathsRaw = WScript.Env("Selected")
End If

'*** �p�X�ϊ��i������^�˔z�񕶎���^�j ***
Dim asTrgtPaths()
Dim bRet
bRet = ConvPathsStringToArray( sTrgtPathsRaw, INCLUDE_DOUBLE_QUOTATION, asTrgtPaths )
If bRet = True Then
    'Do Nothing
Else
    MsgBox "�t�@�C���p�X���w�肳��Ă��܂���I"
    MsgBox "�����𒆒f���܂�"
    WScript.Quit
End If

'*** �p�X�ϊ��i�z�񕶎���^�˕�����^�j ***
Dim sTrgtPaths
ReDim Preserve asTrgtNames( UBound( asTrgtPaths ) )
Dim lIdx
For lIdx = LBound( asTrgtPaths ) To UBound( asTrgtPaths )
    If lIdx = LBound( asTrgtPaths ) Then
        sTrgtPaths = asTrgtPaths( lIdx )
    Else
        sTrgtPaths = sTrgtPaths & vbNewLine & asTrgtPaths( lIdx )
    End If
Next

'*** �N���b�v�{�[�h�փR�s�[ ***
CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sTrgtPaths )

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
