'Option Explicit

'ファイルパスコピーは Clippath:18、ファイル名コピーは Clippath:12 で実行できる。
'しかし先頭に改行が含まれてしまうため、使わない。

Const PROG_NAME = "ファイルパスをコピー"
Const INCLUDE_DOUBLE_QUOTATION = False

Dim sTrgtPathsRaw
If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "デバッグモードです。"
    sTrgtPathsRaw = "C:\Users\draem_000\Desktop\aa ""C:\Users\draem_000\Desktop\b b"""
'   sTrgtPathsRaw = """C:\Users\draem_000\D esktop\aa"" ""C:\Users\draem_000\Desktop\b b"""
'   sTrgtPathsRaw = """C:\Users\draem_000\Desktop\b b"""
'   sTrgtPathsRaw = """C:\Users\draem_000\Desktop\b b.txt"""
'   sTrgtPathsRaw = "C:\Users\draem_000\Desktop\aa"
'   sTrgtPathsRaw = ""
Else
    sTrgtPathsRaw = WScript.Env("Selected")
End If

'*** パス変換（文字列型⇒配列文字列型） ***
Dim asTrgtPaths()
Dim bRet
bRet = ConvPathsStringToArray( sTrgtPathsRaw, INCLUDE_DOUBLE_QUOTATION, asTrgtPaths )
If bRet = True Then
    'Do Nothing
Else
    MsgBox "ファイルパスが指定されていません！"
    MsgBox "処理を中断します"
    WScript.Quit
End If

'*** パス変換（配列文字列型⇒文字列型） ***
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

'*** クリップボードへコピー ***
CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sTrgtPaths )

'空文字列を指定した場合、戻り値 False を返却する
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
