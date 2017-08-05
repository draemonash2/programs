'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

Const PROG_NAME = "リネーム用バッチファイル出力"
Const OUTPUT_BAT_FILE_BASE_NAME = "rename"

Dim bIsContinue
bIsContinue = True

If bIsContinue = True Then
    Dim sOutputBatDirPath
    Dim sExePath
    Dim cFilePaths
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "デバッグモードです。"
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

'*** ファイルパスチェック ***
If bIsContinue = True Then
    If cFilePaths.Count = 0 Then
        MsgBox "ファイルが選択されていません", vbYes, PROG_NAME
        MsgBox "処理を中断します", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** ファイルパスからファイル名取り出し ***
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

'*** 最大の文字列長を取得 ***
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

'*** リネーム前ファイルリスト出力 ***
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

'*** リネーム用バッチファイル出力 ***
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
    
    MsgBox OUTPUT_BAT_FILE_BASE_NAME & ".bat を出力しました。"
Else
    'Do Nothing
End If

' ==================================================================
' = 概要    指定された文字列の文字列長（バイト数）を返却する
' = 引数    sInStr      String  [in]  文字列
' = 戻値                Long          文字列長（バイト数）
' = 覚書    標準で用意されている LenB() 関数は、Unicode における
' =         バイト数を返却するため半角文字も２文字としてカウントする。
' =           （例：LenB("ファイルサイズ ") ⇒ 16）
' =         そのため、半角文字を１文字としてカウントする本関数を用意。
' ==================================================================
Public Function LenByte( _
    ByVal sInStr _
)
    Dim lIdx, sChar
    LenByte = 0
    If Trim(sInStr) <> "" Then
        For lIdx = 1 To Len(sInStr)
            sChar = Mid(sInStr, lIdx, 1)
            '２バイト文字は＋２
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
        Result = Result & vbNewLine & LenByte( "あああ" )   ' 6
        Result = Result & vbNewLine & LenByte( "あああ " )  ' 7
        Result = Result & vbNewLine & LenByte( "ああ あ" )  ' 7
        Result = Result & vbNewLine & LenByte( Chr(9) )     ' 1
        Result = Result & vbNewLine & LenByte( Chr(10) )    ' 1
        MsgBox Result
    End Sub
