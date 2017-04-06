'Option Explicit

Const PROG_NAME = "リネーム用バッチファイル出力"
Const OUTPUT_BAT_FILE_BASE_NAME = "rename"

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

'*** ファイルパスチェック ***
If cFilePaths.Count = 0 Then
    MsgBox "ファイルが選択されていません", vbYes, PROG_NAME
    MsgBox "処理を中断します", vbYes, PROG_NAME
    WScript.Quit
Else
    'Do Nothing
End If

'*** ファイルパスからファイル名取り出し ***
Dim cFileNames
Set cFileNames = CreateObject("System.Collections.ArrayList")
Dim oFilePath
For Each oFilePath In cFilePaths
    cFileNames.Add Mid( oFilePath, InStrRev( oFilePath, "\" ) + 1, Len( oFilePath ) )
Next

'*** 最大の文字列長を取得 ***
Dim lFileNameLenMax
lFileNameLenMax = 0
Dim oFileName
For Each oFileName In cFileNames
    Dim lTrgtFileNameLen
    lTrgtFileNameLen = Len( oFileName )
    If lTrgtFileNameLen > lFileNameLenMax Then
        lFileNameLenMax = lTrgtFileNameLen
    Else
        'Do Nothing
    End If
Next

'*** リネーム前ファイルリスト出力 ***
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

'*** リネーム用バッチファイル出力 ***
Dim sBatFilePath
sBatFilePath = sOutputBatDirPath & "\" & OUTPUT_BAT_FILE_BASE_NAME & ".bat"
Set objTxtFile = objFSO.OpenTextFile( sBatFilePath, 2, True)
For Each oFileName In cFileNames
    objTxtFile.WriteLine _
        "rename " & _
        """" & oFileName & """" & _
        String(lFileNameLenMax - Len( oFileName ) + 1, " ") & _
        """" & oFileName & """"
Next
objTxtFile.WriteLine "pause"
objTxtFile.Close
Set objTxtFile = Nothing

MsgBox OUTPUT_BAT_FILE_BASE_NAME & ".bat を出力しました。"

