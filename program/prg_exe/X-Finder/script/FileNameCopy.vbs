'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'ファイル名コピーは Clippath:12 で実行できるが、
'先頭に改行が含まれてしまうため使わない。

'####################################################################
'### 設定
'####################################################################
Const INCLUDE_DOUBLE_QUOTATION = False

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "ファイル名をコピー"

Dim bIsContinue
bIsContinue = True

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    Dim cFilePaths
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "デバッグモードです。"
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\aabbbbb.txt"
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\b b"
    Else
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

'*** クリップボードへコピー ***
If bIsContinue = True Then
    Dim sOutString
    Dim oFileName
    Dim bFirstStore
    bFirstStore = True
    For Each oFileName In cFileNames
        If bFirstStore = True Then
            sOutString = oFileName
            bFirstStore = False
        Else
            sOutString = sOutString & vbNewLine & oFileName
        End If
    Next
    CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
    'Do Nothing
End If
