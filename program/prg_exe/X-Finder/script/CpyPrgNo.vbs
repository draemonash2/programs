'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "プログラムNo.をコピー"

Dim bIsContinue
bIsContinue = True

Dim cFilePaths

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        Dim sArg
        For Each sArg In WScript.Arguments
            cFilePaths.add sArg
        Next
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        Set cFilePaths = WScript.Col( WScript.Env("Selected") )
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\aabbbbb.txt"
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\b b"
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

'*** クリップボードへコピー ***
If bIsContinue = True Then
    Dim sOutString
    Dim bFirstStore
    bFirstStore = True
    Dim objTxtFile
    Dim sPrgNo
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sOutString = ""
    For Each oFilePath In cFilePaths
        Set objTxtFile = objFSO.OpenTextFile( oFilePath, 1, False)
        sPrgNo = objTxtFile.ReadLine
        sPrgNo = Replace( sPrgNo, "/* ", "" )
        sPrgNo = Replace( sPrgNo, " */", "" )
        If bFirstStore = True Then
            sOutString = sPrgNo
            bFirstStore = False
        Else
            sOutString = sOutString & vbNewLine & sPrgNo
        End If
        objTxtFile.Close
    Next
    Set objTxtFile = Nothing
    Set objFSO = Nothing
    CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
    'Do Nothing
End If
