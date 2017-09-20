'Option Explicit

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "ファイル/フォルダ名末尾文字列付与"

Dim lAnswer
lAnswer = MsgBox ( _
                "ファイル/フォルダ名の末尾に文字列を付与します。よろしいですか？", _
                vbYesNo, _
                PROG_NAME _
            )
If lAnswer = vbYes Then
    'Do Nothing
Else
    MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
    WScript.Quit()
End If

'*******************************************************
'* ファイル/フォルダ名取得
'*******************************************************
Dim cFilePaths
If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "デバッグモードです。"
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

'*** ファイルパスチェック ***
If cFilePaths.Count = 0 Then
    MsgBox "ファイルが選択されていません", vbYes, PROG_NAME
    MsgBox "処理を中断します", vbYes, PROG_NAME
    WScript.Quit
Else
    'Do Nothing
End If

'*******************************************************
'* 追加文字列取得
'*******************************************************
Dim sSearchPattern
Dim sTargetStr
sSearchPattern = "(\w{4})/(\w{1,2})/(\w{1,2}) (\w{1,2}):(\w{1,2}):(\w{1,2})"
sTargetStr = Now()

Dim oRegExp
Set oRegExp = CreateObject("VBScript.RegExp")
oRegExp.Pattern = sSearchPattern                '検索パターンを設定
oRegExp.IgnoreCase = True                       '大文字と小文字を区別しない
oRegExp.Global = True                           '文字列全体を検索
Dim oMatchResult
Set oMatchResult = oRegExp.Execute(sTargetStr)  'パターンマッチ実行
Dim sNowStr
With oMatchResult(0)
    sNowStr = String( 4 - Len( .SubMatches(0) ), "0" ) & .SubMatches(0) & _
              String( 2 - Len( .SubMatches(1) ), "0" ) & .SubMatches(1) & _
              String( 2 - Len( .SubMatches(2) ), "0" ) & .SubMatches(2) & _
              "_" & _
              String( 2 - Len( .SubMatches(3) ), "0" ) & .SubMatches(3) & _
              String( 2 - Len( .SubMatches(4) ), "0" ) & .SubMatches(4) & _
              String( 2 - Len( .SubMatches(5) ), "0" ) & .SubMatches(5)
End With
Set oMatchResult = Nothing
Set oRegExp = Nothing

Dim sAddStr
sAddStr = InputBox( "末尾に付与する文字列を入力してください", PROG_NAME, "_" & sNowStr )

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim oFilePath
For Each oFilePath In cFilePaths
    '*******************************************************
    '* ファイル/フォルダ名判別
    '*******************************************************
    
    Dim lFileOrFolder '1:ファイル、2:フォルダ、0:エラー（存在しないパス）
    Dim bFolderExists
    Dim bFileExists
    bFolderExists = objFSO.FolderExists( oFilePath )
    bFileExists = objFSO.FileExists( oFilePath )
    If bFolderExists = False And bFileExists = True Then
        lFileOrFolder = 1 'ファイル
    ElseIf bFolderExists = True And bFileExists = False Then
        lFileOrFolder = 2 'フォルダー
    Else
        lFileOrFolder = 0 'エラー（存在しないパス）
    End If
    
    '*******************************************************
    '* ファイル/フォルダ名変更
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
        MsgBox "ファイル/フォルダが不正です。", vbOKOnly, PROG_NAME
        WScript.Quit()
    End If
Next

Set objFSO = Nothing
