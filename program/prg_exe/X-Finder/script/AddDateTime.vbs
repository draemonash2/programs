'Option Explicit

Const PROG_NAME = "現在時刻付与"

Dim lAnswer
lAnswer = MsgBox ( _
                "ファイル/フォルダ名に現在時刻を付与します。よろしいですか？", _
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
Dim sTrgtFilePath

If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "デバッグモードです。"
    sTrgtFilePath = WScript.Arguments(0)
Else
    sTrgtFilePath = WScript.Env("Focused")
End If

'*******************************************************
'* 現在日時取得
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

'*******************************************************
'* ファイル/フォルダ名判別
'*******************************************************
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim lFileOrFolder '1:ファイル、2:フォルダ、0:エラー（存在しないパス）
Dim bFolderExists
Dim bFileExists
bFolderExists = objFSO.FolderExists( sTrgtFilePath )
bFileExists = objFSO.FileExists( sTrgtFilePath )
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
sTrgtDirPath = Mid( sTrgtFilePath, 1, InStrRev( sTrgtFilePath, "\" ) - 1 )
sTrgtFileName = Mid( sTrgtFilePath, InStrRev( sTrgtFilePath, "\" ) + 1, Len( sTrgtFilePath ) )

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
            sTrgtFilePath, _
            sTrgtDirPath & "\" & sTrgtFileBaseName & "_" & sNowStr & "." & sTrgtFileExt
    Else
        objFSO.MoveFile _
            sTrgtFilePath, _
            sTrgtDirPath & "\" & sTrgtFileName & "_" & sNowStr
    End If
ElseIf lFileOrFolder = 2 Then
    objFSO.MoveFolder _
        sTrgtFilePath, _
        sTrgtDirPath & "\" & sTrgtFileName & "_" & sNowStr
Else
    MsgBox "ファイル/フォルダが不正です。", vbOKOnly, PROG_NAME
    WScript.Quit()
End If

Set objFSO = Nothing
