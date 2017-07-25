'Option Explicit

'####################################################################
'### 設定
'####################################################################

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "ショートカット＆コピーファイル作成"

Const SHORTCUT_FILE_PREFIX = "CopySource"

Dim bIsContinue
bIsContinue = True

'*** 選択ファイル取得 ***
Dim cSelectedPaths
If bIsContinue = True Then
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "デバッグモードです。"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "X:\100_Documents\200_【学校】共通\大学院\ゼミ出席簿\H20年度 ゼミ出席簿.xls"
        cSelectedPaths.Add "X:\100_Documents\200_【学校】共通\大学院\ゼミ出席簿\H21年度 ゼミ出席簿.xls"
    Else
        Set cSelectedPaths = WScript.Col( WScript.Env("Selected") )
    End If
Else
    'Do Nothing
End If

'*** ファイルパスチェック ***
If bIsContinue = True Then
    If cSelectedPaths.Count = 0 Then
        MsgBox "ファイル/フォルダが選択されていません。", vbOKOnly, PROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** ショートカット作成 ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim oRegExp
Set oRegExp = CreateObject("VBScript.RegExp")
If bIsContinue = True Then
    Dim sSelectedPath
    For Each sSelectedPath In cSelectedPaths
        'ファイル/フォルダ判定
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( sSelectedPath )
        bFileExists = objFSO.FileExists( sSelectedPath )
        If bFolderExists = False And bFileExists = True Then
            '追加文字列取得＆整形
            Dim objFile
            Set objFile = objFSO.GetFile( sSelectedPath )
            Dim sDateLastModifiedRaw
            sDateLastModifiedRaw = objFile.DateLastModified
            
            Dim sSearchPattern
            Dim sTargetStr
            sSearchPattern = "(\w{4})/(\w{1,2})/(\w{1,2}) (\w{1,2}):(\w{1,2}):(\w{1,2})"
            sTargetStr = sDateLastModifiedRaw
            oRegExp.Pattern = sSearchPattern                '検索パターンを設定
            oRegExp.IgnoreCase = True                       '大文字と小文字を区別しない
            oRegExp.Global = True                           '文字列全体を検索
            Dim oMatchResult
            Set oMatchResult = oRegExp.Execute(sTargetStr)  'パターンマッチ実行
            Dim sDateLastModifiedFiltered
            With oMatchResult(0)
                sDateLastModifiedFiltered = String( 4 - Len( .SubMatches(0) ), "0" ) & .SubMatches(0) & _
                                            String( 2 - Len( .SubMatches(1) ), "0" ) & .SubMatches(1) & _
                                            String( 2 - Len( .SubMatches(2) ), "0" ) & .SubMatches(2) & _
                                            "-" & _
                                            String( 2 - Len( .SubMatches(3) ), "0" ) & .SubMatches(3) & _
                                            String( 2 - Len( .SubMatches(4) ), "0" ) & .SubMatches(4) & _
                                            String( 2 - Len( .SubMatches(5) ), "0" ) & .SubMatches(5)
            End With
            Set oMatchResult = Nothing
            
            Dim sFileExt
            Dim sFileBaseName
            Dim sParentDirPath
            Dim sCopyDstFilePath
            Dim sShortcutFilePath
            sFileBaseName = objFSO.GetBaseName( sSelectedPath )
            sFileExt = objFSO.GetExtensionName( sSelectedPath )
            sParentDirPath = objFSO.GetParentFolderName( sSelectedPath )
            'sCopyDstFilePath = sParentDirPath & "\" & sFileBaseName & " [" & sDateLastModifiedFiltered & "]." & sFileExt
            sCopyDstFilePath = sSelectedPath & "_" & sDateLastModifiedFiltered & "." & sFileExt
            sShortcutFilePath = sSelectedPath & "_" & SHORTCUT_FILE_PREFIX & ".lnk"
            
            'ファイルコピー
            objFSO.CopyFile sSelectedPath, sCopyDstFilePath, True
            
            'ショートカット作成
            With objWshShell.CreateShortcut( sShortcutFilePath )
                .TargetPath = sParentDirPath
                .Save
            End With
            
            '選択
            '★
        ElseIf bFolderExists = True And bFileExists = False Then
            'ディレクトリは対象外とする
        Else
            MsgBox "選択されたオブジェクトが存在しません" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, PROG_NAME
            MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
            bIsContinue = False
        End If
    Next
Else
    'Do Nothing
End If

MsgBox "ショートカット＆コピーファイルの作成が完了しました！", vbOKOnly, PROG_NAME

Set oRegExp = Nothing
Set objFSO = Nothing
Set objWshShell = Nothing
