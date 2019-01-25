'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'####################################################################
'### 設定
'####################################################################
Const ADD_DATE_TYPE = 1 '付与する日時の種別（1:現在日時、2:ファイル/フォルダ更新日時）
Const SHORTCUT_FILE_SUFFIX = "#Src#"
Const ORIGINAL_FILE_PREFIX = "#Org#"
Const COPY_FILE_PREFIX     = "#Edt#"

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "ショートカット＆コピーファイル作成"

Dim bIsContinue
bIsContinue = True

'*** 選択ファイル取得 ***
Dim sOrgDirPath
Dim cSelectedPaths
If bIsContinue = True Then
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "デバッグモードです。"
        sOrgDirPath = "X:\100_Documents\200_【学校】共通\大学院\ゼミ出席簿"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "X:\100_Documents\200_【学校】共通\大学院\ゼミ出席簿\H20年度 ゼミ出席簿.xls"
        cSelectedPaths.Add "X:\100_Documents\200_【学校】共通\大学院\ゼミ出席簿\H21年度 ゼミ出席簿.xls"
    Else
        sOrgDirPath = WScript.Env("Current")
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

'*** 上書き確認 ***
If bIsContinue = True Then
    Dim vbAnswer
    vbAnswer = MsgBox( "既にファイルが存在している場合、上書きされます。実行しますか？", vbOkCancel, PROG_NAME )
    If vbAnswer = vbOk Then
        'Do Nothing
    Else
        MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
        bIsContinue = False
    End If
Else
    'Do Nothing
End If

'*** 出力先選択 ***
If bIsContinue = True Then
    Dim sDstParDirPath
    sDstParDirPath = InputBox( "出力先を入力してください。", PROG_NAME, sOrgDirPath )
    If sDstParDirPath = "" Then 'キャンセルの場合
        MsgBox "実行がキャンセルされました。", vbOKOnly, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** ショートカット作成 ***
If bIsContinue = True Then
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    
    Dim sSelectedPath
    For Each sSelectedPath In cSelectedPaths
        'ファイル/フォルダ判定
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( sSelectedPath )
        bFileExists = objFSO.FileExists( sSelectedPath )
        
        Dim sAddDate
        Dim sDstOrgFilePath
        Dim sDstCpyFilePath
        Dim sDstShrtctFilePath
        
        '### ファイル ###
        If bFolderExists = False And bFileExists = True Then
            '追加文字列取得＆整形
            If ADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now() )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFile
                Set objFile = objFSO.GetFile( sSelectedPath )
                sAddDate = ConvDate2String( objFile.DateLastModified )
                Set objFile = Nothing
            Else
                MsgBox "「ADD_DATE_TYPE」の指定が誤っています！", vbOKOnly, PROG_NAME
                MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
                Exit For
            End If
            
            Dim sOrgFileName
            Dim sOrgFileBaseName
            Dim sOrgFileExt
            sOrgFileName = objFSO.GetFileName( sSelectedPath )
            sOrgFileBaseName = objFSO.GetBaseName( sSelectedPath )
            sOrgFileExt = objFSO.GetExtensionName( sSelectedPath )
            sDstOrgFilePath    = sDstParDirPath & "\" & sOrgFileName & "_" & ORIGINAL_FILE_PREFIX & "_" & sAddDate & "." & sOrgFileExt
            sDstCpyFilePath    = sDstParDirPath & "\" & sOrgFileName & "_" & COPY_FILE_PREFIX     & "_" & sAddDate & "." & sOrgFileExt
            sDstShrtctFilePath = sDstParDirPath & "\" & sOrgFileName & "_" & SHORTCUT_FILE_SUFFIX & "_" & sAddDate & ".lnk"
            
            'ファイルコピー
            objFSO.CopyFile sSelectedPath, sDstOrgFilePath, True
            objFSO.CopyFile sSelectedPath, sDstCpyFilePath, True
            
            'ショートカット作成
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### フォルダ ###
        ElseIf bFolderExists = True And bFileExists = False Then
            '追加文字列取得＆整形
            If ADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now() )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFolder
                Set objFolder = objFSO.GetFolder( sSelectedPath )
                sAddDate = ConvDate2String( objFolder.DateLastModified )
                Set objFolder = Nothing
            Else
                MsgBox "「ADD_DATE_TYPE」の指定が誤っています！", vbOKOnly, PROG_NAME
                MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
                Exit For
            End If
            
            Dim sOrgDirName
            sOrgDirName = objFSO.GetFileName( sSelectedPath )
            sDstOrgFilePath    = sDstParDirPath & "\" & sOrgDirName & "_" & ORIGINAL_FILE_PREFIX & "_" & sAddDate
            sDstCpyFilePath    = sDstParDirPath & "\" & sOrgDirName & "_" & COPY_FILE_PREFIX     & "_" & sAddDate
            sDstShrtctFilePath = sDstParDirPath & "\" & sOrgDirName & "_" & SHORTCUT_FILE_SUFFIX & "_" & sAddDate & ".lnk"
            
            'フォルダコピー
            objFSO.CopyFolder sSelectedPath, sDstOrgFilePath, True
            objFSO.CopyFolder sSelectedPath, sDstCpyFilePath, True
            
            'ショートカット作成
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### ファイル/フォルダ以外 ###
        Else
            MsgBox "選択されたオブジェクトが存在しません" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, PROG_NAME
            MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
            bIsContinue = False
        End If
    Next
    
    Set objFSO = Nothing
    Set objWshShell = Nothing
    
    MsgBox "ショートカット＆コピーファイルの作成が完了しました！", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

' ==================================================================
' = 概要    日時文字列をファイル/フォルダ名に適用できる形式に変換する
' = 引数    sDateRaw    String  [in]    日時（例：2017/8/5 12:59:58）
' = 戻値                String          日時（例：20170805_125958）
' = 覚書    なし
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateRaw _
)
    Dim sSearchPattern
    Dim sTargetStr
    sSearchPattern = "(\w{4})/(\w{1,2})/(\w{1,2}) (\w{1,2}):(\w{1,2}):(\w{1,2})"
    sTargetStr = sDateRaw
    
    Dim oRegExp
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = sSearchPattern                '検索パターンを設定
    oRegExp.IgnoreCase = True                       '大文字と小文字を区別しない
    oRegExp.Global = True                           '文字列全体を検索
    Dim oMatchResult
    Set oMatchResult = oRegExp.Execute(sTargetStr)  'パターンマッチ実行
    Dim sDateStr
    With oMatchResult(0)
        sDateStr = String( 4 - Len( .SubMatches(0) ), "0" ) & .SubMatches(0) & _
                   String( 2 - Len( .SubMatches(1) ), "0" ) & .SubMatches(1) & _
                   String( 2 - Len( .SubMatches(2) ), "0" ) & .SubMatches(2) & _
                   "-" & _
                   String( 2 - Len( .SubMatches(3) ), "0" ) & .SubMatches(3) & _
                   String( 2 - Len( .SubMatches(4) ), "0" ) & .SubMatches(4) & _
                   String( 2 - Len( .SubMatches(5) ), "0" ) & .SubMatches(5)
    End With
    Set oMatchResult = Nothing
    Set oRegExp = Nothing
    ConvDate2String = sDateStr
End Function

Public Function SetFileAttributes( _
    ByVal sFilePath, _
    ByVal sDateRaw _
)
	Const SET_ATTR_READONLY	= 1		' 読み取り専用ファイル
	Const SET_ATTR_HIDDEN	= 2		' 隠しファイル
	Const SET_ATTR_SYSTEM	= 4		' システム・ファイル
	Const SET_ATTR_ARCHIVE	= 32	' 前回のバックアップ以降に変更されていれば1
	
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile("test.txt ")
End Function
