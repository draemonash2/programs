'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'<<7-Zip usage>>
'  7z a <zip_file_path> <target_dir_path> -p<password>

'★TODO★：ZIP ファイル以外の圧縮動作確認

'####################################################################
'### 事前処理
'####################################################################
Dim cAcceptFileFormats
Set cAcceptFileFormats = CreateObject("System.Collections.ArrayList")

'####################################################################
'### 設定
'####################################################################
'7-Zip 16.04 圧縮可能形式（/7-ZipPortable/App/7-Zip/7-zip.chm より引用）
'                      [FileExt]    [Format]
cAcceptFileFormats.Add "7z"       ' 7z
cAcceptFileFormats.Add "bz2"      ' BZIP2
cAcceptFileFormats.Add "bzip2"    ' BZIP2
cAcceptFileFormats.Add "tbz2"     ' BZIP2
cAcceptFileFormats.Add "tbz"      ' BZIP2
cAcceptFileFormats.Add "gz"       ' GZIP
cAcceptFileFormats.Add "gzip"     ' GZIP
cAcceptFileFormats.Add "tgz"      ' GZIP
cAcceptFileFormats.Add "tar"      ' TAR
cAcceptFileFormats.Add "wim"      ' WIM
cAcceptFileFormats.Add "swm"      ' WIM
cAcceptFileFormats.Add "xz"       ' XZ
cAcceptFileFormats.Add "txz"      ' XZ
cAcceptFileFormats.Add "zip"      ' ZIP
cAcceptFileFormats.Add "zipx"     ' ZIP
cAcceptFileFormats.Add "jar"      ' ZIP
cAcceptFileFormats.Add "xpi"      ' ZIP
cAcceptFileFormats.Add "odt"      ' ZIP
cAcceptFileFormats.Add "ods"      ' ZIP
cAcceptFileFormats.Add "docx"     ' ZIP
cAcceptFileFormats.Add "xlsx"     ' ZIP
cAcceptFileFormats.Add "epub"     ' ZIP
Const INITIAL_FILE_EXT = "zip"

Const HIDE_PASSWORD = True

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "7-Zip でパスワード圧縮"

Dim sExePath
Dim cSelectedPaths

Dim bIsContinue
bIsContinue = True

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "デバッグモードです。"
        sExePath = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\aa"
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\b b"
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\d.txt"
    Else
        sExePath = WScript.Env("7-Zip")
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

'********************
'*** 圧縮形式選択 ***
'********************
If bIsContinue = True Then
    Dim bIsReEnter
    bIsReEnter = False
    Dim sAcceptFileFormatsStr
    Dim sAcceptFileFormat
    sAcceptFileFormatsStr = ""
    For Each sAcceptFileFormat In cAcceptFileFormats
        sAcceptFileFormatsStr = sAcceptFileFormatsStr & vbNewLine & sAcceptFileFormat
    Next
    Do
        Dim sArchiveFileExt
        sArchiveFileExt = InputBox( _
                            "以下の中から圧縮形式を選択して入力してください。" & vbNewLine & _
                            sAcceptFileFormatsStr & vbNewLine, _
                            PROG_NAME, _
                            INITIAL_FILE_EXT _
                        )
        If sArchiveFileExt = "" Then
            MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
            MsgBox "処理を中断します。", vbYes, PROG_NAME
            bIsReEnter = False
            bIsContinue = False
        Else
            Dim bIsExist
            bIsExist = False
            For Each sAcceptFileFormat In cAcceptFileFormats
                If sAcceptFileFormat = sArchiveFileExt Then
                    bIsExist = True
                Else
                    'Do Nothing
                End If
            Next
            If bIsExist = True Then
                bIsReEnter = False
            Else
                MsgBox "対応する圧縮形式ではありません。" & vbNewLine & vbNewLine & sArchiveFileExt, vbOKOnly, PROG_NAME
                bIsReEnter = True
            End If
            bIsContinue = True
        End If
    Loop While bIsReEnter = True
Else
    'Do Nothing
End If

'************************
'*** 対象ファイル選定 ***
'************************
'*** ファイル選定 ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If bIsContinue = True Then
    Dim cTrgtPaths
    Set cTrgtPaths = CreateObject("System.Collections.ArrayList")
    Dim sSelectedPath
    For Each sSelectedPath In cSelectedPaths
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( sSelectedPath )
        bFileExists = objFSO.FileExists( sSelectedPath )
        If bFolderExists = False And bFileExists = True Then
            cTrgtPaths.Add sSelectedPath
        ElseIf bFolderExists = True And bFileExists = False Then
            cTrgtPaths.Add sSelectedPath
        Else
            MsgBox "選択されたオブジェクトが存在しません" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, PROG_NAME
            MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
            bIsContinue = False
        End If
    Next
Else
    'Do Nothing
End If

'*** ファイルパスチェック ***
If bIsContinue = True Then
    If cTrgtPaths.Count = 0 Then
        MsgBox "対象となるファイル/フォルダが存在しません。", vbYes, PROG_NAME
        MsgBox "処理を中断します。", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'**********************
'*** パスワード設定 ***
'**********************
If bIsContinue = True Then
    Dim sPassword
    Dim sPasswordCheck
    bIsReEnter = False
    Do
        sPassword = InputBox( _
                        "圧縮ファイルの解凍パスワードを設定します。" & _
                        "パスワードを入力してください。" & vbNewLine & _
                        vbNewLine & _
                        "(※) 選択されたファイル/フォルダが全て同じパスワードで圧縮されます。", _
                         PROG_NAME, "" )
        If sPassword = "" Then
            MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
            MsgBox "処理を中断します。", vbYes, PROG_NAME
            bIsReEnter = False
            bIsContinue = False
        Else
            sPasswordCheck = InputBox( "確認のため、もう一度パスワードを入力してください。", PROG_NAME )
            If sPasswordCheck = "" Then
                MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
                MsgBox "処理を中断します。", vbYes, PROG_NAME
                bIsReEnter = False
                bIsContinue = False
            Else
                If sPassword = sPasswordCheck Then
                    bIsReEnter = False
                Else
                    MsgBox "パスワードが一致していません。", vbOKOnly, PROG_NAME
                    bIsReEnter = True
                End If
                bIsContinue = True
            End If
        End If
    Loop While bIsReEnter = True
Else
    'Do Nothing
End If

'****************
'*** 実行確認 ***
'****************
If bIsContinue = True Then
    Dim sTrgtPath
    Dim sTrgtPathsStr
    sTrgtPathsStr = ""
    For Each sTrgtPath In cTrgtPaths
        If sTrgtPathsStr = "" Then
            sTrgtPathsStr = sTrgtPath
        Else
            sTrgtPathsStr = sTrgtPathsStr & vbNewLine & sTrgtPath
        End If
    Next
    Dim sOutputPassword
    If HIDE_PASSWORD = True Then
        sOutputPassword = String( Len( sPassword ), "*" )
    Else
        sOutputPassword = sPassword
    End If
    Dim lAnswer
    lAnswer = MsgBox ( _
                    "以下を【パスワード付き圧縮】して、選択ファイルと同じフォルダに格納します。" & vbNewLine & _
                    "よろしいですか？" & vbNewLine & _
                    vbNewLine & _
                    "<<圧縮形式>>" & vbNewLine & _
                    sArchiveFileExt & vbNewLine & _
                    vbNewLine & _
                    "<<圧縮パスワード>>" & vbNewLine & _
                    sOutputPassword & vbNewLine & _
                    vbNewLine & _
                    "<<対象ファイル/フォルダパス(※)>>" & vbNewLine & _
                    sTrgtPathsStr & vbNewLine & _
                    vbNewLine & _
                    "(※) それぞれのファイル/フォルダが圧縮されます！" & vbNewLine & _
                    "     一つの圧縮ファイルになる訳ではありません！", _
                    vbYesNo, _
                    PROG_NAME _
                )
    If lAnswer = vbYes Then
        'Do Nothing
    Else
        MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
        bIsContinue = False
    End If
Else
    'Do Nothing
End If

'****************
'*** 圧縮実行 ***
'****************
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
If bIsContinue = True Then
    For Each sTrgtPath In cTrgtPaths
        Dim sArchiveFilePath
        Dim bRet
        Dim lAddedPathType
        bRet = GetNotExistPath( sTrgtPath & "." & sArchiveFileExt, sArchiveFilePath, lAddedPathType )
        Dim sExecCmd
        sExecCmd = """" & sExePath & """ a """ & sArchiveFilePath & """ """ & sTrgtPath & """ -p" & sPassword
        objWshShell.Run sExecCmd, 1, True
    Next
    MsgBox "圧縮が完了しました。", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

Set objFSO = Nothing
Set objWshShell = Nothing

' ==================================================================
' = 概要    指定パスが存在する場合、"_XXX" を付与して返却する
' = 引数    sTrgtPath       String      [in]    対象パス
' = 引数    sAddedPath      String      [out]   付与後のパス
' = 引数    lAddedPathType  Long        [out]   付与後のパス種別
' =                                               1: ファイル
' =                                               2: フォルダ
' = 戻値                    Boolean             取得結果
' = 覚書    本関数では、ファイル/フォルダは作成しない。
' ==================================================================
Public Function GetNotExistPath( _
    ByVal sTrgtPath, _
    ByRef sAddedPath, _
    ByRef lAddedPathType _
)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim bFolderExists
    Dim bFileExists
    bFolderExists = objFSO.FolderExists( sTrgtPath )
    bFileExists = objFSO.FileExists( sTrgtPath )
    
    If bFolderExists = False And bFileExists = True Then
        sAddedPath = GetFileNotExistPath( sTrgtPath )
        lAddedPathType = 1
        GetNotExistPath = True
    ElseIf bFolderExists = True And bFileExists = False Then
        sAddedPath = GetFolderNotExistPath( sTrgtPath )
        lAddedPathType = 2
        GetNotExistPath = True
    Else
        sAddedPath = sTrgtPath
        lAddedPathType = 0
        GetNotExistPath = False
    End If
End Function
    'Call Test_GetNotExistPath()
    Private Sub Test_GetNotExistPath()
        Dim sOutStr
        Dim sAddedPath
        Dim lAddedPathType
        Dim bRet
                                                                                           sOutStr = ""
                                                                                           sOutStr = sOutStr & vbNewLine & "*** test start! ***"
        bRet = GetNotExistPath( "C:\codes\vbs\test\a"     , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\a"     , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\a"     , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\b.txt" , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\b.txt" , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\b.txt" , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\c.txt" , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\c.txt" , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\c.txt" , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\d"     , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\d"     , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\d"     , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\e"     , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\e"     , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
        bRet = GetNotExistPath( "C:\codes\vbs\test\e"     , sAddedPath, lAddedPathType ) : sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
                                                                                           sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
        MsgBox sOutStr
    End Sub

Private Function GetFolderNotExistPath( _
    ByVal sTrgtPath _
)
    Dim lIdx
    Dim objFSO
    Dim sCreDirPath
    Dim bIsTrgtPathExists
    lIdx = 0
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sCreDirPath = sTrgtPath
    bIsTrgtPathExists = False
    Do While objFSO.FolderExists( sCreDirPath )
        bIsTrgtPathExists = True
        lIdx = lIdx + 1
        sCreDirPath = sTrgtPath & "_" & String( 3 - len(lIdx), "0" ) & lIdx
    Loop
    If bIsTrgtPathExists = True Then
        GetFolderNotExistPath = sCreDirPath
    Else
        GetFolderNotExistPath = ""
    End If
End Function
'   Call Test_GetFolderNotExistPath()
    Private Sub Test_GetFolderNotExistPath()
        Dim sOutStr
        sOutStr = ""
        sOutStr = sOutStr & vbNewLine & "*** test start! ***"
        sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath( "C:\codes\vbs\test\a" )
        sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath( "C:\codes\vbs\test\b.txt" )
        sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath( "C:\codes\vbs\test\c.txt" )
        sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath( "C:\codes\vbs\test\d" )
        sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath( "C:\codes\vbs\test\e" )
        sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
        MsgBox sOutStr
    End Sub

Private Function GetFileNotExistPath( _
    ByVal sTrgtPath _
)
    Dim lIdx
    Dim objFSO
    Dim sFileParDirPath
    Dim sFileBaseName
    Dim sFileExtName
    Dim sCreFilePath
    Dim bIsTrgtPathExists
    
    lIdx = 0
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sCreFilePath = sTrgtPath
    bIsTrgtPathExists = False
    Do While objFSO.FileExists( sCreFilePath )
        bIsTrgtPathExists = True
        lIdx = lIdx + 1
        sFileParDirPath = objFSO.GetParentFolderName( sTrgtPath )
        sFileBaseName = objFSO.GetBaseName( sTrgtPath ) & "_" & String( 3 - len(lIdx), "0" ) & lIdx
        sFileExtName = objFSO.GetExtensionName( sTrgtPath )
        If sFileExtName = "" Then
            sCreFilePath = sFileParDirPath & "\" & sFileBaseName
        Else
            sCreFilePath = sFileParDirPath & "\" & sFileBaseName & "." & sFileExtName
        End If
    Loop
    If bIsTrgtPathExists = True Then
        GetFileNotExistPath = sCreFilePath
    Else
        GetFileNotExistPath = ""
    End If
End Function
'   Call Test_GetFileNotExistPath()
    Private Sub Test_GetFileNotExistPath()
        Dim sOutStr
        sOutStr = ""
        sOutStr = sOutStr & vbNewLine & "*** test start! ***"
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath( "C:\codes\vbs\test\a" )
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath( "C:\codes\vbs\test\b.txt" )
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath( "C:\codes\vbs\test\c.txt" )
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath( "C:\codes\vbs\test\d" )
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath( "C:\codes\vbs\test\e" )
        sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
        MsgBox sOutStr
    End Sub

