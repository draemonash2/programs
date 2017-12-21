'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "更新日時追加＆リネーム"

Dim bIsContinue
bIsContinue = True

Dim lAnswer
lAnswer = MsgBox ( _
                "ファイル/フォルダ名の末尾に更新日時を付与します。よろしいですか？", _
                vbYesNo, _
                PROG_NAME _
            )
If lAnswer = vbYes Then
    'Do Nothing
Else
    MsgBox "実行をキャンセルしました。", vbOKOnly, PROG_NAME
    bIsContinue = False
End If

'*******************************************************
'* ファイル/フォルダ名取得
'*******************************************************
If bIsContinue = True Then
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
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*******************************************************
'* 追加文字列取得
'*******************************************************
If bIsContinue = True Then
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
        
        Dim vDateRaw
        Dim sDateStr
        Dim sAddStr
        If lFileOrFolder = 1 Then
            Call GetFileInfo( oFilePath, 11, vDateRaw )
            sDateStr = ConvDate2String( vDateRaw )
            sAddStr = "_" & sDateStr
            
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
            Call GetFolderInfo( oFilePath, 11, vDateRaw )
            sDateStr = ConvDate2String( vDateRaw )
            sAddStr = "_" & sDateStr
            
            objFSO.MoveFolder _
                oFilePath, _
                sTrgtDirPath & "\" & sTrgtFileName & sAddStr
        Else
            MsgBox "ファイル/フォルダが不正です。", vbOKOnly, PROG_NAME
            bIsContinue = False
        End If
        
        If bIsContinue = True Then
            'Do Nothing
        Else
            Exit For
        End If
    Next
    
    Set objFSO = Nothing
Else
    'Do Nothing
End If

' ==================================================================
' = 概要    日時形式を変換する。（例：2017/03/22 18:20:14 ⇒ 20170322-182014）
' = 引数    sDateTime   String  [in]  日時（YYYY/MM/DD HH:MM:SS）
' = 戻値                String        日時（YYYYMMDD-HHMMSS）
' = 覚書    主に日時をファイル名やフォルダ名に使用する際に使用する。
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateTime _
)
    On Error Resume Next
    Dim sDateStr
    sDateStr = _
        String(4 - Len(Year(sDateTime)),   "0") & Year(sDateTime)   & _
        String(2 - Len(Month(sDateTime)),  "0") & Month(sDateTime)  & _
        String(2 - Len(Day(sDateTime)),    "0") & Day(sDateTime)    & _
        "-" & _
        String(2 - Len(Hour(sDateTime)),   "0") & Hour(sDateTime)   & _
        String(2 - Len(Minute(sDateTime)), "0") & Minute(sDateTime) & _
        String(2 - Len(Second(sDateTime)), "0") & Second(sDateTime)
    If Err.Number = 0 Then
        ConvDate2String = sDateStr
    Else
        ConvDate2String = ""
    End If
    On Error Goto 0
End Function

' ==================================================================
' = 概要    ファイル情報取得
' = 引数    sTrgtPath       String      [in]    ファイルパス
' = 引数    lGetInfoType    Long        [in]    取得情報種別 (※1)
' = 引数    vFileInfo       Variant     [out]   ファイル情報 (※1)
' = 戻値                    Boolean             取得結果
' = 覚書    以下、参照。
' =     (※1) ファイル情報
' =         [引数]  [説明]                  [プロパティ名]      [データ型]              [Get/Set]   [出力例]
' =         1       ファイル名              Name                vbString    文字列型    Get/Set     03 Ride Featuring Tony Matterhorn.MP3
' =         2       ファイルサイズ          Size                vbLong      長整数型    Get         4286923
' =         3       ファイル種類            Type                vbString    文字列型    Get         MPEG layer 3
' =         4       ファイル格納先ドライブ  Drive               vbString    文字列型    Get         Z:
' =         5       ファイルパス            Path                vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =         6       親フォルダ              ParentFolder        vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         7       MS-DOS形式ファイル名    ShortName           vbString    文字列型    Get         03 Ride Featuring Tony Matterhorn.MP3
' =         8       MS-DOS形式パス          ShortPath           vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =         9       作成日時                DateCreated         vbDate      日付型      Get         2015/08/19 0:54:45
' =         10      アクセス日時            DateLastAccessed    vbDate      日付型      Get         2016/10/14 6:00:30
' =         11      更新日時                DateLastModified    vbDate      日付型      Get         2016/10/14 6:00:30
' =         12      属性                    Attributes          vbLong      長整数型    (※2)       32
' =     (※2) 属性
' =         [値]                [説明]                                      [属性名]    [Get/Set]
' =         1  （0b00000001）   読み取り専用ファイル                        ReadOnly    Get/Set
' =         2  （0b00000010）   隠しファイル                                Hidden      Get/Set
' =         4  （0b00000100）   システム・ファイル                          System      Get/Set
' =         8  （0b00001000）   ディスクドライブ・ボリューム・ラベル        Volume      Get
' =         16 （0b00010000）   フォルダ／ディレクトリ                      Directory   Get
' =         32 （0b00100000）   前回のバックアップ以降に変更されていれば1   Archive     Get/Set
' =         64 （0b01000000）   リンク／ショートカット                      Alias       Get
' =         128（0b10000000）   圧縮ファイル                                Compressed  Get
' ==================================================================
Public Function GetFileInfo( _
    ByVal sTrgtPath, _
    ByVal lGetInfoType, _
    ByRef vFileInfo _
)
    Dim objFSO
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists( sTrgtPath ) Then
        'Do Nothing
    Else
        vFileInfo = ""
        GetFileInfo = False
        Exit Function
    End If
    
    Dim objFile
    Set objFile = objFSO.GetFile(sTrgtPath)
    
    vFileInfo = ""
    GetFileInfo = True
    Select Case lGetInfoType
        Case 1:     vFileInfo = objFile.Name                'ファイル名
        Case 2:     vFileInfo = objFile.Size                'ファイルサイズ
        Case 3:     vFileInfo = objFile.Type                'ファイル種類
        Case 4:     vFileInfo = objFile.Drive               'ファイル格納先ドライブ
        Case 5:     vFileInfo = objFile.Path                'ファイルパス
        Case 6:     vFileInfo = objFile.ParentFolder        '親フォルダ
        Case 7:     vFileInfo = objFile.ShortName           'MS-DOS形式ファイル名
        Case 8:     vFileInfo = objFile.ShortPath           'MS-DOS形式パス
        Case 9:     vFileInfo = objFile.DateCreated         '作成日時
        Case 10:    vFileInfo = objFile.DateLastAccessed    'アクセス日時
        Case 11:    vFileInfo = objFile.DateLastModified    '更新日時
        Case 12:    vFileInfo = objFile.Attributes          '属性
        Case Else:  GetFileInfo = False
    End Select
End Function
'   Call Test_GetFileInfo()
    Private Sub Test_GetFileInfo()
        Dim sBuf
        Dim bRet
        Dim vFileInfo
        sBuf = ""
        Dim sTrgtPath
        sTrgtPath = "C:\codes\vbs\lib\FileSystem.vbs"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFileInfo( sTrgtPath,  1, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイル名："             & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  2, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイルサイズ："         & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  3, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイル種類："           & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  4, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイル格納先ドライブ：" & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  5, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイルパス："           & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  6, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  親フォルダ："             & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  7, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  MS-DOS形式ファイル名："   & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  8, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  MS-DOS形式パス："         & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  9, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  作成日時："               & vFileInfo
        bRet = GetFileInfo( sTrgtPath, 10, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  アクセス日時："           & vFileInfo
        bRet = GetFileInfo( sTrgtPath, 11, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  更新日時："               & vFileInfo
        bRet = GetFileInfo( sTrgtPath, 12, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  属性："                   & vFileInfo
        bRet = GetFileInfo( sTrgtPath, 13, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  ："                       & vFileInfo
        sTrgtPath = "C:\codes\vbs\lib\dummy.vbs"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFileInfo( sTrgtPath,  1, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイル名："             & vFileInfo
        MsgBox sBuf
    End Sub

'ファイル情報は「ファイル名」「属性」が設定可能
'しかし、以下のメソッドにて変更可能なため、実装しない
'  ファイル名： objFSO.MoveFile
'  属性： objFSO.GetFile( "C:\codes\a.txt" ).Attributes
Public Function SetFileInfo( _
   ByVal sTrgtPath, _
   ByVal lSetInfoType, _
   ByVal vFileInfo _
)
    'Do Nothing
End Function

' ==================================================================
' = 概要    フォルダ情報取得
' = 引数    sTrgtPath       String      [in]    フォルダパス
' = 引数    lGetInfoType    Long        [in]    取得情報種別 (※1)
' = 引数    vFolderInfo     Variant     [out]   フォルダ情報 (※1)
' = 戻値                    Boolean             取得結果
' = 覚書    以下、参照。
' =     (※1) フォルダ情報
' =         [引数]  [説明]                  [プロパティ名]      [データ型]              [Get/Set]   [出力例]
' =         1       フォルダ名              Name                vbString    文字列型    Get/Set     Sacrifice
' =         2       フォルダサイズ          Size                vbLong      長整数型    Get         80613775
' =         3       ファイル種類            Type                vbString    文字列型    Get         ファイル フォルダー
' =         4       ファイル格納先ドライブ  Drive               vbString    文字列型    Get         Z:
' =         5       フォルダパス            Path                vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         6       ルート フォルダ         IsRootFolder        vbBoolean   ブール型    Get         False
' =         7       MS-DOS形式ファイル名    ShortName           vbString    文字列型    Get         Sacrifice
' =         8       MS-DOS形式パス          ShortPath           vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         9       作成日時                DateCreated         vbDate      日付型      Get         2015/08/19 0:54:44
' =         10      アクセス日時            DateLastAccessed    vbDate      日付型      Get         2015/08/19 0:54:44
' =         11      更新日時                DateLastModified    vbDate      日付型      Get         2015/04/18 3:38:36
' =         12      属性                    Attributes          vbLong      長整数型    (※2)       16
' =     (※2) 属性
' =         [値]                [説明]                                      [属性名]    [Get/Set]
' =         1  （0b00000001）   読み取り専用ファイル                        ReadOnly    Get/Set
' =         2  （0b00000010）   隠しファイル                                Hidden      Get/Set
' =         4  （0b00000100）   システム・ファイル                          System      Get/Set
' =         8  （0b00001000）   ディスクドライブ・ボリューム・ラベル        Volume      Get
' =         16 （0b00010000）   フォルダ／ディレクトリ                      Directory   Get
' =         32 （0b00100000）   前回のバックアップ以降に変更されていれば1   Archive     Get/Set
' =         64 （0b01000000）   リンク／ショートカット                      Alias       Get
' =         128（0b10000000）   圧縮ファイル                                Compressed  Get
' ==================================================================
Public Function GetFolderInfo( _
    ByVal sTrgtPath, _
    ByVal lGetInfoType, _
    ByRef vFolderInfo _
)
    Dim objFSO
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists( sTrgtPath ) Then
        'Do Nothing
    Else
        vFolderInfo = ""
        GetFolderInfo = False
        Exit Function
    End If
    
    Dim objFolder
    Set objFolder = objFSO.GetFolder(sTrgtPath)
    
    vFolderInfo = ""
    GetFolderInfo = True
    Select Case lGetInfoType
        Case 1:     vFolderInfo = objFolder.Name                'フォルダ名
        Case 2:     vFolderInfo = objFolder.Size                'フォルダサイズ
        Case 3:     vFolderInfo = objFolder.Type                'ファイル種類
        Case 4:     vFolderInfo = objFolder.Drive               'ファイル格納先ドライブ
        Case 5:     vFolderInfo = objFolder.Path                'フォルダパス
        Case 6:     vFolderInfo = objFolder.IsRootFolder        'ルート フォルダ
        Case 7:     vFolderInfo = objFolder.ShortName           'MS-DOS形式ファイル名
        Case 8:     vFolderInfo = objFolder.ShortPath           'MS-DOS形式パス
        Case 9:     vFolderInfo = objFolder.DateCreated         '作成日時
        Case 10:    vFolderInfo = objFolder.DateLastAccessed    'アクセス日時
        Case 11:    vFolderInfo = objFolder.DateLastModified    '更新日時
        Case 12:    vFolderInfo = objFolder.Attributes          '属性
        Case Else:  GetFolderInfo = False
    End Select
End Function
'   Call Test_GetFolderInfo()
    Private Sub Test_GetFolderInfo()
        Dim sBuf
        Dim bRet
        Dim vFolderInfo
        sBuf = ""
        Dim sTrgtPath
        sTrgtPath = "C:\codes\vbs\lib"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFolderInfo( sTrgtPath, 1,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイル名："             & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 2,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイルサイズ："         & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 3,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイル種類："           & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 4,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイル格納先ドライブ：" & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 5,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイルパス："           & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 6,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  親フォルダ："             & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 7,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  MS-DOS形式ファイル名："   & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 8,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  MS-DOS形式パス："         & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 9,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  作成日時："               & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 10, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  アクセス日時："           & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 11, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  更新日時："               & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 12, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  属性："                   & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 13, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  ："                       & vFolderInfo
        sTrgtPath = "C:\codes\vbs\libs"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFolderInfo( sTrgtPath, 1,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイル名："             & vFolderInfo
        MsgBox sBuf
    End Sub
