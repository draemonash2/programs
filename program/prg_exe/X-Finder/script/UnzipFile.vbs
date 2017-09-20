'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'<<7-Zip usage>>
'  7z x -o<target_dir_path> <archive_file_path>

'<<特記事項>>
'  ZipFile.vbs、ZipPasswordFile.vbs は、圧縮時に作成予定の圧縮ファイル名と
'  同名の圧縮ファイル名がすでに存在していた場合、上書きしないよう
'  vbs スクリプト内で圧縮ファイル名_XXX.zip に変更する処理を行う。
'  一方、UnzipFile.vbs（本スクリプト）は上記のような上書きを避けることができない。
'  これは、7-Zip コマンドライン仕様上、解凍後のフォルダ名を変更できないためである。

'★TODO★：ZIP ファイル以外の解凍動作確認

'####################################################################
'### 事前処理
'####################################################################
Dim cAcceptFileFormats
Set cAcceptFileFormats = CreateObject("System.Collections.ArrayList")

'####################################################################
'### 設定
'####################################################################
'7-Zip 16.04 解凍(展開)可能形式（/7-ZipPortable/App/7-Zip/7-zip.chm より引用）
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
cAcceptFileFormats.Add "apm"      ' APM
cAcceptFileFormats.Add "ar"       ' AR
cAcceptFileFormats.Add "a"        ' AR
cAcceptFileFormats.Add "deb"      ' AR
cAcceptFileFormats.Add "lib"      ' AR
cAcceptFileFormats.Add "arj"      ' ARJ
cAcceptFileFormats.Add "cab"      ' CAB
cAcceptFileFormats.Add "chm"      ' CHM
cAcceptFileFormats.Add "chw"      ' CHM
cAcceptFileFormats.Add "chi"      ' CHM
cAcceptFileFormats.Add "chq"      ' CHM
cAcceptFileFormats.Add "msi"      ' COMPOUND
cAcceptFileFormats.Add "msp"      ' COMPOUND
cAcceptFileFormats.Add "doc"      ' COMPOUND
cAcceptFileFormats.Add "xls"      ' COMPOUND
cAcceptFileFormats.Add "ppt"      ' COMPOUND
cAcceptFileFormats.Add "cpio"     ' CPIO
cAcceptFileFormats.Add "cramfs"   ' CramFS
cAcceptFileFormats.Add "dmg"      ' DMG
cAcceptFileFormats.Add "ext"      ' Ext
cAcceptFileFormats.Add "ext2"     ' Ext
cAcceptFileFormats.Add "ext3"     ' Ext
cAcceptFileFormats.Add "ext4"     ' Ext
cAcceptFileFormats.Add "img"      ' Ext
cAcceptFileFormats.Add "fat"      ' FAT
cAcceptFileFormats.Add "img"      ' FAT
cAcceptFileFormats.Add "hfs"      ' HFS
cAcceptFileFormats.Add "hfsx"     ' HFS
cAcceptFileFormats.Add "hxs"      ' HXS
cAcceptFileFormats.Add "hxi"      ' HXS
cAcceptFileFormats.Add "hxr"      ' HXS
cAcceptFileFormats.Add "hxq"      ' HXS
cAcceptFileFormats.Add "hxw"      ' HXS
cAcceptFileFormats.Add "lit"      ' HXS
cAcceptFileFormats.Add "ihex"     ' iHEX
cAcceptFileFormats.Add "iso"      ' ISO
cAcceptFileFormats.Add "img"      ' ISO
cAcceptFileFormats.Add "lzh"      ' LZH
cAcceptFileFormats.Add "lha"      ' LZH
cAcceptFileFormats.Add "lzma"     ' LZMA
cAcceptFileFormats.Add "mbr"      ' MBR
cAcceptFileFormats.Add "mslz"     ' MsLZ
cAcceptFileFormats.Add "mub"      ' Mub
cAcceptFileFormats.Add "nsis"     ' NSIS
cAcceptFileFormats.Add "ntfs"     ' NTFS
cAcceptFileFormats.Add "img"      ' NTFS
cAcceptFileFormats.Add "mbr"      ' MBR
cAcceptFileFormats.Add "rar"      ' RAR
cAcceptFileFormats.Add "r00"      ' RAR
cAcceptFileFormats.Add "rpm"      ' RPM
cAcceptFileFormats.Add "ppmd"     ' PPMD
cAcceptFileFormats.Add "qcow"     ' QCOW2
cAcceptFileFormats.Add "qcow2"    ' QCOW2
cAcceptFileFormats.Add "qcow2c"   ' QCOW2
cAcceptFileFormats.Add "001"      ' SPLIT
cAcceptFileFormats.Add "002"      ' SPLIT
cAcceptFileFormats.Add "squashfs" ' SquashFS
cAcceptFileFormats.Add "udf"      ' UDF
cAcceptFileFormats.Add "iso"      ' UDF
cAcceptFileFormats.Add "img"      ' UDF
cAcceptFileFormats.Add "scap"     ' UEFIc
cAcceptFileFormats.Add "uefif"    ' UEFIs
cAcceptFileFormats.Add "vdi"      ' VDI
cAcceptFileFormats.Add "vhd"      ' VHD
cAcceptFileFormats.Add "vmdk"     ' VMDK
cAcceptFileFormats.Add "wim"      ' WIM
cAcceptFileFormats.Add "esd"      ' WIM
cAcceptFileFormats.Add "xar"      ' XAR
cAcceptFileFormats.Add "pkg"      ' XAR
cAcceptFileFormats.Add "z"        ' Z
cAcceptFileFormats.Add "taz"      ' Z

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "7-Zip で解凍"

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
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\aa.zip"
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\b b.zip"
        cSelectedPaths.Add "C:\Users\draem_000\Desktop\test\cc"
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
            Dim sFileExt
            sFileExt = objFSO.GetExtensionName( sSelectedPath )
            
            Dim bIsTrgtFile
            bIsTrgtFile = False
            Dim sAcceptFileFormat
            For Each sAcceptFileFormat In cAcceptFileFormats
                If sAcceptFileFormat = sFileExt Then
                    bIsTrgtFile = True
                    Exit For
                Else
                    'Do Nothing
                End If
            Next
            If bIsTrgtFile = True Then
                cTrgtPaths.Add sSelectedPath
            Else
                'Do Nothing
            End If
        ElseIf bFolderExists = True And bFileExists = False Then
            'Do Nothing
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
        MsgBox "対象となるファイルが存在しません。", vbYes, PROG_NAME
        MsgBox "処理を中断します。", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
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
    Dim lAnswer
    lAnswer = MsgBox ( _
                    "以下を【解凍】して、選択ファイルと同じフォルダに格納します。" & vbNewLine & _
                    "よろしいですか？" & vbNewLine & _
                    vbNewLine & _
                    "<<対象ファイルパス>>" & vbNewLine & _
                    sTrgtPathsStr, _
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
'*** 解凍実行 ***
'****************
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
If bIsContinue = True Then
    For Each sTrgtPath In cTrgtPaths
        Dim sOutputDirPath
        sOutputDirPath = objFSO.GetParentFolderName( sTrgtPath )
        Dim sExecCmd
        sExecCmd = """" & sExePath & """ x -o""" & sOutputDirPath & """ """ & sTrgtPath & """"
        objWshShell.Run sExecCmd, 1, True
    Next
    MsgBox "解凍が完了しました。", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

Set objFSO = Nothing
Set objWshShell = Nothing
