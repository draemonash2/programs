'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'<<7-Zip usage>>
'  7z x -o<target_dir_path> <archive_file_path>

'<<���L����>>
'  ZipFile.vbs�AZipPasswordFile.vbs �́A���k���ɍ쐬�\��̈��k�t�@�C������
'  �����̈��k�t�@�C���������łɑ��݂��Ă����ꍇ�A�㏑�����Ȃ��悤
'  vbs �X�N���v�g���ň��k�t�@�C����_XXX.zip �ɕύX���鏈�����s���B
'  ����AUnzipFile.vbs�i�{�X�N���v�g�j�͏�L�̂悤�ȏ㏑��������邱�Ƃ��ł��Ȃ��B
'  ����́A7-Zip �R�}���h���C���d�l��A�𓀌�̃t�H���_����ύX�ł��Ȃ����߂ł���B

'��TODO���FZIP �t�@�C���ȊO�̉𓀓���m�F

'####################################################################
'### ���O����
'####################################################################
Dim cAcceptFileFormats
Set cAcceptFileFormats = CreateObject("System.Collections.ArrayList")

'####################################################################
'### �ݒ�
'####################################################################
'7-Zip 16.04 ��(�W�J)�\�`���i/7-ZipPortable/App/7-Zip/7-zip.chm �����p�j
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
'### �{����
'####################################################################
Const PROG_NAME = "7-Zip �ŉ�"

Dim sExePath
Dim cSelectedPaths

Dim bIsContinue
bIsContinue = True

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "�f�o�b�O���[�h�ł��B"
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

'*** �t�@�C���p�X�`�F�b�N ***
If bIsContinue = True Then
    If cSelectedPaths.Count = 0 Then
        MsgBox "�t�@�C��/�t�H���_���I������Ă��܂���B", vbOKOnly, PROG_NAME
        MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'************************
'*** �Ώۃt�@�C���I�� ***
'************************
'*** �t�@�C���I�� ***
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
            MsgBox "�I�����ꂽ�I�u�W�F�N�g�����݂��܂���" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, PROG_NAME
            MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
            bIsContinue = False
        End If
    Next
Else
    'Do Nothing
End If

'*** �t�@�C���p�X�`�F�b�N ***
If bIsContinue = True Then
    If cTrgtPaths.Count = 0 Then
        MsgBox "�ΏۂƂȂ�t�@�C�������݂��܂���B", vbYes, PROG_NAME
        MsgBox "�����𒆒f���܂��B", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'****************
'*** ���s�m�F ***
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
                    "�ȉ����y�𓀁z���āA�I���t�@�C���Ɠ����t�H���_�Ɋi�[���܂��B" & vbNewLine & _
                    "��낵���ł����H" & vbNewLine & _
                    vbNewLine & _
                    "<<�Ώۃt�@�C���p�X>>" & vbNewLine & _
                    sTrgtPathsStr, _
                    vbYesNo, _
                    PROG_NAME _
                )
    If lAnswer = vbYes Then
        'Do Nothing
    Else
        MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
        MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
        bIsContinue = False
    End If
Else
    'Do Nothing
End If

'****************
'*** �𓀎��s ***
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
    MsgBox "�𓀂��������܂����B", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

Set objFSO = Nothing
Set objWshShell = Nothing
