'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'<<7-Zip usage>>
'  7z a <zip_file_path> <target_dir_path> -p<password>

'��TODO���FZIP �t�@�C���ȊO�̈��k����m�F

'####################################################################
'### ���O����
'####################################################################
Dim cAcceptFileFormats
Set cAcceptFileFormats = CreateObject("System.Collections.ArrayList")

'####################################################################
'### �ݒ�
'####################################################################
'7-Zip 16.04 ���k�\�`���i/7-ZipPortable/App/7-Zip/7-zip.chm �����p�j
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
'### �{����
'####################################################################
Const PROG_NAME = "7-Zip �Ńp�X���[�h���k"

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

'********************
'*** ���k�`���I�� ***
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
                            "�ȉ��̒����爳�k�`����I�����ē��͂��Ă��������B" & vbNewLine & _
                            sAcceptFileFormatsStr & vbNewLine, _
                            PROG_NAME, _
                            INITIAL_FILE_EXT _
                        )
        If sArchiveFileExt = "" Then
            MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
            MsgBox "�����𒆒f���܂��B", vbYes, PROG_NAME
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
                MsgBox "�Ή����鈳�k�`���ł͂���܂���B" & vbNewLine & vbNewLine & sArchiveFileExt, vbOKOnly, PROG_NAME
                bIsReEnter = True
            End If
            bIsContinue = True
        End If
    Loop While bIsReEnter = True
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
            cTrgtPaths.Add sSelectedPath
        ElseIf bFolderExists = True And bFileExists = False Then
            cTrgtPaths.Add sSelectedPath
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
        MsgBox "�ΏۂƂȂ�t�@�C��/�t�H���_�����݂��܂���B", vbYes, PROG_NAME
        MsgBox "�����𒆒f���܂��B", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'**********************
'*** �p�X���[�h�ݒ� ***
'**********************
If bIsContinue = True Then
    Dim sPassword
    Dim sPasswordCheck
    bIsReEnter = False
    Do
        sPassword = InputBox( _
                        "���k�t�@�C���̉𓀃p�X���[�h��ݒ肵�܂��B" & _
                        "�p�X���[�h����͂��Ă��������B" & vbNewLine & _
                        vbNewLine & _
                        "(��) �I�����ꂽ�t�@�C��/�t�H���_���S�ē����p�X���[�h�ň��k����܂��B", _
                         PROG_NAME, "" )
        If sPassword = "" Then
            MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
            MsgBox "�����𒆒f���܂��B", vbYes, PROG_NAME
            bIsReEnter = False
            bIsContinue = False
        Else
            sPasswordCheck = InputBox( "�m�F�̂��߁A������x�p�X���[�h����͂��Ă��������B", PROG_NAME )
            If sPasswordCheck = "" Then
                MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
                MsgBox "�����𒆒f���܂��B", vbYes, PROG_NAME
                bIsReEnter = False
                bIsContinue = False
            Else
                If sPassword = sPasswordCheck Then
                    bIsReEnter = False
                Else
                    MsgBox "�p�X���[�h����v���Ă��܂���B", vbOKOnly, PROG_NAME
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
    Dim sOutputPassword
    If HIDE_PASSWORD = True Then
        sOutputPassword = String( Len( sPassword ), "*" )
    Else
        sOutputPassword = sPassword
    End If
    Dim lAnswer
    lAnswer = MsgBox ( _
                    "�ȉ����y�p�X���[�h�t�����k�z���āA�I���t�@�C���Ɠ����t�H���_�Ɋi�[���܂��B" & vbNewLine & _
                    "��낵���ł����H" & vbNewLine & _
                    vbNewLine & _
                    "<<���k�`��>>" & vbNewLine & _
                    sArchiveFileExt & vbNewLine & _
                    vbNewLine & _
                    "<<���k�p�X���[�h>>" & vbNewLine & _
                    sOutputPassword & vbNewLine & _
                    vbNewLine & _
                    "<<�Ώۃt�@�C��/�t�H���_�p�X(��)>>" & vbNewLine & _
                    sTrgtPathsStr & vbNewLine & _
                    vbNewLine & _
                    "(��) ���ꂼ��̃t�@�C��/�t�H���_�����k����܂��I" & vbNewLine & _
                    "     ��̈��k�t�@�C���ɂȂ��ł͂���܂���I", _
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
'*** ���k���s ***
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
    MsgBox "���k���������܂����B", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

Set objFSO = Nothing
Set objWshShell = Nothing

' ==================================================================
' = �T�v    �w��p�X�����݂���ꍇ�A"_XXX" ��t�^���ĕԋp����
' = ����    sTrgtPath       String      [in]    �Ώۃp�X
' = ����    sAddedPath      String      [out]   �t�^��̃p�X
' = ����    lAddedPathType  Long        [out]   �t�^��̃p�X���
' =                                               1: �t�@�C��
' =                                               2: �t�H���_
' = �ߒl                    Boolean             �擾����
' = �o��    �{�֐��ł́A�t�@�C��/�t�H���_�͍쐬���Ȃ��B
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

