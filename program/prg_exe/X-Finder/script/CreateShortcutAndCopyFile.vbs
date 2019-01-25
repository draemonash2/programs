'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'####################################################################
'### �ݒ�
'####################################################################
Const ADD_DATE_TYPE = 1 '�t�^��������̎�ʁi1:���ݓ����A2:�t�@�C��/�t�H���_�X�V�����j
Const SHORTCUT_FILE_SUFFIX = "#Src#"
Const ORIGINAL_FILE_PREFIX = "#Org#"
Const COPY_FILE_PREFIX     = "#Edt#"

'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�V���[�g�J�b�g���R�s�[�t�@�C���쐬"

Dim bIsContinue
bIsContinue = True

'*** �I���t�@�C���擾 ***
Dim sOrgDirPath
Dim cSelectedPaths
If bIsContinue = True Then
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "�f�o�b�O���[�h�ł��B"
        sOrgDirPath = "X:\100_Documents\200_�y�w�Z�z����\��w�@\�[�~�o�ȕ�"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "X:\100_Documents\200_�y�w�Z�z����\��w�@\�[�~�o�ȕ�\H20�N�x �[�~�o�ȕ�.xls"
        cSelectedPaths.Add "X:\100_Documents\200_�y�w�Z�z����\��w�@\�[�~�o�ȕ�\H21�N�x �[�~�o�ȕ�.xls"
    Else
        sOrgDirPath = WScript.Env("Current")
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

'*** �㏑���m�F ***
If bIsContinue = True Then
    Dim vbAnswer
    vbAnswer = MsgBox( "���Ƀt�@�C�������݂��Ă���ꍇ�A�㏑������܂��B���s���܂����H", vbOkCancel, PROG_NAME )
    If vbAnswer = vbOk Then
        'Do Nothing
    Else
        MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
        bIsContinue = False
    End If
Else
    'Do Nothing
End If

'*** �o�͐�I�� ***
If bIsContinue = True Then
    Dim sDstParDirPath
    sDstParDirPath = InputBox( "�o�͐����͂��Ă��������B", PROG_NAME, sOrgDirPath )
    If sDstParDirPath = "" Then '�L�����Z���̏ꍇ
        MsgBox "���s���L�����Z������܂����B", vbOKOnly, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** �V���[�g�J�b�g�쐬 ***
If bIsContinue = True Then
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    
    Dim sSelectedPath
    For Each sSelectedPath In cSelectedPaths
        '�t�@�C��/�t�H���_����
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( sSelectedPath )
        bFileExists = objFSO.FileExists( sSelectedPath )
        
        Dim sAddDate
        Dim sDstOrgFilePath
        Dim sDstCpyFilePath
        Dim sDstShrtctFilePath
        
        '### �t�@�C�� ###
        If bFolderExists = False And bFileExists = True Then
            '�ǉ�������擾�����`
            If ADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now() )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFile
                Set objFile = objFSO.GetFile( sSelectedPath )
                sAddDate = ConvDate2String( objFile.DateLastModified )
                Set objFile = Nothing
            Else
                MsgBox "�uADD_DATE_TYPE�v�̎w�肪����Ă��܂��I", vbOKOnly, PROG_NAME
                MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
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
            
            '�t�@�C���R�s�[
            objFSO.CopyFile sSelectedPath, sDstOrgFilePath, True
            objFSO.CopyFile sSelectedPath, sDstCpyFilePath, True
            
            '�V���[�g�J�b�g�쐬
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### �t�H���_ ###
        ElseIf bFolderExists = True And bFileExists = False Then
            '�ǉ�������擾�����`
            If ADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now() )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFolder
                Set objFolder = objFSO.GetFolder( sSelectedPath )
                sAddDate = ConvDate2String( objFolder.DateLastModified )
                Set objFolder = Nothing
            Else
                MsgBox "�uADD_DATE_TYPE�v�̎w�肪����Ă��܂��I", vbOKOnly, PROG_NAME
                MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
                Exit For
            End If
            
            Dim sOrgDirName
            sOrgDirName = objFSO.GetFileName( sSelectedPath )
            sDstOrgFilePath    = sDstParDirPath & "\" & sOrgDirName & "_" & ORIGINAL_FILE_PREFIX & "_" & sAddDate
            sDstCpyFilePath    = sDstParDirPath & "\" & sOrgDirName & "_" & COPY_FILE_PREFIX     & "_" & sAddDate
            sDstShrtctFilePath = sDstParDirPath & "\" & sOrgDirName & "_" & SHORTCUT_FILE_SUFFIX & "_" & sAddDate & ".lnk"
            
            '�t�H���_�R�s�[
            objFSO.CopyFolder sSelectedPath, sDstOrgFilePath, True
            objFSO.CopyFolder sSelectedPath, sDstCpyFilePath, True
            
            '�V���[�g�J�b�g�쐬
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### �t�@�C��/�t�H���_�ȊO ###
        Else
            MsgBox "�I�����ꂽ�I�u�W�F�N�g�����݂��܂���" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, PROG_NAME
            MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
            bIsContinue = False
        End If
    Next
    
    Set objFSO = Nothing
    Set objWshShell = Nothing
    
    MsgBox "�V���[�g�J�b�g���R�s�[�t�@�C���̍쐬���������܂����I", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

' ==================================================================
' = �T�v    ������������t�@�C��/�t�H���_���ɓK�p�ł���`���ɕϊ�����
' = ����    sDateRaw    String  [in]    �����i��F2017/8/5 12:59:58�j
' = �ߒl                String          �����i��F20170805_125958�j
' = �o��    �Ȃ�
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
    oRegExp.Pattern = sSearchPattern                '�����p�^�[����ݒ�
    oRegExp.IgnoreCase = True                       '�啶���Ə���������ʂ��Ȃ�
    oRegExp.Global = True                           '������S�̂�����
    Dim oMatchResult
    Set oMatchResult = oRegExp.Execute(sTargetStr)  '�p�^�[���}�b�`���s
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
	Const SET_ATTR_READONLY	= 1		' �ǂݎ���p�t�@�C��
	Const SET_ATTR_HIDDEN	= 2		' �B���t�@�C��
	Const SET_ATTR_SYSTEM	= 4		' �V�X�e���E�t�@�C��
	Const SET_ATTR_ARCHIVE	= 32	' �O��̃o�b�N�A�b�v�ȍ~�ɕύX����Ă����1
	
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile("test.txt ")
End Function
