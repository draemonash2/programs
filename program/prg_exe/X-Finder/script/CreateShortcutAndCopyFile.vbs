'Option Explicit

'####################################################################
'### �ݒ�
'####################################################################

'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�V���[�g�J�b�g���R�s�[�t�@�C���쐬"

Const SHORTCUT_FILE_PREFIX = "CopySource"

Dim bIsContinue
bIsContinue = True

'*** �I���t�@�C���擾 ***
Dim cSelectedPaths
If bIsContinue = True Then
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "�f�o�b�O���[�h�ł��B"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "X:\100_Documents\200_�y�w�Z�z����\��w�@\�[�~�o�ȕ�\H20�N�x �[�~�o�ȕ�.xls"
        cSelectedPaths.Add "X:\100_Documents\200_�y�w�Z�z����\��w�@\�[�~�o�ȕ�\H21�N�x �[�~�o�ȕ�.xls"
    Else
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

'*** �V���[�g�J�b�g�쐬 ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim oRegExp
Set oRegExp = CreateObject("VBScript.RegExp")
If bIsContinue = True Then
    Dim sSelectedPath
    For Each sSelectedPath In cSelectedPaths
        '�t�@�C��/�t�H���_����
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( sSelectedPath )
        bFileExists = objFSO.FileExists( sSelectedPath )
        If bFolderExists = False And bFileExists = True Then
            '�ǉ�������擾�����`
            Dim objFile
            Set objFile = objFSO.GetFile( sSelectedPath )
            Dim sDateLastModifiedRaw
            sDateLastModifiedRaw = objFile.DateLastModified
            
            Dim sSearchPattern
            Dim sTargetStr
            sSearchPattern = "(\w{4})/(\w{1,2})/(\w{1,2}) (\w{1,2}):(\w{1,2}):(\w{1,2})"
            sTargetStr = sDateLastModifiedRaw
            oRegExp.Pattern = sSearchPattern                '�����p�^�[����ݒ�
            oRegExp.IgnoreCase = True                       '�啶���Ə���������ʂ��Ȃ�
            oRegExp.Global = True                           '������S�̂�����
            Dim oMatchResult
            Set oMatchResult = oRegExp.Execute(sTargetStr)  '�p�^�[���}�b�`���s
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
            
            '�t�@�C���R�s�[
            objFSO.CopyFile sSelectedPath, sCopyDstFilePath, True
            
            '�V���[�g�J�b�g�쐬
            With objWshShell.CreateShortcut( sShortcutFilePath )
                .TargetPath = sParentDirPath
                .Save
            End With
            
            '�I��
            '��
        ElseIf bFolderExists = True And bFileExists = False Then
            '�f�B���N�g���͑ΏۊO�Ƃ���
        Else
            MsgBox "�I�����ꂽ�I�u�W�F�N�g�����݂��܂���" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, PROG_NAME
            MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
            bIsContinue = False
        End If
    Next
Else
    'Do Nothing
End If

MsgBox "�V���[�g�J�b�g���R�s�[�t�@�C���̍쐬���������܂����I", vbOKOnly, PROG_NAME

Set oRegExp = Nothing
Set objFSO = Nothing
Set objWshShell = Nothing
