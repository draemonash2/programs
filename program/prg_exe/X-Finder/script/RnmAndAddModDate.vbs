'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

'####################################################################
'### �ݒ�
'####################################################################


'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�X�V�����ǉ������l�[��"

Dim bIsContinue
bIsContinue = True

Dim lAnswer
lAnswer = MsgBox ( _
                "�t�@�C��/�t�H���_���̖����ɍX�V������t�^���܂��B��낵���ł����H", _
                vbYesNo, _
                PROG_NAME _
            )
If lAnswer = vbYes Then
    'Do Nothing
Else
    MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
    bIsContinue = False
End If

'*******************************************************
'* �t�@�C��/�t�H���_���擾
'*******************************************************
If bIsContinue = True Then
    Dim cFilePaths
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "�f�o�b�O���[�h�ł��B"
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
    
    '*** �t�@�C���p�X�`�F�b�N ***
    If cFilePaths.Count = 0 Then
        MsgBox "�t�@�C�����I������Ă��܂���", vbYes, PROG_NAME
        MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*******************************************************
'* �ǉ�������擾
'*******************************************************
If bIsContinue = True Then
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim oFilePath
    For Each oFilePath In cFilePaths
        '*******************************************************
        '* �t�@�C��/�t�H���_������
        '*******************************************************
        Dim lFileOrFolder '1:�t�@�C���A2:�t�H���_�A0:�G���[�i���݂��Ȃ��p�X�j
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( oFilePath )
        bFileExists = objFSO.FileExists( oFilePath )
        If bFolderExists = False And bFileExists = True Then
            lFileOrFolder = 1 '�t�@�C��
        ElseIf bFolderExists = True And bFileExists = False Then
            lFileOrFolder = 2 '�t�H���_�[
        Else
            lFileOrFolder = 0 '�G���[�i���݂��Ȃ��p�X�j
        End If
        
        '*******************************************************
        '* �t�@�C��/�t�H���_���ύX
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
            MsgBox "�t�@�C��/�t�H���_���s���ł��B", vbOKOnly, PROG_NAME
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
' = �T�v    �����`����ϊ�����B�i��F2017/03/22 18:20:14 �� 20170322-182014�j
' = ����    sDateTime   String  [in]  �����iYYYY/MM/DD HH:MM:SS�j
' = �ߒl                String        �����iYYYYMMDD-HHMMSS�j
' = �o��    ��ɓ������t�@�C������t�H���_���Ɏg�p����ۂɎg�p����B
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
' = �T�v    �t�@�C�����擾
' = ����    sTrgtPath       String      [in]    �t�@�C���p�X
' = ����    lGetInfoType    Long        [in]    �擾����� (��1)
' = ����    vFileInfo       Variant     [out]   �t�@�C����� (��1)
' = �ߒl                    Boolean             �擾����
' = �o��    �ȉ��A�Q�ƁB
' =     (��1) �t�@�C�����
' =         [����]  [����]                  [�v���p�e�B��]      [�f�[�^�^]              [Get/Set]   [�o�͗�]
' =         1       �t�@�C����              Name                vbString    ������^    Get/Set     03 Ride Featuring Tony Matterhorn.MP3
' =         2       �t�@�C���T�C�Y          Size                vbLong      �������^    Get         4286923
' =         3       �t�@�C�����            Type                vbString    ������^    Get         MPEG layer 3
' =         4       �t�@�C���i�[��h���C�u  Drive               vbString    ������^    Get         Z:
' =         5       �t�@�C���p�X            Path                vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =         6       �e�t�H���_              ParentFolder        vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         7       MS-DOS�`���t�@�C����    ShortName           vbString    ������^    Get         03 Ride Featuring Tony Matterhorn.MP3
' =         8       MS-DOS�`���p�X          ShortPath           vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =         9       �쐬����                DateCreated         vbDate      ���t�^      Get         2015/08/19 0:54:45
' =         10      �A�N�Z�X����            DateLastAccessed    vbDate      ���t�^      Get         2016/10/14 6:00:30
' =         11      �X�V����                DateLastModified    vbDate      ���t�^      Get         2016/10/14 6:00:30
' =         12      ����                    Attributes          vbLong      �������^    (��2)       32
' =     (��2) ����
' =         [�l]                [����]                                      [������]    [Get/Set]
' =         1  �i0b00000001�j   �ǂݎ���p�t�@�C��                        ReadOnly    Get/Set
' =         2  �i0b00000010�j   �B���t�@�C��                                Hidden      Get/Set
' =         4  �i0b00000100�j   �V�X�e���E�t�@�C��                          System      Get/Set
' =         8  �i0b00001000�j   �f�B�X�N�h���C�u�E�{�����[���E���x��        Volume      Get
' =         16 �i0b00010000�j   �t�H���_�^�f�B���N�g��                      Directory   Get
' =         32 �i0b00100000�j   �O��̃o�b�N�A�b�v�ȍ~�ɕύX����Ă����1   Archive     Get/Set
' =         64 �i0b01000000�j   �����N�^�V���[�g�J�b�g                      Alias       Get
' =         128�i0b10000000�j   ���k�t�@�C��                                Compressed  Get
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
        Case 1:     vFileInfo = objFile.Name                '�t�@�C����
        Case 2:     vFileInfo = objFile.Size                '�t�@�C���T�C�Y
        Case 3:     vFileInfo = objFile.Type                '�t�@�C�����
        Case 4:     vFileInfo = objFile.Drive               '�t�@�C���i�[��h���C�u
        Case 5:     vFileInfo = objFile.Path                '�t�@�C���p�X
        Case 6:     vFileInfo = objFile.ParentFolder        '�e�t�H���_
        Case 7:     vFileInfo = objFile.ShortName           'MS-DOS�`���t�@�C����
        Case 8:     vFileInfo = objFile.ShortPath           'MS-DOS�`���p�X
        Case 9:     vFileInfo = objFile.DateCreated         '�쐬����
        Case 10:    vFileInfo = objFile.DateLastAccessed    '�A�N�Z�X����
        Case 11:    vFileInfo = objFile.DateLastModified    '�X�V����
        Case 12:    vFileInfo = objFile.Attributes          '����
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
        bRet = GetFileInfo( sTrgtPath,  1, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C�����F"             & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  2, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���T�C�Y�F"         & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  3, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C����ށF"           & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  4, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���i�[��h���C�u�F" & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  5, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���p�X�F"           & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  6, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �e�t�H���_�F"             & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  7, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  MS-DOS�`���t�@�C�����F"   & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  8, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  MS-DOS�`���p�X�F"         & vFileInfo
        bRet = GetFileInfo( sTrgtPath,  9, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �쐬�����F"               & vFileInfo
        bRet = GetFileInfo( sTrgtPath, 10, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �A�N�Z�X�����F"           & vFileInfo
        bRet = GetFileInfo( sTrgtPath, 11, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �X�V�����F"               & vFileInfo
        bRet = GetFileInfo( sTrgtPath, 12, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �����F"                   & vFileInfo
        bRet = GetFileInfo( sTrgtPath, 13, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �F"                       & vFileInfo
        sTrgtPath = "C:\codes\vbs\lib\dummy.vbs"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFileInfo( sTrgtPath,  1, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C�����F"             & vFileInfo
        MsgBox sBuf
    End Sub

'�t�@�C�����́u�t�@�C�����v�u�����v���ݒ�\
'�������A�ȉ��̃��\�b�h�ɂĕύX�\�Ȃ��߁A�������Ȃ�
'  �t�@�C�����F objFSO.MoveFile
'  �����F objFSO.GetFile( "C:\codes\a.txt" ).Attributes
Public Function SetFileInfo( _
   ByVal sTrgtPath, _
   ByVal lSetInfoType, _
   ByVal vFileInfo _
)
    'Do Nothing
End Function

' ==================================================================
' = �T�v    �t�H���_���擾
' = ����    sTrgtPath       String      [in]    �t�H���_�p�X
' = ����    lGetInfoType    Long        [in]    �擾����� (��1)
' = ����    vFolderInfo     Variant     [out]   �t�H���_��� (��1)
' = �ߒl                    Boolean             �擾����
' = �o��    �ȉ��A�Q�ƁB
' =     (��1) �t�H���_���
' =         [����]  [����]                  [�v���p�e�B��]      [�f�[�^�^]              [Get/Set]   [�o�͗�]
' =         1       �t�H���_��              Name                vbString    ������^    Get/Set     Sacrifice
' =         2       �t�H���_�T�C�Y          Size                vbLong      �������^    Get         80613775
' =         3       �t�@�C�����            Type                vbString    ������^    Get         �t�@�C�� �t�H���_�[
' =         4       �t�@�C���i�[��h���C�u  Drive               vbString    ������^    Get         Z:
' =         5       �t�H���_�p�X            Path                vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         6       ���[�g �t�H���_         IsRootFolder        vbBoolean   �u�[���^    Get         False
' =         7       MS-DOS�`���t�@�C����    ShortName           vbString    ������^    Get         Sacrifice
' =         8       MS-DOS�`���p�X          ShortPath           vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         9       �쐬����                DateCreated         vbDate      ���t�^      Get         2015/08/19 0:54:44
' =         10      �A�N�Z�X����            DateLastAccessed    vbDate      ���t�^      Get         2015/08/19 0:54:44
' =         11      �X�V����                DateLastModified    vbDate      ���t�^      Get         2015/04/18 3:38:36
' =         12      ����                    Attributes          vbLong      �������^    (��2)       16
' =     (��2) ����
' =         [�l]                [����]                                      [������]    [Get/Set]
' =         1  �i0b00000001�j   �ǂݎ���p�t�@�C��                        ReadOnly    Get/Set
' =         2  �i0b00000010�j   �B���t�@�C��                                Hidden      Get/Set
' =         4  �i0b00000100�j   �V�X�e���E�t�@�C��                          System      Get/Set
' =         8  �i0b00001000�j   �f�B�X�N�h���C�u�E�{�����[���E���x��        Volume      Get
' =         16 �i0b00010000�j   �t�H���_�^�f�B���N�g��                      Directory   Get
' =         32 �i0b00100000�j   �O��̃o�b�N�A�b�v�ȍ~�ɕύX����Ă����1   Archive     Get/Set
' =         64 �i0b01000000�j   �����N�^�V���[�g�J�b�g                      Alias       Get
' =         128�i0b10000000�j   ���k�t�@�C��                                Compressed  Get
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
        Case 1:     vFolderInfo = objFolder.Name                '�t�H���_��
        Case 2:     vFolderInfo = objFolder.Size                '�t�H���_�T�C�Y
        Case 3:     vFolderInfo = objFolder.Type                '�t�@�C�����
        Case 4:     vFolderInfo = objFolder.Drive               '�t�@�C���i�[��h���C�u
        Case 5:     vFolderInfo = objFolder.Path                '�t�H���_�p�X
        Case 6:     vFolderInfo = objFolder.IsRootFolder        '���[�g �t�H���_
        Case 7:     vFolderInfo = objFolder.ShortName           'MS-DOS�`���t�@�C����
        Case 8:     vFolderInfo = objFolder.ShortPath           'MS-DOS�`���p�X
        Case 9:     vFolderInfo = objFolder.DateCreated         '�쐬����
        Case 10:    vFolderInfo = objFolder.DateLastAccessed    '�A�N�Z�X����
        Case 11:    vFolderInfo = objFolder.DateLastModified    '�X�V����
        Case 12:    vFolderInfo = objFolder.Attributes          '����
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
        bRet = GetFolderInfo( sTrgtPath, 1,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C�����F"             & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 2,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���T�C�Y�F"         & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 3,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C����ށF"           & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 4,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���i�[��h���C�u�F" & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 5,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���p�X�F"           & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 6,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �e�t�H���_�F"             & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 7,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  MS-DOS�`���t�@�C�����F"   & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 8,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  MS-DOS�`���p�X�F"         & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 9,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �쐬�����F"               & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 10, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �A�N�Z�X�����F"           & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 11, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �X�V�����F"               & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 12, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �����F"                   & vFolderInfo
        bRet = GetFolderInfo( sTrgtPath, 13, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �F"                       & vFolderInfo
        sTrgtPath = "C:\codes\vbs\libs"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFolderInfo( sTrgtPath, 1,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C�����F"             & vFolderInfo
        MsgBox sBuf
    End Sub
