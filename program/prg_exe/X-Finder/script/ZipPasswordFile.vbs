'Option Explicit

'<<7-Zip usage>>
'  7z a <zip_file_path> <target_dir_path> -p<password>

Const PROG_NAME = "7-Zip �Ńp�X���[�h���k (zip)"

Dim sExePath
Dim cFilePaths

If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "�f�o�b�O���[�h�ł��B"
    sExePath = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"
    Set cFilePaths = CreateObject("System.Collections.ArrayList")
    cFilePaths.Add "C:\Users\draem_000\Desktop\aa"
    cFilePaths.Add "C:\Users\draem_000\Desktop\b b"
Else
    sExePath = WScript.Env("X-Finder") & "..\" & WScript.Env("7-Zip")
    Set cFilePaths = WScript.Col( WScript.Env("Selected") )
End If

'*** �t�@�C���p�X�`�F�b�N ***
If cFilePaths.Count = 0 Then
    MsgBox "�t�@�C�����I������Ă��܂���B", vbYes, PROG_NAME
    MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
    WScript.Quit
Else
    'Do Nothing
End If

'****************
'*** ���k���s ***
'****************
Dim sTrgtDirPath
For Each sTrgtDirPath In cFilePaths
    Dim sArchiveFilePath
    sArchiveFilePath = sTrgtDirPath & ".zip"
    
    Dim oFileSys
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    Dim bFolderExists
    Dim bFileExists
    bFolderExists = oFileSys.FolderExists( sTrgtDirPath )
    bFileExists = oFileSys.FileExists( sTrgtDirPath )
    
    If bFolderExists = False And bFileExists = True Then
        MsgBox "�t�H���_���w�肵�Ă�������" & vbNewLine & vbNewLine & sTrgtDirPath, vbOKOnly, PROG_NAME
    ElseIf bFolderExists = True And bFileExists = False Then
        Dim lAnswer
        lAnswer = MsgBox ( _
                        "�ȉ����p�X���[�h�t�� ZIP ���k���܂��B��낵���ł����H" & vbNewLine & _
                        vbNewLine & _
                        "<<�Ώۃt�H���_>>" & vbNewLine & _
                        sTrgtDirPath & vbNewLine & _
                        vbNewLine & _
                        "<<Zip �t�@�C����>>" & vbNewLine & _
                        sArchiveFilePath , _
                        vbYesNo, _
                        PROG_NAME _
                    )
        If lAnswer = vbYes Then
            Dim sPassword
            Dim sPasswordCheck
            Dim sExecCmd
            Dim bIsContinue
            Do
                sPassword = InputBox( "�p�X���[�h����͂��Ă��������B", PROG_NAME )
                If sPassword = "" Then
                    MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
                    bIsContinue = False
                Else
                    sPasswordCheck = InputBox( "������x�p�X���[�h����͂��Ă��������B", PROG_NAME )
                    If sPassword = sPasswordCheck Then
                        sExecCmd = """" & sExePath & """ a """ & sArchiveFilePath & """ """ & sTrgtDirPath & """ -p" & sPassword
                        WScript.CreateObject("WScript.Shell").Run sExecCmd, 1, True
                        bIsContinue = False
                    Else
                        MsgBox "�p�X���[�h����v���܂���B", vbOKOnly, PROG_NAME
                        bIsContinue = True
                    End If
                End If
            Loop While bIsContinue = True
        Else
            MsgBox "���s���L�����Z�����܂����B", vbOKOnly, PROG_NAME
        End If
    Else
        MsgBox "�w�肵���t�H���_�����݂��܂���" & vbNewLine & vbNewLine & sTrgtDirPath, vbOKOnly, PROG_NAME
    End If
Next
Set oFileSys = Nothing
Set objWshShell = Nothing
