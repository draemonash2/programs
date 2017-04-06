'Option Explicit

Const PROG_NAME = "WinMergeで比較"

Dim sTmpPath
Dim sExePath
Dim cSelected
If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "デバッグモードです。"
    sTmpPath = "C:\prg_exe\X-Finder\diff_target_path.tmp"
    sExePath = "C:\prg_exe\WinMerge\WinMergeU.exe"
    Set cSelected = CreateObject("System.Collections.ArrayList")
    cSelected.Add "C:\prg_exe\X-Finder\script\FileNameCopy.vbs"
    cSelected.Add "C:\prg_exe\X-Finder\script\FilePathCopy.vbs"
Else
    sTmpPath = WScript.Env("X-Finder") & "diff_target_path.tmp"
    sExePath = WScript.Env("X-Finder") & "..\" & WScript.Env("WinMerge")
    Set cSelected = WScript.Col(WScript.Env("Selected"))
End If

Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sExecCmd
Dim sDiffPath1
Dim sDiffPath2
Dim objTxtFile
If cSelected.Count > 1 Then
    sExecCmd = """" & sExePath & """ -r """ & cSelected.Item(0) & """ """ & cSelected.Item(1) & """"
    objWshShell.Run sExecCmd, 3, False
ElseIf cSelected.Count = 1 Then
    sDiffPath1 = cSelected.Item(0)
    If  objFSO.FileExists( sTmpPath ) Then
        Set objTxtFile = objFSO.OpenTextFile( sTmpPath, 1 )
        sDiffPath2 = objTxtFile.ReadLine
        objTxtFile.Close
        Set objTxtFile = Nothing
        sExecCmd = """" & sExePath & """ -r """ & sDiffPath2 & """ """ & sDiffPath1 & """"
        objWshShell.Run sExecCmd, 3, False
        objFSO.DeleteFile sTmpPath, True
    Else
        Set objTxtFile = objFSO.OpenTextFile( sTmpPath, 2, True )
        objTxtFile.WriteLine sDiffPath1
        objTxtFile.Close
        Set objTxtFile = Nothing
        MsgBox "以下を比較対象として選択します。" & vbNewLine & vbNewLine & sDiffPath1
    End If
Else
    MsgBox "ファイルが選択されていません", vbYes, PROG_NAME
End If
