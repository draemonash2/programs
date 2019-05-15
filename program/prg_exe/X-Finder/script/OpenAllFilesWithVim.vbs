'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "カレントフォルダ配下の特定ファイルを Vim で全て開く"

Dim bIsContinue
bIsContinue = True

Dim objFSO
Dim sExePath
Dim sCurDirPath

'*** 選択ファイル取得 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Dim sArg
        Dim sDefaultPath
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        For Each sArg In WScript.Arguments
            If sDefaultPath = "" Then
                sDefaultPath = objFSO.GetParentFolderName( sArg )
            End If
        Next
        sExePath = "C:\prg_exe\Vim\gvim.exe"
        sCurDirPath = InputBox( "ファイルパスを指定してください", PROG_NAME, sDefaultPath )
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        sExePath = WScript.Env("Vim")
        sCurDirPath = WScript.Env("Current")
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        sExePath = "C:\prg_exe\Vim\gvim.exe"
        sCurDirPath = "C:\codes\c"
    End If
Else
    'Do Nothing
End If

'*** 拡張子選択 ***
If bIsContinue = True Then
    Dim sExtNames
    sExtNames = InputBox( _
        "拡張子を選択してください。" & vbNewLine & _
        "複数の拡張子を指定する時はスペースで区切ります。" & vbNewLine & _
        "  例１）*.txt *.c" & vbNewLine & _
        "  例２）*.*" & vbNewLine & _
        "" , _
        "title", _
        "*.c *.h" _
    )
    If sExtNames = "" Then
        MsgBox "拡張子が選択されていません", vbYes, PROG_NAME
        MsgBox "処理を中断します", vbYes, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** ファイルリスト作成 ***
If bIsContinue = True Then
    'ファイルリスト出力コマンド実行
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    
    Dim sTmpFilePath
    Dim sExecCmd
    sTmpFilePath = objWshShell.SpecialFolders("Templates") & "\open_file_list.txt"
    'MsgBox sTmpFilePath '★DEBUG★
    sExecCmd = "cd """ & sCurDirPath & """ & dir " & sExtNames & " /b /s /a:a-d > """ & sTmpFilePath & """"
    'MsgBox sExecCmd '★DEBUG★
    objWshShell.Run "cmd /c" & sExecCmd, 0, True
    
    '出力したファイルリスト取り込み
    Dim objFile
    Dim sTextAll
    On Error Resume Next
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    Dim asFileList
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
        'MsgBox Err.Number '★DEBUG★
        If Err.Number = 0 Then
            sTextAll = objFile.ReadAll
            sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '末尾に改行が付与されてしまうため、削除
            asFileList = Split( sTextAll, vbNewLine )
            objFile.Close
        Else
            MsgBox "エラーが発生しました。 [ErrorNo." & Err.Number & "] " & Err.Description, vbYes, PROG_NAME
            MsgBox "処理を中断します", vbYes, PROG_NAME
            bIsContinue = False
        End If
        Set objFile = Nothing   'オブジェクトの破棄
    Else
        MsgBox "エラーが発生しました。 [ErrorNo." & Err.Number & "] " & Err.Description, vbYes, PROG_NAME
        MsgBox "処理を中断します。", vbYes, PROG_NAME
        bIsContinue = False
    End If
    objFSO.DeleteFile sTmpFilePath, True
    Set objFSO = Nothing    'オブジェクトの破棄
    On Error Goto 0
    'MsgBox Ubound(asFileList) '★DEBUG★
Else
    'Do Nothing
End If

'*** ファイルオープン実行 ***
If bIsContinue = True Then
    Dim sFilePathList
    sFilePathList = """"
    Dim lIdx
    lIdx = 0
    For Each sFilePath In asFileList
        If lIdx = 0 Then
            sFilePathList = """" & sFilePath & """"
        Else
            sFilePathList = sFilePathList & " """ & sFilePath & """"
        End If
        lIdx = lIdx + 1
    Next
    'MsgBox sFilePathList '★DEBUG★
    
    objWshShell.Run "cmd /c " & sExePath & " " & sFilePathList, 0, False
Else
    'Do Nothing
End If
