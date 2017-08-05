'Option Explicit
'Const PRODUCTION_ENVIRONMENT = 0

Const PROG_NAME = "隠しファイル表示切り替え"

Dim bIsContinue
bIsContinue = True

If bIsContinue = True Then
    If PRODUCTION_ENVIRONMENT = 0 Then
        MsgBox "このスクリプトはデバッグモードでは実行できません。", vbOKOnly, PROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

If bIsContinue = True Then
    If InStr( WScript.Env("Style"), "h" ) > 0 Then
        MsgBox "隠しファイル、システムファイルを【非表示】にします。", vbOKOnly, PROG_NAME
    Else
        MsgBox _
            "隠しファイル、システムファイルを【表示】します。" & vbNewLine & _
            "" & vbNewLine & _
            "(※) エクスプローラーのフォルダ設定にて「保護されたオペレーティングシステムファイルを表示しない（推奨）」がチェックされている場合、システムファイルは表示されません。" _
            , vbOKOnly, PROG_NAME
    End If
    WScript.Exec("Change:Style ~h")
    WScript.Exec("Refresh:4")
Else
    'Do Nothing
End If
