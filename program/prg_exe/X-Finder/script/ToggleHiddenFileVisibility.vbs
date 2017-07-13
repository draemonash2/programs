'Option Explicit

Const PROG_NAME = "隠しファイル表示切り替え"

If PRODUCTION_ENVIRONMENT = 0 Then
    MsgBox "このスクリプトはデバッグモードでは実行できません。", vbOKOnly, PROG_NAME
    MsgBox "処理を中断します。", vbOKOnly, PROG_NAME
    WScript.Quit
Else
    'Do Nothing
End If

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
