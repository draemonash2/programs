	#NoEnv							; 通常、値が割り当てられていない変数名を参照しようとしたとき、システムの環境変数に同名の変数がないかを調べ、
	
									; もし存在すればその環境変数の値が参照される。スクリプト中に #NoEnv を記述することにより、この動作を無効化できる。
;	#Warn							; Enable warnings to assist with detecting common errors.
	SendMode Input					; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。
;	SetWorkingDir %A_ScriptDir%		; スクリプトの作業ディレクトリを本スクリプトの格納ディレクトリに変更。
	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え

	#Include %A_ScriptDir%\lib\IME.ahk

;* ***************************************************************
;* Settings
;* ***************************************************************
DOC_DIR_PATH = C:\Users\%A_Username%\Dropbox\100_Documents

;* ***************************************************************
;* Keys
;* ***************************************************************
;[参考URL]
;	https://sites.google.com/site/autohotkeyjp/reference/KeyList
;		無変換）vk1Dsc07B
;		変換）	vk1Csc079
;		+）		Shift
;		^）		Control
;		!）		Alt
;		#）		Windowsロゴキー

;ヘルプ表示
!^+F1::
msgbox,
(
Ctrl+Shift+Alt + \：#todo.itmz

Ctrl+Shift+Alt + s：#temp.txt
Ctrl+Shift+Alt + d：#temp.xlsm

Ctrl+Shift+Alt + x：Rapture
Ctrl+Shift+Alt + k：KitchenTimer
Ctrl+Shift+Alt + a：X-Finder.exe
Ctrl+Shift+Alt + c：cCalc.exe

Ctrl+Shift+Alt + _：予算管理.xlsm

Ctrl+Shift+Alt + F12：UserDefHotKey.ahk
Ctrl+Shift+Alt + F10：Bluetoothテザリング起動

Pause：Window最前面On
Alt+Pause：Window最前面Off
（Pause＝Fn+RShift ＠HP Spectre x360）
)
return

;***** Global *****
; *** ファイル起動 ***
	;todo.itmz
		!^+\::
			sExePath = "C:\Program Files (x86)\toketaWare\iThoughts\iThoughts.exe"
			sFilePath = "%DOC_DIR_PATH%\#todo.itmz"
			;StartProgramAndActivate( sExePath, sFilePath )
			Run, %sExePath% %sFilePath%
			return
	;temp.txt
		!^+s::
			sExePath = "C:\prg_exe\Vim\gvim.exe"
			sFilePath = "%DOC_DIR_PATH%\#temp.txt"
			StartProgramAndActivate( sExePath, sFilePath )
			return
	;temp.xlsm
		!^+d::
			sFilePath = "%DOC_DIR_PATH%\#temp.xlsm"
			StartProgramAndActivate( "", sFilePath )
			return
	;#temp.vsdm
		!^+v::
			sFilePath = "%DOC_DIR_PATH%\#temp.vsdm"
			StartProgramAndActivate( "", sFilePath )
			return
	;UserDefHotKey.ahk
		!^+F12::
			sExePath = "C:\prg_exe\Vim\gvim.exe"
			sFilePath = "%A_ScriptFullPath%"
			StartProgramAndActivate( sExePath, sFilePath )
			return
	;予算管理.xlsm
		!^+_::
			sFilePath = "%DOC_DIR_PATH%\210_【衣食住】家計\100_予算管理.xlsm"
			StartProgramAndActivate( "", sFilePath )
			return
	;言語チートシート
		!^+[::
			sFilePath = "C:\codes\lang_cheet_sheet.xlsx"
			StartProgramAndActivate( "", sFilePath )
			return

; *** プログラム起動 ***
	;rapture.exe
		+^!x::
			Run "C:\prg_exe\Rapture\rapture.exe"
			return
	;KitchenTimer.vbs
		+^!k::
			Run "C:\codes\vbs\tools\win\other\KitchenTimer.vbs"
			return
	;xf.exe
		+^!a::
			Run "C:\prg_exe\X-Finder\XF.exe"
			return
	;cCalc.exe
		+^!c::
			RunSuppressMultiStart( "C:\prg_exe\cCalc\cCalc.exe", "" )
			return
	;Window最前面化
		Pause::
			WinSet, AlwaysOnTop, On, A
			MsgBox, 0x43000, Window最前面化, Window最前面On, 2
			Return
		!Pause::
			WinSet, AlwaysOnTop, Off, A
			MsgBox, 0x43000, Window最前面化, Window最前面Off, 2
			Return
	;Bluetoothテザリング起動
		+^!F10::
			Run, control printers
			Sleep 2000
			Send, myp
			Sleep 300
			Send, {AppsKey}
			Sleep 200
			Send, c
			Sleep 200
			Send, a
			Sleep 5000
			Send, !{F4}
			return
	;かなキーをコンテキストメニュー表示へ
		RAlt::AppsKey
			return
	;プリントスクリーン単押しを抑制
		PrintScreen::return
		
	;テスト用
		^Pause::
			MsgBox, ctrlpause
			Return
		+Pause::
			MsgBox, shiftpause
			Return
		
		+^!i::
			Send, ^c
			Sleep 200
			Send, ^e
			Sleep 200
			Send, ^v
			Sleep 200
			Send, {Left 3}
			Sleep 200
			Send, {Backspace 2}
			Sleep 200
			Send, {Space}
			Sleep 200
			Send, {End}
			Sleep 200
			Send, {Space}
			Sleep 200
			Send, openload
			return

;***** Software local *****
	#IfWinActive ahk_exe EXCEL.EXE
		;F1ヘルプ無効化
			F1::return
		;Scroll left.
			+WheelUp::
			SetScrollLockState, On
			SendInput {Left}
			SetScrollLockState, Off
			Return
		;Scroll right.
			+WheelDown::
			SetScrollLockState, On
			SendInput {Right}
			SetScrollLockState, Off
			Return
		;Move prev sheet.
			^+WheelUp::
			SendInput ^{PgUp}
			Return
		;Move next sheet.
			^+WheelDown::
			SendInput ^{PgDn}
			Return
	#IfWinActive
	
	#IfWinActive ahk_exe iThoughts.exe
		;F1ヘルプ無効化
			F1::return
	#IfWinActive
	
	#IfWinActive ahk_exe Rapture.exe
		;Escで終了
			Esc::!F4
			return
	#IfWinActive
	
	#IfWinActive ahk_exe vimrun.exe
		;Escで終了
			Esc::!F4
			return
	#IfWinActive
	
	#IfWinActive AHK_Exe kinza.exe
		;The Great Suspender 用
			F8::^+s
			F9::^+u
			return
	#IfWinActive
	
	#IfWinActive ahk_class MPC-BE
			]::Send, {Space}
			return
	#IfWinActive

;* ***************************************************************
;* Functions
;* ***************************************************************
	RunSuppressMultiStart( path, argument )
	{
		IfInString, path, \
		{
			Loop, Parse, path , \
			{
				sFileName = %A_LoopField%
			}
			Process, Exist, % sFileName
			If ErrorLevel<>0
			{
				WinActivate,ahk_pid %ErrorLevel%
			}
			else
			{
				Run % path . " " . argument
			}
		}
		else
		{
			msgbox path
			MsgBox argument error!
		}
		return
	}
	
	WinSizeChange( size, maxwinx, maxwiny )
	{
		if size = up
		{
			WinGetActiveStats, A, WinWidth, WinHeight, WinX, WinY
			if ( WinX = maxwinx && WinY = maxwiny )
			{
				WinMaximize
			}
			else
			{
				WinMaximize
			}
		}
		else if size = down
		{
			WinGetActiveStats, A, WinWidth, WinHeight, WinX, WinY
			if ( WinX = maxwinx && WinY = maxwiny )
			{
				WinRestore
			}
			else
			{
				WinMinimize
			}
		}
		else if size = max
		{
			WinMaximize
		}
		else if size = restore
		{
			WinRestore
		}
		else if size = min
		{
			WinMinimize
		}
		else
		{
			MsgBox "[error] please select up / down / max / restore / min."
		}
		return
	}
	
	; 既定のショートカットキーとの干渉によりプログラム起動後に
	; ウィンドウがアクティベートされないことがある。(※)
	; 上記問題を対処するため、本関数ではプログラム起動後に
	; ウィンドウをアクティベートする処理を実行する。
	; 
	; (※)例
	; 「Windows キー + 1」はタスクパーに１つ目にピン止め
	; されているプログラムをアクティベートするショートカットキーで
	; あるため、Run 関数を使用してそのまま実行すると、非アクティブ
	; 状態でプログラムが起動してしまう。
	StartProgramAndActivate( sExePath, sFilePath )
	{
		IfInString, sFilePath, \
		{
			;*** extract file name ***
			Loop, Parse, sFilePath , \
			{
				sFileName = %A_LoopField%
			}
			StringReplace, sFileName, sFileName, ", , All
			
			;*** start program ***
			If ( sExePath == "" )
			{
				;MsgBox A ;for debug
				Run, %sFilePath%
			}
			else
			{
				;MsgBox B ;for debug
				Run, %sExePath% %sFilePath%
			}
			
			;*** activate started program ***
			SetTitleMatchMode, 2
			WinWait, %sFileName%
			WinActivate, %sFileName%
			
			;MsgBox %sExePath% ;for debug
			;MsgBox %sFilePath% ;for debug
			;MsgBox %sFileName% ;for debug
		}
		else
		{
			msgbox sFilePath
			MsgBox argument error!
		}
		return
	}
	
