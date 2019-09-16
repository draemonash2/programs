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
DOC_DIR_PATH = C:\Users\%A_Username%\Dropbox\100_Documents\

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

;***** Global *****
; *** ファイル起動（Win + Fn） ***
	;todo.itmz
		#F1::
			sExePath = "C:\Program Files (x86)\toketaWare\iThoughts\iThoughts.exe"
			sFilePath = "%DOC_DIR_PATH%todo.itmz"
			StartProgramAndActivate( sExePath, sFilePath )
			return
	;予算管理.xlsm
		#F2::
			sFilePath = "%DOC_DIR_PATH%210_【衣食住】家計\100_予算管理.xlsm"
			StartProgramAndActivate( "", sFilePath )
			return
	;日記.txtを初期化して起動
		#F3::
			RunWait "%DOC_DIR_PATH%900_【他】\日記を初期化して起動.bat""
			sExePath = "C:\prg_exe\Vim\gvim.exe"
			sFilePath = "%DOC_DIR_PATH%900_【他】\日記.txt"
			StartProgramAndActivate( sExePath, sFilePath )
			return
	;temp.txt
		#F5::
			sExePath = "C:\prg_exe\Vim\gvim.exe"
			sFilePath = "%DOC_DIR_PATH%temp.txt"
			StartProgramAndActivate( sExePath, sFilePath )
			return
	;temp.xlsm
		#F6::
			sFilePath = "%DOC_DIR_PATH%temp.xlsm"
			StartProgramAndActivate( "", sFilePath )
			return
	;other.itmz
		#F11::
			sExePath = "C:\Program Files (x86)\toketaWare\iThoughts\iThoughts.exe"
			sFilePath = "%DOC_DIR_PATH%other.itmz"
			StartProgramAndActivate( sExePath, sFilePath )
			return
	;UserDefHotKey.ahk
		#F12::
			sExePath = "C:\prg_exe\Vim\gvim.exe"
			sFilePath = "%A_ScriptFullPath%"
			StartProgramAndActivate( sExePath, sFilePath )
			return
	;かなキーをコンテキストメニュー表示へ
		RAlt::AppsKey
			return
	;プリントスクリーン単押しを抑制
		PrintScreen::return

; *** プログラム起動（Win + Alt + Fn）***
	;rapture.exe
		#!F1::
			Run "C:\prg_exe\Rapture\rapture.exe"
			return
	;xf.exe
		+^!a::
			Run "C:\prg_exe\X-Finder\XF.exe"
			return
	;cCalc.exe
		#!F2::
			RunSuppressMultiStart( "C:\prg_exe\cCalc\cCalc.exe", "" )
			return
	;Bluetoothテザリング起動
		#!F11::
			Run, control printers
			Sleep 1000
			Send, myp
			Sleep 200
			Send, {AppsKey}
			Sleep 100
			Send, c
			Sleep 100
			Send, a
			Sleep 5000
			Send, !{F4}
			return
	;ウィンドウ最前面切り替え
		#!F12::
			WinSet, AlwaysOnTop, TOGGLE, A
			MsgBox, 0x43000, ウィンドウ最前面切り替え, アクティブウィンドウ最前面化の有効/無効を切り替えます, 5
			Return

;***** Software local *****
	#IfWinActive ahk_exe EXCEL.EXE
		;F1ヘルプ無効化
			F1::return
	#IfWinActive
	
	#IfWinActive ahk_exe iThoughts.exe
		;F1ヘルプ無効化
			F1::return
		;IME OFF
			IME_SET(0)
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
			;Shift＋クリックで新規タブ（バックグラウンド）で開く
			+LButton::Send, ^{LButton}
			;Ctrl＋クリックで新規タブで開く
			^LButton::Send, +^{LButton}
			+^i::
				Send, ^c
				Sleep 100
				Send, !d
				Sleep 100
				Send, ^v
				Sleep 100
				Send, {Home}
				Sleep 100
				Send, ^{Right}
				Sleep 100
				Send, ^{Right}
				Sleep 100
				Send, {Left}
				Sleep 100
				Send, {Delete}
				Send, {Delete}
				Sleep 100
				Send, {Enter}
				return
	#IfWinActive
	
	#IfWinActive ahk_class MPC-BE
		;★
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
	
