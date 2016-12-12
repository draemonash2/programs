	#NoEnv							; Recommended for performance and compatibility with future AutoHotkey releases.
;	#Warn							; Enable warnings to assist with detecting common errors.
	SendMode Input					; Recommended for new scripts due to its superior speed and reliability.
;	SetWorkingDir %A_ScriptDir%		; Ensures a consistent starting directory.
;	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え

;[参考URL]
;	https://sites.google.com/site/autohotkeyjp/home
;		無変換）vk1Dsc07B
;		変換）	vk1Csc079
;		+）		Shift
;		^）		Control
;		!）		Alt
;		#）		Windowsロゴキー

;* ***************************************************************
;* Keys
;* ***************************************************************
;*** Global ***
vk1Dsc07B & 1::	Run "C:\prg_exe\Vim\gvim.exe" "%A_MyDocuments%\Dropbox\000_ToDo.txt" "%A_MyDocuments%\Dropbox\920_Music.txt" "%A_MyDocuments%\Dropbox\999_Other.txt"
vk1Dsc07B & 2::	Run "%A_MyDocuments%\Dropbox\300_Mny_AccountsBook.xlsm"
vk1Dsc07B & 3::	Run "C:\prg_exe\Vim\gvim.exe" "%A_ScriptFullPath%"

vk1Dsc07B & c::	RunSuppressMultiStart( "C:\prg_exe\cCalc\cCalc.exe", "" )
vk1Dsc07B & f::	Run "C:\prg_exe\Everything\Everything.exe"
vk1Dsc07B & v::	Run "C:\prg_exe\Vim\gvim.exe" "%A_Desktop%\temp.txt"

;*** Software local ***
#IfWinActive ahk_exe kinza.exe
	;The Great Suspender 用
	F8::^+s
	F9::^+u
	;Shift＋クリックで新規タブ（バックグラウンド）で開く
	+LButton::Send, ^{LButton}
	;Ctrl＋クリックで新規タブで開く
	^LButton::Send, +^{LButton}
#IfWinActive

#IfWinActive ahk_exe CherryPlayer.exe
	win_title = "ahk_exe CherryPlayer.exe"
	#Up::	WinSizeChange( win_title, "up" )
	#Down::	WinSizeChange( win_title, "down" )
	[::		WinSizeChange( win_title, "up" )
	]::		WinSizeChange( win_title, "down" )
#IfWinActive

#IfWinActive ahk_exe EXCEL.EXE
	;F1ヘルプ無効化
	F1::return
#IfWinActive

#IfWinActive ahk_exe mpc-hc.exe
	win_title = "ahk_exe mpc-hc.exe"
	[::	WinSizeChange( win_title, "up" )
	]::	WinSizeChange( win_title, "down" )
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
			filename = %A_LoopField%
		}
		Process, Exist, % filename
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
		MsgBox argument error!
	}
	return
}

WinSizeChange( win_title, size )
{
	if size = up
	{
		WinGetActiveStats, A, WinWidth, WinHeight, WinX, WinY
		if WinX = -32000
		{
			WinRestore, % win_title
		}
		else if WinX = 0
		{
			WinMaximize, % win_title
		}
		else
		{
			WinMaximize, % win_title
		}
	}
	else if size = down
	{
		WinGetActiveStats, A, WinWidth, WinHeight, WinX, WinY
		if WinX = -32000
		{
			WinMinimize, % win_title
		}
		else if WinX = 0
		{
			WinRestore, % win_title
		}
		else
		{
			WinMinimize, % win_title
		}
	}
	else if size = max
	{
		WinMaximize, % win_title
	}
	else if size = restore
	{
		WinRestore, % win_title
	}
	else if size = min
	{
		WinMinimize, % win_title
	}
	else
	{
		MsgBox "[error] please select up or down."
	}
	return
}
