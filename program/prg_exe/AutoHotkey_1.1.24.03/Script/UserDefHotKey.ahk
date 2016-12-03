#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;#SingleInstance force	;このスクリプトが再度呼び出されたらリロードして置き換え

;[参考]
;  https://sites.google.com/site/autohotkeyjp/home

;* ***************************************************************
;* keys
;* ***************************************************************
;*** Global ***
;  無変換）	vk1Dsc07B
;  変換）	vk1Csc079
;  +）		Shift
;  ^）		Control
;  !）		Alt
;  #）		Windowsロゴキー
vk1Dsc07B & F5::Run "%A_ScriptFullPath%"

vk1Dsc07B & 1::Run "C:\prg_exe\Vim\vim80-kaoriya-win64\gvim.exe" "%A_MyDocuments%\Dropbox\000_ToDo.txt" "%A_MyDocuments%\Dropbox\920_Music.txt" "%A_MyDocuments%\Dropbox\999_Other.txt"
vk1Dsc07B & 2::Run "%A_MyDocuments%\Dropbox\300_Mny_AccountsBook.xlsm"
vk1Dsc07B & 3::Run "C:\prg_exe\Vim\vim80-kaoriya-win64\gvim.exe" "%A_ScriptFullPath%"

vk1Dsc07B & c::RunSuppressMultiStart( "C:\prg_exe\cCalc\cCalc.exe" )
vk1Dsc07B & f::Run "C:\prg\Everything\Everything.exe"
vk1Dsc07B & v::Run "C:\prg_exe\Vim\vim80-kaoriya-win64\gvim.exe" "%A_Desktop%\temp.txt"

#IfWinActive ahk_exe kinza.exe
;	::^+u
#IfWinActive

#IfWinActive ahk_exe CherryPlayer.exe
	#Up::WinMaximize, ahk_exe CherryPlayer.exe
	#Right::WinRestore, ahk_exe CherryPlayer.exe
	#Down::WinMinimize, ahk_exe CherryPlayer.exe
#IfWinActive

;* ***************************************************************
;* functions
;* ***************************************************************
RunSuppressMultiStart( path )
{
	IfInString, path, \
	{
		Loop, Parse, path , \
		{
			filename = %A_LoopField%
		}
		Process, Exist, % filename
		If ErrorLevel<>0
			WinActivate,ahk_pid %ErrorLevel%
		else
			Run % path
	}
	else
	{
		MsgBox argument error!
	}
	return
}
