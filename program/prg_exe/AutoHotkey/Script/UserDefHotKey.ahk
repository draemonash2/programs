	#NoEnv							; 通常、値が割り当てられていない変数名を参照しようとしたとき、システムの環境変数に同名の変数がないかを調べ、
									; もし存在すればその環境変数の値が参照される。スクリプト中に #NoEnv を記述することにより、この動作を無効化できる。
;	#Warn							; Enable warnings to assist with detecting common errors.
	SendMode Input					; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。
;	SetWorkingDir %A_ScriptDir%		; スクリプトの作業ディレクトリを本スクリプトの格納ディレクトリに変更。
;	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え

;[参考URL]
;	https://sites.google.com/site/autohotkeyjp/home
;		無変換）vk1Dsc07B
;		変換）	vk1Csc079
;		+）		Shift
;		^）		Control
;		!）		Alt
;		#）		Windowsロゴキー
;	http://ahkwiki.net/Top

;* ***************************************************************
;* Keys
;* ***************************************************************
;*** Global ***
^+1::
^+\::
	Run "C:\prg_exe\Vim\gvim.exe" "%A_MyDocuments%\Dropbox\100_Documents\900_【その他】\100_ToDo.txt" "%A_MyDocuments%\Dropbox\100_Documents\132_【生活】＜趣味＞音楽\音楽ストック.txt" "%A_MyDocuments%\Dropbox\100_Documents\900_【その他】\999_その他.txt" "%A_ScriptFullPath%"
	return
^+2::
^+^::
	Run "%A_MyDocuments%\Dropbox\100_Documents\141_【生活】＜衣食住＞家計\100_家計簿.xlsm"
	return
^+v::
	Run "C:\prg_exe\Vim\gvim.exe" "%A_Desktop%\temp.txt"
	return
^+e::
	Run "C:\codes\vbs\500_CreateExcelFile.vbs"
	return
^+m::
	RunSuppressMultiStart( "C:\prg_exe\cCalc\cCalc.exe", "" )
	return

;*** Software local ***
#IfWINaCTIVE AHK_Exe kinza.exe
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
		Send, {Delete}
		Send, {Delete}
		Sleep 100
		Send, {Enter}
		return
#IfWinActive

#IfWinActive ahk_exe EXCEL.EXE
	;F1ヘルプ無効化
	F1::return
#IfWinActive

#IfWinActive ahk_class MPC-BE
	]::Send, {Space}
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
