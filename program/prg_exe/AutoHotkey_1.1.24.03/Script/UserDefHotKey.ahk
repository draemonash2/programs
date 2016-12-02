#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;*** Global ***
vk1Dsc07B & 1::Run "%A_MyDocuments%\Dropbox\000_ToDo.txt"
vk1Dsc07B & 2::Run "%A_MyDocuments%\Dropbox\300_Mny_AccountsBook.xlsm"
vk1Dsc07B & c::Run "C:\prg_exe\cCalc\cCalc.exe"
vk1Dsc07B & f::Run "C:\prg\Everything\Everything.exe"
vk1Dsc07B & v::Run "C:\prg_exe\Vim\vim80-kaoriya-win64\gvim.exe" "%A_Desktop%\temp.txt"
vk1Dsc07B & h::Run "C:\prg_exe\Vim\vim80-kaoriya-win64\gvim.exe" "%A_ScriptFullPath%"
vk1Dsc07B & u::Run "%A_ScriptFullPath%"

;*** Kinza ***
#IfWinActive ahk_exe kinza.exe
;	::^+u
#IfWinActive
