/*
	=========================================================================
	Title:							Library Creator
	=========================================================================
	Descrition:						Take a folder and save all the ahk files
									to an '#Include' format for quick storage
	-------------------------------------------------------------------------
	AutoHotkey Version:				1.1.27.02
	Language:						English
	Tested Platform(s):				Win 10
	Author:							Delta
	Contact information:			octalblockmc@gmail.com

	|=======================================================================|
	|	Hotkeys:															|
	|		None															|
	|=======================================================================|
	=========================================================================
*/

#NoEnv
#SingleInstance Force
#MaxHotkeysPerInterval 99000000
#HotkeyInterval 99000000
#NoTrayIcon
#ErrorStdOut
Process, Priority,, H
SetBatchLines, -1
SetKeyDelay, -1, -1
SetMouseDelay, -1
SetWinDelay, -1
SetControlDelay, -1
SetDefaultMouseSpeed, 0
SendMode, Input
CoordMode, Mouse, Screen
SetWorkingDir, %A_ScriptDir%

For Each, Item in A_Args
	If (SplitPath(Item).Ext = "AHK")
		Files .= "#Include, `%A_ScriptDir`%\Lib\" SplitPath(Item).FileName

Menu, File, Add, Select Folder`tCtrl+O, OpenFolder
Menu, File, Add,
Menu, File, Add, Edit`tCtrl+W, Exit

Menu, Edit, Add, Add On..., OpenFolder
Menu, Edit, Add, Add AHK Path Lib, AddLib
Menu, Edit, Add, Clear, ControlClear
Menu, Edit, Add,
Menu, Edit, Add, Copy`tCtrl+C, ControlCopy
Menu, Edit, Add, Cut`tCtrl+X, ControlCut

Menu, Options, Add, Add Onto Clipboard On Copy, Option1
Menu, Options, Check, Add Onto Clipboard On Copy
Check1 := True

Menu, Other, Add, About, About

Menu, Main, Add, File, :File
Menu, Main, Add, Edit, :Edit
Menu, Main, Add, Options, :Options
Menu, Main, Add, Other, :Other

Menu, Tray, Icon, Shell32.dll, 55
Gui, 1:New, +Resize
Gui, 1:Menu, Main
Gui, 1:Font,, Consolas
Gui, 1:Add, Edit, x0 y0 vEdit, % Files
Gui, 1:Show, w560 h333, % "Library Helper, by Delta"
Return

; =======================================
; All Events
; =======================================
GuiClose:
Exit:
ExitApp

GuiSize:
	GuiControl, Move, Edit, % "w" A_GuiWidth " h" A_GuiHeight
	Return

ControlCut:
	Gui, Submit, NoHide
	If (Check1)
		Clipboard .= Edit
	Else
		Clipboard := Edit
ControlClear:
	GuiControl,, Edit, % ""
	Return

ControlCopy:
	Gui, Submit, NoHide
	If (Check1)
		Clipboard .= Edit
	Else
		Clipboard := Edit
	Return

OpenFolder:
	Folder := SelectFolder("", "Please select a folder to use.",, "Select")
	If (!Folder)
		Return
	Gui, Submit, NoHide
	If (A_ThisMenuItem != "Add On...")
		GuiControl,, Edit, % OutStr := ""
	If (!(Folder ~= "\\$"))
		Folder .= "\"
	Loop, Files, % Folder "*.*", R
	{
		If (SplitPath(A_LoopFileFullPath).Ext = "AHK") {
			TrayTip,, % Info .= SplitPath(A_LoopFileLongPath).Dir "`n"
			OutStr .= "#Include, " StrReplace(A_LoopFileLongPath, SplitPath(A_LoopFileLongPath).Dir, "`%A_ScriptDir`%\Lib") "`n"
		}
	}
	GuiControl,, Edit, % OutStr
	Return

AddLib:
	A_AHKDir := SplitPath(A_AhkPath).Dir
	Loop, Files, % A_AHKDir "\*.*", R
	{
		If (SplitPath(A_LoopFileFullPath).Ext = "AHK") {
			TrayTip,, % Info .= SplitPath(A_LoopFileLongPath).Dir "`n"
			OutStr .= "#Include, " StrReplace(StrReplace(A_LoopFileLongPath, SplitPath(A_LoopFileLongPath).Dir, "<"), "<\", "<") ">`n"
		}
	}
	Gui, Submit, NoHide
	GuiControl,, Edit, % OutStr
	Return

Option1:
	Check1 := !Check1
	Menu, Options, ToggleCheck, Add Onto Clipboard On Copy
	Return

About:
	About :=
	(LTrim
	"Created by Delta
	Thank you to"
	)
	MsgBox,
	, % SplitPath(A_ScriptName).NameNoExt
	, %
	Return

; ==============================================================================================
; Functions
; ==============================================================================================

/*
	================================================================================================
	SelectFolder
	================================================================================================

	OutFolder := SelectFolder([ StartingFolder, Prompt, OwnerHwnd, OkBtnLabel ])

	================================================================================================
	OutFolder		[Out, Return]
		The Folder that is sent out, that the user selects.
	StartFolder		[In, Folder]
		The starting folder to use, if ommited, defaults to the last location the user selects from computer-wide.
	Prompt			[In, Str]
		The Prompt (Title), that you can give the user.
	OwnerHWND		[In, HWND]
		The HWND of the window you want to force the dialog to hook to, if blank, uses the script's main window.
	OkBtnLabel		[In, Str]
		The button name you'd like to use, defaults to 'Select Folder'
	================================================================================================
*/

SelectFolder(StartingFolder := "", Prompt := "", OwnerHwnd := 0, OkBtnLabel := "") {
	OwnerHwnd := WinExist(App.Name)
	Static OsVersion := DllCall("GetVersion", "UChar")
		  , IID_IShellItem := 0
		  , InitIID := VarSetCapacity(IID_IShellItem, 16, 0)
						& DllCall("Ole32.dll\IIDFromString", "WStr", "{43826d1e-e718-42ee-bc55-a1e261c37bfe}", "Ptr", &IID_IShellItem)
		  , Show := A_PtrSize * 3
		  , SetOptions := A_PtrSize * 9
		  , SetFolder := A_PtrSize * 12
		  , SetTitle := A_PtrSize * 17
		  , SetOkButtonLabel := A_PtrSize * 18
		  , GetResult := A_PtrSize * 20
	SelectedFolder := ""
	If (OsVersion < 6) { ; IFileDialog requires Win Vista+, so revert to FileSelectFolder
		FileSelectFolder, SelectedFolder, *%StartingFolder%, 3, %Prompt%
		Return SelectedFolder
	}
	OwnerHwnd := DllCall("IsWindow", "Ptr", OwnerHwnd, "UInt") ? OwnerHwnd : 0
	If !(FileDialog := ComObjCreate("{DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7}", "{42f85136-db7e-439c-85f1-e4075d135fc8}"))
		Return ""
	VTBL := NumGet(FileDialog + 0, "UPtr")
	; FOS_CREATEPROMPT | FOS_NOCHANGEDIR | FOS_PICKFOLDERS
	DllCall(NumGet(VTBL + SetOptions, "UPtr"), "Ptr", FileDialog, "UInt", 0x00002028, "UInt")
	If (StartingFolder <> "")
		If !DllCall("Shell32.dll\SHCreateItemFromParsingName", "WStr", StartingFolder, "Ptr", 0, "Ptr", &IID_IShellItem, "PtrP", FolderItem)
			DllCall(NumGet(VTBL + SetFolder, "UPtr"), "Ptr", FileDialog, "Ptr", FolderItem, "UInt")
	If (Prompt <> "")
		DllCall(NumGet(VTBL + SetTitle, "UPtr"), "Ptr", FileDialog, "WStr", Prompt, "UInt")
	If (OkBtnLabel <> "")
		DllCall(NumGet(VTBL + SetOkButtonLabel, "UPtr"), "Ptr", FileDialog, "WStr", OkBtnLabel, "UInt")
	If !DllCall(NumGet(VTBL + Show, "UPtr"), "Ptr", FileDialog, "Ptr", OwnerHwnd, "UInt") {
		If !DllCall(NumGet(VTBL + GetResult, "UPtr"), "Ptr", FileDialog, "PtrP", ShellItem, "UInt") {
			GetDisplayName := NumGet(NumGet(ShellItem + 0, "UPtr"), A_PtrSize * 5, "UPtr")
			If !DllCall(GetDisplayName, "Ptr", ShellItem, "UInt", 0x80028000, "PtrP", StrPtr) ; SIGDN_DESKTOPABSOLUTEPARSING
				SelectedFolder := StrGet(StrPtr, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "Ptr", StrPtr)
			ObjRelease(ShellItem)
	}  }
	If (FolderItem)
		ObjRelease(FolderItem)
	ObjRelease(FileDialog)
	Return, SelectedFolder
}

/*
	============================================
	SplitPath
	============================================

	Usage:

	--------------------------------------------
	OutObj := SplitPath(File)
	--------------------------------------------
	Vars:
		OutObj		[Out, Return]
			The variable to push the object to.
		File		[In, File/URL]
			The File or URL to parse from.

	============================================

	--------------------------------------------
	OutData := SplitPath(File).ObjName
	--------------------------------------------
	Vars:
		OutData		[Out, Return]
			The variable to push the string of data.
		File		[In, File/URL]
			The File or URL to parse from.
		ObjName		[Attach, VarName]
			The object name to use when grabbing data.
			Usable names:
				FileName: The Filename with the extention.
				Dir: The directory the file/URL is from.
				Ext: The extention of the file/URL
				NameNoExt: Same as FileName, but no extention.
				Drive: The drive/domain of the file/URL.
*/
SplitPath(File) {
	SplitPath, File, F, D, E, N, D_
	Return, {FileName: F, Dir: D, Ext: E, NameNoExt: N, Drive: D_}
}
