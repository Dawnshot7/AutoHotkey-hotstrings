#Persistent
#SingleInstance force
SetCapsLockState, alwaysoff
SetNumLockState, alwayson
tabskip=0
CoordMode, mouse, Screen
;;;;;;;;;;;;PERSONAL AUTOREMOTE URL: Append your message to this to send it to your phone;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
autoremotemessageurl:="https://autoremotejoaomgcd.appspot.com/sendmessage?key=APA91bFM-ChfXwLNzUtIxP3LsFdidob4bk1h4p2hHQDrWSGsfu2n5vO0Hoy0JvZBNuyXucYcQZqvf-VQVAe3NTjK6B05LzSz5duuLOnuHoSbyaXMl5WwVrt-9ZsH4Pzoi079WJAzuEbH&message="


;;;;;;;;;;;;;;;;HOTSTRINGS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;HOTSTRINGS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

CapsLock & g::	
	input, thismessage,T45,g							; Inputs what you type into a variable, finishes when "g" is pressed  					
	if (thismessage=="n") {
		IfWinExist ahk_class Notepad++
			Winactivate, ahk_class Notepad++
		else {
			IfWinNotExist ahk_class Notepad
				Run, Notepad.exe
			Winactivate, ahk_class Notepad
		}	
	}
	if (thismessage=="f") {
		IfWinNotExist ahk_class MozillaWindowClass 
			if (a_computername == "LEGEND")
				run C:\Program Files\Mozilla Firefox\firefox.exe
			else
				run D:\Program Files\firefox.exe
		Winactivate, ahk_class MozillaWindowClass
	}
	if (thismessage=="x") 
		Winactivate, ahk_class XLMAIN
	if (thismessage=="w") 
		Winactivate, ahk_class OpusApp
	if (thismessage=="p") 
		Winactivate, ahk_class PPTFrameClass
	if (thismessage=="e") {
		IfWinNotExist ahk_class CabinetWClass
			run, Explorer.exe
		Winactivate, ahk_class CabinetWClass
	}
	if (thismessage=="v") {
		IfWinNotExist ahk_class Chrome_WidgetWin_1
			run, D:\Microsoft VS Code\Code.exe
		Winactivate, ahk_class Chrome_WidgetWin_1
	}
	if (thismessage=="a") 
		Winactivate, ahk_class AcrobatSDIWindow
	if (thismessage=="t") {
		IfWinNotExist ahk_class #32770
			run, C:\Program Files (x86)\TeamViewer\TeamViewer.exe
		Winactivate, ahk_class #32770
		Winactivate, ahk_class TV_CClientWindowClass
	}
	if (thismessage=="h") {
		IfWinNotExist "Hud"
			run, C:\Program Files (x86)\Fonality\HUD3.5\HUD3.exe
		Winactivate, "Hud"
	}
	if (thismessage=="fm") 
		Winactivate, Dedham
	if (thismessage=="cl") {
		wingetclass, thisclass, A
		msgbox, %thisclass%
	}
	if (thismessage=="c") 
		sendinput, !{F4}
	if (thismessage=="ct") 
		sendinput, ^{F4}
	if (thismessage=="ss" or thismessage=="sv") 
		sendinput, ^s
	return


;;;;;;;;;;;;;;;;HOTKEYS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;HOTKEYS;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;VSCODE or NOTEPAD;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#if WinActive("ahk_class Chrome_WidgetWin_1") or WinActive("ahk_class Notepad") or WinActive("ahk_class Notepad++")
:?*:rld::											; Reload script
	sendinput, ^s
	reload
	return
^+r::												; Reload script
	sendinput, ^s
	reload
	return
#if
^+t::	
	ControlGetFocus, ctrl, ahk_class XLMAIN
	if (ctrl="EXCEL<1" OR ctrl="EXCEL61")
	{
		sendinput, {enter}
		Return
	}
	ComObjActive("Excel.Application").ActiveWindow.SmallScroll(0,0,0,8)  ; Scroll left. 
	return
;VSCODE or TEAMVIEWER MAC;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#if WinActive("ahk_class Chrome_WidgetWin_1") or WinActive("ahk_class TV_CClientWindowClass")
:?*:runahk::											; Run current active file
	sendinput, ^!r
	return
CapsLock & s::sendinput, {Home}						; Go to beginning of line
CapsLock & f::sendinput, {End}						; Go to end of line
CapsLock & w::sendinput, {Shift up}^+k				; Delete line
CapsLock & r::sendinput, {Shift up}^{Enter}			; Insert line below
CapsLock & '::sendinput, {Shift up}^/				; Toggle line comment
CapsLock & t::sendinput, ^+9						; Focus terminal
CapsLock & 4::sendinput, ^+6						; Focus previous editor
CapsLock & 5::sendinput, ^+7						; Focus next editor
#if

;FIREFOX;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#if WinActive("ahk_class MozillaWindowClass")
CapsLock & s::sendinput, {Shift up}^+{Tab}			; Next tab
CapsLock & f::sendinput, {Shift up}^{Tab}			; Previous tab
CapsLock & w::sendinput, {Shift up}!{Left}			; Browser back
CapsLock & r::sendinput, {Shift up}!{Right}			; Browser forward
#if

;EXCEL;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#if WinActive("ahk_class XLMAIN")
CapsLock & 4::sendinput, !4							; Open cell dropdown
CapsLock & 5::sendinput, !5							; Drag series down
#if

;ANY WINDOW;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
:?*:eml::											; Run current active file
	sendinput, paul.cameron.haines+2gmail.com
	return
:*?:jj::{Esc}										; Press Escape
CapsLock & /::sendinput, {Shift up}^f				; Find
CapsLock & 1::sendinput, ^{-}						; Zoom out
CapsLock & 2::sendinput, ^{+}						; Zoom in
CapsLock & Space::
	if (GetKeyState("Shift"))
		sendinput, {Shift up}						; Start select
	else
		sendinput, {Shift down}						; End select
	return
CapsLock & H::sendinput, ^{Left}					; Go to previous word
CapsLock & SC027::sendinput, ^{Right}				; Go to next Word
CapsLock & u::sendinput, {Shift up}{BS}				; Backspace
CapsLock & o::sendinput, {Shift up}{Delete}			; Delete
CapsLock & y::										; Backspace word
	if (!WinActive("ahk_class TV_CClientWindowClass"))
		sendinput, {Shift up}^{BS}					
	else
		sendinput, !{BS}
	return 
CapsLock & p:: 										; Delete word
	if (!WinActive("ahk_class TV_CClientWindowClass"))
		sendinput, {Shift up}^{Delete}				
	else
		sendinput, !{Delete}
	return
CapsLock & c::sendinput, {Shift up}^c				; Copy
CapsLock & v::
	clipboard=%clipboard%
	sendinput, {Shift up}^v				; Paste
	return
CapsLock & n::sendinput, {Shift up}^z				; Undo
CapsLock & m::sendinput, {Shift up}^y				; Redo
CapsLock & i::										; Up
	if (not GetKeyState("Shift")) 
		sendinput, {Shift up}
	sendinput, {Up}				
	return
CapsLock & k::										; Down
	if (not GetKeyState("Shift")) 
		sendinput, {Shift up}
	sendinput, {Down}				
	return
CapsLock & j::										; Left
	if (not GetKeyState("Shift")) 
		sendinput, {Shift up}
	sendinput, {Left}				
	return
CapsLock & l::										; Right
	if (not GetKeyState("Shift")) 
		sendinput, {Shift up}
	sendinput, {Right}				
	return
CapsLock & e::
	if (not GetKeyState("Shift")) 					; Up 10
		sendinput, {Shift up}
	sendinput, {Up 10}								
	return
CapsLock & d::										; Down 10
	if (not GetKeyState("Shift")) 
		sendinput, {Shift up}
	sendinput, {Down 10}							
	return
CapsLock & SC00D::									; Volume up
	SoundGet, oldVolume
    SoundSet, oldVolume+10
    return
CapsLock & SC00C::									; Volume down
	SoundGet, oldVolume
    SoundSet, oldVolume-10
    return
Tab::return
Tab Up::
	if (tabskip==0) 
		sendinput, {Tab}
	tabskip=0
	return
CapsLock & lbutton::
	clicknum++
	mousegetpos, mx,my,mwin,mnn
	wingetclass, mclass, A
	gui,mldata:default
	lv_add("",mx,my,mclass)
	lv_modifycol()
	fileappend, %mx%`,%my%`,%mclass%`n, %pathscriptdir%\MLtraining.txt  
	iniwrite, %mx%|%my%|%mclass%, %pathscriptdir%\MLtraining.ini, clicks, %clicknum%
	return
CapsLock & rbutton::
	pathscriptdir:="C:\Users\PaulC\Documents\Scripts"
	filedelete, %pathscriptdir%\MLtraining.txt
	inidelete, %pathscriptdir%\MLtraining.ini, clicks
	clicknum:=0
	return
#if (GetKeyState("Tab","P"))
s::
	WinGetPos, X, Y, Width, Height, A
	WinMove, A, , X, Y, Width-40, Height
	tabskip=1
	return
f::
	WinGetPos, X, Y, Width, Height, A
	WinMove, A, , X, Y, Width+40, Height
	tabskip=1
	return
e::
	WinGetPos, X, Y, Width, Height, A
	WinMove, A, , X, Y, Width, Height-40
	tabskip=1
	return
d::
	WinGetPos, X, Y, Width, Height, A
	WinMove, A, , X, Y, Width, Height+40
	tabskip=1
	return
i::
	WinGetPos, X, Y, Width, Height, A
	WinMove, A, , X, Y-40
	tabskip=1
	return
k::
	WinGetPos, X, Y, Width, Height, A
	WinMove, A, , X, Y+40
	tabskip=1
	return
j::
	WinGetPos, X, Y, Width, Height, A
	WinMove, A, , X-40, Y
	tabskip=1
	return
l::
	WinGetPos, X, Y, Width, Height, A
	WinMove, A, , X+40, Y
	tabskip=1
	return		
g::
	WinRestore, A
	tabskip=1
	return		
h::
	WinMaximize, A
	tabskip=1
	return		
#if

;ANY MOUSE;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
rbutton & wheeldown::sendinput, {wheeldown 10}		; Fast scroll down
rbutton & wheelup::sendinput, {wheelup 10}			; Fast scroll up
rbutton::send, {rbutton down}{rbutton up}			; Preserves rbutton alone functionality

;GAMING MOUSE;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#if (a_computername="LEGEND") 						; The mouse buttons on my Roccat are mapped to numpad keys
	numpadadd::sendinput, {alt down}{left}{alt up}		
	numpadsub::sendinput, {alt down}{right}{alt up}
	numpad0::sendinput, {ctrl down}t{ctrl up}
	numpad1::sendinput, {shift down}{ctrl down}{tab}{ctrl up}{shift up}
	numpad2::sendinput, {delete}
	numpad3::sendinput, {ctrl down}{tab}{ctrl up}
	numpad4::sendinput, {ctrl down}c{ctrl up}
	numpad6::sendinput, {enter}
	numpad7::
		coordmode,mouse,screen 
		settitlematchmode, 2
		MouseGetPos, xr, yr, MouseWindowUID, MouseControlID ;gathers mouse location info and mouseovered window info
		WinGetClass, class, ahk_id %MouseWindowUID%
		DetectHiddenWindows, On
		WinGet, OutputVar, ID, LaunchPadMin.ahk
		coordmode,mouse,window 
		if (MouseControlID="MSTaskListWClass1") { ;click to close mouseovered item in taskbar
			send, {esc}{Rbutton down}{rbutton up}
			sleep, 200	
			send, {tab}{up}{enter}
			return 
		}
		if (class="CabinetWClass") { ;click to close mouseovered explorer folder
			click, ,left, 1
			sendinput, {alt down}{F4}{alt up}
			return 
		}
		if (MouseControlID="SysListView321") ;click to close mouseovered item in taskbar
		{
			if (xr<78 && yr>900) {
				click,,right,1
				sendinput, b
				sleep,200
				sendinput, {space}
				return
			}
		}
		if (class="IEFrame") ;close tab in internet explorer
		{ 
			settitlematchmode, 2
			ifwinnotactive, Toodledo
			{
				ifwinnotactive, Kerio
				{
					sendinput, ^{F4}
					return
				}
			}
		}
		if (class="Chrome_WidgetWin_1" || class="MozillaWindowClass") ;close tab in internet explorer
		{ 
			settitlematchmode, 2
			ifwinnotactive, Toodledo
			{
				ifwinnotactive, Kerio
				{
					sendinput, ^{F4}
					return
				}
			}
		}
		send, {tab} ;tab in all other windows
		coordmode,mouse,window 
		return	
	numpad8::sendinput, {bs}
	numpad9::sendinput, {ctrl down}v{ctrl up}	

	
	rbutton & numpad5::send, {alt down}{tab}{alt up}
	rbutton & numpad6::	send, {CTRL down}{shift down}{tab}{shift up}{CTRL up}	;activate next tab in internet explorer
	rbutton & numpad7::	send, {CTRL down}{tab}{CTRL up}	;activate previous tab in internet explorer
	rbutton & lbutton up::send, {ctrl up}{lbutton up}	;holds control while clicking if rbutton is held down
	rbutton & lbutton::send, {ctrl down}{lbutton down}	;holds control while clicking if rbutton is held down

	
	numpad5 & numpad4::	;save and reload this AutoHotkey file
		send, {ctrl down}s{ctrl up}
		reload
		return
	numpad5 & numpad7::send, {shift down}{tab}{shift up}
	numpad5 & numpad8::send, {ctrl down}{shift down}z{shift up}{ctrl up} ;redo
	numpad5 & numpad9::send, {ctrl down}z{ctrl up}	;undo
	numpad5 & rbutton::	
		ifwinactive, ahk_class IEFrame
			sendinput, {rbutton}{down 2}{enter}
		ifwinactive, ahk_class MozillaWindowClass
		{
			sendinput, {rbutton down}{rbutton up}
			coordmode,mouse,screen
			MouseGetPos, xb, yb
			Mousemove, xb+10, yb+5
			click,,left,1
		}
		return
	numpad5 & wheeldown::
		SoundGet, oldVolume
    	SoundSet, oldVolume-10
    	return
	numpad5 & wheelup::
		SoundGet, oldVolume
    	SoundSet, oldVolume+10
    	return
#if

;AUTOREMOTE;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
:?*:ddf::	;communicate with my phone hotstring
	input, thismessage,T45,{enter}
	ifnotinstring,thismessage,aaa
		gosub, sendAR
	return
sendAR:	;communicate with phone subroutine to be processed as a voice command
	AR := ComObjCreate("InternetExplorer.Application")
	thismessage:=autoremotemessageurl . "processvoice=:=" . thismessage
	AR.Navigate(thismessage)
	sleep, 500
	AR.Quit
	return