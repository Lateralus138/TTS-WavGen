; Text to speech gui using AutoHotkey and VBScript.

; Init
SetBatchLines,-1
#SingleInstance,Force
#NoTrayIcon
RunAsAdmin(A_ScriptFullPath)

; Vars
_title:="TTS WavGen"
_hover_msga:="Use periods '...' to create delays in time."
_help_msg=
(
%_title% by Ian Pride @ New Pride Services 2017

%_title% is a program that uses AutoHotkey to generate a
VBScript that converts text to speech in Windows wav 
format.

To use:
Simply type or paste any text into the provided edit box and
click the [Convert Text To Speech] button.

You can select what folder you would like to save your file in
and name your file.

To create delays in words or sentences add E.g. "." where you
would  like your delay. This delay method is a hack method as
true VBS delays in the message would use the:
<wscript.sleep time>
method. This abilty will be added in the next version.
)
maxInnerW:=524
mrgnX:=mrgnY:=8
spc:=mrgnX/2
example_script=
(
Const SAFT48kHz16BitStereo = 39
Const SSFMCreateForWrite = 3
Dim oFileStream, oVoice
Set oFileStream = CreateObject("SAPI.SpFileStream")
oFileStream.Format.Type = SAFT48kHz16BitStereo
oFileStream.Open "replace_fdreplace_fl.wav", SSFMCreateForWrite
Set oVoice = CreateObject("SAPI.SpVoice")
Set oVoice.AudioOutputStream = oFileStream
oVoice.Speak "Example Text..."
oFileStream.Close
)
default_folder:=A_MyDocuments "\"
default_filename:="Example File Name"
folderTxtW:=maxInnerW-88
enterTxtW:=maxInnerW-26
help_y:=spc-2
; Gui
Gui,Font,s11 c0xFFFFFF,Segoe UI
Gui,Color,0x1D1D1D,0xFFFFFF
Gui,Margin,0,0
GuiButton(_title,"DummyVar",0x2D89EF,0x3091ff,,,540,32) ; This button is a dummy and serves as the title bar
MinButton("488","22",FFFFFF,FFFFFF)
CloseButton("514","6",FFFFFF,FFFFFF)
Gui,Add,Text,x8 y+8 Section w%maxInnerW%,Select folder to save file:
Gui,Font,c0x1D1D1D w600,Consolas
Gui,Add,Edit,xs y+4 w%folderTxtW% vSaveFolder,%default_folder%
Gui,Font,c0xFFFFFF w500,Segoe UI
GuiButton("&Browse","BrowseVar",0x2D89EF,0x3091ff,"+0","p",88,26)
Gui,Add,Text,xs y+4 w%maxInnerW%,Choose a file name:
Gui,Font,c0x1D1D1D w600,Consolas
Gui,Add,Edit,xs y+4 w%maxInnerW% vFileName,%default_filename%
Gui,Font,c0xFFFFFF w500,Segoe UI
Gui,Add,Text,xs y+4 w%enterTxtW%,Enter text to convert to speech below:
Gui,Add,Picture,Icon24 w18 h18 x+8 yp+%help_y% gHelp,shell32.dll
Gui,Font,c0x1D1D1D w600,Consolas
Gui,Add,Text,xs y+0 w%maxInnerW% h%spc%
Gui,Add,Edit,xs y+0 w%maxInnerW% vTxtToCnvrt
Gui,Font,c0xFFFFFF w500,Segoe UI
GuiButton("Con&vert Text To Speech","ConvertVar",0x2D89EF,0x3091ff,"p","+" spc,maxInnerW,24)
GuiButton("Hide &Status","StatusVar",0x2D89EF,0x3091ff,"p","+" spc,100,24)
Gui,Add,Text,xs y+0 w%maxInnerW% h%mrgnY%
Gui,Add,StatusBar,vStatus,Current working folder: %default_folder%
Gui,Font,c0xEE1111,Consolas
GuiControl,Font,Status
Gui,+LastFound
Gui,Show,,%_title%
gui_id:=WinExist()
WinSet,Style,-0xC00000,ahk_id %gui_id%
Gui,Show,AutoSize,%_title%
GetControls(_title)
; Window messages
OnMessage(0x201,"WM_LBUTTONDOWN")
OnMessage(0x200,"WM_MOUSEHOVER")
OnMessage(0x4E,"Help")
SetTimer,ToolTipOff,-1
GuiControl,+Background0x1D1D1D,Status
SetTimer,HideShowStatus,-1
Return

; Hotkeys
If WinActive("ahk_id " gui_id)
	!v::SetTimer,Convert,-1
	!b::SetTimer,Browse,-1
	!/::SetTimer,Help,-1
	!s::SetTimer,HideShowStatus,-1
#IfWinActive

; Functions
;#Include,texttalk_funcs.ahk
MinButton(x,y,lcolor,dcolor,small:=False){
	small:=small?3:4
	big:=small*3
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% x%x% y%y% w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% yp x+0 w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% yp x+0 w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% yp x+0 w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% yp x+0 w%small% h%small%, 100
}
CloseButton(x,y,lcolor,dcolor,small:=False){
	small:=small?3:4
	big:=small*3
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% x%x% y%y% w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% y+0 x+0 w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% y+0 x+0 w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% y+0 xp-%small% w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% y+0 xp-%small% w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% yp-%big% xp+%big% w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% yp-%small% x+0 w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% yp+%big% xp-%small% w%small% h%small%, 100
	Gui, Add, Progress, Background0x%lcolor% c0x%dcolor% y+0 x+0 w%small% h%small%, 100
}
TrayTip(title,text,seconds:=0,options:=0){
	TrayTip,%title%,%text%,%seconds%,%options%
}
RunAsAdmin(file){
	If FileExist(file){
		If ! A_IsAdmin {
			Try, Run *RunAs %file%
			Catch
				Return 1
			If (file=A_ScriptFullPath Or file=A_ScriptName)
				ExitApp
		}
	}
}
Toggle(var){
	Global
	%var%:=!%var%
	If (%var%=1)
		Return True
}
WM_MOUSEHOVER(){
	Global
	If ! A_TimeIdlePhysical {
		ToolTip
	}
	Sleep,1500
	ToolTip %	MouseOver(E3X,E3Y,E3X2,E3Y2)
			?	"It may be easier to write your message in an editor and paste it here..."
			:	MouseOver(MP17X,MP17Y,MP17X2,MP17Y2)
			?	"Click to convert text to speech..."
			:	""
	If ! A_TimeIdlePhysical
		ToolTip
}
WM_LBUTTONDOWN(){
	Global
	If MouseOver(MP1X,MP1Y,MP1X2,MP1Y2)
		PostMessage, 0xA1, 2,,,ahk_id %gui_id%
	If MouseOver(MP2X,MP2Y,MP6X2,MP6Y2)
		WinMinimize,ahk_id %gui_id%
	If MouseOver(MP7X,MP7Y,MP15X2,MP15Y2)
		ExitApp
	If MouseOver(MP16X,MP16Y,MP16X2,MP16Y2)
		SetTimer,Browse,-1
	If MouseOver(MP17X,MP17Y,MP17X2,MP17Y2)
		SetTimer,Convert,-1
	If MouseOver(MP18X,MP18Y,MP18X2,MP18Y2)
		SetTimer,HideShowStatus,-1
}
; Need to rewrite EnumVars() and EnumControlPos()
;ControlList
GetControls(title,control:=0,posvar:=0){
	If (control && posvar)
		{
			namenum:=EnumVarName(control)
			ControlGetPos,x,y,w,h,%control%,%title%
			pos:=(posvar == "X")?x
			:(posvar == "Y")?y
			:(posvar == "W")?w
			:(posvar == "H")?h
			:(posvar == "X2")?x+w
			:(posvar == "Y2")?Y+H
			:0
			Globals.SetGlobal(namenum posvar,pos)
			Return pos
		}
	Else If !(control && posvar)
		{
			WinGet,a,ControlList,%title%
			Loop,Parse,a,`n
				{
					namenum:=EnumVarName(A_LoopField)
					If namenum
						{
							ControlGetPos,x,y,w,h,%A_LoopField%,%title%
							Globals.SetGlobal(namenum "X",x)
							Globals.SetGlobal(namenum "Y",y)
							Globals.SetGlobal(namenum "W",w)
							Globals.SetGlobal(namenum "H",h)
							Globals.SetGlobal(namenum "X2",x+w)
							Globals.SetGlobal(namenum "Y2",y+h)				
						}
				}
			Return a
		}
}
EnumVarName(control){
	name:=InStr(control,"msctls_p")?"MP"
	:InStr(control,"Static")?"S"
	:InStr(control,"Button")?"B"
	:InStr(control,"Edit")?"E"
	:InStr(control,"ListBox")?"LB"
	:InStr(control,"msctls_u")?"UD"
	:InStr(control,"ComboBox")?"CB"
	:InStr(control,"ListView")?"LV"
	:InStr(control,"SysTreeView")?"TV"
	:InStr(control,"SysLink")?"L"
	:InStr(control,"msctls_h")?"H"
	:InStr(control,"SysDate")?"TD"
	:InStr(control,"SysMonthCal")?"MC"
	:InStr(control,"msctls_t")?"SL"
	:InStr(control,"msctls_s")?"SB"
	:InStr(control,"327701")?"AX"
	:InStr(control,"SysTabC")?"T"
	:0
	num:=(name == "MP")?SubStr(control,18)
	:(name == "S")?SubStr(control,7)
	:(name == "B")?SubStr(control,7)
	:(name == "E")?SubStr(control,5)
	:(name == "LB")?SubStr(control,8)
	:(name == "UD")?SubStr(control,15)
	:(name == "CB")?SubStr(control,9)
	:(name == "LV")?SubStr(control,14)
	:(name == "TV")?SubStr(control,14)
	:(name == "L")?SubStr(control,8)
	:(name == "H")?SubStr(control,16)
	:(name == "TD")?SubStr(control,18)
	:(name == "MC")?SubStr(control,14)
	:(name == "SL")?SubStr(control,18)
	:(name == "SB")?SubStr(control,19)
	:(name == "AX")?SubStr(control,5)
	:(name == "T")?SubStr(control,16)
	:0
	Return name num
}
MouseOver(x1,y1,x2,y2){
	MouseGetPos,_x,_y
	Return (_x>=x1 AND _x<=x2 AND _y>=y1 AND _y<=y2)
}
Debug(title:="Debug Pause",msg:=0,type:=64,exit:=0){
	type:=(Mod(type,16)=0)?type:64
	If (type = 64)
		MsgBox,64,%title%,%msg%
	If (type = 48)
		MsgBox,48,%title%,%msg%
	If (type = 32)
		MsgBox,32,%title%,%msg%
	If (type = 16)
		MsgBox,16,%title%,%msg%
	If exit
		ExitApp
}
GuiButton(title,txtvar,color:=0x1D1D1D,border:=0x1D1D1D,x:="+0",y:="+0",w:="",h=""){
	Global
	%txtvar%:=title
	Gui,Add,Progress,x%x% y%y% w%w% h%h% Background%border% c%color% ,100
	Gui,Add,Text,w%w% h%h% xp yp Center +BackgroundTrans 0x200 v%txtvar%,%title%
}

; Classes
Class Globals { ; my favorite way to set and retrive global tions. Good for
	SetGlobal(name,value=""){ ; setting globals from other tions
		Global
		%name%:=value
		Return
	}
	GetGlobal(name){	
		Global
		Local var:=%name%
		Return var
	}
}

; Subs
HideShowStatus:
	GuiControl,% Toggle("HideShow")?"Hide":"Show",Status
	GuiControl,,StatusVar,% Toggle("ShowHide")?"Show &Status":"Hide &Status"
	Gui,Show,AutoSize,%_title%
Return
Browse:
	GuiControl,,BrowseVar,&Browsing...
	lastFolder:=NewFolder?NewFolder:default_folder
	FileSelectFolder,NewFolder
	GuiControl,,BrowseVar,&Browse
	GuiControl,,SaveFolder,% NewFolder?NewFolder "\":lastFolder
	SB_SetText(NewFolder?"Current working folder: " NewFolder "\":"Current working folder: " lastFolder)
Return
Convert:
	GuiControl,,ConvertVar,Con&verting Text To Speech...
	Gui,Submit,NoHide
	SB_SetText("Checking passed parameters...")
	If ! TxtToCnvrt {
		SB_SetText("No text provided...")
		Debug(_title " " A_ThisLabel " Error","No text provided...",16)
		GuiControl,,ConvertVar,Con&vert Text To Speech
		Return
	}
	If ! InStr(FileExist(SaveFolder),"D") {
		SB_SetText("No valid folder provided...")
		Debug(_title " " A_ThisLabel " Error","No valid folder provided...",16)
		GuiControl,,ConvertVar,Con&vert Text To Speech
		Return	
	}
	If ! FileName {
		SB_SetText("No file name provided...")
		Debug(_title " " A_ThisLabel " Error","No file name provided...",16)
		GuiControl,,ConvertVar,Con&vert Text To Speech
		Return
	}
	SB_SetText("Creating VBScript text...")
	replacement:=StrReplace(example_script,"Example Text...",TxtToCnvrt)
	replacement:=StrReplace(replacement,"replace_fd",SaveFolder)
	replacement:=StrReplace(replacement,"replace_fl",FileName)
	SB_SetText("Checking if old script exists...")
	If FileExist(SaveFolder FileName ".vbs"){
		SB_SetText("Attempting to delete old script...")
		Gosub,DeleteVBS	
	}
	SB_SetText("Writing VBScript to file: " SaveFolder FileName ".vbs")
	Try,FileAppend,%replacement%,%SaveFolder%%FileName%.vbs
	Catch {
		SB_SetText("Could not write " SaveFolder FileName ".vbs to file...")
		Debug(A_ThisLabel " Error","Could not save VBScript to the provided folder.`nMake sure you have the appropriate permissions.",16)
		GuiControl,,ConvertVar,Con&vert Text To Speech
		Return
	}
	SB_SetText(SaveFolder FileName ".vbs successfully written to file...")
	Sleep,2000
	SB_SetText("Attempting to run file "  SaveFolder FileName ".vbs")
	Try,RunWait,%SaveFolder%%FileName%.vbs
	Catch {
		SB_SetText("Could not run file "  SaveFolder FileName ".vbs")
		Debug(A_ThisLabel " Error","Could not run then generated VBScript.`nMake sure you have the appropriate permissions.",16)
		GuiControl,,ConvertVar,Con&vert Text To Speech
		If FileExist(SaveFolder FileName ".vbs"){
			SB_SetText("Attempting to delete old script...")
			Gosub,DeleteVBS	
		}
		Return	
	}
	SB_SetText("Successfully converted your text to speech...")
	TrayTip,%_title% Info,You can find your wav file at: %SaveFolder%%FileName%.wav,10,1
	If FileExist(SaveFolder FileName ".vbs"){
		SB_SetText("Attempting to delete old script...")
		Gosub,DeleteVBS	
	}
	SB_SetText("Successfully converted your text to speech...")
	Sleep,1000
	Gui,Flash
	SB_SetText("Find your file @ " SaveFolder FileName ".wav...")
	GuiControl,,ConvertVar,Con&vert Text To Speech
Return
DeleteVBS:
	Try,FileDelete,%SaveFolder%%FileName%.vbs
	Catch {
		Debug(A_ThisLabel " Error","Could not delete:`n" SaveFolder FileName ".vbs",16)
	}
Return
ToolTipOff:
	If (A_TimeIdle > 3000)
		ToolTip
	SetTimer,%A_ThisLabel%,-1
Return
Help:
	Debug(_title " " A_ThisLabel,_help_msg)
Return
GuiClose:
	ExitApp