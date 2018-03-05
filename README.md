# TTS-WavGen
TTS WavGen is a program that uses AutoHotkey to generate a VBScript that converts text to speech in Windows wav format.

## Current Release
[TTS WavGen 32 Bit](https://github.com/Lateralus138/TTS-WavGen/releases/download/1.12.22.17/TTS.WavGen.32bit.exe)
[TTS WavGen 64 Bit](https://github.com/Lateralus138/TTS-WavGen/releases/download/1.12.22.17/TTS.WavGen.64bit.exe)
[Release Page - Source Files](https://github.com/Lateralus138/TTS-WavGen/releases/latest)

## Example Code - method I use to convert text to speech.
```
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
```
## Motivation

I use VBScript to convert text to speech very often and so I wrote this to help automate the process.

## Installation

Portable program (Plans for installer and portable option).


## Test
I have tested on Windows 10 64 Bit

## Contributors

Ian Pride @ faithnomoread@yahoo.com - [Lateralus138] @ New Pride Services 

## License

	This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

	License provided in gpl.txt

