!define MyName "QuickExport"

SetCompressor /SOLID lzma ; Define the compression we will use for .exe
Name "${MyName}" ; The name of the installer
OutFile "${MyName}.exe" ; The name of the file to write
XPStyle on ; Make the UI Pretty
InstallColors /windows ; Make the UI Pretty
ShowInstDetails hide ; Shows the details of the installer, can be set to 'hide' or 'nevershow'
SetDateSave on ; Saves the last write date and time of the file
CRCCheck on ; Perfrom CRC Check before install

BrandingText "http://cdrpro.ru/"

LoadLanguageFile "${NSISDIR}\Contrib\Language files\English.nlf"
LoadLanguageFile "${NSISDIR}\Contrib\Language files\Dutch.nlf"
LoadLanguageFile "${NSISDIR}\Contrib\Language files\French.nlf"
LoadLanguageFile "${NSISDIR}\Contrib\Language files\German.nlf"
LoadLanguageFile "${NSISDIR}\Contrib\Language files\Italian.nlf"
LoadLanguageFile "${NSISDIR}\Contrib\Language files\Portuguese.nlf"
LoadLanguageFile "${NSISDIR}\Contrib\Language files\Spanish.nlf"
LoadLanguageFile "${NSISDIR}\Contrib\Language files\Russian.nlf"

RequestExecutionLevel admin ; Request application privilege for Windows Vista / Windows 7

LicenseData "license.txt"

page license
Page components ;components selection page
Page instfiles ;installation page where the sections are executed

var cdr17
var cdr17x64

var cdrts17
var cdrts17x64

!include Sections.nsh
!include LogicLib.nsh
!include WinMessages.nsh

; =======================================================================================

;Check if CorelDRAW is running
!macro CheckRunningCorelDRAW appVer appVerFriendly
	${DO}
		FindWindow $0 "CorelDRAW ${appVer}"
		${IF} $0 = 0
			${BREAK}
		${ENDIF}
		ShowWindow $0 ${SW_SHOW}
		BringtoFront
		MessageBox MB_OKCANCEL  "Close CorelDRAW ${appVerFriendly} and press OK" IDOK +2
		abort "Installation aborted"
	${LOOP}
!macroend

; =======================================================================================

!macro execSection appVer appVerFriendly
	detailprint "————————————————— CorelDRAW ${appVerFriendly} ———————————"
	!insertmacro CheckRunningCorelDRAW ${appVer} ${appVerFriendly}
	
	strcpy $0 "$cdr${appVer}\Programs\Addons\${MyName}\${MyName}.dll"
	SetOutPath "$cdr${appVer}\Programs\Addons\${MyName}\"
	File ${MyName}.dll
	
	strcpy $0 "$APPDATA\Corel\CorelDRAW Graphics Suite ${appVerFriendly}\Draw\GMS\${MyName}Installer.gms"
	SetOutPath "$APPDATA\Corel\CorelDRAW Graphics Suite ${appVerFriendly}\Draw\GMS\"
	File ${MyName}Installer.gms
!macroend

!macro execSectionx64 appVer appVerFriendly
	detailprint "—————————————— CorelDRAW ${appVerFriendly} (64-Bit) ———————————"
	!insertmacro CheckRunningCorelDRAW ${appVer} ${appVerFriendly}
	
	strcpy $0 "$cdr${appVer}x64\Programs64\Addons\${MyName}\${MyName}.dll"
	SetOutPath "$cdr${appVer}x64\Programs64\Addons\${MyName}\"
	File ${MyName}.dll
	
	strcpy $0 "$APPDATA\Corel\CorelDRAW Graphics Suite ${appVerFriendly}\Draw\GMS\${MyName}Installer.gms"
	SetOutPath "$APPDATA\Corel\CorelDRAW Graphics Suite ${appVerFriendly}\Draw\GMS\"
	File ${MyName}Installer.gms
!macroend

!macro execSectionTS appVer appVerFriendly
	detailprint "————————————————— CorelDRAW ${appVerFriendly} TS ——————————"
	!insertmacro CheckRunningCorelDRAW ${appVer} ${appVerFriendly}
	
	strcpy $0 "$cdrts${appVer}\Programs\Addons\${MyName}\${MyName}.dll"
	SetOutPath "$cdrts${appVer}\Programs\Addons\${MyName}\"
	File ${MyName}.dll
	
	strcpy $0 "$APPDATA\Corel\CorelDRAW Technical Suite ${appVerFriendly}\Draw\GMS\${MyName}Installer.gms"
	SetOutPath "$APPDATA\Corel\CorelDRAW Technical Suite ${appVerFriendly}\Draw\GMS\"
	File ${MyName}Installer.gms
!macroend

!macro execSectionTSx64 appVer appVerFriendly
	detailprint "—————————————— CorelDRAW ${appVerFriendly} TS (64-Bit) ———————————"
	!insertmacro CheckRunningCorelDRAW ${appVer} ${appVerFriendly}
	
	strcpy $0 "$cdrts${appVer}x64\Programs64\Addons\${MyName}\${MyName}.dll"
	SetOutPath "$cdrts${appVer}x64\Programs64\Addons\${MyName}\"
	File ${MyName}.dll
	
	strcpy $0 "$APPDATA\Corel\CorelDRAW Technical Suite ${appVerFriendly}\Draw\GMS\${MyName}Installer.gms"
	SetOutPath "$APPDATA\Corel\CorelDRAW Technical Suite ${appVerFriendly}\Draw\GMS\"
	File ${MyName}Installer.gms
!macroend

; =======================================================================================

;SectionGroup /e "!${MyName}"
	Section /o "" sec17
		${IF} $cdr17 != "CorelDRAW X7"
		${AndIf} $cdr17 != ""
			!insertmacro execSection 17 X7
		${ENDIF}
	SectionEnd
	
	Section /o "" sec17x64
		${IF} $cdr17x64 != "CorelDRAW X7 (64-Bit)"
		${AndIf} $cdr17x64 != ""
			!insertmacro execSectionx64 17 X7
		${ENDIF}
	SectionEnd
	
	Section /o "" sects17
		${IF} $cdrts17 != "CorelDRAW TS X7"
		${AndIf} $cdrts17 != ""
			!insertmacro execSectionTS 17 X7
		${ENDIF}
	SectionEnd
	
	Section /o "" sects17x64
		${IF} $cdrts17x64 != "CorelDRAW TS X7 (64-Bit)"
		${AndIf} $cdrts17x64 != ""
			!insertmacro execSectionTSx64 17 X7
		${ENDIF}
	SectionEnd
;SectionGroupEnd

; =======================================================================================

!macro regCheckCorel appVer
	ReadRegStr $0 HKLM "SOFTWARE\Corel\CorelDRAW\${appVer}.0" "ConfigDir"
	strCpy $0 $0 -6
	strCpy $1 "$0Programs\coreldrw.exe"
	strCpy $0 $0 -1
	strcpy $cdr${appVer} "$0"
	${IF} ${FileExists} $1
		!insertmacro SelectSection "${sec${appVer}}"
	${ENDIF}
!macroend

!macro regCheckCorel64 appVer
	SetRegView 64
	ReadRegStr $0 HKLM "SOFTWARE\Corel\CorelDRAW\${appVer}.0" "ConfigDir"
	strcpy $0 $0 -6
	strcpy $1 "$0Programs64\coreldrw.exe"
	strCpy $0 $0 -1
	strcpy $cdr${appVer}x64 "$0"
	${IF} ${FileExists} $1
		!insertmacro SelectSection "${sec${appVer}x64}"
	${ENDIF}
!macroend

!macro regCheckCorelTS appVer
	ReadRegStr $0 HKLM "SOFTWARE\Corel\Corel DESIGNER\${appVer}.0" "ConfigDir"
	strCpy $0 $0 -6
	strCpy $1 "$0Programs\coreldrw.exe"
	strCpy $0 $0 -1
	strcpy $cdrts${appVer} "$0"
	${IF} ${FileExists} $1
		!insertmacro SelectSection "${sects${appVer}}"
	${ENDIF}
!macroend

!macro regCheckCorelTS64 appVer
	SetRegView 64
	ReadRegStr $0 HKLM "SOFTWARE\Corel\Corel DESIGNER\${appVer}.0" "ConfigDir"
	strcpy $0 $0 -6
	strcpy $1 "$0Programs64\coreldrw.exe"
	strCpy $0 $0 -1
	strcpy $cdrts${appVer}x64 "$0"
	${IF} ${FileExists} $1
		!insertmacro SelectSection "${sects${appVer}x64}"
	${ENDIF}
!macroend

; =======================================================================================

Function .onInit
	;splash
	InitPluginsDir
	File /oname=$PLUGINSDIR\splash.bmp "cdrpro.bmp"
	advsplash::show 1000 600 400 0x00FF00 $PLUGINSDIR\splash
	Pop $0
	;splash
	
	!insertmacro regCheckCorel "17"
	!insertmacro regCheckCorel64 "17"
	
	!insertmacro regCheckCorelTS "17"
	!insertmacro regCheckCorelTS64 "17"
	
	;CorelDRAW GS
	${IF} ${SectionIsSelected} ${sec17}
		SectionSetText ${sec17} "CorelDRAW X7"
	${ENDIF}
	${IF} ${SectionIsSelected} ${sec17x64}
		SectionSetText ${sec17x64} "CorelDRAW X7 (64-Bit)"
	${ENDIF}
	
	;CorelDRAW TS
	${IF} ${SectionIsSelected} ${sects17}
		SectionSetText ${sects17} "CorelDRAW TS X7"
	${ENDIF}
	${IF} ${SectionIsSelected} ${sects17x64}
		SectionSetText ${sects17x64} "CorelDRAW TS X7 (64-Bit)"
	${ENDIF}
FunctionEnd
