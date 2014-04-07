!define PRODUCT_NAME		"Open Twebst"
!define PRODUCT_VERSION		"1.4.0.0"
!define SETUP_NAME			"OpenTwebstSetup_1.4.exe"
!define PRODUCT_PUBLISHER	"Codecentrix Software"
!define PRODUCT_WEB_SITE	"http://www.codecentrix.com/"


!include    "WMI.nsh"
!include	"MUI.nsh"
Name		"${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile		${SETUP_NAME}
InstallDir	"$PROGRAMFILES\Codecentrix\OpenTwebst"

; Request application privileges.
RequestExecutionLevel admin

;Get installation folder from registry if available
InstallDirRegKey HKLM "Software\CodeCentrix\OpenTwebst" "InstallDir"

;--------------------------------
;Variables
  Var MUI_TEMP
  Var STARTMENU_FOLDER
  Var MYDOC

;--------------------------------
;Interface Settings
	!define MUI_ABORTWARNING
	!define MUI_ICON	".\twebst.ico"
	!define MUI_UNICON	".\twebst.ico"

;--------------------------------
;Pages
	!define MUI_WELCOMEPAGE_TITLE "${PRODUCT_NAME} ${PRODUCT_VERSION}"
	!define MUI_WELCOMEPAGE_TEXT  "Twebst Recorder requires .NET Framework 2.0 while Twebst Library does not; macros can run without .NET Framework but you cannot use the recorder.\r\n\r\nPlease uninstall any previous Open Twebst version first!\r\n\r\n\r\nClick Next to continue."
	!insertmacro MUI_PAGE_WELCOME
	!insertmacro MUI_PAGE_LICENSE "License.rtf"
	!insertmacro MUI_PAGE_COMPONENTS
	!insertmacro MUI_PAGE_DIRECTORY

;Start Menu Folder Page Configuration
	!define MUI_STARTMENUPAGE_REGISTRY_ROOT "HKCU" 
	!define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\CodeCentrix\OpenTwebst" 
	!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "Start Menu Folder"

	!insertmacro MUI_PAGE_STARTMENU Application $STARTMENU_FOLDER
	!insertmacro MUI_PAGE_INSTFILES
	!insertmacro MUI_UNPAGE_CONFIRM
	!insertmacro MUI_UNPAGE_INSTFILES

;--------------------------------
;Languages
	!insertmacro MUI_LANGUAGE "English"

;--------------------------------
;Installer Sections
Section ""
	SetShellVarContext all
	SetOutPath "$INSTDIR"

	;Create uninstaller
	WriteUninstaller "$INSTDIR\Uninstall.exe"

	;Store installation folder
	WriteRegStr HKLM "Software\CodeCentrix\Twebst" "InstallDir" $INSTDIR
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenTwebst" "DisplayName" "Open Twebst"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenTwebst" "Publisher" "Codecentrix Software"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenTwebst" "DisplayVersion" "${PRODUCT_VERSION}"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenTwebst" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenTwebst" "UninstallString" "$INSTDIR\Uninstall.exe"

	!insertmacro MUI_STARTMENU_WRITE_BEGIN Application
	CreateDirectory "$SMPROGRAMS\$STARTMENU_FOLDER"
	!insertmacro MUI_STARTMENU_WRITE_END
SectionEnd


Section "Twebst Library" SecTwebstLib
	SetShellVarContext all
	SectionIn RO
	SetOutPath "$INSTDIR\Bin"

	File "..\outdir\release\OTWBSTPlugin.dll"
	File "..\outdir\release\OpenTwebstLib.dll"
	File "..\outdir\release\Dbgserv.dll"
	File "..\outdir\release\OTWBSTInjector.dll"
	
	File "..\outdir\release\OTWBSTPlugin_x64.dll"
	File "..\outdir\release\OpenTwebstLib_x64.dll"
	File "..\outdir\release\Dbgserv_x64.dll"
	File "..\outdir\release\OTWBSTInjector_x64.dll"
	
	File "..\outdir\release\OTwbstXbit_x86.exe"
	File "..\outdir\release\OTwbstXbit_x64.exe"

	File "..\License.txt"

	RegDLL "$INSTDIR\Bin\OTWBSTPlugin.dll"
	RegDLL "$INSTDIR\Bin\OpenTwebstLib.dll"

	ExecWait 'regsvr32.exe /s "$INSTDIR\Bin\OTWBSTPlugin_x64.dll"'
	ExecWait 'regsvr32.exe /s "$INSTDIR\Bin\OpenTwebstLib_x64.dll"'
SectionEnd


Section "Twebst Web Recorder" SecTwebstRecorder
	SetShellVarContext all
	SetOutPath "$INSTDIR\Bin"

	File "..\outdir\release\Microsoft.mshtml.dll"
	File "..\outdir\release\Interop.SHDocVw.dll"
	File "..\outdir\release\Interop.OpenTwebstLib.dll"
	File "..\outdir\release\OpenTwebst.exe"
	
	CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\Open Twebst.lnk" "$INSTDIR\Bin\OpenTwebst.exe"
	CreateShortCut "$DESKTOP\Open Twebst.lnk" "$INSTDIR\Bin\OpenTwebst.exe"
SectionEnd


Section "Sample Files" SecSamples
	SetShellVarContext current
	SetOutPath "$DOCUMENTS\OpenTwebstSamples\JScript"
	File "..\Samples\JScript\*.js"

	SetOutPath "$DOCUMENTS\OpenTwebstSamples\VBScript"
	File "..\Samples\VBScript\*.vbs"

	SetOutPath "$DOCUMENTS\OpenTwebstSamples\C#"
	File "..\Samples\C#\*.*"

	SetOutPath "$DOCUMENTS\OpenTwebstSamples\VB.Net"
	File "..\Samples\VB.Net\*.*"

	SetOutPath "$DOCUMENTS\OpenTwebstSamples\C++"
	File "..\Samples\C++\*.*"
	
	SetOutPath "$DOCUMENTS\OpenTwebstSamples\WebBrowserCSharp"
	File "..\Samples\WebBrowserCSharp\*.*"

	SetOutPath "$INSTDIR\Bin"
	File ".\tbwstsamples.ico"

	StrCpy $MYDOC "$DOCUMENTS"
	SetShellVarContext all
	CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\Samples.lnk" "$MYDOC\OpenTwebstSamples\" "" "$INSTDIR\Bin\tbwstsamples.ico"
SectionEnd


Section "Help" SecHelp
    SetShellVarContext all
	SetOutPath "$INSTDIR\Bin"
	File "..\Help\OpenTwebst.chm"
	File ".\twbstonlinedoc.ico"

	CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\Help.lnk" "$INSTDIR\Bin\OpenTwebst.chm"
	CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\Online Documentation.lnk" "http://doc.codecentrix.com/" "" "$INSTDIR\Bin\twbstonlinedoc.ico"
SectionEnd


;--------------------------------
;Descriptions
	LangString DESC_SecTwebstLib		${LANG_ENGLISH} "Open Twebst Library"
	LangString DESC_SecTwebstRecorder	${LANG_ENGLISH} "Open Twebst Web Recorder (requires MS.Net Framework 2.0)"
	LangString DESC_SecHelp				${LANG_ENGLISH} "Help File in CHM format"  
	LangString DESC_SecSamples			${LANG_ENGLISH} "Samples: JScript, VBScript, VB.Net, C#, C++"

;Assign language strings to sections
	!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
		!insertmacro MUI_DESCRIPTION_TEXT ${SecTwebstLib}		$(DESC_SecTwebstLib)
		!insertmacro MUI_DESCRIPTION_TEXT ${SecTwebstRecorder}	$(DESC_SecTwebstRecorder)
		!insertmacro MUI_DESCRIPTION_TEXT ${SecHelp}			$(DESC_SecHelp)	
		!insertmacro MUI_DESCRIPTION_TEXT ${SecSamples}			$(DESC_SecSamples)
	!insertmacro MUI_FUNCTION_DESCRIPTION_END

;--------------------------------
;Uninstaller Section
Section "Uninstall"
	SetShellVarContext all
	
	; http://nsis.sourceforge.net/WMI_header
	${WMIKillProcess} OTwbstXbit_x86.exe
	${WMIKillProcess} OTwbstXbit_x64.exe

	UnRegDLL "$INSTDIR\Bin\OTWBSTPlugin.dll"
	UnRegDLL "$INSTDIR\Bin\OpenTwebstLib.dll"
	
	ExecWait 'regsvr32.exe /s /u "$INSTDIR\Bin\OTWBSTPlugin_x64.dll"'
	ExecWait 'regsvr32.exe /s /u "$INSTDIR\Bin\OpenTwebstLib_x64.dll"'

	;Delete binaries and some icons
	Delete "$INSTDIR\Bin\*.dll"
	Delete "$INSTDIR\Bin\*.exe"
	Delete "$INSTDIR\Bin\*.ico"
	Delete "$INSTDIR\Bin\*.txt"
	Delete "$INSTDIR\Uninstall.exe"

	;Delete help and samples
	Delete "$INSTDIR\Bin\OpenTwebst.chm"
	
	SetShellVarContext current
	Delete "$DOCUMENTS\OpenTwebstSamples\JScript\*.js"
	Delete "$DOCUMENTS\OpenTwebstSamples\VBScript\*.vbs"
	Delete "$DOCUMENTS\OpenTwebstSamples\VB.Net\*.*"
	Delete "$DOCUMENTS\OpenTwebstSamples\C#\*.*"
	Delete "$DOCUMENTS\OpenTwebstSamples\C++\*.*"
	Delete "$DOCUMENTS\OpenTwebstSamples\WebBrowserCSharp\*.*"
	SetShellVarContext all

	;Delete folders
	RMDir "$INSTDIR\Bin"
	
	SetShellVarContext current
	RMDir "$DOCUMENTS\OpenTwebstSamples\C++"
	RMDir "$DOCUMENTS\OpenTwebstSamples\C#"
	RMDir "$DOCUMENTS\OpenTwebstSamples\VB.Net"
	RMDir "$DOCUMENTS\OpenTwebstSamples\VB6"
	RMDir "$DOCUMENTS\OpenTwebstSamples\JScript"
	RMDir "$DOCUMENTS\OpenTwebstSamples\VBScript"
	RMDir "$DOCUMENTS\OpenTwebstSamples"
	SetShellVarContext all

	RMDir "$INSTDIR"
	RMDir "$INSTDIR\.."

	;Delete Start/Programs shortcuts
	!insertmacro MUI_STARTMENU_GETFOLDER Application $MUI_TEMP
	Delete "$SMPROGRAMS\$MUI_TEMP\Samples.lnk"
	Delete "$SMPROGRAMS\$MUI_TEMP\Help.lnk"
	Delete "$SMPROGRAMS\$MUI_TEMP\Online Documentation.lnk"
	Delete "$SMPROGRAMS\$MUI_TEMP\Open Twebst.lnk"
	Delete "$DESKTOP\Open Twebst.lnk"

	;Delete empty start menu parent diretories
	StrCpy $MUI_TEMP "$SMPROGRAMS\$MUI_TEMP"

	startMenuDeleteLoop:
	ClearErrors
	RMDir $MUI_TEMP
	GetFullPathName $MUI_TEMP "$MUI_TEMP\.."

	IfErrors startMenuDeleteLoopDone

	StrCmp $MUI_TEMP $SMPROGRAMS startMenuDeleteLoopDone startMenuDeleteLoop
	startMenuDeleteLoopDone:

	DeleteRegKey /ifempty HKCU "Software\Modern UI Test"
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenTwebst"
	DeleteRegKey HKCU "Software\CodeCentrix\OpenTwebst"
	DeleteRegKey HKLM "Software\CodeCentrix\OpenTwebst"
	
	; To refresh desktop after short-cuts are deleted.
	System::Call 'Shell32::SHChangeNotify(i 0x8000000, i 0, i 0, i 0)'	
SectionEnd
