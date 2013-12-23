/* =================================================================================
   MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
   Licenced under a Creative Commons 3.0 Attribution Licence
   --You may use and modify this code how you see fit as long as you give credit
   ================================================================================= */
/* This is a NSIS (Nullsoft Scriptable Install System) script, download and install
   NSIS to compile this into an installer. This script written using HM NIS Edit,
   a text editor for NSIS scripts */

;This installer does not run as Admin by default and the app can be installed locally,
; or portably. Huge thanks goes to Anders for an excellent sample of this approach:
; <stackoverflow.com/questions/13777988/use-nsis-to-create-both-normal-install-and-portable-install>

;--------------------------------------------------------------------------------------

!include LogicLib.nsh                                   ;If / Then logic
!include FileFunc.nsh                                   ;File system operations
!include WordFunc.nsh                                   ;String comparisons
!include MUI2.nsh                                       ;Modern interface

;Include the UAC plugin for handling user / admin rights. You won't need to install
; the plugin yourself, it's provided in a sub-folder thanks to its permissable licence
; <nsis.sourceforge.net/UAC_plug-in>
!addplugindir UAC\Ansi
!include UAC\UAC.nsh

;--------------------------------------------------------------------------------------
;We begin with a series of global constants which will help us avoid a lot of
; duplicate values that, when changed, we will want to be reflected throughout

;Windows new-line sequence
!define CRLF "$\r$\n"

!define PRODUCT_NAME "MaSS1VE"
!define EXE_NAME "${PRODUCT_NAME}.exe"

;Meta data
!define PRODUCT_DESCRIPTION "Create new adventures with Sonic on the Master System!"
!define PRODUCT_PUBLISHER "Camen Design"
!define PRODUCT_WEB_SITE "http://camendesign.com"

/* NOTE: To save having to manually update this script every time there's a new version
   of MaSS1VE, we will want to extract the version number from MaSS1VE and apply it to
   the installer. We cannot do that here as `VIProductVersion` is declarative, and is
   compiled before we get a chance to run any code to extract the version number.

   Therefore, our compiler batch script "Build Installer.bat", compiles MaSS1VE,
   extracts its version number and passes it to the NSIS compiler with this script.

   (Take note that this means that you will not be able to compile this script on its
    own, you should always run "Build Installer.bat" to do that.)

   The batch script defines `PRODUCT_VERSION_WIN` ("1.2.3.4" format, required for
   the `VIProductVersion` setting) and `PRODUCT_VERSION_VB6` ("1.2,4" format)
*/
!ifndef PRODUCT_VERSION_WIN
        !warning "Version number not defined, please run 'Build Installer.bat' \
                  instead of compiling the script on its own"
        ;We'll let the script continue, for testing purposes only
        !define PRODUCT_VERSION_WIN "0.0.0.0"
        !define PRODUCT_VERSION_VB6 "0.0,0"
!endif

;Name the installer with the version number from MaSS1VE.exe
!define INSTALLER_NAME "Install_${PRODUCT_NAME}_v${PRODUCT_VERSION_VB6}"
;The uninstaller is more guessable
!define UNINSTALLER_EXE_NAME "Uninstall.exe"

;define default installation paths for Local and Portable mode
!define INSTDIR_LOCAL_DEFAULT "$APPDATA\${PRODUCT_NAME}"
!define INSTDIR_PORTABLE_DEFAULT "$DESKTOP\${PRODUCT_NAME}"

;We'll put a single shortcut directly into the start menu (no sub-folder)
!define START_MENU_SHORTCUT "$SMPROGRAMS\${PRODUCT_NAME}.lnk"

;Registry key for the uninstaller info ("Add/Remove Programs")
!define REG_UNINSTALL "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
;Allow running the app from the run box using its name
!define REG_APPPATH "Software\Microsoft\Windows\CurrentVersion\App Paths\${EXE_NAME}"

;======================================================================================
;Installer configuration:

;Give names to things...
Name "${PRODUCT_NAME}"
OutFile "${INSTALLER_NAME}.exe"

;Don't run this installer as Admin. We want users to be able to install MaSS1VE even
; if their account is a limited one and especially be able to update it without Admin
; rights. Thanks goes to Lorenz Cuno for general details to this approach:
; <klopfenstein.net/lorenz.aspx/simple-nsis-installer-with-user-execution-level>
RequestExecutionLevel user

;We default installation to the portable directory because it can be chosen, but the
; Local mode always installs to $LOCALAPPDATA (this gets set on the install files page)
; This ensures that if the user selects Portable mode, changes the directory and then
; goes back, the path is remembered, but also not used if they select Local mode
InstallDir "${INSTDIR_PORTABLE_DEFAULT}"

;Since we're going to allow portable installations, we want users to be able to
; install to the root of a drive (e.g. "F:\")
AllowRootDirInstall true

;Compression settings
SetCompressor /SOLID lzma                               ;LZMA compression works best
SetDatablockOptimize on                                 ;Intelligent file ordering
CRCCheck force                                          ;Check for corrupt download

;General user interface settings
XPStyle on                                              ;Native theming
BrandingText /TRIMRIGHT " "                             ;Hide "NSIS Installer" text
ShowInstDetails nevershow
ShowUnInstDetails nevershow

;Define installer and uninstaller icons. We use the app's icon so that it's easily
; identifiable in the Add/Remove Programs list
!define MUI_ICON "InstallIcon.ico"
!define MUI_UNICON "InstallIcon.ico"

;--------------------------------------------------------------------------------------
;Set the version info on the installer .exe

VIProductVersion "${PRODUCT_VERSION_WIN}"
VIAddVersionKey "ProductName" "${PRODUCT_NAME} Installer"
VIAddVersionKey "Comments" "${PRODUCT_NAME} Installer"
VIAddVersionKey "CompanyName" "${PRODUCT_PUBLISHER}"
VIAddVersionKey "LegalTrademarks" ""
VIAddVersionKey "LegalCopyright" "© Kroc Camen of ${PRODUCT_PUBLISHER} 2013"
VIAddVersionKey "FileDescription" "${PRODUCT_NAME} Installer"
VIAddVersionKey "FileVersion" "${PRODUCT_VERSION_VB6}"

;======================================================================================

var /GLOBAL IsAdmin
var /GLOBAL PortableMode
var /GLOBAL UpdateMode

;Define installation pages: -----------------------------------------------------------

Caption "Install ${PRODUCT_NAME}"

!define MUI_WELCOMEFINISHPAGE_BITMAP "Welcome.bmp"
!define MUI_WELCOMEPAGE_TITLE "${PRODUCT_NAME} v${PRODUCT_VERSION_VB6}"
!define MUI_WELCOMEPAGE_TEXT "\
        Create your own levels for Sonic the Hedgehog™ for the Master System™ with \
        MaSS1VE, the first and most comprehensive editing tool for modifying the game.\
        ${CRLF}${CRLF}\
        This program can install MaSS1VE either locally, or in any folder as a \
        portable application (such as on a USB-stick) so that you can carry it \
        around with you!\
        ${CRLF}${CRLF}\
        Sonic the Hedgehog™, Master System™ © Sega Enterprises. \
        MaSS1VE is © Kroc Camen, with a Creative Commons Attribution 3.0 licence."
!insertmacro MUI_PAGE_WELCOME

;Define our custom portable-mode choice page
Page Custom PortableModePageCreate PortableModePageLeave

;Allow skipping of the Directory page (if portable mode is on),
; with thanks to <forums.winamp.com/showpost.php?p=2358237&postcount=4>
!define MUI_PAGE_CUSTOMFUNCTION_PRE DirectoryPagePre
;Check the chosen directory before continuing, we won't be able to install to an
; Admin-only direcoty (such as %PROGRAMFILES%) and need to warn the user
!define MUI_PAGE_CUSTOMFUNCTION_LEAVE DirectoryPageLeave
!insertmacro MUI_PAGE_DIRECTORY

;Do the actual installation
!insertmacro MUI_PAGE_INSTFILES

!define MUI_FINISHPAGE_TITLE "Installation Complete"
!define MUI_FINISHPAGE_TEXT "\
        Please note that you will need a Master System emulator (such as 'Kega Fusion') \
        to play the $\".sms$\" files MaSS1VE produces."
!define MUI_FINISHPAGE_LINK "Download 'Kega Fusion' Emulator"
!define MUI_FINISHPAGE_LINK_LOCATION "http://kega.eidolons-inn.net/"
/* Remember that you shouldn't launch the executable from admin-privliges
   <mdb-blog.blogspot.co.uk/2013/01/nsis-lunch-program-as-user-from-uac.html>
   MaSS1VE has a bug where OLE drag-and-drop won't work when running as Administrator,
   but regardless we don't want it saving data to %PROGRAMFILES% and then being unable
   to change it the next time the program is run! */
!define MUI_FINISHPAGE_RUN
!define MUI_FINISHPAGE_RUN_FUNCTION FinishPageRun
;This saves some installer space since we won't request a reboot
!define MUI_FINISHPAGE_NOREBOOTSUPPORT
;Skip the finish page if updating (`/UPDATE`)
!define MUI_PAGE_CUSTOMFUNCTION_PRE FinishPagePre
!insertmacro MUI_PAGE_FINISH

;Uninstallation pages:
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE English


;=== INSTALLATION =====================================================================

;--------------------------------------------------------------------------------------
Section Install
        ;If local installation has been selected,
        ; choose the deafult installation path
        ${If} $PortableMode = 0
                StrCpy $INSTDIR ${INSTDIR_LOCAL_DEFAULT}
        ${EndIf}

        SetOverwrite on
        SetOutPath "$INSTDIR"

        ;Package the app
        File "..\RELEASE\${EXE_NAME}"
        
        ;Remove the old update data that might confuse things when upgrading version
        Delete "$INSTDIR\App Data\Update.ini"
        Delete "$INSTDIR\App Data\Update.html"
        Delete "$INSTDIR\App Data\Update.exe"
SectionEnd

;--------------------------------------------------------------------------------------
Section Local
        ;Don't create shortcut / modify registry if portable mode selected
        ${If} $PortableMode = 0

        ;Is this Windows 8.1?
        ;<forums.winamp.com/showthread.php?t=365416>
        ReadRegStr $R0 HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion" CurrentVersion
        ${VersionCompare} "6.2" "$R0" $R1
        ${If} $R1 = 2
                File "..\RELEASE\${PRODUCT_NAME}.VisualElementsManifest.xml"
                File "..\RELEASE\Resources.pri"
                File "..\RELEASE\Resources.scale-140.pri"
                File "..\RELEASE\Resources.scale-180.pri"

                ;Install the images for the custom Windows 8.1 Start screen tile
                SetOutPath "$INSTDIR\VisualElements"
                File "..\RELEASE\VisualElements\70x70Logo.scale-80.png"
                File "..\RELEASE\VisualElements\70x70Logo.scale-100.png"
                File "..\RELEASE\VisualElements\70x70Logo.scale-140.png"
                File "..\RELEASE\VisualElements\70x70Logo.scale-180.png"
                File "..\RELEASE\VisualElements\150x150Logo.scale-80.png"
                File "..\RELEASE\VisualElements\150x150Logo.scale-100.png"
                File "..\RELEASE\VisualElements\150x150Logo.scale-140.png"
                File "..\RELEASE\VisualElements\150x150Logo.scale-180.png"
        ${EndIf}
        
        ;We're going to install a single shortcut directly into the Start Menu,
        ; no sub-folder -- that's so passe
        CreateShortCut "${START_MENU_SHORTCUT}" "$INSTDIR\${EXE_NAME}" "" "" "" \
                       SW_SHOWNORMAL "" "${PRODUCT_DESCRIPTION}"

        ;Place Uninstall.exe
        WriteUninstaller "$INSTDIR\${UNINSTALLER_EXE_NAME}"

        ;Write the uninstaller info to the registry
        WriteRegStr   SHCTX "${REG_UNINSTALL}" "DisplayName" "${PRODUCT_NAME}"
        WriteRegStr   SHCTX "${REG_UNINSTALL}" "UninstallString" "$\"$INSTDIR\${UNINSTALLER_EXE_NAME}$\""
        WriteRegStr   SHCTX "${REG_UNINSTALL}" "DisplayIcon" "$\"$INSTDIR\${EXE_NAME}$\""
        WriteRegStr   SHCTX "${REG_UNINSTALL}" "DisplayVersion" "${PRODUCT_VERSION_VB6}"
        WriteRegStr   SHCTX "${REG_UNINSTALL}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
        WriteRegStr   SHCTX "${REG_UNINSTALL}" "Publisher" "${PRODUCT_PUBLISHER}"
        WriteRegDWORD SHCTX "${REG_UNINSTALL}" "NoModify" 1
        WriteRegDWORD SHCTX "${REG_UNINSTALL}" "NoRepair" 1

        ;Allow running the app from the run box (WIN+R)
        WriteRegStr   SHCTX "${REG_APPPATH}" "" "$INSTDIR\${EXE_NAME}"

        ;Add the program size to the uninstall info
        ; (this measures the size of the install directory)
        ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
        IntFmt $0 "0x%08X" $0
        WriteRegDWORD SHCTX "${REG_UNINSTALL}" "EstimatedSize" "$0"
        
        ;Refresh Explorer to reflect shortcut / icon changes
        ${RefreshShellIcons}
        
        ${EndIf}
SectionEnd

;=== UNINSTALLATION ===================================================================

Section un.Install
        ;Remove the Start shortcut
        Delete "${START_MENU_SHORTCUT}"

        ;NOTE: We should only delete our own files and not delete the whole folder
        ;      in case the user installed into a location with other files
        Delete "$INSTDIR\${EXE_NAME}"
        Delete "$INSTDIR\${UNINSTALLER_EXE_NAME}"
        ;Delete the Windows 8.1 tile icon (when installed locally)
        Delete "$INSTDIR\${PRODUCT_NAME}.VisualElementsManifest.xml"
        Delete "$INSTDIR\Resources.pri"
        Delete "$INSTDIR\Resources.scale-140.pri"
        Delete "$INSTDIR\Resources.scale-180.pri"
        RMDir /r "$INSTDIR\VisualElements"
        ;Remove the install directory only if it's empty
        RMDir  "$INSTDIR"
        
        ;Clean up the registry keys
        DeleteRegKey SHCTX "${REG_APPPATH}"
        DeleteRegKey SHCTX "${REG_UNINSTALL}"
SectionEnd

;=== FUNCTIONS ========================================================================

;Initialise the installer:
;--------------------------------------------------------------------------------------
Function .onInit
        ;Get the command line parameters
        ; <nsis.sourceforge.net/Docs/AppendixE.html#E.1.11>
        ${GetParameters} $9

        ;Test for the "/?" help switch, and display parameter info
        ClearErrors
        ${GetOptions} $9 "/?" $8
        ${IfNot} ${Errors}
                MessageBox MB_ICONINFORMATION|MB_SETFOREGROUND "\
                        /PORTABLE : Install without shortuct / uninstaller${CRLF}\
                        /UPDATE : Update an existing installation, implies `/PORTABLE` \
                        and `/D` should be used to specify the directory${CRLF}\
                        /D=%directory% : Specify destination directory (must be last \
                        and not contain any quotes, even if there are spaces in the \
                        path${CRLF}"
                Quit
        ${EndIf}
	
        ;Is the installer being "run as administrator"? We don't support running as
        ; admin just yet, it would prevent us from automatically updating
;;NOT WORKING IN WIN7??
;        !insertmacro UAC_IsAdmin
;        StrCpy $IsAdmin $0
;        ${If} $IsAdmin = 1
;                MessageBox MB_ICONSTOP|MB_SETFOREGROUND "\
;                        ${PRODUCT_NAME} does not support being installed as \
;                        Administrator (it prevents updates installing correctly). \
;                        Please re-run the installer normally."
;                Abort
;        ${EndIf}

        ClearErrors
        ${GetOptions} $9 "/PORTABLE" $8
        ${IfNot} ${Errors}
            StrCpy $PortableMode 1
        ${Else}
            StrCpy $PortableMode 0
        ${EndIf}
        
        ClearErrors
        ${GetOptions} $9 "/UPDATE" $8
        ${IfNot} ${Errors}
            StrCpy $PortableMode 1
            StrCpy $UpdateMode 1
        ${Else}
            StrCpy $PortableMode 0
            StrCpy $UpdateMode 0
        ${EndIf}
FunctionEnd

;--------------------------------------------------------------------------------------
;The Update Mode (use switch `/UPDATE`) skips all but the install files page
Function FinishPagePre
        ${If} $UpdateMode = 1
                Call FinishPageRun
                Abort
        ${EndIf}
FunctionEnd

;--------------------------------------------------------------------------------------
Function PortableModePageCreate
        ;Skip the local / portable mode select page if running an update
        ${If} $UpdateMode = 1
                Abort
        ${EndIf}

        !insertmacro MUI_HEADER_TEXT "Install Mode" ""

        nsDialogs::Create 1018
        Pop $0
        ${NSD_CreateLabel} 0 10u 100% 24u "How would you like to install MaSS1VE?"
        Pop $0
        ${NSD_CreateRadioButton} 30u 40u -30u 8u "Local, with shortcut and uninstaller"
        Pop $1
        ${NSD_CreateRadioButton} 30u 60u -30u 8u "Portable (such as on a USB drive)"
        Pop $2
        ${If} $PortableMode = 0
                SendMessage $1 ${BM_SETCHECK} ${BST_CHECKED} 0
        ${Else}
                SendMessage $2 ${BM_SETCHECK} ${BST_CHECKED} 0
        ${EndIf}
        nsDialogs::Show
FunctionEnd

;Set the portable mode based on the user selection:
;--------------------------------------------------------------------------------------
Function PortableModePageLeave
        ${NSD_GetState} $1 $0
        ${If} $0 <> ${BST_UNCHECKED}
                StrCpy $PortableMode 0
        ${Else}
                StrCpy $PortableMode 1
        ${EndIf}
FunctionEnd

;Allow skipping the Directory page when local install is selected:
;--------------------------------------------------------------------------------------
Function DirectoryPagePre
         ;If local mode selected, skip the Directory page, installation happens
         ; automatically to $LOCALAPPDATA
        ${If} $PortableMode = 0
                Abort
        ${EndIf}
        
        ;If update mode enabled (`/UPDATE`), skip this page, `/D` will specify the
        ; installation directory (If `/D` is not used, will default to $LOCALAPPDATA)
        ${If} $UpdateMode = 1
                Abort
        ${EndIf}
FunctionEnd

;Check the chosen directory is writable:
;--------------------------------------------------------------------------------------
Function DirectoryPageLeave
        ;Attempt to create a temporary file in the installation directory,
        ; (creating the installation directory if it doesn't exist) --
        ; if this fails the user doesn't have write-permissions to this folder
        ; and we need to warn the user
        ClearErrors
        CreateDirectory "$INSTDIR"
        FileOpen $R0 "$INSTDIR\tmp.dat" w
        FileClose $R0
        ;Do not proceed if the directory can't be written to
        ${If} ${Errors}
                MessageBox MB_ICONSTOP|MB_SETFOREGROUND "\
                        You don't have write permission to install ${PRODUCT_NAME} to \
                        '$INSTDIR', please select a location that is accessible, e.g. \
                        Desktop, Documents.${CRLF}\
                        ${CRLF}\
                        ${PRODUCT_NAME} does not support installation as the \
                        Administrator account yet."
                Abort
        ${EndIf}
        ;Clean up. We delete the directory (if empty) in case the user changes their
        ; mind and selects a different location
        Delete "$INSTDIR\tmp.dat"
        RmDir "$INSTDIR"
FunctionEnd

;--------------------------------------------------------------------------------------
Function FinishPageRun
        ;Ensure the program is not launched as Administrator (otherwise the program
        ; will write data somewhere where it won't be able to modify it later!)
        !insertmacro UAC_AsUser_ExecShell "" "$INSTDIR\${EXE_NAME}" "" "" ""
FunctionEnd
