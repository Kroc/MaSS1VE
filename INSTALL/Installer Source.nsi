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

;--- Includes -------------------------------------------------------------------------

/* NOTE: To save having to manually update this script every time there's a new version
   of MaSS1VE, we will want to extract the version number from MaSS1VE and apply it to
   the installer. We cannot do that here as `VIProductVersion` is declarative, and is
   compiled before we get a chance to run any code to extract the version number.

   Therefore, our compiler batch script "Build Installer.bat", compiles MaSS1VE,
   extracts its version number and passes it to the NSIS compiler with this script.

   (Take note that this means that you will not be able to compile this script on its
    own, you should always run "Build Installer.bat" to do that.)

   The batch script defines `PRODUCT_VERSION_WIN` ("1.2.3.4" format, required for
   the `VIProductVersion` setting) and `PRODUCT_VERSION_VB6` ("1.2.4" format)
*/
!ifndef PRODUCT_VERSION_WIN
        !warning "Version number not defined, please run 'Build Installer.bat' \
                  instead of compiling the script on its own"
        ;We'll let the script continue, for testing purposes only
        !define PRODUCT_VERSION_WIN "0.0.0.0"
        !define PRODUCT_VERSION_VB6 "0.0.0"
!endif

!include LogicLib.nsh                                  ;If / Then logic
!include FileFunc.nsh                                  ;File system operations
!include MUI2.nsh                                      ;Modern interface

;--------------------------------------------------------------------------------------
;We begin with a series of global constants which will help us avoid a lot of
; duplicate values that, when changed, we will want to be reflected throughout

!define PRODUCT_NAME "MaSS1VE"
!define PRODUCT_DESCRIPTION "Create new adventures with Sonic on the Master System!"
!define PRODUCT_PUBLISHER "Camen Design"
!define PRODUCT_WEB_SITE "http://camendesign.com"
!define EXE_NAME "${PRODUCT_NAME}.exe"

;Name the installer with the version number from MaSS1VE.exe
!define INSTALLER_NAME "Install ${PRODUCT_NAME} v${PRODUCT_VERSION_VB6}"
;The uninstaller is more guessable
!define UNINSTALLER_EXE_NAME "Uninstall.exe"

;define default installation paths for Local and Portable mode
!define INSTDIR_LOCAL_DEFAULT "$LOCALAPPDATA\${PRODUCT_NAME}"
!define INSTDIR_PORTABLE_DEFAULT "$DESKTOP\${PRODUCT_NAME}"

;We'll put a single shortcut directly into the start menu (no sub-folder)
!define START_MENU_SHORTCUT "$STARTMENU\Programs\${PRODUCT_NAME}.lnk"

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
var /GLOBAL PortableMode

;Define installation pages: -----------------------------------------------------------

!insertmacro MUI_PAGE_WELCOME

;Define our custom portable-mode choice page
Page Custom PortableModePageCreate PortableModePageLeave

;Allow skipping of the Directory page (if portable mode is on),
; with thanks to <forums.winamp.com/showpost.php?p=2358237&postcount=4>
!define MUI_PAGE_CUSTOMFUNCTION_PRE DirectoryPagePre
!insertmacro MUI_PAGE_DIRECTORY

;Do the actual installation
!insertmacro MUI_PAGE_INSTFILES

/* Remember that you shouldn't launch the executable from admin-privliges
   <mdb-blog.blogspot.co.uk/2013/01/nsis-lunch-program-as-user-from-uac.html>
   We will have to fix the Admin OLE drag-and-drop bug in MaSS1VE to be safe in case
   people run the installer as Admin out of need / habit */
!define MUI_FINISHPAGE_RUN "$INSTDIR\${EXE_NAME}"
!insertmacro MUI_PAGE_FINISH

;Uninstallation pages:
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE English


;=== INSTALLATION =====================================================================

Section Install
        ${If} $PortableMode = 0
               StrCpy $INSTDIR ${INSTDIR_LOCAL_DEFAULT}
        ${EndIf}

        SetOutPath "$INSTDIR"
        SetOverwrite ifnewer

        File "..\BUILD\${EXE_NAME}"
SectionEnd

;--------------------------------------------------------------------------------------
Section Local
        ;We're going to install a single shortcut directly into the Start Menu,
        ; no sub-folder -- that's so passe
        CreateShortCut "${START_MENU_SHORTCUT}" "$INSTDIR\${EXE_NAME}" "" "" "" SW_SHOWNORMAL "" "${PRODUCT_DESCRIPTION}"

        ;Place Uninstall.exe
        WriteUninstaller "$INSTDIR\${UNINSTALLER_EXE_NAME}"
        
        ;Allow running the app from the run box (WIN+R)
        WriteRegStr HKCU "${REG_APPPATH}" "" "$INSTDIR\${EXE_NAME}"
        
        ;Write the uninstaller info to the registry
        WriteRegStr HKCU "${REG_UNINSTALL}" "DisplayName" "${PRODUCT_NAME}"
        WriteRegStr HKCU "${REG_UNINSTALL}" "UninstallString" "$\"$INSTDIR\${UNINSTALLER_EXE_NAME}$\""
        WriteRegStr HKCU "${REG_UNINSTALL}" "DisplayIcon" "$INSTDIR\${EXE_NAME}"
        WriteRegStr HKCU "${REG_UNINSTALL}" "DisplayVersion" "${PRODUCT_VERSION_VB6}"
        WriteRegStr HKCU "${REG_UNINSTALL}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
        WriteRegStr HKCU "${REG_UNINSTALL}" "Publisher" "${PRODUCT_PUBLISHER}"
        WriteRegDWORD HKCU "${REG_UNINSTALL}" "NoModify" 1
        WriteRegDWORD HKCU "${REG_UNINSTALL}" "NoRepair" 1
        
        ;Add the program size to the uninstall info
        ; (this measures the size of the install directory)
        ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
        IntFmt $0 "0x%08X" $0
        WriteRegDWORD HKCU "${REG_UNINSTALL}" "EstimatedSize" "$0"
SectionEnd

;=== UNINSTALLATION ===================================================================

Section un.Install
        ;Remove the Start shortcut
        Delete "${START_MENU_SHORTCUT}"

        ;NOTE: We should only delete our own files and not delete the whole folder
        ;      in case the user installed into a location with other files
        Delete "$INSTDIR\${EXE_NAME}"
        Delete "$INSTDIR\${UNINSTALLER_EXE_NAME}"
        ;Remove the install directory only if it's empty
        RMDir "$INSTDIR"
        
        ;Clean up the registry keys
        DeleteRegKey HKCU "${REG_APPPATH}"
        DeleteRegKey HKCU "${REG_UNINSTALL}"
SectionEnd

;=== FUNCTIONS ========================================================================

Function .onInit
        ;Get the command line parameters
        ; <nsis.sourceforge.net/Docs/AppendixE.html#E.1.11>
        ${GetParameters} $9

        ;Test for the "/?" help switch, and display parameter info
        ClearErrors
        ${GetOptions} $9 "/?" $8
        ${IfNot} ${Errors}
                MessageBox MB_ICONINFORMATION|MB_SETFOREGROUND "\
                           /PORTABLE : Install without shortuct / uninstaller$\n\
                           /S : Silent install$\n\
                           /D=%directory% : Specify destination directory$\n"
                Quit
        ${EndIf}

        ClearErrors
        ${GetOptions} $9 "/PORTABLE" $8
        ${IfNot} ${Errors}
            StrCpy $PortableMode 1
;            StrCpy $0 $PortableDestDir
        ${Else}
            StrCpy $PortableMode 0
;            StrCpy $0 $NormalDestDir
;            ${If} ${Silent}
;                Call RequireAdmin
;            ${EndIf}
        ${EndIf}
;
;        ${If} $InstDir == ""
;            ; User did not use /D to specify a directory,
;            ; we need to set a default based on the install mode
;            StrCpy $InstDir $0
;        ${EndIf}
;        Call SetModeDestinationFromInstdir
FunctionEnd

;--------------------------------------------------------------------------------------
Function PortableModePageCreate
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

;--------------------------------------------------------------------------------------
Function PortableModePageLeave
        ${NSD_GetState} $1 $0
        ${If} $0 <> ${BST_UNCHECKED}
                StrCpy $PortableMode 0
                ;StrCpy $InstDir $NormalDestDir
                ;Call RequireAdmin
        ${Else}
               StrCpy $PortableMode 1
               ;StrCpy $InstDir $PortableDestDir
        ${EndIf}
FunctionEnd

;Allow skipping the Directory page when local install is selected
;--------------------------------------------------------------------------------------
Function DirectoryPagePre
         ;If local mode selected, skip the Directory page, installation happens
         ; automatically to $LOCALAPPDATA
        ${If} $PortableMode = 0
              Abort
        ${EndIf}
FunctionEnd