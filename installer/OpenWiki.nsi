#
# Script to create the installation program for OpenWiki software.
#

Name "OpenWiki"
BrandingText "OpenWiki v0.78 sp1"
ComponentText "This will install OpenWiki 0.78 sp1"
DirText "Setup has determined the optimal location to install. If you would like to change the directory, do so now."

LicenseText "OpenWiki is a free product. Please read the license terms below before installing."
LicenseData ..\LICENSE.txt

OutFile OpenWiki_078sp1.exe

CRCCheck on

EnabledBitmap checked.bmp
DisabledBitmap unchecked.bmp
Icon "main.ico"
UninstallIcon uninst.ico
UninstallText "Uninstall OpenWiki Now!"

InstType Full

Section "Programs (required)"
    SectionIn 1

    SetOverwrite on
    SetOutPath $INSTDIR
    File ..\README.txt
    File ..\LICENSE.txt
    File ..\INSTALL.txt
    File ..\CHANGES.txt
    SetOutPath $INSTDIR\installer
    File ..\installer\build.bat
    File ..\installer\checked.bmp
    File ..\installer\unchecked.bmp
    File ..\installer\main.ico
    File ..\installer\uninst.ico
    File ..\installer\OpenWiki.nsi
    File ..\installer\virtdir.vbs
    File ..\installer\virtdir.html
    SetOutPath $INSTDIR\data
    File ..\data\*.sql
    File ..\data\OpenWikiDist.mdb
    SetOutPath $INSTDIR\owbase
    File ..\owbase\*.*
    SetOutPath $INSTDIR\owbase\ow
    File ..\owbase\ow\*.*
    SetOutPath $INSTDIR\owbase\ow\css
    File ..\owbase\ow\css\*.*
    SetOutPath $INSTDIR\owbase\ow\images
    File ..\owbase\ow\images\*.*
    SetOutPath $INSTDIR\owbase\ow\images\icons
    File ..\owbase\ow\images\icons\*.*
    SetOutPath $INSTDIR\owbase\ow\images\icons\doc
    File ..\owbase\ow\images\icons\doc\*.*
    SetOutPath $INSTDIR\owbase\ow\xsl
    File ..\owbase\ow\xsl\*.*

    SetOverwrite off
    SetOutPath $INSTDIR\owbase\ow\my
    File ..\owbase\ow\my\*.*

    SetOverwrite on
    SetOutPath $INSTDIR\web1
    File ..\web1\*.*
    CreateDirectory $INSTDIR\web1\attachments

#    SetOutPath $INSTDIR\web2
#    File ..\web2\*.*
#    CreateDirectory $INSTDIR\web2\attachments

    WriteRegStr HKEY_LOCAL_MACHINE "Software\OpenWiki" "InstallDir" $INSTDIR

    WriteUninstaller Uninstall.exe

    CreateShortCut "$SMPROGRAMS\OpenWiki\Uninstall.lnk" "$INSTDIR\Uninstall.exe" "" "" 0
SectionEnd


Section "Sample Application"
    SectionIn 1
    WriteRegStr HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\OpenWikiDist" "DBQ" $INSTDIR\data\OpenWikiDist.mdb
    WriteRegStr HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\OpenWikiDist" "Description" "OpenWiki Distribution"
    WriteRegStr HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\OpenWikiDist" "Driver" $SYSDIR\odbcjt32.dll
    WriteRegDWORD HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\OpenWikiDist" "DriverId" 25
    WriteRegStr HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\OpenWikiDist" "FIL" "MS Access;"
    WriteRegDWORD HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\OpenWikiDist" "SafeTransactions" 0
    WriteRegStr HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\OpenWikiDist" "UID" ""
    WriteRegStr HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\OpenWikiDist\Engines\Jet" "ImplicitCommitSync" ""
    WriteRegStr HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\OpenWikiDist\Engines\Jet" "UserCommitSync" "Yes"
    WriteRegDWORD HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\OpenWikiDist\Engines\Jet" "Threads" 3
    WriteRegStr HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\ODBC Data Sources" "OpenWikiDist" "Microsoft Access Driver (*.mdb)"

    ClearErrors
    ExecWait 'cscript //Nologo "$INSTDIR\installer\virtdir.vbs" -i'
    IfErrors VirtdirError VirtdirNoError
    VirtdirNoError:
        StrCpy $5 "1"
    VirtdirError:
SectionEnd


Section -PostInstall
    WriteRegStr HKEY_LOCAL_MACHINE "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenWiki" "DisplayName" "OpenWiki"
    WriteRegStr HKEY_LOCAL_MACHINE "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenWiki" "UninstallString" '"$INSTDIR\Uninstall.exe"'
SectionEnd


Section Uninstall
    DeleteRegValue HKEY_LOCAL_MACHINE "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenWiki" "DisplayName"
    DeleteRegValue HKEY_LOCAL_MACHINE "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenWiki" "UninstallString"
    DeleteRegKey HKEY_LOCAL_MACHINE "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenWiki"

    DeleteRegValue HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\ODBC Data Sources" "OpenWikiDist"
    DeleteRegKey HKEY_LOCAL_MACHINE "Software\ODBC\ODBC.INI\OpenWikiDist"

    ExecWait 'cscript //Nologo "$INSTDIR\installer\virtdir.vbs" -uq'

    DeleteRegKey HKEY_LOCAL_MACHINE "Software\OpenWiki"

    RMDir  /r $INSTDIR
SectionEnd


Function .onInit
    StrCpy $1 $PROGRAMFILES 1 0
    StrCpy $INSTDIR $1:\OpenWiki
FunctionEnd


Function .onInstSuccess
    MessageBox MB_YESNO "View readme file?" IDNO NoReadme
    Exec 'notepad "$INSTDIR\README.txt"'
    NoReadme:
    StrCmp $5 "1" StartBrowser End
    StartBrowser:
    ExecShell "open" "$INSTDIR\installer\virtdir.html"
    End:
FunctionEnd


Function .onInstFailed
    Exec 'notepad "$INSTDIR\INSTALL.txt"'
FunctionEnd


#Function .onInit
#  StrCpy $9 0 ; we start on page 0
#FunctionEnd
#
#Function .onNextPage
#  StrCmp $9 1 "" noabort
#  MessageBox MB_YESNO "Before installing OpenWiki you must have the following components installed:$\n$\n        MS ADO.DB v2.5 or higher$\n        VBScript v5.5 or higher$\n        MSXML v3 SP2 or higher$\n        A Web Server (i.e. IIS)$\n$\nDo you want to continue?" IDYES noabort
#      Abort
#  noabort:
#    IntOp $9 $9 + 1
#FunctionEnd

