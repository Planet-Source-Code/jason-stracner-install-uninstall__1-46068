VERSION 5.00
Begin VB.Form frmInstall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Install"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5340
   Icon            =   "Install.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInstall 
      Caption         =   "&Install"
      Default         =   -1  'True
      Height          =   435
      Left            =   3840
      TabIndex        =   0
      Top             =   1260
      Width           =   1455
   End
   Begin VB.OptionButton optJustThisUser 
      Caption         =   "Install just for this user."
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Top             =   960
      Width           =   2895
   End
   Begin VB.OptionButton optAllUsers 
      Caption         =   "Install for all users on this computer"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   780
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.CheckBox chkStrartup 
      Caption         =   "Startup menu shortcut"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   2115
   End
   Begin VB.CheckBox chkDesktop 
      Caption         =   "Desktop shortcut"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   2115
   End
   Begin VB.CheckBox chkStartMenu 
      Caption         =   "Start menu shortcut"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Value           =   1  'Checked
      Width           =   2115
   End
   Begin VB.TextBox txtProgramGroup 
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Top             =   420
      Width           =   3915
   End
   Begin VB.TextBox txtInstallPath 
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Top             =   60
      Width           =   3915
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Program Group:"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Install path:"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "frmInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'in order for this to work all the files that need to be installed
'must be placed in a seperate folder.  In the zip file they should
'be inside a seperate folder also.
'There should be a data.ini file in this folder.
'Here is an example of what should be in the data.ini file for this
'to work:
'
'[Settings]
'App Name=My Program's Name
'Main exe = MyProgram.exe
'Other shortcut file1=Uninstall.exe
'Other shortcut title1=Uninstall This Program
'Other shortcut file2=Code.vbp
'Other shortcut title2=Edit the code in vb
'Other shortcut file3=InternetUpdate.exe
'Other shortcut title3=Update this program from the internet
'Ask to run=yes

Public sNameOfProgram4ShorcutsForMainExe As String
Public fAskToRunAfterInstall As Boolean
Public sPath2Desktop As String
Public sPath2StartMenuPrograms As String
Public sPath2StartMenuStartupFolder As String
Public sMainExeName As String
Dim sOtherFilesShorcuts() As String
Dim sOtherFilesShorcutsTitles() As String

'apis for adding the uninstall data to the registry so
'the user can uninstall from the control panel
Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
          ByVal Hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" _
          Alias "RegCreateKeyA" ( _
          ByVal Hkey As Long, _
          ByVal lpSubKey As String, _
          phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" _
          Alias "RegDeleteKeyA" ( _
          ByVal Hkey As Long, _
          ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" _
          Alias "RegDeleteValueA" ( _
          ByVal Hkey As Long, _
          ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" _
          Alias "RegOpenKeyA" ( _
          ByVal Hkey As Long, _
          ByVal lpSubKey As String, _
          phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
          Alias "RegQueryValueExA" ( _
          ByVal Hkey As Long, _
          ByVal lpValueName As String, _
          ByVal lpReserved As Long, _
          lpType As Long, _
          lpData As Any, _
          lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" _
          Alias "RegSetValueExA" ( _
          ByVal Hkey As Long, _
          ByVal lpValueName As String, _
          ByVal Reserved As Long, _
          ByVal dwType As Long, _
          lpData As Any, _
          ByVal cbData As Long) As Long
Const ERROR_SUCCESS = 0&
Const REG_SZ = 1 ' Unicode nul terminated String
Const REG_DWORD = 4 ' 32-bit number
Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

'apis for registering activex dlls, ocxs and exes.
Private Declare Function LoadLibraryRegister Lib "kernel32" _
          Alias "LoadLibraryA" ( _
          ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddressRegister Lib "kernel32" _
          Alias "GetProcAddress" ( _
          ByVal hModule As Long, _
          ByVal lpProcName As String) As Long
Private Declare Function FreeLibraryRegister Lib "kernel32" _
          Alias "FreeLibrary" ( _
          ByVal hLibModule As Long) As Long
Private Declare Function CreateThreadForRegister Lib "kernel32" _
          Alias "CreateThread" ( _
          lpThreadAttributes As Any, _
          ByVal dwStackSize As Long, _
          ByVal lpStartAddress As Long, _
          ByVal lParameter As Long, _
          ByVal dwCreationFlags As Long, _
          lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" _
          (ByVal hHandle As Long, _
          ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" _
          (ByVal hThread As Long, _
          lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" _
          (ByVal dwExitCode As Long)
Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long

Public Enum SHOWCMDFLAGS
    SHOWNORMAL = 5
    SHOWMAXIMIZE = 3
    SHOWMINIMIZE = 7
End Enum
          
'read ini files
Private Declare Function GetPrivateProfileString Lib "kernel32" _
          Alias "GetPrivateProfileStringA" ( _
          ByVal lpApplicationName As String, _
          ByVal lpKeyName As Any, _
          ByVal lpDefault As String, _
          ByVal lpReturnedString As String, _
          ByVal nSize As Long, _
          ByVal lpFileName As String) As Long
'write ini files
Private Declare Function WritePrivateProfileString Lib "kernel32" _
          Alias "WritePrivateProfileStringA" ( _
          ByVal lpApplicationName As String, _
          ByVal lpKeyName As Any, _
          ByVal lpString As Any, _
          ByVal lpFileName As String) As Long

'api for special folder paths
Private Type SHELLITEMID
  cb As Long
  abID As Byte
End Type

Private Type ITEMIDLIST
  mkid As SHELLITEMID
End Type

Public Enum SpecialFolderTypes
    sftCDBurningCache = 59&
    sftCommonAdminTools = 47&
    sftCommonApplicationData = 35&
    sftCommonDesktop = 25&
    sftCommonDocumentTemplates = 45&
    sftCommonFavorites = 31&
    sftCommonMyDocuments = 46&
    sftCommonMyPictures = 54&
    sftCommonProgramFiles = 43&
    sftCommonStartMenu = 22&
    sftCommonStartMenuPrograms = 23&
    sftCommonStartup = 24&
    sftFonts = 20&
    sftProgramFiles = 38&
    sftSystem32Folder = 41&
    sftSystemFolder = 37&
    sftThemes = 56&
    sftUserAdminTools = 48&
    sftUserApplicationData = 26&
    sftUserCookies = 33&
    sftUserDesktop = 16&
    sftUserDocumentTemplates = 21&
    sftUserFavorites = 6&
    sftUserHistory = 34&
    sftUserLocalApplicationData = 28&
    sftUserMyDocuments = 5&
    sftUserMyMusic = 13&
    sftUserMyPictures = 39&
    sftUserNetHood = 19&
    sftUserPrintHood = 27&
    sftUserProfileFolder = 40&
    sftUserRecentDocuments = 8&
    sftUserSendTo = 9&
    sftUserStartMenu = 11&
    sftUserStartMenuPrograms = 2&
    sftUserStartup = 7&
    sftUserTempInternetFiles = 32&
    sftWindowsFolder = 36&
End Enum

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
          (ByVal hwndOwner As Long, _
          ByVal nFolder As Long, _
          pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
          Alias "SHGetPathFromIDListA" ( _
          ByVal pidl As Long, _
          ByVal pszPath As String) As Long

'shell replacement
Private Declare Function ShellExecute _
      Lib "shell32.dll" _
      Alias "ShellExecuteA" _
      (ByVal hwnd As Long, _
      ByVal lpOperation As String, _
      ByVal lpFile As String, _
      ByVal lpParameters As String, _
      ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long


Public Function VBShellExecute(sFile As String, _
      Optional Args As String, _
      Optional Show As VbAppWinStyle = vbNormalFocus, _
      Optional InitDir As String, _
      Optional Verb As String, _
      Optional hwnd As Long = vbNull) As Long
    Call ShellExecute(hwnd, Verb, sFile, Args, InitDir, Show)
End Function

Public Function SpecialFolderPath(ByVal lngFolderType As SpecialFolderTypes) As String
    Dim strPath As String
    Dim IDL As ITEMIDLIST
    Dim MAX_PATH As Integer
    
    MAX_PATH = 255
    SpecialFolderPath = ""
    If SHGetSpecialFolderLocation(0&, lngFolderType, IDL) = 0& Then
        strPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath) Then
            SpecialFolderPath = Left$(strPath, InStr(strPath, vbNullChar) - 1&) & "\"
        End If
    End If
End Function

Public Function WriteINIString( _
          ByVal strSection As String, _
          ByVal strKeyName As String, _
          ByVal strValue As String, _
          ByVal strFile As String) As Long
 Dim lngStatus As Long

 lngStatus& = WritePrivateProfileString( _
            strSection, _
            strKeyName, _
            strValue, _
            strFile)
 WriteINIString& = (lngStatus& <> 0)
End Function

Public Function GetINIString( _
          ByVal strSection As String, _
          ByVal strKeyName As String, _
          ByVal strFile As String, _
          Optional ByVal strDefault As String = "") As String
 Dim strBuffer         As String * 256, lngSize As Long

 lngSize& = GetPrivateProfileString( _
            strSection$, _
            strKeyName$, _
            strDefault$, _
            strBuffer$, _
            CLng(256), _
            strFile$)
 GetINIString$ = Left$(strBuffer$, lngSize&)
End Function

Public Sub cmdInstall_Click()
  Dim sCurrentLocationPath As String
  Dim sPath4Shorcuts As String
  Dim sInstalledMainExe As String
  Dim sInstallLocationPath As String
  Dim sFile As String
  Dim iEachShortcut As Long
  Dim sFilesToCopy As New Collection
  Dim iEachFile As Long
  Dim iEachOtherShortcut As Long
  Dim sShortcutFileName As String
  Dim sShortchutTitle As String
  Dim iUninstallFileCounter As Long
  
  Me.Enabled = False
  Me.MousePointer = vbHourglass
  iUninstallFileCounter = 1
  fAskToRunAfterInstall = GetINIString("Settings", "Ask to run", App.Path & "\data.ini", "") <> ""
  
  If optAllUsers.Value Then
    sPath2Desktop = funcCheckPathSlash(SpecialFolderPath(sftCommonDesktop))
    sPath2StartMenuPrograms = funcCheckPathSlash(SpecialFolderPath(sftCommonStartMenuPrograms))
    sPath2StartMenuStartupFolder = funcCheckPathSlash(SpecialFolderPath(sftCommonStartup))
  Else
    sPath2Desktop = funcCheckPathSlash(SpecialFolderPath(sftUserDesktop))
    sPath2StartMenuPrograms = funcCheckPathSlash(SpecialFolderPath(sftUserStartMenuPrograms))
    sPath2StartMenuStartupFolder = funcCheckPathSlash(SpecialFolderPath(sftUserStartup))
  End If
          
  ReDim sOtherFilesShorcuts(0) As String
  ReDim sOtherFilesShorcutsTitles(0) As String
  iEachOtherShortcut = 0
  Do
     sShortcutFileName = GetINIString("Settings", "Other shortcut file" & iEachOtherShortcut + 1, App.Path & "\data.ini", "")
     sShortchutTitle = GetINIString("Settings", "Other shortcut title" & iEachOtherShortcut + 1, App.Path & "\data.ini", "")
     If sShortcutFileName <> "" And sShortchutTitle <> "" Then
      ReDim Preserve sOtherFilesShorcuts(iEachOtherShortcut) As String
      ReDim Preserve sOtherFilesShorcutsTitles(iEachOtherShortcut) As String
      sOtherFilesShorcuts(iEachOtherShortcut) = sShortcutFileName
      sOtherFilesShorcutsTitles(iEachOtherShortcut) = sShortchutTitle
      iEachOtherShortcut = iEachOtherShortcut + 1
    Else
      Exit Do
    End If
  Loop
  
  sInstallLocationPath = funcCheckPathSlash(Me.txtInstallPath)
  sCurrentLocationPath = funcCheckPathSlash(App.Path)
  If sInstallLocationPath <> "" Then
    'Dir will return "" if the location is not valid.
    If FileExists(sInstallLocationPath) = False Then
      Call subEnsurePath(sInstallLocationPath)
    End If
    If FileExists(sInstallLocationPath) = False Then
        MsgBox "Error while trying to build the path: " & vbNewLine & _
               sInstallLocationPath & "." & vbNewLine & _
               "Cannot continue with installation to this location.", _
               vbCritical, "Install error."
        Me.Enabled = True
        Exit Sub
    End If
    sPath4Shorcuts = sInstallLocationPath
    'copy the rest of the files in the install folder
    sFile = Dir(sCurrentLocationPath & "*.*", vbNormal + vbHidden + vbReadOnly + vbSystem + vbArchive)
    If sFile <> "" And sFile <> "." And sFile <> ".." Then
      If sFile <> "data.ini" Then
        sFilesToCopy.Add sFile
      End If
      Do
        sFile = Dir()
        If sFile <> "" And sFile <> "." And sFile <> ".." Then
          If sFile <> "data.ini" Then
            sFilesToCopy.Add sFile
          End If
        Else
          Exit Do
        End If
      Loop
    End If
  Else
    sPath4Shorcuts = sCurrentLocationPath
  End If
  
  '>>remove: this is taken care of by the uninstaller
  ''this will be used by the uninstaller
  'Call WriteINIString("Settings", "Install location", sPath4Shorcuts, App.Path & "\data.ini")
  
  For iEachFile = 1 To sFilesToCopy.Count
    sFile = sFilesToCopy(iEachFile)
    On Error Resume Next
    'unreginster dll so we can delete it. (Just in case.)
    Call RegSvr32(sInstallLocationPath & sFile, False)
    Call RegSvr32(sCurrentLocationPath & sFile, False)
    Err.Clear
    'del old version
    If FileExists(sInstallLocationPath & sFile) Then
      If FileDateTime(sInstallLocationPath & sFile) <= FileDateTime(sCurrentLocationPath & sFile) Then
        Call Kill(sInstallLocationPath & sFile)
        If Err And Err.Description <> "File not found" Then
          MsgBox "Error while trying to delete the file: " & vbNewLine & _
                 sInstallLocationPath & sFile & "." & vbNewLine & _
                 "Cannot continue with installation to this location." & vbNewLine & _
                 "Please make sure that you are not currently running an older version of the program.", _
                 vbCritical, "Install error."
          Err.Clear
          Me.MousePointer = vbNormal
          Me.Enabled = True
          Exit Sub
        End If
        Err.Clear
      Else
        MsgBox "Error while trying to install the file: " & vbNewLine & _
               sInstallLocationPath & sFile & "." & vbNewLine & _
               "Cannot install this file because it would " & vbNewLine & _
               "replace a file that is that is newer.", _
               vbInformation, "Install error."
      End If
    End If
    On Error GoTo 0
    If FileExists(sInstallLocationPath & sFile) = False Then
      Call FileCopy(sCurrentLocationPath & sFile, sInstallLocationPath & sFile)
      Call WriteINIString("Uninstall info", "Install file" & iUninstallFileCounter, sFile, App.Path & "\data.ini")
      iUninstallFileCounter = iUninstallFileCounter + 1
    End If
    On Error Resume Next
    'register all the files as if they are dlls (just incase) and ignore errors.
    Call RegSvr32(sInstallLocationPath & sFile, False)
    Err.Clear
    On Error GoTo 0
  Next iEachFile
  
  On Error Resume Next
  Call subRemoveProgramToControlPanelUninstallList(sNameOfProgram4ShorcutsForMainExe)
  Call subAddProgramToControlPanelUninstallList(sNameOfProgram4ShorcutsForMainExe, sPath4Shorcuts & "Uninstall.exe")
  On Error GoTo 0
  
  'Start menu shortcut.
  If chkStartMenu.Value = vbChecked Then
    Call WriteINIString("Uninstall info", "Startmenu group", sPath2StartMenuPrograms & sNameOfProgram4ShorcutsForMainExe, App.Path & "\data.ini")
    'This will build this folder structure if needed.
    Call subEnsurePath(sPath2StartMenuPrograms & sNameOfProgram4ShorcutsForMainExe)
    'This makes the shortcut file using a windows API call.
    Call fCreateShellLink(sPath2StartMenuPrograms & sNameOfProgram4ShorcutsForMainExe & "\" & sNameOfProgram4ShorcutsForMainExe & ".lnk", _
          sPath4Shorcuts & sMainExeName, sPath4Shorcuts, "", "", 0, SHOWNORMAL)
    'create the other shorcuts that are named in the ini file.
    For iEachShortcut = LBound(sOtherFilesShorcuts) To UBound(sOtherFilesShorcuts)
      If sOtherFilesShorcutsTitles(iEachShortcut) <> "" And sOtherFilesShorcuts(iEachShortcut) <> "" Then
        Call fCreateShellLink(sPath2StartMenuPrograms & sNameOfProgram4ShorcutsForMainExe & "\" & sOtherFilesShorcutsTitles(iEachShortcut) & ".lnk", _
                  sPath4Shorcuts & sOtherFilesShorcuts(iEachShortcut), sPath4Shorcuts, "", "", 0, SHOWNORMAL)
      End If
    Next iEachShortcut
  Else
    Call WriteINIString("Uninstall info", "Startmenu group", "", App.Path & "\data.ini")
  End If

  'Desktop shortcut prompt.
  If chkDesktop.Value = vbChecked Then
    Call WriteINIString("Uninstall info", "Desktop shortcut", sPath2Desktop & sNameOfProgram4ShorcutsForMainExe & ".lnk", App.Path & "\data.ini")
    Call fCreateShellLink(sPath2Desktop & sNameOfProgram4ShorcutsForMainExe & ".lnk", _
          sPath4Shorcuts & sMainExeName, sPath4Shorcuts, "", "", 0, SHOWNORMAL)
  Else
    Call WriteINIString("Uninstall info", "Desktop shortcut", "", App.Path & "\data.ini")
  End If
  
  'Startup shortcut prompt.
  If chkStrartup.Value = vbChecked Then
    Call WriteINIString("Uninstall info", "Startup shortcut", sPath2StartMenuStartupFolder & sNameOfProgram4ShorcutsForMainExe & ".lnk", App.Path & "\data.ini")
    Call fCreateShellLink(sPath2StartMenuStartupFolder & sNameOfProgram4ShorcutsForMainExe & ".lnk", _
          sPath4Shorcuts & sMainExeName, sPath4Shorcuts, "", "", 0, SHOWNORMAL)
  Else
    Call WriteINIString("Uninstall info", "Startup shortcut", "", App.Path & "\data.ini")
  End If

  'we have to wait until the end to copy the data.ini file because we are still writting to it.
  If FileExists(sInstallLocationPath & "data.ini") Then
    Call Kill(sInstallLocationPath & "data.ini")
  End If
  Call WriteINIString("Uninstall info", "Install file" & iUninstallFileCounter, "data.ini", App.Path & "\data.ini")
  DoEvents
  Call FileCopy(sCurrentLocationPath & "data.ini", sInstallLocationPath & "data.ini")
    
  Me.MousePointer = vbNormal
  If fAskToRunAfterInstall Then
    If MsgBox("Install is complete.  Run '" & sNameOfProgram4ShorcutsForMainExe & "'?", vbQuestion Or vbYesNo, "Installer") = vbYes Then
      Call VBShellExecute("""" & sPath4Shorcuts & sMainExeName & """")
    End If
  Else
    MsgBox sNameOfProgram4ShorcutsForMainExe & " is now installed.", vbInformation, "Installer"
  End If
  
  Unload Me
End Sub

Private Function funcCheckPathSlash(sPath As String) As String
  funcCheckPathSlash = sPath
  If Right(funcCheckPathSlash, 1) <> "\" Then
    funcCheckPathSlash = funcCheckPathSlash & "\"
  End If
End Function


Private Sub subEnsurePath(sPath2Make)
  Dim iInitialindex As Long
  Dim iSlashPos As Long
  Dim Cnt As Long
  Dim sBasePath As String
  Dim sCurrFolder As String
  Dim sDirOfBasePath As String
  Dim sParamPath2Make As String
  
  sParamPath2Make = sPath2Make
  If Right(sParamPath2Make, 1) <> "\" Then
    sParamPath2Make = sParamPath2Make & "\"
  End If
  iInitialindex = 4
  iSlashPos = InStr(1, sParamPath2Make, "\", 1)
  For Cnt = 1 To Len(sParamPath2Make)
    iSlashPos = InStr((iSlashPos + 1), sParamPath2Make, "\", 1)
    If iSlashPos = 0 Then Exit For 'Last slash
    sBasePath = Left(sParamPath2Make, (iSlashPos - 1))
    sCurrFolder = Mid(sBasePath, iInitialindex)
    iInitialindex = iInitialindex + Len(sCurrFolder) + 1
    sDirOfBasePath = Dir(sBasePath, vbDirectory)
    If StrComp(sDirOfBasePath, sCurrFolder, 1) <> 0 Then
      MkDir (sBasePath)
    End If
  Next Cnt
End Sub

'requires "VB 5 - IShellLinkA Interface(ANSI)"
'usually in the file (SHELLLNK.TLB)
'does this require any dependencies?  i know that the .tlb is not required on the client.
Private Function fCreateShellLink(sLnkFile As String, sExeFile As String, sWorkDir As String, _
       sExeArgs As String, sIconFile As String, lIconIdx As Long, ShowCmd As SHOWCMDFLAGS) As Long

    Dim cShellLink   As ShellLinkA   ' An explorer IShellLinkA(Win 9x/Win NT) instance
    Dim cPersistFile As IPersistFile ' An explorer IPersistFile instance
    
    If (sLnkFile = "") Or (sExeFile = "") Then
        Exit Function
    End If

    On Error GoTo fCreateShellLinkError
    Set cShellLink = New ShellLinkA   'Create new IShellLink interface
    Set cPersistFile = cShellLink     'Implement cShellLink's IPersistFile interface
    
    With cShellLink
        .SetPath sExeFile
        If sWorkDir <> "" Then .SetWorkingDirectory sWorkDir
        If sExeArgs <> "" Then .SetArguments sExeArgs
        .SetDescription "" & vbNullChar
        If sIconFile <> "" Then .SetIconLocation sIconFile, lIconIdx
        .SetShowCmd ShowCmd
    End With

    cShellLink.Resolve 0, SLR_UPDATE
    cPersistFile.Save StrConv(sLnkFile, vbUnicode), 0 'Unicode conversion that must be done!
    fCreateShellLink = True 'Return Success
fCreateShellLinkError:
    Set cPersistFile = Nothing
    Set cShellLink = Nothing
End Function

Private Sub Form_Load()
  Me.txtInstallPath.ToolTipText = "Tip: Make this blank to just add shortcuts to the current location of the progam."
    
  sNameOfProgram4ShorcutsForMainExe = GetINIString("Settings", "App Name", App.Path & "\data.ini", "")
  sMainExeName = GetINIString("Settings", "Main exe", App.Path & "\data.ini", "")
  If sNameOfProgram4ShorcutsForMainExe = "" Or sMainExeName = "" Then
    MsgBox "Error: Couldn't read the file '" & App.Path & "\data.ini' to get intall information.", vbCritical, "Installer"
    End
  End If
  
  Me.txtInstallPath.Text = "C:\Program Files\" & sNameOfProgram4ShorcutsForMainExe & "\"
  Me.txtProgramGroup.Text = sNameOfProgram4ShorcutsForMainExe
End Sub

'thanks to 'Thomas Sturm' for showing how to do this activex registration stuff
Private Function RegSvr32(ByVal FileName As String, bUnReg As Boolean) As Boolean
  Dim lLib As Long
  Dim lProcAddress As Long
  Dim lThreadID As Long
  Dim lSuccess As Long
  Dim lExitCode As Long
  Dim lThread As Long
  Dim bAns As Boolean
  Dim sPurpose As String
  
  On Error Resume Next
  RegSvr32 = False 'assume failure incase the blows up
  
  sPurpose = IIf(bUnReg, "DllUnregisterServer", _
    "DllRegisterServer")
  
  If Dir(FileName) = "" Then Exit Function
  
  lLib = LoadLibraryRegister(FileName)
  'could load file
  If lLib = 0 Then Exit Function
  
  lProcAddress = GetProcAddressRegister(lLib, sPurpose)
  
  If lProcAddress = 0 Then
    'Not an ActiveX Component
     FreeLibraryRegister lLib
     Exit Function
  Else
     lThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lProcAddress, ByVal 0&, 0&, lThread)
     If lThread Then
          lSuccess = (WaitForSingleObject(lThread, 10000) = 0)
          If Not lSuccess Then
             Call GetExitCodeThread(lThread, lExitCode)
             Call ExitThread(lExitCode)
             bAns = False
             Exit Function
          Else
             bAns = True
          End If
          CloseHandle lThread
          FreeLibraryRegister lLib
     End If
  End If
  RegSvr32 = bAns
End Function

Public Function funcHandleError(ByVal sModule As String, ByVal sProcedure As String, ByVal oErr As ErrObject) As Long
  Dim Msg As String
  
  Msg = oErr.Source & " caused error '" & _
          oErr.Description & "' (" & _
          oErr.Number & ")" & vbCrLf & _
          "in module " & sModule & _
          " procedure " & sProcedure & _
          ", line " & Erl & "."
  funcHandleError = MsgBox( _
          Msg, vbAbortRetryIgnore + vbMsgBoxHelpButton + vbCritical, _
          "What to you want to do?", _
          oErr.HelpFile, _
          oErr.HelpContext)
End Function

Public Function FileExists(ByVal strFileName As String) As Boolean
  Dim intLen As Integer
  On Error Resume Next

  If strFileName$ <> "" Then
    intLen% = Len(Dir$(strFileName$))
    If intLen = 0 Then
      intLen% = Len(Dir$(strFileName$, vbDirectory))
    End If
    FileExists = (Not Err And intLen% > 0)
  Else
    FileExists = False
  End If
  Err.Clear
End Function

'Thanks to "Rabid Nerd Productions" for showing me this registry stuff.
Public Sub subAddProgramToControlPanelUninstallList(ProgramName As String, UninstallCommand As String)
  'Add a program to the 'Add/Remove Programs' registry keys
  Call SaveRegistryString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + ProgramName, "DisplayName", ProgramName)
  Call SaveRegistryString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + ProgramName, "UninstallString", UninstallCommand)
End Sub

Public Sub subRemoveProgramToControlPanelUninstallList(ProgramName As String)
  'Remove a program from the 'Add/Remove Programs' registry keys
  Call DeleteRegistryKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + ProgramName)
End Sub
    
Public Sub SaveRegistryString(Hkey As HKeyTypes, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    'Call SaveRegistryString(HKEY_CURRENT_USER, "Software\VBW\Registry", "String", text1.text)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub

Public Function DeleteRegistryKey(ByVal Hkey As HKeyTypes, ByVal strPath As String) As Long
    'EXAMPLE:
    'Call DeleteRegistryKey(HKEY_CURRENT_USER, "Software\VBW\Registry")
    Call RegDeleteKey(Hkey, strPath)
End Function

