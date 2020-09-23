VERSION 5.00
Begin VB.Form frmUninstall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uninstall"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmUninstall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   2940
      Width           =   1455
   End
   Begin VB.CommandButton cmdUninstall 
      Caption         =   "&Uninstall"
      Default         =   -1  'True
      Height          =   495
      Left            =   3180
      TabIndex        =   0
      Top             =   2940
      Width           =   1455
   End
   Begin VB.ListBox lstDisplay 
      Height          =   2790
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4515
   End
End
Attribute VB_Name = "frmUninstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sFiles2Delete As New Collection
Public sFolders2Delete As New Collection
Public sProgramName As String

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

'api for adding horizontal scroll to listbox
Private Declare Function SendMessage Lib "user32" _
          Alias "SendMessageA" ( _
          ByVal hwnd As Long, _
          ByVal wMsg As Long, _
          ByVal wParam As Long, _
          lParam As Any) As Long
Private Const LB_SETHORIZONTALEXTENT = &H194

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

Private Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" _
          (ByVal hwndOwner As Long, _
          ByVal nFolder As Long, _
          pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" _
          Alias "SHGetPathFromIDListA" ( _
          ByVal pidl As Long, _
          ByVal pszPath As String) As Long

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

Private Function funcCheckPathSlash(sPath As String) As String
  funcCheckPathSlash = sPath
  If Right(funcCheckPathSlash, 1) <> "\" Then
    funcCheckPathSlash = funcCheckPathSlash & "\"
  End If
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

Private Sub cmdExit_Click()
  Unload Me
  End
End Sub

Private Sub cmdUninstall_Click()
  Dim sItem2Delete As String
  Dim iEachItem As Long
  Dim iReply As Long
  
  On Error Resume Next
  'kill files
  For iEachItem = 1 To sFiles2Delete.Count
    Kill sFiles2Delete(iEachItem)
    Do While Err
      iReply = MsgBox("Error: Couldn't delete the file " & sFiles2Delete(iEachItem) & ".  Would you like to try again?  " & _
                Err.Description & " " & Err.Number, vbCritical + vbRetryCancel + vbQuestion, "Uninstall error")
      Err.Clear
      If iReply = vbRetry Then
        Kill sFiles2Delete(iEachItem)
      End If
    Loop
  Next iEachItem

  'kill folders
  For iEachItem = 1 To sFolders2Delete.Count
    ChDir "\" 'important
    RmDir sFolders2Delete(iEachItem)
    Do While Err
      iReply = MsgBox("Error: Couldn't delete the folder " & sFolders2Delete(iEachItem) & ".  Would you like to try again?  " & _
                Err.Description & " " & Err.Number, vbCritical + vbRetryCancel + vbQuestion, "Uninstall error")
      Err.Clear
      If iReply = vbRetry Then
        RmDir sFolders2Delete(iEachItem)
      End If
    Loop
  Next iEachItem
  Call subRemoveProgramToControlPanelUninstallList(sProgramName)
  MsgBox "Uninstall of '" & sProgramName & "' is complete.", vbInformation, "Uninstaller"
  Unload Me
  End
End Sub

Private Sub Form_Activate()
  Call AddScroll(lstDisplay)
End Sub

Public Sub AddScroll(lstList As ListBox)
    Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
    'Find Longest Text in Listbox

    For i = 0 To lstList.ListCount - 1
        If Len(lstList.List(i)) > Len(lstList.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i
    'Get Twips
    lngGreatestWidth = lstList.Parent.TextWidth(lstList.List(intGreatestLen) + Space(1))
    'Space(1) is used to prevent the last Character from being cut off
    'Convert to Pixels
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    'Use api to add scrollbar
    SendMessage lstList.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
End Sub

Private Sub Form_Load()
  Dim sInstallPath As String
  Dim sPathToTempFolder As String
  Dim sPath2Desktop As String
  Dim sPath2Programs As String
  Dim sPath2StartMenu As String
  Dim sPath2StartupFolder As String
  Dim sStartmenuGroupPath As String
  Dim sDesktopShortcutPath As String
  Dim sStartupShortcutPath As String
  Dim sFile As String
  Dim iEachFile As Long
  Dim iUninstallFileCounter As Long
  'the story:
  'see if i am in the Temporary Internet Files
  'if not copy myself there and then
  'run that copy and end this copy
  'that way I (this exe) will get thrown out with the
  'temp files at some later time
  
  sPathToTempFolder = SpecialFolderPath(sftUserTempInternetFiles)
          
  Me.WindowState = FormWindowStateConstants.vbNormal

  If LCase(funcCheckPathSlash(App.Path)) <> LCase(funcCheckPathSlash(sPathToTempFolder)) Then
    On Error Resume Next
    Call WriteINIString("Uninstall info", "Install location", funcCheckPathSlash(App.Path), App.Path & "\data.ini")
    Kill funcCheckPathSlash(sPathToTempFolder) & "Uninstall.exe"
    Kill funcCheckPathSlash(sPathToTempFolder) & "data.ini"
    Err.Clear
    On Error GoTo 0
    FileCopy funcCheckPathSlash(App.Path) & "Uninstall.exe", funcCheckPathSlash(sPathToTempFolder) & "Uninstall.exe"
    FileCopy funcCheckPathSlash(App.Path) & "data.ini", funcCheckPathSlash(sPathToTempFolder) & "data.ini"
    Shell """" & funcCheckPathSlash(sPathToTempFolder) & "Uninstall.exe" & """"
    End
  End If
  
  sProgramName = GetINIString("Settings", "App Name", App.Path & "\data.ini", "")
  If sProgramName = "" Then
    MsgBox "Error: Couldn't read the file '" & App.Path & "\data.ini' to get unintall information.", vbCritical, "Uninstaller"
    End
  End If
  
  sInstallPath = funcCheckPathSlash(GetINIString("Uninstall info", "Install location", App.Path & "\data.ini", ""))
  If sInstallPath = "" Then
    MsgBox "Error: Couldn't determine where the program '" & sProgramName & "' was installed.", vbCritical, "Uninstaller"
    End
  End If
  
  sStartmenuGroupPath = GetINIString("Uninstall info", "Startmenu group", App.Path & "\data.ini", "")
  sStartmenuGroupPath = funcCheckPathSlash(sStartmenuGroupPath)
  sDesktopShortcutPath = GetINIString("Uninstall info", "Desktop shortcut", App.Path & "\data.ini", "")
  sStartupShortcutPath = GetINIString("Uninstall info", "Startup shortcut", App.Path & "\data.ini", "")
  
  If sDesktopShortcutPath <> "" Then
    sFiles2Delete.Add sDesktopShortcutPath
  End If
  If sStartupShortcutPath <> "" Then
    sFiles2Delete.Add sStartupShortcutPath
  End If
  If sStartmenuGroupPath <> "" Then
    sFiles2Delete.Add sStartmenuGroupPath & "*.*"
    sFolders2Delete.Add sStartmenuGroupPath
  End If
  'get the list of files that the installer intalled into the app's folder.
  iUninstallFileCounter = 1
  Do
    sFile = GetINIString("Uninstall info", "Install file" & iUninstallFileCounter, App.Path & "\data.ini", "")
    If sFile <> "" Then
      sFiles2Delete.Add sInstallPath & sFile
      iUninstallFileCounter = iUninstallFileCounter + 1
    Else
      Exit Do
    End If
  Loop
  sFolders2Delete.Add sInstallPath
  
  lstDisplay.AddItem "This program will uninstall "
  lstDisplay.AddItem "'" & sProgramName & "'?"
  lstDisplay.AddItem ""
  lstDisplay.AddItem "Clicking uninstall will delete"
  lstDisplay.AddItem "the following files and folders:"
  lstDisplay.AddItem "Files:"
  
  For iEachFile = 1 To sFiles2Delete.Count
    lstDisplay.AddItem sFiles2Delete(iEachFile)
  Next iEachFile
  lstDisplay.AddItem ""
  lstDisplay.AddItem "Folders:"
  For iEachFile = 1 To sFolders2Delete.Count
    lstDisplay.AddItem sFolders2Delete(iEachFile)
  Next iEachFile
End Sub

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

