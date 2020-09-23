Attribute VB_Name = "FileExtension"
'Associate .lmf extension to LogoMan using

'In my case i use a other icon than the programm icon
'the save icon stored in SaveIcon.res
'the call is
'FileExt ".lmf", AppPath & "LogoMan.exe", "LogoMan Saves", AppPath & "LogoMan.exe,1"
'".lmf" is the extension we want to associate
'AppPath & "LogoMan.exe" is the exe windows should start
'"LogoMan Saves" is the filetype like "Word Document"...
'AppPath & "LogoMan.exe,1" is the icon (take icon 1 from logoman.exe)
'                          u also could write "C:\.....MyIcon.ico"
'                          but i wantet to store the icon in my exe



Option Explicit

'Registry
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 260&
Private Const REG_SZ = 1

'Tell Windows there is a new extension
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0&

Public Sub FileExt(Extension As String, ExePath As String, FileType As String, IconPlace As String)

 Dim KeyName As String
 Dim KeyValue As String
 Dim KeyHandle As Long

 'Create the Root Entry
 KeyName = FileType
 KeyValue = FileType
 RegCreateKey HKEY_CLASSES_ROOT, KeyName, KeyHandle
 RegSetValue KeyHandle, "", REG_SZ, KeyValue, 0&
 'The open command
 KeyValue = ExePath & " %1"
 RegCreateKey HKEY_CLASSES_ROOT, KeyName, KeyHandle
 RegSetValue KeyHandle, "shell\open\command", REG_SZ, KeyValue, MAX_PATH
 'What icon should the extension get
 KeyValue = IconPlace
 RegCreateKey HKEY_CLASSES_ROOT, KeyName, KeyHandle
 RegSetValue KeyHandle, "DefaultIcon", REG_SZ, KeyValue, MAX_PATH
 'Root Entry for extension
 KeyName = Extension
 KeyValue = FileType
 RegCreateKey HKEY_CLASSES_ROOT, KeyName, KeyHandle
 RegSetValue KeyHandle, "", REG_SZ, KeyValue, 0&
 'Tell the system we changed something
 SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub






