Attribute VB_Name = "Misc"

Private Declare Function RegCreateKey Lib "advapi32.dll" _
 Alias "RegCreateKeyA" (ByVal hKey As Long, _
                        ByVal lpSubKey As String, _
                        phkResult As Long) As Long
                 
Private Declare Function RegSetValue Lib "advapi32.dll" _
 Alias "RegSetValueA" (ByVal hKey As Long, _
                       ByVal lpSubKey As String, _
                       ByVal dwType As Long, _
                       ByVal lpData As String, _
                       ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

' Reg Create Type Values...
Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

' Return codes from Registration functions.
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 260&
Private Const REG_EXPAND_SZ = 2
Private Const REG_DWORD = 4
Private Const REG_SZ = 1

Private Declare Sub SHChangeNotify Lib "shell32.dll" _
           (ByVal wEventId As Long, _
            ByVal uFlags As Long, _
            dwItem1 As Any, _
            dwItem2 As Any)

Const SHCNE_ASSOCCHANGED = &H8000000
Const SHCNF_IDLIST = &H0&

Public Type struct_File_Entry
   FileName As String * 64
   crc As Long
   Size As Long
End Type

Public Type struct_FIX_header
   sig     As String * 5    'FIX signature
   version As Long          'FIX version
   orgnum  As Long          'number of original files
   fixnum  As Long          'number of FIX files
   datsz   As Long          'FIX repair block data size
   crc     As Long          'CRC for this FIX file
   fullcrc As Boolean
End Type

Public Const FIX_SIG As String = "FIX0 "
Public Const FIX_VER As String = 1 * 512 + 1
Public Const FIX_MAX_DATA As Long = 255
Public Const FIX_MAX_SIZE As Long = &H6400000   'max 100MB per source file
Public Const REAL_APP_NAME As String = "Rapid-Repair"
Public Const REAL_CONTEXT_NAME As String = REAL_APP_NAME & " Repair archive"
Public Const REAL_CONTEXT_ACTION_NAME_VERIFY As String = "Verify with " & REAL_APP_NAME
Public Const REAL_CONTEXT_ACTION_NAME_REPAIR As String = "Repair with " & REAL_APP_NAME
Public Const REAL_EXTENSION As String = ".fix"
Public Const CMD_VERIFY_ONLY As String = ":verifyonly"
Public Const CMD_REMOVE_ASSOC As String = ":noassoc"
Public Const FIX_SHORT_CRC_LENGTH As Long = &H4000    '16k crc

Public Function FormatSize(Size As Long) As String
   Select Case Size
   Case &H0 To &H3FF
     FormatSize = Format(Size, "###0 B")
   Case &H400 To &HFFFFF
     FormatSize = Format(Size / &H400, "####.0 KB")
   Case &H100000 To &H3FFFFFFF
     FormatSize = Format(Size / &H100000, "####.0 MB")
   Case Is > &H40000000
     FormatSize = Format(Size / &H40000000, "####.0 GB")
   End Select
End Function

'Filter filename with extension from path (path is also resturned)
Public Function GetFilename(FileName As String, ByRef path As String) As String
   Dim S As String
   S = InStrRev(FileName, "\")
   path = ""
   If S > 0 Then
     path = Left(FileName, S - 1)
     GetFilename = Mid(FileName, S + 1)
   Else
     GetFilename = FileName
   End If
End Function

'Filter filename without extension (extension is also returned)
Public Function GetBaseFilename(FileName As String, ByRef Extension As String) As String
   Dim S As String
   S = InStrRev(FileName, ".")
   Extension = ""
   If S > 0 Then
     Extension = Mid(FileName, S)
     GetBaseFilename = Left(FileName, S - 1)
   Else
     GetBaseFilename = FileName
   End If
End Function

Public Function getCRC32(ByVal Size As Long) As Long
 Dim CRCv As Long
 CRCv = &HFFFFFFFF
 crcBuffer VarPtr(CRCv), _
           VarPtr(dataBuffer(LBound(dataBuffer))), _
           VarPtr(dataCRCt(LBound(dataCRCt))), _
           Size
 getCRC32 = CRCv
End Function

'Get a value from the registry
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' Search Data Types...
    Case REG_SZ, REG_EXPAND_SZ                              ' String Registry Key Data Type
        sKeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = sKeyVal                                   ' Return Value
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString                              ' Set Return Val To Empty String
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Function Associate(AppTitle As String, FileExtension As String, FileType As String, IconFileName As String, Optional Parameters As String)
   If AppTitle = "" Then Exit Function
   Dim sKeyName As String   ' Holds Key Name in registry.
   Dim sKeyValue As String  ' Holds Key Value in registry.
   Dim ret&           ' Holds error status if any from API calls.
   Dim lphKey&        ' Holds  key handle from RegCreateKey.
   Dim path As String

   path = App.path
   If Right(path, 1) <> "\" Then
      path = path & "\"
   End If

' This creates a Root entry that called as the string of AppTitle.
   sKeyName = AppTitle
   sKeyValue = FileType
   ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
   ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)

' This creates a Root entry called as the string of FileExtension associated with AppTitle.
   sKeyName = FileExtension
   sKeyValue = AppTitle
   ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
   ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)

' This sets the command line for AppTitle.
   sKeyName = AppTitle
   If Parameters <> "" Then
    sKeyValue = path & App.EXEName & ".exe " & Trim(Parameters)
   Else
    sKeyValue = path & App.EXEName & ".exe %1"
   End If
   ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
   ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, _
                       sKeyValue, MAX_PATH)

' This sets the icon for the file extension
   If IconFileName <> "" Then
    sKeyName = AppTitle
    sKeyValue = IconFileName
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, sKeyValue, MAX_PATH)
   End If
 
' This notifies the shell that the icon has changed
  SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
 
End Function

'Removes file association (that was created by Associate function)
Function RemoveAssociate(AppTitle As String, FileExtension As String)
    'Delete all keys
    RegDeleteKey HKEY_CLASSES_ROOT, FileExtension
    RegDeleteKey HKEY_CLASSES_ROOT, AppTitle & "\DefaultIcon"
    RegDeleteKey HKEY_CLASSES_ROOT, AppTitle & "\Shell\Open\Command"
    RegDeleteKey HKEY_CLASSES_ROOT, AppTitle & "\Shell\Open"
    RegDeleteKey HKEY_CLASSES_ROOT, AppTitle & "\Shell"
    RegDeleteKey HKEY_CLASSES_ROOT, AppTitle
    'Notify shell on the delete and refresh the icons
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Function

