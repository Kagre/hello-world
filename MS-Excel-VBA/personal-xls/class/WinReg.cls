'Additional Resources:
'all Win32 API by DLL
'http://msdn.microsoft.com/en-us/library/ee391646(v=vs.85).aspx

'registry functions
'http://msdn.microsoft.com/en-us/library/ms724875(v=vs.85).aspx


#If Win32 Then
Const REG_CREATED_NEW_KEY = &H1        ' A new key was created
Const REG_OPENED_EXISTING_KEY = &H2    ' An existing key was opened

Public Enum KeyHandle
    KEY_NOT_SET = 0
End Enum

Public Enum HiveKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum WinErrorCode              'winerror.h
    ERROR_SUCCESS = 0&
    ERROR_FILE_NOT_FOUND = 2&         ' Registry path does not exist
    ERROR_ACCESS_DENIED = 5&          ' Requested permissions not available
    ERROR_INVALID_HANDLE = 6&         ' Invalid handle or top-level key
    ERROR_BAD_NETPATH = 53            ' Network path not found
    ERROR_INVALID_PARAMETER = 87      ' Bad parameter to a Win32 API function
    ERROR_CALL_NOT_IMPLEMENTED = 120& ' Function valid only in WinNT/2000?XP
    ERROR_INSUFFICIENT_BUFFER = 122   ' Buffer too small to hold data
    ERROR_BAD_PATHNAME = 161          ' Registry path does not exist
    ERROR_NO_MORE_ITEMS = 259&        ' Invalid enumerated value
    ERROR_BADDB = 1009                ' Corrupted registry
    ERROR_BADKEY = 1010               ' Invalid registry key
    ERROR_CANTOPEN = 1011&            ' Cannot open registry key
    ERROR_CANTREAD = 1012&            ' Cannot read from registry key
    ERROR_CANTWRITE = 1013&           ' Cannot write to registry key
    ERROR_REGISTRY_RECOVERED = 1014&  ' Recovery of part of registry successful
    ERROR_REGISTRY_CORRUPT = 1015&    ' Corrupted registry
    ERROR_REGISTRY_IO_FAILED = 1016&  ' Input/output operation failed
    ERROR_NOT_REGISTRY_FILE = 1017&   ' Input file not in registry file format
    ERROR_KEY_DELETED = 1018&         ' Key already deleted
    ERROR_KEY_HAS_CHILDREN = 1020&    ' Key has subkeys & cannot be deleted

    ERROR_MORE_DATA = 234             ' More data is available.
End Enum

Public Enum SecurityAccessRight
    STANDARD_RIGHTS_ALL = &H1F0000
    STANDARD_RIGHTS_READ = &H20000
    STANDARD_RIGHTS_WRITE = &H20000
    SYNCHRONIZE = &H100000

    KEY_CREATE_LINK = &H20
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_NOTIFY = &H10
    KEY_QUERY_VALUE = &H1
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    KEY_SET_VALUE = &H2
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
End Enum

Public Enum RegKeyType
    REG_OPTION_BACKUP_RESTORE = 4
    REG_OPTION_NON_VOLATILE = 0
    REG_OPTION_VOLATILE = 1
End Enum

Public Enum RegKeyValueType
    REG_SZ = 1
    REG_EXPAND_SZ = 1
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
End Enum

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'    ByVal hKey As Long               Handle of already open key
'    ByVal lpSubKey As String, _      Path from hKey to key to open
'    ByVal ulOptions As Long, _       Reserved, must be 0
'    ByVal samDesired As Long, _      SAM constant(s)
'    phkResult As Long) As Long       Handle of newly opened key
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
       (ByVal hKey As HiveKey, ByVal lpSubKey As String, ByVal ulOptions As Long, _
       ByVal samDesired As SecurityAccessRight, phkResult As KeyHandle) As WinErrorCode

'       (ByVal hKey As Long, _                 ' Handle of open key
Private Declare Function RegCloseKey Lib "advapi32.dll" _
       (ByVal hKey As KeyHandle) As WinErrorCode

'       (ByVal hKey As Long, _                ' Handle of already open key
'       ByVal lpSubKey As String, _           ' Path from hkey to key to open/create
'       ByVal Reserved As Long, _             ' Reserved, must be 0
'       ByVal lpClass As String, _            ' Reserved, must be a null string
'       ByVal dwOptions As Long, _            ' Type of key, or backup/restore
'       ByVal samDesired As Long, _           ' SAM constant(s)
'       lpSecurityAttributes As SECURITY_ATTRIBUTES, _
'       phkResult As Long, _                  ' Handle of opened/created key
'       lpdwDisposition As Long) As Long      '
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
       (ByVal hKey As HiveKey, ByVal lpSubKey As String, ByVal Reserved As Long, _
       ByVal lpClass As String, ByVal dwOptions As RegKeyType, ByVal samDesired As SecurityAccessRight, _
       lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As KeyHandle, _
       lpdwDisposition As Long) As WinErrorCode

'       (ByVal hKey As Long, _               ' Handle of key to query
'        ByVal lpClass As String, _          ' Class (unsued except for remote keys)
'        lpcbClass As Long, _                ' Length of class string (unused)
'        ByVal lpReserved As Long, _         ' Reserved, must be 0
'        lpcSubKeys As Long, _               ' Number of subkeys
'        lpcbMaxSubKeyLen As Long, _         ' Length of longest subkey name
'        lpcbMaxClassLen As Long, _          ' Length of longest class name (unused)
'        lpcValues As Long, _                ' Number of values
'        lpcbMaxValueNameLen As Long, _      ' Length of longest value name
'        lpcbMaxValueLen As Long, _          ' Length of longest value
'        lpcbSecurityDescriptor As Long, _   ' Security Descriptor
'        lpftLastWriteTime As FILETIME _     ' Date/time of last write to key
'       ) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
       (ByVal hKey As KeyHandle, ByVal lpClass As String, lpcbClass As Long, _
        ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, _
        lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, _
        lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, _
        lpftLastWriteTime As FILETIME) As WinErrorCode

'      (ByVal hKey As Long, _           ' Handle of open key
'       ByVal dwIndex As Long, _        ' Index of subkey
'       ByVal lpName As String, _       ' Name of key
'       lpcbName As Long, _             ' Length of name buffer
'       ByVal lpReserved As Long, _     ' Reserved, must be 0
'       ByVal lpClass As String, _      ' Class name (usually usused)
'       lpcbClass As Long, _            ' Length of class name
'       lpftLastWriteTime As FILETIME _ ' Date of last write
'      ) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
      (ByVal hKey As KeyHandle, ByVal dwIndex As Long, ByVal lpName As String, _
       lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
       lpcbClass As Long, lpftLastWriteTime As FILETIME) As WinErrorCode

'       (ByVal hKey As Long, _                 ' Handle of open key
'        ByVal lpSubKey As String _            ' Path from handle to key
'       ) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
       (ByVal hKey As KeyHandle, ByVal lpSubKey As String) As WinErrorCode

Private Declare Function SHDeleteKey Lib "SHLWAPI.DLL" Alias "SHDeleteKeyA" _
       (ByVal hKey As KeyHandle, ByVal pszSubKey As String) As WinErrorCode

Private Declare Function SHRegGetValue Lib "SHLWAPI.DLL" Alias "SHGetValueA" ( _
    ByVal hKey As KeyHandle, ByVal pszSubKey As String, ByVal pszValue As String, _
    pdwType As RegKeyValueType, pvData As Any, pcbData As Long) As WinErrorCode

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
    ByVal hKey As KeyHandle, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As RegKeyValueType, lpData As Any, cbData As Long)
    


'***************************************************************************
Private oHive As HiveKey
'Private oHandle As KeyHandle
Private oSubKey As String
Private oValueName As String

Private oSubKeys() As String
Private oKeysSet As Boolean

Private oValues() As Variant
Private oValuesSet As Boolean
'Private oAccess As SecurityAccessRight

Private Sub Class_Initialize()
    oHive = HKEY_CURRENT_USER
    oSubKey = ""
    oValueName = ""
    oKeysSet = False
    oValuesSet = False
    'oAccess = KEY_READ
End Sub

Public Static Function RegKeyExists(hKey As HiveKey, sKeyPath As String) As Boolean
    Dim lResult As WinErrorCode, hSubKey As KeyHandle
    
    lResult = RegOpenKeyEx(hKey, sKeyPath, 0, KEY_EXECUTE, hSubKey)
    ' Determine if handle is valid
    If hSubKey <> KEY_NOT_SET Then
       RegKeyExists = True
       RegCloseKey hSubKey
    End If
End Function

Public Static Function RegKeyValueExists(hKey As KeyHandle, sKeyPath As String, sValueName As String) As Boolean
    RegKeyValueExists = False
    If sValueName = "" Then Exit Function
    Dim e As WinErrorCode, hSubKey As KeyHandle, s As String
    e = RegOpenKeyEx(hKey, sKeyPath, 0, KEY_READ, hSubKey)
    If hSubKey <> KEY_NOT_SET Then
        e = SHRegGetValue(hSubKey, "", sValueName, 0, ByVal s, Len(s))
        RegKeyValueExists = (e <> ERROR_FILE_NOT_FOUND)
        RegCloseKey hSubKey
    End If
End Function

Public Property Let Hive(NewHive As HiveKey)
    Dim k As Boolean: k = False
    For Each mbr In Array(HiveKey.HKEY_CLASSES_ROOT, HiveKey.HKEY_CURRENT_CONFIG, HiveKey.HKEY_CURRENT_USER, HiveKey.HKEY_LOCAL_MACHINE, HiveKey.HKEY_USERS)
        k = (mbr = NewHive)
        If k Then Exit For
    Next mbr
    If k Then oHive = NewHive Else Err.Raise (vbObjectError + WinErrorCode.ERROR_FILE_NOT_FOUND)
    oKeysSet = False: oValuesSet = False
End Property
Public Static Property Get Hive() As HiveKey
    Hive = oHive
End Property

Public Function upPath(Optional count As Long = 1) As String
    If count < 1 Then Exit Function
    Dim tmp As String, at As Long, i As Long
    i = 0
    tmp = oSubKey
    at = InStrRev(oSubKey, "\")
    Do While i < count And at > 0
        i = i + 1
        tmp = Left(tmp, at - 1)
        at = InStrRev(oSubKey, "\")
    Loop
    If at = 0 Then oSubKey = "" Else oSubKey = tmp
    upPath = oSubKey
    oKeysSet = False: oValuesSet = False
End Function

Public Function downPath(AppendPath As String) As Boolean
    downPath = False
    If AppendPath = "" Then Exit Function
    Dim NewKey As String
    If Left(AppendPath, 1) = "\" Then NewKey = oSubKey & AppendPath Else NewKey = oSubKey & "\" & AppendPath
    downPath = RegKeyExists(oHive, NewKey)
    If Not downPath Then Exit Function
    oSubKey = NewKey
    oKeysSet = False: oValuesSet = False
End Function

Public Property Let SubKey(sSubKey As String)
    Dim NewKey As String, st As Long, en As Long
    NewKey = Trim(sSubKey)
    If Left(NewKey, 1) = "\" Then NewKey = Mid(NewKey, 2)
    If Right(NewKey, 1) = "\" Then NewKey = Left(NewKey, Len(NewKey) - 1)
    If RegKeyExists(oHive, NewKey) Then oSubKey = NewKey Else Err.Raise (vbObjectError + WinErrorCode.ERROR_FILE_NOT_FOUND), "Windows Registry", "Key '" & HiveKeyName(oHive) & "\" & NewKey & "' not found."
    oKeysSet = False: oValuesSet = False
End Property

Public Static Property Get SubKey() As String
    SubKey = oSubKey
End Property

Public Property Let ValueName(sValueName As String)
    Dim vn As String: vn = Trim(sValueName)
    If RegKeyValueExists(oHive, oSubKey, vn) Then
        oValueName = vn
    Else
        Err.Raise WinErrorCode.ERROR_FILE_NOT_FOUND, "Windows Registry", "Value '" & vn & "' not found in key '" & HiveKeyName(oHive) & "\" & oSubKey & "'."
    End If
End Property

Public Static Property Get ValueName() As String
    ValueName = oValueName
End Property

Public Static Property Get ValueType() As RegKeyValueType
    If oValueName = "" Then Err.Raise 425, "ValueName" '(425) - Invalid Object Use
    Dim vt As RegKeyValueType
    
    Dim e As WinErrorCode, hSubKey As KeyHandle, s As String
    e = RegOpenKeyEx(oHive, oSubKey, 0, KEY_READ, hSubKey)
    If hSubKey <> KEY_NOT_SET Then
        e = SHRegGetValue(hSubKey, "", oValueName, vt, ByVal s, Len(s))
        'RegKeyValueExists = (e <> ERROR_FILE_NOT_FOUND)
        RegCloseKey hSubKey
    End If
    ValueType = vt
End Property

Public Static Property Get ValueBytes() As Long
    If oValueName = "" Then Err.Raise 425, "ValueName" '(425) - Invalid Object Use
    Dim vt As RegKeyValueType, vb As Long
    
    Dim e As WinErrorCode, hSubKey As KeyHandle, s As String
    e = RegOpenKeyEx(oHive, oSubKey, 0, KEY_READ, hSubKey)
    If hSubKey <> KEY_NOT_SET Then
        e = SHRegGetValue(hSubKey, "", oValueName, vt, ByVal s, vb)
        'RegKeyValueExists = (e <> ERROR_FILE_NOT_FOUND)
        RegCloseKey hSubKey
    End If
    ValueBytes = vb
End Property

Public Static Property Get Value() As Variant
    If oValueName = "" Then Err.Raise vbObjectError + WinErrorCode.ERROR_FILE_NOT_FOUND
    Dim e As WinErrorCode, t As RegKeyValueType, hSubKey As KeyHandle
    Dim L As Long, V As Long, s As String, B() As Byte
    t = ValueType
    e = RegOpenKeyEx(oHive, oSubKey, 0, KEY_READ, hSubKey)
    If e <> ERROR_SUCCESS Then Err.Raise e, "WinReg", "Object defined WinErrorCode."
    
    Select Case t
     Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
        e = SHRegGetValue(hSubKey, "", oValueName, REG_DWORD, V, 4)
        If e <> ERROR_SUCCESS Then Err.Raise e, "WinReg:Value Get", "Object defined WinErrorCode."
        Value = V
     Case REG_SZ, REG_EXPAND_SZ, REG_BINARY
        L = ValueBytes + 4
        s = String(L, vbNullChar)
        e = SHRegGetValue(hSubKey, "", oValueName, REG_SZ, ByVal s, Len(s))
        If e <> ERROR_SUCCESS Then Err.Raise e, "WinReg:Value Get", "Object defined WinErrorCode."
        If t = REG_BINARY Then Value = Left(s, L - 4) Else Value = Left(s, InStr(s, vbNullChar) - 1)
'     Case REG_BINARY
'        l = ValueBytes + 4
'        ReDim b(l) As Byte
'        e = SHRegGetValue(hSubKey, "", oValueName, REG_BINARY, ByVal b, UBound(b))
'        Do While e <> ERROR_SUCCESS
'            Err.Raise e, "WinReg:Value Get", "Object defined WinErrorCode."
'            e = SHRegGetValue(hSubKey, "", oValueName, REG_BINARY, ByVal b, UBound(b))
'        Loop
'        RegCloseKey hSubKey
'        ReDim Preserve b(l - 4)
'        Value = b
     Case Else
        '(31032) -- Unable to create embedded object
        Err.Raise 31032, "WinReg", "Unknown registry key type encountered, please add ValueType:" & t & " to WinReg:RegKeyValueType and WinReg:Value code."
    End Select
    RegCloseKey hSubKey
End Property

Public Function setKey(val As Variant, typ As RegKeyValueType)
    If oValueName = "" Then Err.Raise vbObjectError + WinErrorCode.ERROR_FILE_NOT_FOUND
    Dim e As WinErrorCode, hSubKey As KeyHandle
    Dim L As Long, V As Long, s As String
    e = RegOpenKeyEx(oHive, oSubKey, 0, KEY_WRITE, hSubKey)
    If e <> ERROR_SUCCESS Then Err.Raise e, "WinReg", "Object defined WinErrorCode."
    
    Select Case typ
     Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
        V = val
        e = RegSetValueEx(hSubKey, oValueName, 0, REG_DWORD, V, 4)
        If e <> ERROR_SUCCESS Then Err.Raise e, "WinReg:Value Get", "Object defined WinErrorCode."
     Case REG_SZ, REG_EXPAND_SZ, REG_BINARY
        s = val + vbNullChar
        e = RegSetValueEx(hSubKey, oValueName, 0, REG_SZ, V, Len(s))
        If e <> ERROR_SUCCESS Then Err.Raise e, "WinReg:Value Get", "Object defined WinErrorCode."
'     Case REG_BINARY
'        l = ValueBytes + 4
'        ReDim b(l) As Byte
'        e = SHRegGetValue(hSubKey, "", oValueName, REG_BINARY, ByVal b, UBound(b))
'        Do While e <> ERROR_SUCCESS
'            Err.Raise e, "WinReg:Value Get", "Object defined WinErrorCode."
'            e = SHRegGetValue(hSubKey, "", oValueName, REG_BINARY, ByVal b, UBound(b))
'        Loop
'        RegCloseKey hSubKey
'        ReDim Preserve b(l - 4)
'        Value = b
     Case Else
        '(31032) -- Unable to create embedded object
        Err.Raise 31032, "WinReg", "Unknown registry key type encountered, please add ValueType:" & t & " to WinReg:RegKeyValueType and WinReg:Value code."
    End Select
    RegCloseKey hSubKey
End Function

Public Property Get SubKeys() As Variant
    'todo - add code that returns array of keys

End Property

Public Property Get Values() As Variant
    If oValuesSet Then
        Values = oValues
        Exit Property
    End If
    
    Dim a() As Variant
    ReDim a(2, 5)
    'todo - add code that returns 2d array of values


    Values = TransposeArray(a)
    oValues = Values
    oValuesSet = True
End Property

Private Static Function TransposeArray(a As Variant) As Variant
    ReDim TransposeArray(LBound(a, 2) To UBound(a, 2), LBound(a) To UBound(a))
    For i = LBound(a, 2) To UBound(a, 2)
        For j = LBound(a) To UBound(a)
            TransposeArray(i, j) = a(0, j)
        Next j
    Next i
End Function

Public Static Function HiveKeyName(hKey As HiveKey) As String
    Select Case hKey
     Case HiveKey.HKEY_CLASSES_ROOT
        HiveKeyName = "HKEY_CLASSES_ROOT"
     Case HiveKey.HKEY_CURRENT_CONFIG
        HiveKeyName = "HKEY_CURRENT_CONFIG"
     Case HiveKey.HKEY_CURRENT_USER
        HiveKeyName = "HKEY_CURRENT_USER"
     Case HiveKey.HKEY_LOCAL_MACHINE
        HiveKeyName = "HKEY_LOCAL_MACHINE"
     Case HiveKey.HKEY_USERS
        HiveKeyName = "HKEY_USERS"
     Case Else
        HiveKeyName = "Unknown_Key"
    End Select
End Function
#End If
