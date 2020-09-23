Attribute VB_Name = "MAPFILEMEM"
Option Explicit

'API DECs
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long

'MEMORY MAPPING APIs
Private Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByVal lpBaseAddress As Long) As Boolean

'STRUCTS FOR THE SAFEARRAY:
Private Type SafeBound
    cElements As Long
    lLbound As Long
End Type

Private Type SafeArray
    cDim As Integer
    fFeature As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgsabound As SafeBound
End Type

'MISC CONSTs
Private Const VT_BY_REF = &H4000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const MOVEFILE_REPLACE_EXISTING = &H1
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_BEGIN = &H0
Private Const CREATE_NEW = &H1
Private Const OPEN_EXISTING = &H3
Private Const OPEN_ALWAYS = &H4
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const PAGE_READWRITE = &H4
Private Const PAGE_READONLY = &H2
Private Const FILE_MAP_WRITE = &H2
Private Const FILE_MAP_READ = &H4
Private Const FADF_FIXEDSIZE = &H10

Public hFile As Long
Public hFileMap As Long
Public lPointer As Long

Public dataBuffer() As Byte

Public Function MapFileMemory(sFile As String, Optional Size As Long = -1, Optional mapId As String = "FileMap") As Boolean

    Dim lFileLen As Long
    Dim uTemp As SafeArray
    
    ReDim dataBuffer(0)
    
    hFile = CreateFile(sFile, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = 0 Then
       Exit Function
    End If
    hFileMap = CreateFileMapping(hFile, 0, PAGE_READWRITE, 0, 0, mapId)
    If hFileMap = 0 Then
       UnMapFileMemory
       Exit Function
    End If
    lPointer = MapViewOfFile(hFileMap, FILE_MAP_WRITE, 0, 0, 0)
    If lPointer = 0 Then
       UnMapFileMemory
       Exit Function
    End If

    If Size > -1 Then
      lFileLen = Size
    Else
      lFileLen = FileLen(sFile)                           'Find the length of the target file.
    End If
    
    If GetArrayInfo(dataBuffer, uTemp) Then                 'Load the UDT with the array info.
        uTemp.cbElements = 1                            'Set element size to a byte.
        uTemp.rgsabound.cElements = lFileLen            'Set the UBound of the array.
        uTemp.fFeature = uTemp.fFeature And FADF_FIXEDSIZE  'Set the "Fixed size" flag, SHOULD MAKE REDIM FAIL!
        uTemp.pvData = lPointer                         'Point it to the memory mapped file as it's data.
        Call AlterArray(dataBuffer, uTemp)                  'Write the UDT over the old array.
    End If
   
    MapFileMemory = True
   
End Function

Public Sub UnMapFileMemory()
    If lPointer <> 0 Then UnmapViewOfFile lPointer                 'Release the memory map.
    If hFileMap <> 0 Then CloseHandle hFileMap                     'Close the openen filemap
    If hFile <> 0 Then CloseHandle hFile                           'Close the opened file.
    lPointer = 0
    hFileMap = 0
    hFile = 0
    ReDim dataBuffer(0)
End Sub


Private Function GetArrayInfo(vArray As Variant, uInfo As SafeArray) As Boolean
    
    'NOTE, the array is passed as a variant so we can get it's absolute memory address.  This function
    'loads a copy of the SafeArray structure into the UDT.
    
    Dim lPointer As Long, iVType As Integer
    
    If Not IsArray(vArray) Then Exit Function               'Need to work with a safearray here.

    With uInfo
        CopyMemory iVType, vArray, 2                        'First 2 bytes are the subtype.
        CopyMemory lPointer, ByVal VarPtr(vArray) + 8, 4    'Get the pointer.

        If (iVType And VT_BY_REF) <> 0 Then                 'Test for subtype "pointer"
            CopyMemory lPointer, ByVal lPointer, 4          'Get the real address.
        End If
        
        CopyMemory uInfo.cDim, ByVal lPointer, 16           'Write the safearray to the passed UDT.
        
        If uInfo.cDim = 1 Then                              'Can't do multi-dimensional
            CopyMemory .rgsabound, ByVal lPointer + 16, LenB(.rgsabound)
            GetArrayInfo = True
        End If
    End With

End Function

Private Function AlterArray(vArray As Variant, uInfo As SafeArray) As Boolean
    
    'NOTE, the array is passed as a variant so we can get it's absolute memory address.  This function
    'writes the SafeArray UDT information into the actual memory address of the passed array.
    
    Dim lPointer As Long, iVType As Integer

    If Not IsArray(vArray) Then Exit Function

    With uInfo
        CopyMemory iVType, vArray, 2                        'Get the variant subtype
        CopyMemory lPointer, ByVal VarPtr(vArray) + 8, 4    'Get the pointer.

        If (iVType And VT_BY_REF) <> 0 Then                 'Test for subtype "pointer"
            CopyMemory lPointer, ByVal lPointer, 4          'Get the real address.
        End If

        CopyMemory ByVal lPointer, uInfo.cDim, 16           'Overwrite the array with the UDT.

        If uInfo.cDim = 1 Then                              'Multi-dimensions might wipe out other memory.
            CopyMemory ByVal lPointer + 16, .rgsabound, LenB(.rgsabound)
            AlterArray = True
        End If

    End With

End Function

Private Function GetPointer(vVariant As Variant) As Long

    Dim lPointer As Long, iVType As Integer

    CopyMemory iVType, vVariant, 2                          'Get the variant subtype
    CopyMemory lPointer, ByVal VarPtr(vVariant) + 8, 4      'Get the pointer
    If (iVType And VT_BY_REF) <> 0 Then                     'Test for subtype "pointer"
        CopyMemory lPointer, ByVal lPointer, 4              'Get and return the real pointer.
    End If
    
    GetPointer = lPointer

End Function

Public Function TruncateFile(ByVal sFile As String, ByVal TruncateToSize As Long) As Boolean
    Dim hpFile As Long
    hpFile = CreateFile(sFile, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hpFile = 0 Then
       Exit Function
    End If
    TruncateFile = SetFilePointer(hpFile, TruncateToSize, 0&, FILE_BEGIN)
    SetEndOfFile hpFile
    CloseHandle hpFile
End Function


