<div align="center">

## List a File's Dependencies


</div>

### Description

Get a list of a files dependencies. For example, calling ListDependencies "C:\WINDOWS\SYSTEM32\MSVBVM60.DLL", results in "KERNEL32.dll USER32.dll GDI32.dll ADVAPI32.dll ole32.dll OLEAUT32.dll".
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[syntax\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/syntax.md)
**Level**          |Intermediate
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/syntax-list-a-file-s-dependencies__1-55638/archive/master.zip)

### API Declarations

```
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function ImageDirectoryEntryToData Lib "imagehlp" (ByVal Base As Long, ByVal MappedAsImage As Byte, ByVal DirectoryEntry As Integer, Size As Long) As Long
Public Declare Function ImageRvaToVa Lib "imagehlp" (NtHeaders As Any, ByVal Base As Long, ByVal RVA As Long, Optional LastRvaSection As Long) As Long
Public Declare Function MapAndLoad Lib "imagehlp" (ByVal ImageName As String, ByVal DllPath As String, LoadedImage As LOADED_IMAGE, ByVal DotDll As Long, ByVal ReadOnly As Long) As Long
Public Declare Function UnMapAndLoad Lib "imagehlp" (LoadedImage As LOADED_IMAGE) As Long
Public Declare Function CheckSumMappedFile Lib "imagehlp" (ByVal BaseAddress As Long, ByVal FileLength As Long, HeaderSum As Long, CheckSum As Long) As Long
Public Declare Function UnDecorateSymbolName Lib "imagehlp" (ByVal DecoratedName As String, ByVal UnDecoratedName As String, ByVal UndecoratedLength As Long, ByVal Flags As Long) As Long
Public Declare Function lstrlenA Lib "kernel32" (ByVal pString As Any) As Long
Public Declare Function lstrcpyA Lib "kernel32" (ByVal pString1 As Any, ByVal pString2 As Any) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal L As Long)
Public Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16
Public Type IMAGE_IMPORT_DESCRIPTOR
  OriginalFirstThunk As Long
  TimeDateStamp As Long
  ForwarderChain As Long
  Name As Long
  FirstThunk As Long
End Type
Public Type IMAGE_DATA_DIRECTORY
  VirtualAddress As Long
  Size As Long
End Type
Public Type LIST_ENTRY
  Flk As Long
  Blk As Long
End Type
Public Type LOADED_IMAGE
  ModuleName As Long
  hFile As Long
  MappedAddress As Long
  FileHeader As Long
  LastRvaSection As Long
  NumberOfSections As Long
  Sections As Long
  Characteristics As Long
  fSystemImage As Byte
  fDOSImage As Byte
  Lks As LIST_ENTRY
  SizeOfImage As Long
End Type
Public Type IMAGE_OPTIONAL_HEADER
  Magic As Integer
  MajorLinkerVersion As Byte
  MinorLinkerVersion As Byte
  SizeOfCode As Long
  SizeOfInitializedData As Long
  SizeOfUnitializedData As Long
  AddressOfEntryPot As Long
  BaseOfCode As Long
  BaseOfData As Long
  ImageBase As Long
  SectionAlignment As Long
  FileAlignment As Long
  MajorOperatingSystemVersion As Integer
  MinorOperatingSystemVersion As Integer
  MajorImageVersion As Integer
  MinorImageVersion As Integer
  MajorSubsystemVersion As Integer
  MinorSubsystemVersion As Integer
  W32VersionValue As Long
  SizeOfImage As Long
  SizeOfHeaders As Long
  CheckSum As Long
  SubSystem As Integer
  DllCharacteristics As Integer
  SizeOfStackReserve As Long
  SizeOfStackCommit As Long
  SizeOfHeapReserve As Long
  SizeOfHeapCommit As Long
  LoaderFlags As Long
  NumberOfRvaAndSizes As Long
  DataDirectory(0 To IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) As IMAGE_DATA_DIRECTORY
End Type
Public Type IMAGE_FILE_HEADER
  Machine As Integer
  NumberOfSections As Integer
  TimeDateStamp As Long
  PointerToSymbolTable As Long
  NumberOfSymbols As Long
  SizeOfOptionalHeader As Integer
  Characteristics As Integer
End Type
Public Type IMAGE_NT_HEADERS
  Signature As Long
  FileHeader As IMAGE_FILE_HEADER
  OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type
```


### Source Code

```
'Lists the dependencies of sFile
Public Function ListDependencies(sFile As String)
  If (Dir(sFile, vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "") Then
    Dim lRet As Long
    Dim m_LI As LOADED_IMAGE
    Dim m_NTHdr As IMAGE_NT_HEADERS
    Dim sWholeFile As String
    sWholeFile = sFile
    DoEvents
    If MapAndLoad(sWholeFile, vbNullString, m_LI, True, True) Then
      MoveMemory m_NTHdr, ByVal m_LI.FileHeader, Len(m_NTHdr)
      Dim aModules() As String
      Dim lNumModules As Long
      Dim ImpDir As IMAGE_IMPORT_DESCRIPTOR
      Dim lNamePtr As Long, lIdx As Long
      Dim lImpPtr As Long, lSize As Long
        lNumModules = 0
        Erase aModules
        lImpPtr = ImageDirectoryEntryToData(m_LI.MappedAddress, 0, 1, lSize)
        If lImpPtr Then
         MoveMemory ImpDir, ByVal lImpPtr, Len(ImpDir)
         Do Until ImpDir.Name = 0
           ReDim Preserve aModules(0 To lIdx)
           aModules(lIdx) = sStringFromRVA(ImpDir.Name, m_LI)
           lIdx = lIdx + 1
           MoveMemory ImpDir, ByVal lImpPtr + (Len(ImpDir) * lIdx), Len(ImpDir)
           DoEvents
         Loop
         lNumModules = lIdx
        End If
      For lIdx = 0 To lNumModules - 1
        Debug.Print aModules(lIdx)
      Next lIdx
      UnMapAndLoad m_LI
    End If
  Else
    Debug.Print "File " & sFile & " Doesn't Exist"
  End If
End Function
'sStringFromRVA(ByVal RVA As Long) As String
'This will grab a string located at a given RVA
Private Function sStringFromRVA(ByVal RVA As Long, m_LI As LOADED_IMAGE) As String
  Dim lVA As Long
  lVA = ImageRvaToVa(ByVal m_LI.FileHeader, m_LI.MappedAddress, RVA)
  sStringFromRVA = String$(lstrlenA(lVA) + 1, 0)
  lstrcpyA sStringFromRVA, lVA
  If InStr(sStringFromRVA, vbNullChar) Then
    sStringFromRVA = Left$(sStringFromRVA, InStr(sStringFromRVA, vbNullChar) - 1)
  End If
End Function
```

