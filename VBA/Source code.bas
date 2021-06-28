Option Explicit

'==============================================================================
' Data section
'==============================================================================

'------------------------------------------------------------------------------
' Constants
'------------------------------------------------------------------------------

Private Const ValidSignature   As String = "E11AB1A1E011CFD0"
Private Const SectorMAXREGSECT As Long = &HFFFFFFFA
Private Const SectorReserved   As Long = &HFFFFFFFB
Private Const SectorDIFSECT    As Long = &HFFFFFFFC
Private Const SectorFATSECT    As Long = &HFFFFFFFD
Private Const SectorENDOFCHAIN As Long = &HFFFFFFFE
Private Const SectorFREESECT   As Long = &HFFFFFFFF
Private Const cmtModuleStandard As Integer = &H21
Private Const cmtModuleClass As Integer = &H22

'------------------------------------------------------------------------------
' Record structure types
'------------------------------------------------------------------------------

Private Enum TModuleType
    mtModuleStandard = cmtModuleStandard
    mtModuleClass = cmtModuleClass
End Enum

Private Type TModule
    ModuleName          As String ' Id = $0019
    ModuleStreamName    As String ' Id = $001A
    TextOffset          As Long ' Id = $0031
    ModuleType          As TModuleType ' Id = $0021 or $0022
    SourceCode          As String
    StreamSize          As Long
End Type

' Structure for CFB Header, according to [MS-CFB] section 2.2
Private Type TCFBHeader
    ' Header Signature (8 bytes): Identification signature for the compound file structure, and MUST be
    ' set to the value 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1.
    Signature(7) As Byte
    ' Header CLSID (16 bytes): Reserved and unused class ID that MUST be set to all zeroes
    '  (CLSID_NULL).
    HeaderCLSID(15) As Byte
    ' Minor Version (2 bytes): Version number for nonbreaking changes. This field SHOULD be set to
    '  0x003E if the major version field is either 0x0003 or 0x0004.
    MinorVersion As Integer
    ' Major Version (2 bytes): Version number for breaking changes. This field MUST be set to either
    ' 0x0003 (version 3) or 0x0004 (version 4).
    MajorVersion As Integer
    ' Byte Order (2 bytes): This field MUST be set to 0xFFFE. This field is a byte order mark for all integer
    ' fields, specifying little-endian byte order.
    ByteOrder As Integer
    ' Sector Shift (2 bytes): This field MUST be set to 0x0009, or 0x000c, depending on the Major
    ' Version field. This field specifies the sector size of the compound file as a power of 2.
    ' If Major Version is 3, the Sector Shift MUST be 0x0009, specifying a sector size of 512 bytes.
    ' If Major Version is 4, the Sector Shift MUST be 0x000C, specifying a sector size of 4096 bytes.
    SectorShift As Integer
    ' Mini Sector Shift (2 bytes): This field MUST be set to 0x0006. This field specifies the sector size of
    ' the Mini Stream as a power of 2. The sector size of the Mini Stream MUST be 64 bytes.
    MiniSectorShift As Integer
    ' Reserved (6 bytes): This field MUST be set to all zeroes.
    Reserved(5) As Byte
    ' Number of Directory Sectors (4 bytes): This integer field contains the count of the number of
    ' directory sectors in the compound file.
    ' If Major Version is 3, the Number of Directory Sectors MUST be zero. This field is not
    ' supported for version 3 compound files.
    NumberOfDirectorySectors As Long
    ' Number of FAT Sectors (4 bytes): This integer field contains the count of the number of FAT
    ' sectors in the compound file.
    NumberOfFATSectors As Long
    ' First Directory Sector Location (4 bytes): This integer field contains the starting sector number for
    ' the directory stream.
    FirstDirectorySectorLocation As Long
    ' Transaction Signature Number (4 bytes): This integer field MAY contain a sequence number that
    ' is incremented every time the compound file is saved by an implementation that supports file
    ' transactions. This is the field that MUST be set to all zeroes if file transactions are not
    ' implemented.
    TransactionSignatureNumber As Long
    ' Mini Stream Cutoff Size (4 bytes): This integer field MUST be set to 0x00001000. This field
    ' specifies the maximum size of a user-defined data stream that is allocated from the mini FAT
    ' and mini stream, and that cutoff is 4,096 bytes. Any user-defined data stream that is greater than
    ' or equal to this cutoff size must be allocated as normal sectors from the FAT.
    MiniStreamCutoffSize As Long
    ' First Mini FAT Sector Location (4 bytes): This integer field contains the starting sector number for
    ' the mini FAT.
    FirstMiniFATSectorLocation As Long
    ' Number of Mini FAT Sectors (4 bytes): This integer field contains the count of the number of mini
    '  FAT sectors in the compound file.
    NumberOfMiniFATSectors As Long
    ' First DIFAT Sector Location (4 bytes): This integer field contains the starting sector number for
    ' the DIFAT.
    FirstDIFATSectorLocation As Long
    ' Number of DIFAT Sectors (4 bytes): This integer field contains the count of the number of DIFAT
    ' sectors in the compound file.
    NumberOfDIFATSectors As Long
    ' DIFAT (436 bytes): This array of 32-bit integer fields contains the first 109 FAT sector locations of
    ' the compound file.
    DIFAT(108) As Long
End Type

' Structure for Compound File Directory Entry, according to [MS-CFB] section 2.6
Private Type TDirectoryEntry
    ' Directory Entry Name (64 bytes): This field MUST contain a Unicode string for the storage or
    '  stream name encoded in UTF-16. The name MUST be terminated with a UTF-16 terminating null
    '  character. Thus, storage and stream names are limited to 32 UTF-16 code points, including the
    '  terminating null character. When locating an object in the compound file except for the root
    '  storage, the directory entry name is compared by using a special case-insensitive uppercase
    '  mapping, described in Red-Black Tree. The following characters are illegal and MUST NOT be part
    '  of the name: '/', '\', ':', '!'. }
    DirectoryEntryName As String
    ' Directory Entry Name Length (2 bytes): This field MUST match the length of the Directory Entry
    '  Name Unicode string in bytes. The length MUST be a multiple of 2 and include the terminating null
    '  character in the count. This length MUST NOT exceed 64, the maximum size of the Directory Entry
    '  Name field.
    DirectoryEntryNameLength As Integer
    ' Object Type (1 byte): This field MUST be 0x00, 0x01, 0x02, or 0x05, depending on the actual type
    '  of object. All other values are not valid.
    '  $00 Unknown or unallocated
    '  $01 Storage Object
    '  $02 Stream Object
    '  $03 Root Storage Object
    ObjectType As Byte
    ' Color Flag (1 byte): This field MUST be 0x00 (red) or 0x01 (black). All other values are not valid.
    ColorFlag As Byte
    ' Left Sibling ID (4 bytes): This field contains the stream ID of the left sibling. If there is no left
    '  sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF).
    LeftSiblingID As Long
    ' Right Sibling ID (4 bytes): This field contains the stream ID of the right sibling. If there is no right
    '  sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF).
    RightSiblingID As Long
    ' Child ID (4 bytes): This field contains the stream ID of a child object. If there is no child object,
    '  including all entries for stream objects, the field MUST be set to NOSTREAM (0xFFFFFFFF).
    ChildID As Long
    ' CLSID (16 bytes): This field contains an object class GUID, if this entry is for a storage object or
    '  root storage object. For a stream object, this field MUST be set to all zeroes. A value containing all
    '  zeroes in a storage or root storage directory entry is valid, and indicates that no object class is
    '  associated with the storage. If an implementation of the file format enables applications to create
    '  storage objects without explicitly setting an object class GUID, it MUST write all zeroes by default.
    '  If this value is not all zeroes, the object class GUID can be used as a parameter to start
    '  applications.
    CLSID(15) As Byte
    ' State Bits (4 bytes): This field contains the user-defined flags if this entry is for a storage object or
    '  root storage object. For a stream object, this field SHOULD be set to all zeroes because many
    '  implementations provide no way for applications to retrieve state bits from a stream object. If an
    '  implementation of the file format enables applications to create storage objects without explicitly
    '  setting state bits, it MUST write all zeroes by default.
    StateBits As Long
    ' Creation Time (8 bytes): This field contains the creation time for a storage object, or all zeroes to
    '  indicate that the creation time of the storage object was not recorded. The Windows FILETIME
    '  structure is used to represent this field in UTC. For a stream object, this field MUST be all zeroes.
    '  For a root storage object, this field MUST be all zeroes, and the creation time is retrieved or set on
    '  the compound file itself.
    CreationTime As Double
    ' Modified Time (8 bytes): This field contains the modification time for a storage object, or all
    '  zeroes to indicate that the modified time of the storage object was not recorded. The Windows
    '  FILETIME structure is used to represent this field in UTC. For a stream object, this field MUST be
    '  all zeroes. For a root storage object, this field MAY<2> be set to all zeroes, and the modified time
    '  is retrieved or set on the compound file itself.
    ModifiedTime As Double
    ' Starting Sector Location (4 bytes): This field contains the first sector location if this is a stream
    '  object. For a root storage object, this field MUST contain the first sector of the mini stream, if the
    '  mini stream exists. For a storage object, this field MUST be set to all zeroes.
    StartingSectorLocation As Long
    ' Stream Size (8 bytes): This 64-bit integer field contains the size of the user-defined data if this is
    '  a stream object. For a root storage object, this field contains the size of the mini stream. For a
    '  storage object, this field MUST be set to all zeroes.
    '  For a version 3 compound file 512-byte sector size, the value of this field MUST be less than
    '  or equal to 0x80000000. (Equivalently, this requirement can be stated: the size of a stream or
    '  of the mini stream in a version 3 compound file MUST be less than or equal to 2 gigabytes
    '  (GB).) Note that as a consequence of this requirement, the most significant 32 bits of this field
    '  MUST be zero in a version 3 compound file. However, implementers should be aware that
    '  some older implementations did not initialize the most significant 32 bits of this field, and
    '  these bits might therefore be nonzero in files that are otherwise valid version 3 compound
    '  files. Although this document does not normatively specify parser behavior, it is recommended
    '  that parsers ignore the most significant 32 bits of this field in version 3 compound files,
    '  treating it as if its value were zero, unless there is a specific reason to do otherwise (for
    '  example, a parser whose purpose is to verify the correctness of a compound file).
    StreamSize As Long
    StreamSizeHigh As Long
End Type

Private Type TCFBStream
    Content() As Byte
    ContentName As String
    FullName As String
    StartingSector As Long
    StorageName As String
    StreamSize As Long
End Type

'------------------------------------------------------------------------------
' Global variables
'------------------------------------------------------------------------------

Private CFBHeader As TCFBHeader
Private CFBDiFATArray() As Long
Private CFBDiFATLength As Long
Private CFBFATArray() As Long
Private CFBFATLength As Long
Private CFBDirArray() As TDirectoryEntry
Private CFBDirLength As Long
Private CFBMiniFATArray() As Long
Private CFBMiniFATLength As Long
Private CFBMiniStream() As Byte
Private CFBStreamArray() As TCFBStream
Private CFBStreamCount As Long
Private CFBStreamLength As Long
Private SectorSize As Long ' Number of Bytes in a sector
Private SectorLength As Long ' Number of UInt32s in a sector
Private CodePage As Integer ' Actually ignored
Private ModulesArray() As TModule
Private ModulesLength As Integer

'==============================================================================
' Code section
'==============================================================================

'------------------------------------------------------------------------------
' Common functions
'------------------------------------------------------------------------------

Private Function ExtractFileExtension(ByVal FileName As String) As String
    Dim LastDotPosition As Integer
    LastDotPosition = InStrRev(FileName, ".")
    If LastDotPosition = 0 Then
        ExtractFileExtension = vbNullString
    Else
        ExtractFileExtension = Right$(FileName, Len(FileName) - LastDotPosition)
    End If
End Function

Private Function ExtractFileName(ByVal FileName As String) As String
    Dim LastSlashPosition As Integer
    LastSlashPosition = InStrRev(FileName, "\")
    If LastSlashPosition = 0 Then
        ExtractFileName = vbNullString
    Else
        ExtractFileName = Right$(FileName, Len(FileName) - LastSlashPosition)
    End If
End Function

Private Function LeftShift(ByVal Number As Long, ByVal Shifts As Integer) As Long
    Dim I As Integer
    Dim Result As Long
    Result = Number
    For I = 1 To Shifts
        Result = Result * 2
    Next
    LeftShift = Result
End Function

Private Function RightShift(ByVal Number As Long, ByVal Shifts As Integer) As Long
    Dim I As Integer
    Dim Result As Long
    Result = Number
    For I = 1 To Shifts
        Result = Result \ 2
    Next
    RightShift = Result
End Function

'------------------------------------------------------------------------------
' Microsoft Compound File Binary management functions
'------------------------------------------------------------------------------

Private Function GetStreamByName(ByVal Name As String) As Byte()
    Dim Index As Integer
    Dim NotFound As Boolean
    Index = 0
    NotFound = True
    Do While (NotFound) And (Index < CFBStreamCount)
        If CFBStreamArray(Index).ContentName = Name Then
            GetStreamByName = CFBStreamArray(Index).Content
            NotFound = False
        Else
            Index = Index + 1
        End If
    Loop
    ' Default result is zero, if not found!
End Function

Private Sub AppendStreamList(ByVal StorageName As String, ByVal NewIndex As Long)
    Dim ContentName As String
    If NewIndex <> SectorFREESECT Then
        ContentName = CFBDirArray(NewIndex).DirectoryEntryName
        Select Case CFBDirArray(NewIndex).ObjectType
            Case 1:
                AppendStreamList StorageName + "\" + ContentName, CFBDirArray(NewIndex).ChildID
                ' Recursively call AppendStreamList for the two siblings
                AppendStreamList StorageName, CFBDirArray(NewIndex).LeftSiblingID
                AppendStreamList StorageName, CFBDirArray(NewIndex).RightSiblingID
            Case 2:
                CFBStreamArray(CFBStreamCount).ContentName = ContentName
                CFBStreamArray(CFBStreamCount).StorageName = StorageName
                CFBStreamArray(CFBStreamCount).FullName = StorageName + "\" + ContentName
                CFBStreamArray(CFBStreamCount).StartingSector = CFBDirArray(NewIndex).StartingSectorLocation
                CFBStreamArray(CFBStreamCount).StreamSize = CFBDirArray(NewIndex).StreamSize
                CFBStreamCount = CFBStreamCount + 1
                ' Recursively call AppendStreamList for the two siblings
                AppendStreamList StorageName, CFBDirArray(NewIndex).LeftSiblingID
                AppendStreamList StorageName, CFBDirArray(NewIndex).RightSiblingID
        End Select
    End If
End Sub

Private Sub ReadSectorAsByte(ByVal FileChannel As Integer, ByVal SectorNumber As Long, ByRef SectorData() As Byte)
    Dim I As Long
    Seek FileChannel, (SectorNumber + 1) * SectorSize + 1
    For I = 0 To SectorSize - 1
        Get FileChannel, , SectorData(I)
    Next
End Sub

Private Sub ReadSectorAsLong(ByVal FileChannel As Integer, ByVal SectorNumber As Long, ByRef SectorData() As Long)
    Dim I As Long
    Seek FileChannel, (SectorNumber + 1) * SectorSize + 1
    For I = 0 To SectorLength - 1
        Get FileChannel, , SectorData(I)
    Next
End Sub

Private Sub ReadSectorAsDirectory(ByVal FileChannel As Integer, ByVal SectorNumber As Long, ByRef SectorData() As TDirectoryEntry)
    Dim DirectoryEntry As TDirectoryEntry
    Dim DirectoryEntryName(63) As Byte
    Dim I As Long
    Seek FileChannel, (SectorNumber + 1) * SectorSize + 1
    For I = 0 To RightShift(SectorSize, 7) - 1
        Get FileChannel, , DirectoryEntryName
        Get FileChannel, , DirectoryEntry.DirectoryEntryNameLength
        If DirectoryEntry.DirectoryEntryNameLength = 0 Then Exit For
        Get FileChannel, , DirectoryEntry.ObjectType
        Get FileChannel, , DirectoryEntry.ColorFlag
        Get FileChannel, , DirectoryEntry.LeftSiblingID
        Get FileChannel, , DirectoryEntry.RightSiblingID
        Get FileChannel, , DirectoryEntry.ChildID
        Get FileChannel, , DirectoryEntry.CLSID
        Get FileChannel, , DirectoryEntry.StateBits
        Get FileChannel, , DirectoryEntry.CreationTime
        Get FileChannel, , DirectoryEntry.ModifiedTime
        Get FileChannel, , DirectoryEntry.StartingSectorLocation
        Get FileChannel, , DirectoryEntry.StreamSize
        Get FileChannel, , DirectoryEntry.StreamSizeHigh
        DirectoryEntry.DirectoryEntryName = DirectoryEntryName
        DirectoryEntry.DirectoryEntryName = Left$(DirectoryEntry.DirectoryEntryName, RightShift(DirectoryEntry.DirectoryEntryNameLength, 1) - 1)
        SectorData(I) = DirectoryEntry
    Next
End Sub

' ReadStream: reads a chained block of sectors
Function ReadStream(ByVal FileChannel As Integer, StartingSectorLocation As Long, StreamSize As Long) As Byte()
    Dim DataBuffer()   As Byte ' Result
    Dim DataSize       As Long ' Calculated data size from FAT chain
    Dim DataRemaining  As Long ' Remaining bytes to read
    Dim DataChunkSize  As Long ' Bytes to read during the step
    Dim DataOffset     As Long ' Pointer to the current reading location
    Dim NextFATSector  As Long ' Next chain sector
    Dim SectorCount    As Long ' Calculated sector count from FAT chain
    Dim SectorSize     As Long ' Calculated sector size in Byte
    Dim SectorBuffer() As Byte ' Reading buffer for the sector
    Dim I As Long
    SectorSize = LeftShift(1, CFBHeader.SectorShift)
    If StreamSize = 0 Then
        SectorCount = 0
        NextFATSector = StartingSectorLocation
        Do While NextFATSector <> SectorENDOFCHAIN
            SectorCount = SectorCount + 1
            NextFATSector = CFBFATArray(NextFATSector)
If NextFATSector < 0 Then Exit Do
        Loop
        DataSize = LeftShift(SectorCount, CFBHeader.SectorShift)
    Else
        DataSize = StreamSize
    End If
    ReDim DataBuffer(DataSize - 1)
    ReDim SectorBuffer(SectorSize - 1)
    DataRemaining = DataSize
    NextFATSector = StartingSectorLocation
    DataOffset = 0
    Do While NextFATSector <> SectorENDOFCHAIN
        DataChunkSize = DataRemaining
        If DataChunkSize > SectorSize Then
            DataChunkSize = SectorSize
        End If
        ReadSectorAsByte FileChannel, NextFATSector, SectorBuffer
        For I = 0 To DataChunkSize - 1
            DataBuffer(DataOffset) = SectorBuffer(I)
            DataOffset = DataOffset + 1
        Next
        DataRemaining = DataRemaining - DataChunkSize
        NextFATSector = CFBFATArray(NextFATSector)
If NextFATSector < 0 Then Exit Do
    Loop
    ReadStream = DataBuffer
End Function

' ReadMiniStream: reads a chained block of mini sectors
Function ReadMiniStream(ByVal FileChannel As Integer, StartingSectorLocation As Long, StreamSize As Long) As Byte()
    Dim DataBuffer()  As Byte ' Result
    Dim DataSize      As Long ' Calculated data size from FAT chain
    Dim DataRemaining As Long ' Remaining bytes to read
    Dim DataChunkSize As Long ' Bytes to read during the step
    Dim DataOffset    As Long ' Pointer to the current reading location
    Dim NextFATSector As Long ' Next chain mini sector
    Dim SectorCount   As Long ' Calculated sector count from FAT chain
    Dim SectorSize    As Long ' Calculated sector size in Byte
    Dim SectorBuffer() As Byte ' Reading buffer for the sector
    Dim I As Long
    SectorSize = LeftShift(1, CFBHeader.MiniSectorShift)
    If StreamSize = 0 Then
        SectorCount = 0
        NextFATSector = StartingSectorLocation
        Do While NextFATSector <> SectorENDOFCHAIN
            SectorCount = SectorCount + 1
            NextFATSector = CFBMiniFATArray(NextFATSector)
        Loop
        DataSize = LeftShift(SectorCount, CFBHeader.MiniSectorShift)
    Else
        DataSize = StreamSize
    End If
    ReDim DataBuffer(DataSize - 1)
    DataRemaining = DataSize
    NextFATSector = StartingSectorLocation
    DataOffset = 0
    Do While NextFATSector <> SectorENDOFCHAIN
        DataChunkSize = DataRemaining
        If DataChunkSize > SectorSize Then
            DataChunkSize = SectorSize
        End If
        For I = 0 To DataChunkSize - 1
            DataBuffer(DataOffset) = CFBMiniStream(LeftShift(NextFATSector, CFBHeader.MiniSectorShift) + I)
            DataOffset = DataOffset + 1
        Next
        DataRemaining = DataRemaining - DataChunkSize
        NextFATSector = CFBMiniFATArray(NextFATSector)
    Loop
    ReadMiniStream = DataBuffer
End Function

Private Function ReadCFB(ByVal FileName As String) As Boolean
    Dim DataSize As Long
    Dim FileChannel As Integer
    Dim I As Long
    Dim J As Long
    Dim N As Long
    Dim NextFATSector As Long
    Dim Position As Long
    Dim SectorDataAsLong() As Long
    Dim SectorDataAsDirectory() As TDirectoryEntry
    FileChannel = FreeFile()
    Open FileName For Binary Access Read Shared As FileChannel
    ' Step 1: read and check the Compound File Header
    Seek FileChannel, 1
    Get FileChannel, , CFBHeader
    If (CFBHeader.Signature(0) = &HD0) Or (CFBHeader.Signature(1) = &HCF) _
    Or (CFBHeader.Signature(2) = &H11) Or (CFBHeader.Signature(3) = &HE0) _
    Or (CFBHeader.Signature(4) = &HA1) Or (CFBHeader.Signature(5) = &HB1) _
    Or (CFBHeader.Signature(6) = &H1A) Or (CFBHeader.Signature(7) = &HE1) Then
        If (CFBHeader.HeaderCLSID(0) = &H0) Or (CFBHeader.HeaderCLSID(1) = &H0) _
        Or (CFBHeader.HeaderCLSID(2) = &H0) Or (CFBHeader.HeaderCLSID(3) = &H0) _
        Or (CFBHeader.HeaderCLSID(4) = &H0) Or (CFBHeader.HeaderCLSID(5) = &H0) _
        Or (CFBHeader.HeaderCLSID(6) = &H0) Or (CFBHeader.HeaderCLSID(7) = &H0) _
        Or (CFBHeader.HeaderCLSID(8) = &H0) Or (CFBHeader.HeaderCLSID(9) = &H0) _
        Or (CFBHeader.HeaderCLSID(10) = &H0) Or (CFBHeader.HeaderCLSID(11) = &H0) _
        Or (CFBHeader.HeaderCLSID(12) = &H0) Or (CFBHeader.HeaderCLSID(13) = &H0) _
        Or (CFBHeader.HeaderCLSID(14) = &H0) Or (CFBHeader.HeaderCLSID(15) = &H0) Then
            SectorSize = LeftShift(1, CFBHeader.SectorShift)
            SectorLength = RightShift(SectorSize, 2)
            ReDim SectorDataAsLong(SectorLength - 1)
            ReDim SectorDataAsDirectory(RightShift(SectorSize, 7) - 1)
            ' Step 2: read the DIFAT
            N = 0
            DataSize = CFBHeader.NumberOfDIFATSectors * SectorSize
            CFBDiFATLength = RightShift(DataSize, 2) + 109
            ReDim CFBDiFATArray(CFBDiFATLength - 1)
            For I = 0 To 108
                CFBDiFATArray(N) = CFBHeader.DIFAT(I)
                N = N + 1
            Next
            NextFATSector = CFBHeader.FirstDIFATSectorLocation
            Do While NextFATSector <> SectorENDOFCHAIN
                ReadSectorAsLong FileChannel, NextFATSector, SectorDataAsLong
                For I = 0 To SectorLength - 2
                    CFBDiFATArray(N) = SectorDataAsLong(I)
                    N = N + 1
                Next
                NextFATSector = SectorDataAsLong(SectorLength - 1)
            Loop
            ' Step 3: read the FAT
            N = 0
            DataSize = CFBHeader.NumberOfFATSectors * SectorSize
            CFBFATLength = RightShift(DataSize, 2)
            ReDim CFBFATArray(CFBFATLength - 1)
            For I = 0 To CFBHeader.NumberOfFATSectors - 1
                ReadSectorAsLong FileChannel, CFBDiFATArray(I), SectorDataAsLong
                For J = 0 To SectorLength - 1
                    CFBFATArray(N) = SectorDataAsLong(J)
                    N = N + 1
                Next
            Next
            ' Step 4: read the Directory Entry Array
            N = 0
            NextFATSector = CFBHeader.FirstDirectorySectorLocation
            Do While NextFATSector <> SectorENDOFCHAIN
                N = N + 1
                NextFATSector = CFBFATArray(NextFATSector)
            Loop
            DataSize = N * SectorSize
            CFBDirLength = RightShift(DataSize, 7)
            ReDim CFBDirArray(CFBDirLength - 1)
            N = 0
            NextFATSector = CFBHeader.FirstDirectorySectorLocation
            Do While NextFATSector <> SectorENDOFCHAIN
                ReadSectorAsDirectory FileChannel, NextFATSector, SectorDataAsDirectory
                For I = 0 To RightShift(SectorSize, 7) - 1
                    CFBDirArray(N) = SectorDataAsDirectory(I)
                    N = N + 1
                Next
                NextFATSector = CFBFATArray(NextFATSector)
            Loop
            ReDim Preserve CFBDirArray(CFBDirLength - 1)
            ' Step 5: read the mini FAT
            N = 0
            DataSize = CFBHeader.NumberOfMiniFATSectors * SectorSize
            CFBMiniFATLength = RightShift(DataSize, 2)
            ReDim CFBMiniFATArray(CFBMiniFATLength - 1)
            NextFATSector = CFBHeader.FirstMiniFATSectorLocation
            Do While NextFATSector <> SectorENDOFCHAIN
                ReadSectorAsLong FileChannel, NextFATSector, SectorDataAsLong
                For J = 0 To SectorLength - 1
                    CFBMiniFATArray(N) = SectorDataAsLong(J)
                    N = N + 1
                Next
                NextFATSector = CFBFATArray(NextFATSector)
            Loop
            ' Step 6: read the mini stream
            CFBMiniStream = ReadStream(FileChannel, CFBDirArray(0).StartingSectorLocation, 0)
            ' Step 7: build the stream list
            ' Initialize the array of streams
            CFBStreamCount = 0
            CFBStreamLength = 0
            For I = 0 To CFBDirLength - 1
                If CFBDirArray(I).ObjectType = 2 Then
                    CFBStreamLength = CFBStreamLength + 1
                End If
            Next
            ReDim CFBStreamArray(CFBStreamLength - 1)
            ' Start with the root entry, use AppendStreamList, to add the stream objects recursively
            AppendStreamList "", CFBDirArray(0).ChildID
            ' Read the streams
            For I = 0 To CFBStreamLength - 1
                If (InStr(CFBStreamArray(I).StorageName, "VBA") > 0) And (CFBStreamArray(I).StreamSize > 0) Then
                    If CFBStreamArray(I).StreamSize <= CFBHeader.MiniStreamCutoffSize Then
                        CFBStreamArray(I).Content = ReadMiniStream(FileChannel, CFBStreamArray(I).StartingSector, CFBStreamArray(I).StreamSize)
                    Else
                        CFBStreamArray(I).Content = ReadStream(FileChannel, CFBStreamArray(I).StartingSector, CFBStreamArray(I).StreamSize)
                    End If
                End If
            Next
            ReadCFB = True
        Else
            ReadCFB = False
        End If
    Else
        ReadCFB = False
    End If
    Close FileChannel
End Function

'------------------------------------------------------------------------------
' Microsoft VBA streams management functions
'------------------------------------------------------------------------------

Private Sub DecompressContainer(ByRef CompressedContainer() As Byte, ByRef CompressedStart As Long, ByRef DecompressedData() As Byte)
    Dim DecompressedIndex  As Long
    Dim DecompressedLength As Long
    Dim ChunkHeader        As Long
    Dim ChunkSignature     As Long
    Dim ChunkFlag          As Long
    Dim ChunkSize          As Long
    Dim ChunkEnd           As Long
    Dim BitFlags           As Byte
    Dim Token              As Long
    Dim BitCount           As Long
    Dim BitMask            As Long
    Dim CopyLength         As Long
    Dim CopyOffset         As Long
    Dim I                  As Long
    Dim J                  As Long
    ' Use an array to speedup calculations
    Dim PowerOf2(0 To 16)  As Long
    PowerOf2(0) = 1
    For I = 1 To UBound(PowerOf2)
        PowerOf2(I) = PowerOf2(I - 1) * 2
    Next
    Do
        If Not (Not DecompressedData) Then
            ReDim Preserve DecompressedData(UBound(DecompressedData) + 4096)
        Else
            ReDim DecompressedData(0 To 4095)
            DecompressedIndex = 0
        End If
        ChunkHeader = CompressedContainer(CompressedStart) + 256& * CompressedContainer(CompressedStart + 1)
        CompressedStart = CompressedStart + 2
        ChunkSize = (ChunkHeader And &HFFF)
        ChunkEnd = CompressedStart + ChunkSize
        ChunkSignature = (ChunkHeader And &H7000) \ &H1000&
        ChunkFlag = (ChunkHeader And &H8000) \ &H8000&
        If ChunkFlag = 0 Then
            For J = 0 To 4095
                DecompressedData(DecompressedIndex + J) = CompressedContainer(CompressedStart + J)
            Next J
            CompressedStart = CompressedStart + 4096
            DecompressedIndex = DecompressedIndex + 4096
        Else
            Do
                BitFlags = CompressedContainer(CompressedStart)
                CompressedStart = CompressedStart + 1
                For I = 0 To 7
                    If CompressedStart > ChunkEnd Then Exit Do
                    If (BitFlags And PowerOf2(I)) = 0 Then
                        DecompressedData(DecompressedIndex) = CompressedContainer(CompressedStart)
                        CompressedStart = CompressedStart + 1
                        DecompressedIndex = DecompressedIndex + 1
                    Else
                        Token = CompressedContainer(CompressedStart) + CompressedContainer(CompressedStart + 1) * 256&
                        CompressedStart = CompressedStart + 2
                        DecompressedLength = DecompressedIndex Mod 4096
                        For BitCount = 4 To 11
                            If DecompressedLength <= PowerOf2(BitCount) Then Exit For
                        Next BitCount
                        BitMask = PowerOf2(16) - PowerOf2(16 - BitCount)
                        CopyOffset = (Token And BitMask) \ PowerOf2(16 - BitCount) + 1
                        BitMask = PowerOf2(16 - BitCount) - 1
                        CopyLength = (Token And BitMask) + 3
                        For J = 0 To CopyLength - 1
                            DecompressedData(DecompressedIndex + J) = DecompressedData(DecompressedIndex - CopyOffset + J)
                        Next J
                        DecompressedIndex = DecompressedIndex + CopyLength
                    End If
                Next
            Loop
        End If
        If CompressedStart > UBound(CompressedContainer) Then Exit Do
    Loop
    ReDim Preserve DecompressedData(0 To DecompressedIndex - 1)
End Sub

Private Function ReadBYTE(ByRef DecompressedStream() As Byte, ByRef Offset As Long) As Byte
    ReadBYTE = DecompressedStream(Offset)
    Offset = Offset + 1
End Function

Private Function ReadWORD(ByRef DecompressedStream() As Byte, ByRef Offset As Long) As Long
    Dim Half1 As Long
    Dim Half2 As Long
    Half1 = ReadBYTE(DecompressedStream, Offset)
    Half2 = ReadBYTE(DecompressedStream, Offset)
    ReadWORD = Half1 + Half2 * 256
End Function

Private Function ReadDWORD(ByRef DecompressedStream() As Byte, ByRef Offset As Long) As Long
    Dim Half1 As Long
    Dim Half2 As Long
    Half1 = ReadWORD(DecompressedStream, Offset)
    Half2 = ReadWORD(DecompressedStream, Offset)
    ReadDWORD = Half1 + Half2 * 65536
End Function

Private Function ReadString(ByRef DecompressedStream() As Byte, ByRef Offset As Long, ByVal NumberOfBytes As Long) As String
    Dim Result As String
    Dim StringData() As Byte
    Dim StringIndex As Long
    Result = vbNullString
    If NumberOfBytes > 0 Then
        ReDim StringData(NumberOfBytes - 1)
        For StringIndex = 0 To NumberOfBytes - 1
            StringData(StringIndex) = DecompressedStream(Offset)
            Offset = Offset + 1
        Next
        Result = StrConv(StringData, vbUnicode)
    End If
    ReadString = Result
End Function

Private Function ParseDirStream() As Boolean
    Dim CompressedStream() As Byte
    Dim DecompressedStream() As Byte
    Dim ModuleIndex As Integer
    Dim RecordId As Integer
    Dim RecordSize As Long
    Dim StreamEnd As Long
    Dim StreamOffset As Long
    CompressedStream = GetStreamByName("dir")
    If Not (Not CompressedStream) Then
        DecompressContainer CompressedStream, 1, DecompressedStream
        CodePage = 0
        StreamOffset = 0
        StreamEnd = UBound(DecompressedStream)
        Do While StreamOffset <= StreamEnd
            ' Step 1: get the RecordId
            RecordId = ReadWORD(DecompressedStream, StreamOffset)
            ' Step 2: get the RecordId
            RecordSize = ReadDWORD(DecompressedStream, StreamOffset)
            ' Step 3: parse the RecordId field
            Select Case RecordId
                Case &H1:    ' SysKindRecord
                    StreamOffset = StreamOffset + RecordSize
                Case &H2:    ' LcidRecord
                    StreamOffset = StreamOffset + RecordSize
                Case &H14:   ' LcidInvokeRecord
                    StreamOffset = StreamOffset + RecordSize
                Case &H3:    ' CodePageRecord
                    CodePage = ReadWORD(DecompressedStream, StreamOffset)
                Case &H4:    ' NameRecord
                    StreamOffset = StreamOffset + RecordSize
                Case &H5:    ' DocStringRecord
                    StreamOffset = StreamOffset + RecordSize
                    StreamOffset = StreamOffset + 2
                    RecordSize = ReadDWORD(DecompressedStream, StreamOffset)
                    StreamOffset = StreamOffset + RecordSize
                Case &H6:    ' HelpFilePathRecord 1
                    StreamOffset = StreamOffset + RecordSize
                Case &H3D:   ' HelpFilePathRecord 2
                    StreamOffset = StreamOffset + RecordSize
                Case &H7:    ' HelpContextRecord
                    StreamOffset = StreamOffset + RecordSize
                Case &H8:    ' LibFlagsRecord
                    StreamOffset = StreamOffset + RecordSize
                Case &H9:    ' VersionRecord
                    StreamOffset = StreamOffset + 6
                Case &HC:    ' ConstantsRecord
                    StreamOffset = StreamOffset + RecordSize
                    StreamOffset = StreamOffset + 2
                    RecordSize = ReadDWORD(DecompressedStream, StreamOffset)
                    StreamOffset = StreamOffset + RecordSize
                Case &H16:   ' Reference name
                    StreamOffset = StreamOffset + RecordSize
                    StreamOffset = StreamOffset + 2
                    RecordSize = ReadDWORD(DecompressedStream, StreamOffset)
                    StreamOffset = StreamOffset + RecordSize
                Case &H2F:   ' Reference Control
                    RecordSize = ReadDWORD(DecompressedStream, StreamOffset)
                    StreamOffset = StreamOffset + RecordSize
                    StreamOffset = StreamOffset + 6
                Case &H30:   ' Reference extended
                    RecordSize = ReadDWORD(DecompressedStream, StreamOffset)
                    StreamOffset = StreamOffset + RecordSize
                    StreamOffset = StreamOffset + 6
                    StreamOffset = StreamOffset + 16
                    StreamOffset = StreamOffset + 4
                Case &H33:   ' Reference Original
                    StreamOffset = StreamOffset + RecordSize
                Case &HD:    ' Reference Registered
                    RecordSize = ReadDWORD(DecompressedStream, StreamOffset)
                    StreamOffset = StreamOffset + RecordSize
                    StreamOffset = StreamOffset + 6
                Case &HE:    ' Reference Project
                    RecordSize = ReadDWORD(DecompressedStream, StreamOffset)
                    StreamOffset = StreamOffset + RecordSize
                    RecordSize = ReadDWORD(DecompressedStream, StreamOffset)
                    StreamOffset = StreamOffset + RecordSize
                    StreamOffset = StreamOffset + 6
                Case &HF:    ' Modules
                    ModulesLength = ReadWORD(DecompressedStream, StreamOffset)
                    ReDim ModulesArray(ModulesLength - 1)
                    StreamOffset = StreamOffset + 8
                    ModuleIndex = -1
                Case &H19:   ' Module name
                    ModuleIndex = ModuleIndex + 1
                    ModulesArray(ModuleIndex).ModuleName = ReadString(DecompressedStream, StreamOffset, RecordSize)
                Case &H47:   ' Module name in Unicode
                    StreamOffset = StreamOffset + RecordSize
                Case &H1A:   ' Module stream name
                    ModulesArray(ModuleIndex).ModuleStreamName = ReadString(DecompressedStream, StreamOffset, RecordSize)
                Case &H32:   ' Module stream name in Unicode
                    StreamOffset = StreamOffset + RecordSize
                Case &H1C:   ' Module doc string
                    StreamOffset = StreamOffset + RecordSize
                Case &H48:   ' Module doc string in Unicode
                    StreamOffset = StreamOffset + RecordSize
                Case &H31:   ' Module offset
                    ModulesArray(ModuleIndex).TextOffset = ReadDWORD(DecompressedStream, StreamOffset)
                Case &H1E:   ' Module help context
                    StreamOffset = StreamOffset + RecordSize
                Case &H2C:   ' Module cookie
                    StreamOffset = StreamOffset + RecordSize
                Case &H21:   ' Module type standard
                Case &H22:   ' Module type class
                Case &H25:   ' Module read only
                Case &H28:   ' Module private
                Case &H10:   ' Terminator of dir stream
                Case &H2B:   ' Terminator of Module record
                Case &H4A:   ' Mystery field?
                    StreamOffset = StreamOffset + 4
            End Select
        Loop
        ParseDirStream = True
    Else
        ParseDirStream = False
    End If
End Function

Private Sub ParseModuleStream()
    Dim CompressedStream() As Byte
    Dim DecompressedStream() As Byte
    Dim SourceCode As String
    Dim SourceLine As String
    Dim SourceLines() As String
    Dim I As Long
    Dim J As Long
    For I = 0 To ModulesLength - 1
        CompressedStream = GetStreamByName(ModulesArray(I).ModuleName)
        DecompressContainer CompressedStream, ModulesArray(I).TextOffset + 1, DecompressedStream
        SourceCode = StrConv(DecompressedStream, vbUnicode)
        SourceLines = Split(SourceCode, vbNewLine)
        SourceCode = ""
        For J = LBound(SourceLines) To UBound(SourceLines)
            SourceLine = SourceLines(J)
            If Left$(SourceLine, 9) <> "Attribute" Then
                SourceCode = SourceCode & SourceLine & vbNewLine
            End If
        Next
        ModulesArray(I).SourceCode = SourceCode
    Next
End Sub

Private Function ParseFile(ByVal FileName As String) As Boolean
    Dim FileExtension As String
    FileExtension = UCase(ExtractFileExtension(FileName))
    If (FileExtension = "BIN") Or (FileExtension = "DOC") _
    Or (FileExtension = "OTM") Or (FileExtension = "XLS") Then
        If ReadCFB(FileName) Then
            ' Step 1: parse the 'dir' stream
            If ParseDirStream() Then
                ' Step 2: read the modules streams
                ParseModuleStream
                ParseFile = True
            Else
                ParseFile = False
            End If
        Else
            ParseFile = False
        End If
    Else
        ParseFile = False
    End If
End Function

'------------------------------------------------------------------------------
' Gui management
'------------------------------------------------------------------------------

Private Sub Reset()
    TextBoxFileName.Text = ""
    ListBoxModules.Clear
    TextBoxSourceCode.Text = ""
End Sub

Private Sub Update()
    Dim I As Integer
    For I = 0 To ModulesLength - 1
        ListBoxModules.AddItem ModulesArray(I).ModuleName
    Next
End Sub

Private Sub CommandButtonOpen_Click()
    Dim SelectFileDialog As FileDialog
    Dim SelectedFileName As String
    Reset
    Set SelectFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    SelectFileDialog.AllowMultiSelect = False
    SelectFileDialog.Filters.Clear
    SelectFileDialog.Filters.Add "All supported files", "*.bin; *.doc; *.otm; *.xls"
    SelectFileDialog.Filters.Add "Microsoft Excel", "*.xls"
    SelectFileDialog.Filters.Add "Microsoft Word", "*.doc"
    SelectFileDialog.Filters.Add "Microsoft Outlook", "*.otm"
    SelectFileDialog.Filters.Add "Microsoft VBA project from XML documents", "*.bin"
    SelectFileDialog.Title = "Choose an Office file"
    If SelectFileDialog.Show() Then
        SelectedFileName = SelectFileDialog.SelectedItems.Item(1)
        If ParseFile(SelectedFileName) Then
            TextBoxFileName.Text = SelectedFileName
            Update
        End If
    End If
End Sub

Private Sub ListBoxModules_Click()
    TextBoxSourceCode.Text = ModulesArray(ListBoxModules.ListIndex).SourceCode
    Repaint
End Sub

Private Sub UserForm_Terminate()
    ' Safety: close all opened files
    Close
End Sub
