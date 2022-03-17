object FormVBASettings: TFormVBASettings
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'FormVBASettings'
  ClientHeight = 359
  ClientWidth = 548
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object LabelSynEdit: TLabel
    Left = 8
    Top = 8
    Width = 127
    Height = 13
    Caption = 'Syntax Highlighting Theme'
  end
  object LabelFont: TLabel
    Left = 375
    Top = 8
    Width = 81
    Height = 13
    Caption = 'Code Editor Font'
  end
  object LabelSize: TLabel
    Left = 491
    Top = 8
    Width = 19
    Height = 13
    Caption = 'Size'
  end
  object ButtonApply: TButton
    Left = 271
    Top = 25
    Width = 98
    Height = 25
    Caption = 'Apply'
    TabOrder = 0
  end
  object ComboBoxFont: TComboBox
    Left = 375
    Top = 27
    Width = 110
    Height = 21
    Style = csDropDownList
    TabOrder = 1
  end
  object ComboBoxTheme: TComboBox
    Left = 11
    Top = 27
    Width = 254
    Height = 21
    Style = csDropDownList
    TabOrder = 2
  end
  object EditFontSize: TEdit
    Left = 491
    Top = 27
    Width = 32
    Height = 21
    ReadOnly = True
    TabOrder = 3
    Text = '10'
    OnChange = EditFontSizeChange
  end
  object UpDownSize: TUpDown
    Left = 523
    Top = 27
    Width = 16
    Height = 21
    Associate = EditFontSize
    Min = 8
    Max = 48
    Position = 10
    TabOrder = 4
  end
  object SynEditVB: TSynEdit
    Left = 8
    Top = 56
    Width = 531
    Height = 297
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Courier New'
    Font.Style = []
    TabOrder = 5
    CodeFolding.GutterShapeSize = 11
    CodeFolding.CollapsedLineColor = clGrayText
    CodeFolding.FolderBarLinesColor = clGrayText
    CodeFolding.IndentGuidesColor = clGray
    CodeFolding.IndentGuides = True
    CodeFolding.ShowCollapsedLine = False
    CodeFolding.ShowHintMark = True
    UseCodeFolding = False
    Gutter.Font.Charset = DEFAULT_CHARSET
    Gutter.Font.Color = clWindowText
    Gutter.Font.Height = -11
    Gutter.Font.Name = 'Courier New'
    Gutter.Font.Style = []
    Gutter.ShowLineNumbers = True
    Highlighter = SynVBSyn
    Lines.Strings = (
      'Option Explicit'
      ''
      
        #39'===============================================================' +
        '==============='
      #39' Data section'
      
        #39'===============================================================' +
        '==============='
      ''
      
        #39'---------------------------------------------------------------' +
        '---------------'
      #39' Constants'
      
        #39'---------------------------------------------------------------' +
        '---------------'
      ''
      'Private Const ValidSignature   As String = "E11AB1A1E011CFD0"'
      'Private Const SectorMAXREGSECT As Long = &HFFFFFFFA'
      'Private Const SectorReserved   As Long = &HFFFFFFFB'
      'Private Const SectorDIFSECT    As Long = &HFFFFFFFC'
      'Private Const SectorFATSECT    As Long = &HFFFFFFFD'
      'Private Const SectorENDOFCHAIN As Long = &HFFFFFFFE'
      'Private Const SectorFREESECT   As Long = &HFFFFFFFF'
      'Private Const cmtModuleStandard As Integer = &H21'
      'Private Const cmtModuleClass As Integer = &H22'
      ''
      
        #39'---------------------------------------------------------------' +
        '---------------'
      #39' Record structure types'
      
        #39'---------------------------------------------------------------' +
        '---------------'
      ''
      'Private Enum TModuleType'
      '    mtModuleStandard = cmtModuleStandard'
      '    mtModuleClass = cmtModuleClass'
      'End Enum'
      ''
      'Private Type TModule'
      '    ModuleName          As String '#39' Id = $0019'
      '    ModuleStreamName    As String '#39' Id = $001A'
      '    TextOffset          As Long '#39' Id = $0031'
      '    ModuleType          As TModuleType '#39' Id = $0021 or $0022'
      '    SourceCode          As String'
      '    StreamSize          As Long'
      'End Type'
      ''
      #39' Structure for CFB Header, according to [MS-CFB] section 2.2'
      'Private Type TCFBHeader'
      
        '    '#39' Header Signature (8 bytes): Identification signature for t' +
        'he compound file structure, and MUST be'
      
        '    '#39' set to the value 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A,' +
        ' 0xE1.'
      '    Signature(7) As Byte'
      
        '    '#39' Header CLSID (16 bytes): Reserved and unused class ID that' +
        ' MUST be set to all zeroes'
      '    '#39'  (CLSID_NULL).'
      '    HeaderCLSID(15) As Byte'
      
        '    '#39' Minor Version (2 bytes): Version number for nonbreaking ch' +
        'anges. This field SHOULD be set to'
      
        '    '#39'  0x003E if the major version field is either 0x0003 or 0x0' +
        '004.'
      '    MinorVersion As Integer'
      
        '    '#39' Major Version (2 bytes): Version number for breaking chang' +
        'es. This field MUST be set to either'
      '    '#39' 0x0003 (version 3) or 0x0004 (version 4).'
      '    MajorVersion As Integer'
      
        '    '#39' Byte Order (2 bytes): This field MUST be set to 0xFFFE. Th' +
        'is field is a byte order mark for all integer'
      '    '#39' fields, specifying little-endian byte order.'
      '    ByteOrder As Integer'
      
        '    '#39' Sector Shift (2 bytes): This field MUST be set to 0x0009, ' +
        'or 0x000c, depending on the Major'
      
        '    '#39' Version field. This field specifies the sector size of the' +
        ' compound file as a power of 2.'
      
        '    '#39' If Major Version is 3, the Sector Shift MUST be 0x0009, sp' +
        'ecifying a sector size of 512 bytes.'
      
        '    '#39' If Major Version is 4, the Sector Shift MUST be 0x000C, sp' +
        'ecifying a sector size of 4096 bytes.'
      '    SectorShift As Integer'
      
        '    '#39' Mini Sector Shift (2 bytes): This field MUST be set to 0x0' +
        '006. This field specifies the sector size of'
      
        '    '#39' the Mini Stream as a power of 2. The sector size of the Mi' +
        'ni Stream MUST be 64 bytes.'
      '    MiniSectorShift As Integer'
      '    '#39' Reserved (6 bytes): This field MUST be set to all zeroes.'
      '    Reserved(5) As Byte'
      
        '    '#39' Number of Directory Sectors (4 bytes): This integer field ' +
        'contains the count of the number of'
      '    '#39' directory sectors in the compound file.'
      
        '    '#39' If Major Version is 3, the Number of Directory Sectors MUS' +
        'T be zero. This field is not'
      '    '#39' supported for version 3 compound files.'
      '    NumberOfDirectorySectors As Long'
      
        '    '#39' Number of FAT Sectors (4 bytes): This integer field contai' +
        'ns the count of the number of FAT'
      '    '#39' sectors in the compound file.'
      '    NumberOfFATSectors As Long'
      
        '    '#39' First Directory Sector Location (4 bytes): This integer fi' +
        'eld contains the starting sector number for'
      '    '#39' the directory stream.'
      '    FirstDirectorySectorLocation As Long'
      
        '    '#39' Transaction Signature Number (4 bytes): This integer field' +
        ' MAY contain a sequence number that'
      
        '    '#39' is incremented every time the compound file is saved by an' +
        ' implementation that supports file'
      
        '    '#39' transactions. This is the field that MUST be set to all ze' +
        'roes if file transactions are not'
      '    '#39' implemented.'
      '    TransactionSignatureNumber As Long'
      
        '    '#39' Mini Stream Cutoff Size (4 bytes): This integer field MUST' +
        ' be set to 0x00001000. This field'
      
        '    '#39' specifies the maximum size of a user-defined data stream t' +
        'hat is allocated from the mini FAT'
      
        '    '#39' and mini stream, and that cutoff is 4,096 bytes. Any user-' +
        'defined data stream that is greater than'
      
        '    '#39' or equal to this cutoff size must be allocated as normal s' +
        'ectors from the FAT.'
      '    MiniStreamCutoffSize As Long'
      
        '    '#39' First Mini FAT Sector Location (4 bytes): This integer fie' +
        'ld contains the starting sector number for'
      '    '#39' the mini FAT.'
      '    FirstMiniFATSectorLocation As Long'
      
        '    '#39' Number of Mini FAT Sectors (4 bytes): This integer field c' +
        'ontains the count of the number of mini'
      '    '#39'  FAT sectors in the compound file.'
      '    NumberOfMiniFATSectors As Long'
      
        '    '#39' First DIFAT Sector Location (4 bytes): This integer field ' +
        'contains the starting sector number for'
      '    '#39' the DIFAT.'
      '    FirstDIFATSectorLocation As Long'
      
        '    '#39' Number of DIFAT Sectors (4 bytes): This integer field cont' +
        'ains the count of the number of DIFAT'
      '    '#39' sectors in the compound file.'
      '    NumberOfDIFATSectors As Long'
      
        '    '#39' DIFAT (436 bytes): This array of 32-bit integer fields con' +
        'tains the first 109 FAT sector locations of'
      '    '#39' the compound file.'
      '    DIFAT(108) As Long'
      'End Type'
      ''
      
        #39' Structure for Compound File Directory Entry, according to [MS-' +
        'CFB] section 2.6'
      'Private Type TDirectoryEntry'
      
        '    '#39' Directory Entry Name (64 bytes): This field MUST contain a' +
        ' Unicode string for the storage or'
      
        '    '#39'  stream name encoded in UTF-16. The name MUST be terminate' +
        'd with a UTF-16 terminating null'
      
        '    '#39'  character. Thus, storage and stream names are limited to ' +
        '32 UTF-16 code points, including the'
      
        '    '#39'  terminating null character. When locating an object in th' +
        'e compound file except for the root'
      
        '    '#39'  storage, the directory entry name is compared by using a ' +
        'special case-insensitive uppercase'
      
        '    '#39'  mapping, described in Red-Black Tree. The following chara' +
        'cters are illegal and MUST NOT be part'
      '    '#39'  of the name: '#39'/'#39', '#39'\'#39', '#39':'#39', '#39'!'#39'. }'
      '    DirectoryEntryName As String'
      
        '    '#39' Directory Entry Name Length (2 bytes): This field MUST mat' +
        'ch the length of the Directory Entry'
      
        '    '#39'  Name Unicode string in bytes. The length MUST be a multip' +
        'le of 2 and include the terminating null'
      
        '    '#39'  character in the count. This length MUST NOT exceed 64, t' +
        'he maximum size of the Directory Entry'
      '    '#39'  Name field.'
      '    DirectoryEntryNameLength As Integer'
      
        '    '#39' Object Type (1 byte): This field MUST be 0x00, 0x01, 0x02,' +
        ' or 0x05, depending on the actual type'
      '    '#39'  of object. All other values are not valid.'
      '    '#39'  $00 Unknown or unallocated'
      '    '#39'  $01 Storage Object'
      '    '#39'  $02 Stream Object'
      '    '#39'  $03 Root Storage Object'
      '    ObjectType As Byte'
      
        '    '#39' Color Flag (1 byte): This field MUST be 0x00 (red) or 0x01' +
        ' (black). All other values are not valid.'
      '    ColorFlag As Byte'
      
        '    '#39' Left Sibling ID (4 bytes): This field contains the stream ' +
        'ID of the left sibling. If there is no left'
      '    '#39'  sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF).'
      '    LeftSiblingID As Long'
      
        '    '#39' Right Sibling ID (4 bytes): This field contains the stream' +
        ' ID of the right sibling. If there is no right'
      '    '#39'  sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF).'
      '    RightSiblingID As Long'
      
        '    '#39' Child ID (4 bytes): This field contains the stream ID of a' +
        ' child object. If there is no child object,'
      
        '    '#39'  including all entries for stream objects, the field MUST ' +
        'be set to NOSTREAM (0xFFFFFFFF).'
      '    ChildID As Long'
      
        '    '#39' CLSID (16 bytes): This field contains an object class GUID' +
        ', if this entry is for a storage object or'
      
        '    '#39'  root storage object. For a stream object, this field MUST' +
        ' be set to all zeroes. A value containing all'
      
        '    '#39'  zeroes in a storage or root storage directory entry is va' +
        'lid, and indicates that no object class is'
      
        '    '#39'  associated with the storage. If an implementation of the ' +
        'file format enables applications to create'
      
        '    '#39'  storage objects without explicitly setting an object clas' +
        's GUID, it MUST write all zeroes by default.'
      
        '    '#39'  If this value is not all zeroes, the object class GUID ca' +
        'n be used as a parameter to start'
      '    '#39'  applications.'
      '    CLSID(15) As Byte'
      
        '    '#39' State Bits (4 bytes): This field contains the user-defined' +
        ' flags if this entry is for a storage object or'
      
        '    '#39'  root storage object. For a stream object, this field SHOU' +
        'LD be set to all zeroes because many'
      
        '    '#39'  implementations provide no way for applications to retrie' +
        've state bits from a stream object. If an'
      
        '    '#39'  implementation of the file format enables applications to' +
        ' create storage objects without explicitly'
      '    '#39'  setting state bits, it MUST write all zeroes by default.'
      '    StateBits As Long'
      
        '    '#39' Creation Time (8 bytes): This field contains the creation ' +
        'time for a storage object, or all zeroes to'
      
        '    '#39'  indicate that the creation time of the storage object was' +
        ' not recorded. The Windows FILETIME'
      
        '    '#39'  structure is used to represent this field in UTC. For a s' +
        'tream object, this field MUST be all zeroes.'
      
        '    '#39'  For a root storage object, this field MUST be all zeroes,' +
        ' and the creation time is retrieved or set on'
      '    '#39'  the compound file itself.'
      '    CreationTime As Double'
      
        '    '#39' Modified Time (8 bytes): This field contains the modificat' +
        'ion time for a storage object, or all'
      
        '    '#39'  zeroes to indicate that the modified time of the storage ' +
        'object was not recorded. The Windows'
      
        '    '#39'  FILETIME structure is used to represent this field in UTC' +
        '. For a stream object, this field MUST be'
      
        '    '#39'  all zeroes. For a root storage object, this field MAY<2> ' +
        'be set to all zeroes, and the modified time'
      '    '#39'  is retrieved or set on the compound file itself.'
      '    ModifiedTime As Double'
      
        '    '#39' Starting Sector Location (4 bytes): This field contains th' +
        'e first sector location if this is a stream'
      
        '    '#39'  object. For a root storage object, this field MUST contai' +
        'n the first sector of the mini stream, if the'
      
        '    '#39'  mini stream exists. For a storage object, this field MUST' +
        ' be set to all zeroes.'
      '    StartingSectorLocation As Long'
      
        '    '#39' Stream Size (8 bytes): This 64-bit integer field contains ' +
        'the size of the user-defined data if this is'
      
        '    '#39'  a stream object. For a root storage object, this field co' +
        'ntains the size of the mini stream. For a'
      '    '#39'  storage object, this field MUST be set to all zeroes.'
      
        '    '#39'  For a version 3 compound file 512-byte sector size, the v' +
        'alue of this field MUST be less than'
      
        '    '#39'  or equal to 0x80000000. (Equivalently, this requirement c' +
        'an be stated: the size of a stream or'
      
        '    '#39'  of the mini stream in a version 3 compound file MUST be l' +
        'ess than or equal to 2 gigabytes'
      
        '    '#39'  (GB).) Note that as a consequence of this requirement, th' +
        'e most significant 32 bits of this field'
      
        '    '#39'  MUST be zero in a version 3 compound file. However, imple' +
        'menters should be aware that'
      
        '    '#39'  some older implementations did not initialize the most si' +
        'gnificant 32 bits of this field, and'
      
        '    '#39'  these bits might therefore be nonzero in files that are o' +
        'therwise valid version 3 compound'
      
        '    '#39'  files. Although this document does not normatively specif' +
        'y parser behavior, it is recommended'
      
        '    '#39'  that parsers ignore the most significant 32 bits of this ' +
        'field in version 3 compound files,'
      
        '    '#39'  treating it as if its value were zero, unless there is a ' +
        'specific reason to do otherwise (for'
      
        '    '#39'  example, a parser whose purpose is to verify the correctn' +
        'ess of a compound file).'
      '    StreamSize As Long'
      '    StreamSizeHigh As Long'
      'End Type'
      ''
      'Private Type TCFBStream'
      '    Content() As Byte'
      '    ContentName As String'
      '    FullName As String'
      '    StartingSector As Long'
      '    StorageName As String'
      '    StreamSize As Long'
      'End Type'
      ''
      
        #39'---------------------------------------------------------------' +
        '---------------'
      #39' Global variables'
      
        #39'---------------------------------------------------------------' +
        '---------------'
      ''
      'Private CFBHeader As TCFBHeader'
      'Private CFBDiFATArray() As Long'
      'Private CFBDiFATLength As Long'
      'Private CFBFATArray() As Long'
      'Private CFBFATLength As Long'
      'Private CFBDirArray() As TDirectoryEntry'
      'Private CFBDirLength As Long'
      'Private CFBMiniFATArray() As Long'
      'Private CFBMiniFATLength As Long'
      'Private CFBMiniStream() As Byte'
      'Private CFBStreamArray() As TCFBStream'
      'Private CFBStreamCount As Long'
      'Private CFBStreamLength As Long'
      'Private SectorSize As Long '#39' Number of Bytes in a sector'
      'Private SectorLength As Long '#39' Number of UInt32s in a sector'
      'Private CodePage As Integer '#39' Actually ignored'
      'Private ModulesArray() As TModule'
      'Private ModulesLength As Integer'
      ''
      
        #39'===============================================================' +
        '==============='
      #39' Code section'
      
        #39'===============================================================' +
        '==============='
      ''
      
        #39'---------------------------------------------------------------' +
        '---------------'
      #39' Common functions'
      
        #39'---------------------------------------------------------------' +
        '---------------'
      ''
      
        'Private Function ExtractFileExtension(ByVal FileName As String) ' +
        'As String'
      '    Dim LastDotPosition As Integer'
      '    LastDotPosition = InStrRev(FileName, ".")'
      '    If LastDotPosition = 0 Then'
      '        ExtractFileExtension = vbNullString'
      '    Else'
      
        '        ExtractFileExtension = Right$(FileName, Len(FileName) - ' +
        'LastDotPosition)'
      '    End If'
      'End Function'
      ''
      
        'Private Function ExtractFileName(ByVal FileName As String) As St' +
        'ring'
      '    Dim LastSlashPosition As Integer'
      '    LastSlashPosition = InStrRev(FileName, "\")'
      '    If LastSlashPosition = 0 Then'
      '        ExtractFileName = vbNullString'
      '    Else'
      
        '        ExtractFileName = Right$(FileName, Len(FileName) - LastS' +
        'lashPosition)'
      '    End If'
      'End Function'
      ''
      
        'Private Function LeftShift(ByVal Number As Long, ByVal Shifts As' +
        ' Integer) As Long'
      '    Dim I As Integer'
      '    Dim Result As Long'
      '    Result = Number'
      '    For I = 1 To Shifts'
      '        Result = Result * 2'
      '    Next'
      '    LeftShift = Result'
      'End Function'
      ''
      
        'Private Function RightShift(ByVal Number As Long, ByVal Shifts A' +
        's Integer) As Long'
      '    Dim I As Integer'
      '    Dim Result As Long'
      '    Result = Number'
      '    For I = 1 To Shifts'
      '        Result = Result \ 2'
      '    Next'
      '    RightShift = Result'
      'End Function'
      ''
      
        #39'---------------------------------------------------------------' +
        '---------------'
      #39' Microsoft Compound File Binary management functions'
      
        #39'---------------------------------------------------------------' +
        '---------------'
      ''
      'Private Function GetStreamByName(ByVal Name As String) As Byte()'
      '    Dim Index As Integer'
      '    Dim NotFound As Boolean'
      '    Index = 0'
      '    NotFound = True'
      '    Do While (NotFound) And (Index < CFBStreamCount)'
      '        If CFBStreamArray(Index).ContentName = Name Then'
      '            GetStreamByName = CFBStreamArray(Index).Content'
      '            NotFound = False'
      '        Else'
      '            Index = Index + 1'
      '        End If'
      '    Loop'
      '    '#39' Default result is zero, if not found!'
      'End Function'
      ''
      
        'Private Sub AppendStreamList(ByVal StorageName As String, ByVal ' +
        'NewIndex As Long)'
      '    Dim ContentName As String'
      '    If NewIndex <> SectorFREESECT Then'
      '        ContentName = CFBDirArray(NewIndex).DirectoryEntryName'
      '        Select Case CFBDirArray(NewIndex).ObjectType'
      '            Case 1:'
      
        '                AppendStreamList StorageName + "\" + ContentName' +
        ', CFBDirArray(NewIndex).ChildID'
      
        '                '#39' Recursively call AppendStreamList for the two ' +
        'siblings'
      
        '                AppendStreamList StorageName, CFBDirArray(NewInd' +
        'ex).LeftSiblingID'
      
        '                AppendStreamList StorageName, CFBDirArray(NewInd' +
        'ex).RightSiblingID'
      '            Case 2:'
      
        '                CFBStreamArray(CFBStreamCount).ContentName = Con' +
        'tentName'
      
        '                CFBStreamArray(CFBStreamCount).StorageName = Sto' +
        'rageName'
      
        '                CFBStreamArray(CFBStreamCount).FullName = Storag' +
        'eName + "\" + ContentName'
      
        '                CFBStreamArray(CFBStreamCount).StartingSector = ' +
        'CFBDirArray(NewIndex).StartingSectorLocation'
      
        '                CFBStreamArray(CFBStreamCount).StreamSize = CFBD' +
        'irArray(NewIndex).StreamSize'
      '                CFBStreamCount = CFBStreamCount + 1'
      
        '                '#39' Recursively call AppendStreamList for the two ' +
        'siblings'
      
        '                AppendStreamList StorageName, CFBDirArray(NewInd' +
        'ex).LeftSiblingID'
      
        '                AppendStreamList StorageName, CFBDirArray(NewInd' +
        'ex).RightSiblingID'
      '        End Select'
      '    End If'
      'End Sub'
      ''
      
        'Private Sub ReadSectorAsByte(ByVal FileChannel As Integer, ByVal' +
        ' SectorNumber As Long, ByRef SectorData() As Byte)'
      '    Dim I As Long'
      '    Seek FileChannel, (SectorNumber + 1) * SectorSize + 1'
      '    For I = 0 To SectorSize - 1'
      '        Get FileChannel, , SectorData(I)'
      '    Next'
      'End Sub'
      ''
      
        'Private Sub ReadSectorAsLong(ByVal FileChannel As Integer, ByVal' +
        ' SectorNumber As Long, ByRef SectorData() As Long)'
      '    Dim I As Long'
      '    Seek FileChannel, (SectorNumber + 1) * SectorSize + 1'
      '    For I = 0 To SectorLength - 1'
      '        Get FileChannel, , SectorData(I)'
      '    Next'
      'End Sub'
      ''
      
        'Private Sub ReadSectorAsDirectory(ByVal FileChannel As Integer, ' +
        'ByVal SectorNumber As Long, ByRef SectorData() As TDirectoryEntr' +
        'y)'
      '    Dim DirectoryEntry As TDirectoryEntry'
      '    Dim DirectoryEntryName(63) As Byte'
      '    Dim I As Long'
      '    Seek FileChannel, (SectorNumber + 1) * SectorSize + 1'
      '    For I = 0 To RightShift(SectorSize, 7) - 1'
      '        Get FileChannel, , DirectoryEntryName'
      
        '        Get FileChannel, , DirectoryEntry.DirectoryEntryNameLeng' +
        'th'
      
        '        If DirectoryEntry.DirectoryEntryNameLength = 0 Then Exit' +
        ' For'
      '        Get FileChannel, , DirectoryEntry.ObjectType'
      '        Get FileChannel, , DirectoryEntry.ColorFlag'
      '        Get FileChannel, , DirectoryEntry.LeftSiblingID'
      '        Get FileChannel, , DirectoryEntry.RightSiblingID'
      '        Get FileChannel, , DirectoryEntry.ChildID'
      '        Get FileChannel, , DirectoryEntry.CLSID'
      '        Get FileChannel, , DirectoryEntry.StateBits'
      '        Get FileChannel, , DirectoryEntry.CreationTime'
      '        Get FileChannel, , DirectoryEntry.ModifiedTime'
      '        Get FileChannel, , DirectoryEntry.StartingSectorLocation'
      '        Get FileChannel, , DirectoryEntry.StreamSize'
      '        Get FileChannel, , DirectoryEntry.StreamSizeHigh'
      '        DirectoryEntry.DirectoryEntryName = DirectoryEntryName'
      
        '        DirectoryEntry.DirectoryEntryName = Left$(DirectoryEntry' +
        '.DirectoryEntryName, RightShift(DirectoryEntry.DirectoryEntryNam' +
        'eLength, 1) - 1)'
      '        SectorData(I) = DirectoryEntry'
      '    Next'
      'End Sub'
      ''
      #39' ReadStream: reads a chained block of sectors'
      
        'Function ReadStream(ByVal FileChannel As Integer, StartingSector' +
        'Location As Long, StreamSize As Long) As Byte()'
      '    Dim DataBuffer()   As Byte '#39' Result'
      
        '    Dim DataSize       As Long '#39' Calculated data size from FAT c' +
        'hain'
      '    Dim DataRemaining  As Long '#39' Remaining bytes to read'
      '    Dim DataChunkSize  As Long '#39' Bytes to read during the step'
      
        '    Dim DataOffset     As Long '#39' Pointer to the current reading ' +
        'location'
      '    Dim NextFATSector  As Long '#39' Next chain sector'
      
        '    Dim SectorCount    As Long '#39' Calculated sector count from FA' +
        'T chain'
      '    Dim SectorSize     As Long '#39' Calculated sector size in Byte'
      '    Dim SectorBuffer() As Byte '#39' Reading buffer for the sector'
      '    Dim I As Long'
      '    SectorSize = LeftShift(1, CFBHeader.SectorShift)'
      '    If StreamSize = 0 Then'
      '        SectorCount = 0'
      '        NextFATSector = StartingSectorLocation'
      '        Do While NextFATSector <> SectorENDOFCHAIN'
      '            SectorCount = SectorCount + 1'
      '            NextFATSector = CFBFATArray(NextFATSector)'
      'If NextFATSector < 0 Then Exit Do'
      '        Loop'
      '        DataSize = LeftShift(SectorCount, CFBHeader.SectorShift)'
      '    Else'
      '        DataSize = StreamSize'
      '    End If'
      '    ReDim DataBuffer(DataSize - 1)'
      '    ReDim SectorBuffer(SectorSize - 1)'
      '    DataRemaining = DataSize'
      '    NextFATSector = StartingSectorLocation'
      '    DataOffset = 0'
      '    Do While NextFATSector <> SectorENDOFCHAIN'
      '        DataChunkSize = DataRemaining'
      '        If DataChunkSize > SectorSize Then'
      '            DataChunkSize = SectorSize'
      '        End If'
      
        '        ReadSectorAsByte FileChannel, NextFATSector, SectorBuffe' +
        'r'
      '        For I = 0 To DataChunkSize - 1'
      '            DataBuffer(DataOffset) = SectorBuffer(I)'
      '            DataOffset = DataOffset + 1'
      '        Next'
      '        DataRemaining = DataRemaining - DataChunkSize'
      '        NextFATSector = CFBFATArray(NextFATSector)'
      'If NextFATSector < 0 Then Exit Do'
      '    Loop'
      '    ReadStream = DataBuffer'
      'End Function'
      ''
      #39' ReadMiniStream: reads a chained block of mini sectors'
      
        'Function ReadMiniStream(ByVal FileChannel As Integer, StartingSe' +
        'ctorLocation As Long, StreamSize As Long) As Byte()'
      '    Dim DataBuffer()  As Byte '#39' Result'
      
        '    Dim DataSize      As Long '#39' Calculated data size from FAT ch' +
        'ain'
      '    Dim DataRemaining As Long '#39' Remaining bytes to read'
      '    Dim DataChunkSize As Long '#39' Bytes to read during the step'
      
        '    Dim DataOffset    As Long '#39' Pointer to the current reading l' +
        'ocation'
      '    Dim NextFATSector As Long '#39' Next chain mini sector'
      
        '    Dim SectorCount   As Long '#39' Calculated sector count from FAT' +
        ' chain'
      '    Dim SectorSize    As Long '#39' Calculated sector size in Byte'
      '    Dim SectorBuffer() As Byte '#39' Reading buffer for the sector'
      '    Dim I As Long'
      '    SectorSize = LeftShift(1, CFBHeader.MiniSectorShift)'
      '    If StreamSize = 0 Then'
      '        SectorCount = 0'
      '        NextFATSector = StartingSectorLocation'
      '        Do While NextFATSector <> SectorENDOFCHAIN'
      '            SectorCount = SectorCount + 1'
      '            NextFATSector = CFBMiniFATArray(NextFATSector)'
      '        Loop'
      
        '        DataSize = LeftShift(SectorCount, CFBHeader.MiniSectorSh' +
        'ift)'
      '    Else'
      '        DataSize = StreamSize'
      '    End If'
      '    ReDim DataBuffer(DataSize - 1)'
      '    DataRemaining = DataSize'
      '    NextFATSector = StartingSectorLocation'
      '    DataOffset = 0'
      '    Do While NextFATSector <> SectorENDOFCHAIN'
      '        DataChunkSize = DataRemaining'
      '        If DataChunkSize > SectorSize Then'
      '            DataChunkSize = SectorSize'
      '        End If'
      '        For I = 0 To DataChunkSize - 1'
      
        '            DataBuffer(DataOffset) = CFBMiniStream(LeftShift(Nex' +
        'tFATSector, CFBHeader.MiniSectorShift) + I)'
      '            DataOffset = DataOffset + 1'
      '        Next'
      '        DataRemaining = DataRemaining - DataChunkSize'
      '        NextFATSector = CFBMiniFATArray(NextFATSector)'
      '    Loop'
      '    ReadMiniStream = DataBuffer'
      'End Function'
      ''
      'Private Function ReadCFB(ByVal FileName As String) As Boolean'
      '    Dim DataSize As Long'
      '    Dim FileChannel As Integer'
      '    Dim I As Long'
      '    Dim J As Long'
      '    Dim N As Long'
      '    Dim NextFATSector As Long'
      '    Dim Position As Long'
      '    Dim SectorDataAsLong() As Long'
      '    Dim SectorDataAsDirectory() As TDirectoryEntry'
      '    FileChannel = FreeFile()'
      '    Open FileName For Binary Access Read Shared As FileChannel'
      '    '#39' Step 1: read and check the Compound File Header'
      '    Seek FileChannel, 1'
      '    Get FileChannel, , CFBHeader'
      
        '    If (CFBHeader.Signature(0) = &HD0) Or (CFBHeader.Signature(1' +
        ') = &HCF) _'
      
        '    Or (CFBHeader.Signature(2) = &H11) Or (CFBHeader.Signature(3' +
        ') = &HE0) _'
      
        '    Or (CFBHeader.Signature(4) = &HA1) Or (CFBHeader.Signature(5' +
        ') = &HB1) _'
      
        '    Or (CFBHeader.Signature(6) = &H1A) Or (CFBHeader.Signature(7' +
        ') = &HE1) Then'
      
        '        If (CFBHeader.HeaderCLSID(0) = &H0) Or (CFBHeader.Header' +
        'CLSID(1) = &H0) _'
      
        '        Or (CFBHeader.HeaderCLSID(2) = &H0) Or (CFBHeader.Header' +
        'CLSID(3) = &H0) _'
      
        '        Or (CFBHeader.HeaderCLSID(4) = &H0) Or (CFBHeader.Header' +
        'CLSID(5) = &H0) _'
      
        '        Or (CFBHeader.HeaderCLSID(6) = &H0) Or (CFBHeader.Header' +
        'CLSID(7) = &H0) _'
      
        '        Or (CFBHeader.HeaderCLSID(8) = &H0) Or (CFBHeader.Header' +
        'CLSID(9) = &H0) _'
      
        '        Or (CFBHeader.HeaderCLSID(10) = &H0) Or (CFBHeader.Heade' +
        'rCLSID(11) = &H0) _'
      
        '        Or (CFBHeader.HeaderCLSID(12) = &H0) Or (CFBHeader.Heade' +
        'rCLSID(13) = &H0) _'
      
        '        Or (CFBHeader.HeaderCLSID(14) = &H0) Or (CFBHeader.Heade' +
        'rCLSID(15) = &H0) Then'
      '            SectorSize = LeftShift(1, CFBHeader.SectorShift)'
      '            SectorLength = RightShift(SectorSize, 2)'
      '            ReDim SectorDataAsLong(SectorLength - 1)'
      
        '            ReDim SectorDataAsDirectory(RightShift(SectorSize, 7' +
        ') - 1)'
      '            '#39' Step 2: read the DIFAT'
      '            N = 0'
      
        '            DataSize = CFBHeader.NumberOfDIFATSectors * SectorSi' +
        'ze'
      '            CFBDiFATLength = RightShift(DataSize, 2) + 109'
      '            ReDim CFBDiFATArray(CFBDiFATLength - 1)'
      '            For I = 0 To 108'
      '                CFBDiFATArray(N) = CFBHeader.DIFAT(I)'
      '                N = N + 1'
      '            Next'
      '            NextFATSector = CFBHeader.FirstDIFATSectorLocation'
      '            Do While NextFATSector <> SectorENDOFCHAIN'
      
        '                ReadSectorAsLong FileChannel, NextFATSector, Sec' +
        'torDataAsLong'
      '                For I = 0 To SectorLength - 2'
      '                    CFBDiFATArray(N) = SectorDataAsLong(I)'
      '                    N = N + 1'
      '                Next'
      
        '                NextFATSector = SectorDataAsLong(SectorLength - ' +
        '1)'
      '            Loop'
      '            '#39' Step 3: read the FAT'
      '            N = 0'
      '            DataSize = CFBHeader.NumberOfFATSectors * SectorSize'
      '            CFBFATLength = RightShift(DataSize, 2)'
      '            ReDim CFBFATArray(CFBFATLength - 1)'
      '            For I = 0 To CFBHeader.NumberOfFATSectors - 1'
      
        '                ReadSectorAsLong FileChannel, CFBDiFATArray(I), ' +
        'SectorDataAsLong'
      '                For J = 0 To SectorLength - 1'
      '                    CFBFATArray(N) = SectorDataAsLong(J)'
      '                    N = N + 1'
      '                Next'
      '            Next'
      '            '#39' Step 4: read the Directory Entry Array'
      '            N = 0'
      
        '            NextFATSector = CFBHeader.FirstDirectorySectorLocati' +
        'on'
      '            Do While NextFATSector <> SectorENDOFCHAIN'
      '                N = N + 1'
      '                NextFATSector = CFBFATArray(NextFATSector)'
      '            Loop'
      '            DataSize = N * SectorSize'
      '            CFBDirLength = RightShift(DataSize, 7)'
      '            ReDim CFBDirArray(CFBDirLength - 1)'
      '            N = 0'
      
        '            NextFATSector = CFBHeader.FirstDirectorySectorLocati' +
        'on'
      '            Do While NextFATSector <> SectorENDOFCHAIN'
      
        '                ReadSectorAsDirectory FileChannel, NextFATSector' +
        ', SectorDataAsDirectory'
      '                For I = 0 To RightShift(SectorSize, 7) - 1'
      '                    CFBDirArray(N) = SectorDataAsDirectory(I)'
      '                    N = N + 1'
      '                Next'
      '                NextFATSector = CFBFATArray(NextFATSector)'
      '            Loop'
      '            ReDim Preserve CFBDirArray(CFBDirLength - 1)'
      '            '#39' Step 5: read the mini FAT'
      '            N = 0'
      
        '            DataSize = CFBHeader.NumberOfMiniFATSectors * Sector' +
        'Size'
      '            CFBMiniFATLength = RightShift(DataSize, 2)'
      '            ReDim CFBMiniFATArray(CFBMiniFATLength - 1)'
      '            NextFATSector = CFBHeader.FirstMiniFATSectorLocation'
      '            Do While NextFATSector <> SectorENDOFCHAIN'
      
        '                ReadSectorAsLong FileChannel, NextFATSector, Sec' +
        'torDataAsLong'
      '                For J = 0 To SectorLength - 1'
      '                    CFBMiniFATArray(N) = SectorDataAsLong(J)'
      '                    N = N + 1'
      '                Next'
      '                NextFATSector = CFBFATArray(NextFATSector)'
      '            Loop'
      '            '#39' Step 6: read the mini stream'
      
        '            CFBMiniStream = ReadStream(FileChannel, CFBDirArray(' +
        '0).StartingSectorLocation, 0)'
      '            '#39' Step 7: build the stream list'
      '            '#39' Initialize the array of streams'
      '            CFBStreamCount = 0'
      '            CFBStreamLength = 0'
      '            For I = 0 To CFBDirLength - 1'
      '                If CFBDirArray(I).ObjectType = 2 Then'
      '                    CFBStreamLength = CFBStreamLength + 1'
      '                End If'
      '            Next'
      '            ReDim CFBStreamArray(CFBStreamLength - 1)'
      
        '            '#39' Start with the root entry, use AppendStreamList, t' +
        'o add the stream objects recursively'
      '            AppendStreamList "", CFBDirArray(0).ChildID'
      '            '#39' Read the streams'
      '            For I = 0 To CFBStreamLength - 1'
      
        '                If (InStr(CFBStreamArray(I).StorageName, "VBA") ' +
        '> 0) And (CFBStreamArray(I).StreamSize > 0) Then'
      
        '                    If CFBStreamArray(I).StreamSize <= CFBHeader' +
        '.MiniStreamCutoffSize Then'
      
        '                        CFBStreamArray(I).Content = ReadMiniStre' +
        'am(FileChannel, CFBStreamArray(I).StartingSector, CFBStreamArray' +
        '(I).StreamSize)'
      '                    Else'
      
        '                        CFBStreamArray(I).Content = ReadStream(F' +
        'ileChannel, CFBStreamArray(I).StartingSector, CFBStreamArray(I).' +
        'StreamSize)'
      '                    End If'
      '                End If'
      '            Next'
      '            ReadCFB = True'
      '        Else'
      '            ReadCFB = False'
      '        End If'
      '    Else'
      '        ReadCFB = False'
      '    End If'
      '    Close FileChannel'
      'End Function'
      ''
      
        #39'---------------------------------------------------------------' +
        '---------------'
      #39' Microsoft VBA streams management functions'
      
        #39'---------------------------------------------------------------' +
        '---------------'
      ''
      
        'Private Sub DecompressContainer(ByRef CompressedContainer() As B' +
        'yte, ByRef CompressedStart As Long, ByRef DecompressedData() As ' +
        'Byte)'
      '    Dim DecompressedIndex  As Long'
      '    Dim DecompressedLength As Long'
      '    Dim ChunkHeader        As Long'
      '    Dim ChunkSignature     As Long'
      '    Dim ChunkFlag          As Long'
      '    Dim ChunkSize          As Long'
      '    Dim ChunkEnd           As Long'
      '    Dim BitFlags           As Byte'
      '    Dim Token              As Long'
      '    Dim BitCount           As Long'
      '    Dim BitMask            As Long'
      '    Dim CopyLength         As Long'
      '    Dim CopyOffset         As Long'
      '    Dim I                  As Long'
      '    Dim J                  As Long'
      '    '#39' Use an array to speedup calculations'
      '    Dim PowerOf2(0 To 16)  As Long'
      '    PowerOf2(0) = 1'
      '    For I = 1 To UBound(PowerOf2)'
      '        PowerOf2(I) = PowerOf2(I - 1) * 2'
      '    Next'
      '    Do'
      '        If Not (Not DecompressedData) Then'
      
        '            ReDim Preserve DecompressedData(UBound(DecompressedD' +
        'ata) + 4096)'
      '        Else'
      '            ReDim DecompressedData(4095)'
      '            DecompressedIndex = 0'
      '        End If'
      
        '        ChunkHeader = CompressedContainer(CompressedStart) + 256' +
        '& * CompressedContainer(CompressedStart + 1)'
      '        CompressedStart = CompressedStart + 2'
      '        ChunkSize = (ChunkHeader And &HFFF)'
      '        ChunkEnd = CompressedStart + ChunkSize'
      '        ChunkSignature = (ChunkHeader And &H7000) \ &H1000&'
      '        ChunkFlag = (ChunkHeader And &H8000) \ &H8000&'
      '        If ChunkFlag = 0 Then'
      '            For J = 0 To 4095'
      
        '                DecompressedData(DecompressedIndex + J) = Compre' +
        'ssedContainer(CompressedStart + J)'
      '            Next'
      '            CompressedStart = CompressedStart + 4096'
      '            DecompressedIndex = DecompressedIndex + 4096'
      '        Else'
      '            Do'
      '                BitFlags = CompressedContainer(CompressedStart)'
      '                CompressedStart = CompressedStart + 1'
      '                For I = 0 To 7'
      '                    If CompressedStart > ChunkEnd Then Exit Do'
      '                    If (BitFlags And PowerOf2(I)) = 0 Then'
      
        '                        DecompressedData(DecompressedIndex) = Co' +
        'mpressedContainer(CompressedStart)'
      '                        CompressedStart = CompressedStart + 1'
      
        '                        DecompressedIndex = DecompressedIndex + ' +
        '1'
      '                    Else'
      
        '                        Token = CompressedContainer(CompressedSt' +
        'art) + CompressedContainer(CompressedStart + 1) * 256&'
      '                        CompressedStart = CompressedStart + 2'
      
        '                        DecompressedLength = DecompressedIndex M' +
        'od 4096'
      '                        For BitCount = 4 To 11'
      
        '                            If DecompressedLength <= PowerOf2(Bi' +
        'tCount) Then Exit For'
      '                        Next'
      
        '                        BitMask = PowerOf2(16) - PowerOf2(16 - B' +
        'itCount)'
      
        '                        CopyOffset = (Token And BitMask) \ Power' +
        'Of2(16 - BitCount) + 1'
      '                        BitMask = PowerOf2(16 - BitCount) - 1'
      '                        CopyLength = (Token And BitMask) + 3'
      '                        For J = 0 To CopyLength - 1'
      
        '                            DecompressedData(DecompressedIndex +' +
        ' J) = DecompressedData(DecompressedIndex - CopyOffset + J)'
      '                        Next'
      
        '                        DecompressedIndex = DecompressedIndex + ' +
        'CopyLength'
      '                    End If'
      '                Next'
      '            Loop'
      '        End If'
      
        '        If CompressedStart > UBound(CompressedContainer) Then Ex' +
        'it Do'
      '    Loop'
      '    ReDim Preserve DecompressedData(DecompressedIndex - 1)'
      'End Sub'
      ''
      
        'Private Function ReadBYTE(ByRef DecompressedStream() As Byte, By' +
        'Ref Offset As Long) As Byte'
      '    ReadBYTE = DecompressedStream(Offset)'
      '    Offset = Offset + 1'
      'End Function'
      ''
      
        'Private Function ReadWORD(ByRef DecompressedStream() As Byte, By' +
        'Ref Offset As Long) As Long'
      '    Dim Half1 As Long'
      '    Dim Half2 As Long'
      '    Half1 = ReadBYTE(DecompressedStream, Offset)'
      '    Half2 = ReadBYTE(DecompressedStream, Offset)'
      '    ReadWORD = Half1 + Half2 * 256'
      'End Function'
      ''
      
        'Private Function ReadDWORD(ByRef DecompressedStream() As Byte, B' +
        'yRef Offset As Long) As Long'
      '    Dim Half1 As Long'
      '    Dim Half2 As Long'
      '    Half1 = ReadWORD(DecompressedStream, Offset)'
      '    Half2 = ReadWORD(DecompressedStream, Offset)'
      '    ReadDWORD = Half1 + Half2 * 65536'
      'End Function'
      ''
      
        'Private Function ReadString(ByRef DecompressedStream() As Byte, ' +
        'ByRef Offset As Long, ByVal NumberOfBytes As Long) As String'
      '    Dim Result As String'
      '    Dim StringData() As Byte'
      '    Dim StringIndex As Long'
      '    Result = vbNullString'
      '    If NumberOfBytes > 0 Then'
      '        ReDim StringData(NumberOfBytes - 1)'
      '        For StringIndex = 0 To NumberOfBytes - 1'
      '            StringData(StringIndex) = DecompressedStream(Offset)'
      '            Offset = Offset + 1'
      '        Next'
      '        Result = StrConv(StringData, vbUnicode)'
      '    End If'
      '    ReadString = Result'
      'End Function'
      ''
      'Private Function ParseDirStream() As Boolean'
      '    Dim CompressedStream() As Byte'
      '    Dim DecompressedStream() As Byte'
      '    Dim ModuleIndex As Integer'
      '    Dim RecordId As Integer'
      '    Dim RecordSize As Long'
      '    Dim StreamEnd As Long'
      '    Dim StreamOffset As Long'
      '    CompressedStream = GetStreamByName("dir")'
      '    If Not (Not CompressedStream) Then'
      
        '        DecompressContainer CompressedStream, 1, DecompressedStr' +
        'eam'
      '        CodePage = 0'
      '        StreamOffset = 0'
      '        StreamEnd = UBound(DecompressedStream)'
      '        Do While StreamOffset <= StreamEnd'
      '            '#39' Step 1: get the RecordId'
      
        '            RecordId = ReadWORD(DecompressedStream, StreamOffset' +
        ')'
      '            '#39' Step 2: get the RecordId'
      
        '            RecordSize = ReadDWORD(DecompressedStream, StreamOff' +
        'set)'
      '            '#39' Step 3: parse the RecordId field'
      '            Select Case RecordId'
      '                Case &H1:    '#39' SysKindRecord'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H2:    '#39' LcidRecord'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H14:   '#39' LcidInvokeRecord'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H3:    '#39' CodePageRecord'
      
        '                    CodePage = ReadWORD(DecompressedStream, Stre' +
        'amOffset)'
      '                Case &H4:    '#39' NameRecord'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H5:    '#39' DocStringRecord'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                    StreamOffset = StreamOffset + 2'
      
        '                    RecordSize = ReadDWORD(DecompressedStream, S' +
        'treamOffset)'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H6:    '#39' HelpFilePathRecord 1'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H3D:   '#39' HelpFilePathRecord 2'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H7:    '#39' HelpContextRecord'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H8:    '#39' LibFlagsRecord'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H9:    '#39' VersionRecord'
      '                    StreamOffset = StreamOffset + 6'
      '                Case &HC:    '#39' ConstantsRecord'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                    StreamOffset = StreamOffset + 2'
      
        '                    RecordSize = ReadDWORD(DecompressedStream, S' +
        'treamOffset)'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H16:   '#39' Reference name'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                    StreamOffset = StreamOffset + 2'
      
        '                    RecordSize = ReadDWORD(DecompressedStream, S' +
        'treamOffset)'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H2F:   '#39' Reference Control'
      
        '                    RecordSize = ReadDWORD(DecompressedStream, S' +
        'treamOffset)'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                    StreamOffset = StreamOffset + 6'
      '                Case &H30:   '#39' Reference extended'
      
        '                    RecordSize = ReadDWORD(DecompressedStream, S' +
        'treamOffset)'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                    StreamOffset = StreamOffset + 6'
      '                    StreamOffset = StreamOffset + 16'
      '                    StreamOffset = StreamOffset + 4'
      '                Case &H33:   '#39' Reference Original'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &HD:    '#39' Reference Registered'
      
        '                    RecordSize = ReadDWORD(DecompressedStream, S' +
        'treamOffset)'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                    StreamOffset = StreamOffset + 6'
      '                Case &HE:    '#39' Reference Project'
      
        '                    RecordSize = ReadDWORD(DecompressedStream, S' +
        'treamOffset)'
      '                    StreamOffset = StreamOffset + RecordSize'
      
        '                    RecordSize = ReadDWORD(DecompressedStream, S' +
        'treamOffset)'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                    StreamOffset = StreamOffset + 6'
      '                Case &HF:    '#39' Modules'
      
        '                    ModulesLength = ReadWORD(DecompressedStream,' +
        ' StreamOffset)'
      '                    ReDim ModulesArray(ModulesLength - 1)'
      '                    StreamOffset = StreamOffset + 8'
      '                    ModuleIndex = -1'
      '                Case &H19:   '#39' Module name'
      '                    ModuleIndex = ModuleIndex + 1'
      
        '                    ModulesArray(ModuleIndex).ModuleName = ReadS' +
        'tring(DecompressedStream, StreamOffset, RecordSize)'
      '                Case &H47:   '#39' Module name in Unicode'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H1A:   '#39' Module stream name'
      
        '                    ModulesArray(ModuleIndex).ModuleStreamName =' +
        ' ReadString(DecompressedStream, StreamOffset, RecordSize)'
      '                Case &H32:   '#39' Module stream name in Unicode'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H1C:   '#39' Module doc string'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H48:   '#39' Module doc string in Unicode'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H31:   '#39' Module offset'
      
        '                    ModulesArray(ModuleIndex).TextOffset = ReadD' +
        'WORD(DecompressedStream, StreamOffset)'
      '                Case &H1E:   '#39' Module help context'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H2C:   '#39' Module cookie'
      '                    StreamOffset = StreamOffset + RecordSize'
      '                Case &H21:   '#39' Module type standard'
      '                Case &H22:   '#39' Module type class'
      '                Case &H25:   '#39' Module read only'
      '                Case &H28:   '#39' Module private'
      '                Case &H10:   '#39' Terminator of dir stream'
      '                Case &H2B:   '#39' Terminator of Module record'
      '                Case &H4A:   '#39' Mystery field?'
      '                    StreamOffset = StreamOffset + 4'
      '            End Select'
      '        Loop'
      '        ParseDirStream = True'
      '    Else'
      '        ParseDirStream = False'
      '    End If'
      'End Function'
      ''
      'Private Sub ParseModuleStream()'
      '    Dim CompressedStream() As Byte'
      '    Dim DecompressedStream() As Byte'
      '    Dim SourceCode As String'
      '    Dim SourceLine As String'
      '    Dim SourceLines() As String'
      '    Dim I As Long'
      '    Dim J As Long'
      '    For I = 0 To ModulesLength - 1'
      
        '        CompressedStream = GetStreamByName(ModulesArray(I).Modul' +
        'eName)'
      
        '        DecompressContainer CompressedStream, ModulesArray(I).Te' +
        'xtOffset + 1, DecompressedStream'
      '        SourceCode = StrConv(DecompressedStream, vbUnicode)'
      '        SourceLines = Split(SourceCode, vbNewLine)'
      '        SourceCode = ""'
      '        For J = LBound(SourceLines) To UBound(SourceLines)'
      '            SourceLine = SourceLines(J)'
      '            If Left$(SourceLine, 9) <> "Attribute" Then'
      '                SourceCode = SourceCode & SourceLine & vbNewLine'
      '            End If'
      '        Next'
      '        ModulesArray(I).SourceCode = SourceCode'
      '    Next'
      'End Sub'
      ''
      'Private Function ParseFile(ByVal FileName As String) As Boolean'
      '    Dim FileExtension As String'
      '    FileExtension = UCase(ExtractFileExtension(FileName))'
      '    If (FileExtension = "BIN") Or (FileExtension = "DOC") _'
      '    Or (FileExtension = "OTM") Or (FileExtension = "XLS") Then'
      '        If ReadCFB(FileName) Then'
      '            '#39' Step 1: parse the '#39'dir'#39' stream'
      '            If ParseDirStream() Then'
      '                '#39' Step 2: read the modules streams'
      '                ParseModuleStream'
      '                ParseFile = True'
      '            Else'
      '                ParseFile = False'
      '            End If'
      '        Else'
      '            ParseFile = False'
      '        End If'
      '    Else'
      '        ParseFile = False'
      '    End If'
      'End Function'
      ''
      
        #39'---------------------------------------------------------------' +
        '---------------'
      #39' Gui management'
      
        #39'---------------------------------------------------------------' +
        '---------------'
      ''
      'Private Sub Reset()'
      '    TextBoxFileName.Text = ""'
      '    ListBoxModules.Clear'
      '    TextBoxSourceCode.Text = ""'
      'End Sub'
      ''
      'Private Sub Update()'
      '    Dim I As Integer'
      '    For I = 0 To ModulesLength - 1'
      '        ListBoxModules.AddItem ModulesArray(I).ModuleName'
      '    Next'
      'End Sub'
      ''
      'Private Sub CommandButtonOpen_Click()'
      '    Dim SelectFileDialog As FileDialog'
      '    Dim SelectedFileName As String'
      '    Reset'
      
        '    Set SelectFileDialog = Application.FileDialog(msoFileDialogF' +
        'ilePicker)'
      '    SelectFileDialog.AllowMultiSelect = False'
      '    SelectFileDialog.Filters.Clear'
      
        '    SelectFileDialog.Filters.Add "All supported files", "*.bin; ' +
        '*.doc; *.otm; *.xls"'
      '    SelectFileDialog.Filters.Add "Microsoft Excel", "*.xls"'
      '    SelectFileDialog.Filters.Add "Microsoft Word", "*.doc"'
      '    SelectFileDialog.Filters.Add "Microsoft Outlook", "*.otm"'
      
        '    SelectFileDialog.Filters.Add "Microsoft VBA project from XML' +
        ' documents", "*.bin"'
      '    SelectFileDialog.Title = "Choose an Office file"'
      '    If SelectFileDialog.Show() Then'
      
        '        SelectedFileName = SelectFileDialog.SelectedItems.Item(1' +
        ')'
      '        If ParseFile(SelectedFileName) Then'
      '            TextBoxFileName.Text = SelectedFileName'
      '            Update'
      '        End If'
      '    End If'
      'End Sub'
      ''
      'Private Sub ListBoxModules_Click()'
      
        '    TextBoxSourceCode.Text = ModulesArray(ListBoxModules.ListInd' +
        'ex).SourceCode'
      '    Repaint'
      'End Sub'
      ''
      'Private Sub UserForm_Terminate()'
      '    '#39' Safety: close all opened files'
      '    Close'
      'End Sub'
      '')
    Options = [eoAltSetsColumnMode, eoAutoIndent, eoDragDropEditing, eoEnhanceEndKey, eoGroupUndo, eoScrollPastEol, eoShowScrollHint, eoSmartTabDelete, eoSmartTabs, eoTabsToSpaces]
    ReadOnly = True
    FontSmoothing = fsmNone
  end
  object SynVBSyn: TSynVBSyn
    Options.AutoDetectEnabled = False
    Options.AutoDetectLineLimit = 0
    Options.Visible = False
    StringAttri.Foreground = clHotLight
    Left = 461
    Top = 64
  end
end
