unit FilesCFB;
{$ALIGN 1}

interface

uses System.Classes, System.Generics.Collections, Common;

const
  ValidSignature   : UInt64 = $E11AB1A1E011CFD0;
  SectorMAXREGSECT : UInt32 = $FFFFFFFA;
  SectorReserved   : UInt32 = $FFFFFFFB;
  SectorDIFSECT    : UInt32 = $FFFFFFFC;
  SectorFATSECT    : UInt32 = $FFFFFFFD;
  SectorENDOFCHAIN : UInt32 = $FFFFFFFE;
  SectorFREESECT   : UInt32 = $FFFFFFFF;

type
  { Structure for CFB Header, according to [MS-CFB] section 2.2 }
  TCFBHeader  = record
    { Header Signature (8 bytes): Identification signature for the compound file structure, and MUST be
      set to the value 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1. }
    Signature : UInt64;
    { Header CLSID (16 bytes): Reserved and unused class ID that MUST be set to all zeroes
      (CLSID_NULL). }
    HeaderCLSID: TGUID;
    { Minor Version (2 bytes): Version number for nonbreaking changes. This field SHOULD be set to
      0x003E if the major version field is either 0x0003 or 0x0004.}
    MinorVersion: UInt16;
    { Major Version (2 bytes): Version number for breaking changes. This field MUST be set to either
      0x0003 (version 3) or 0x0004 (version 4). }
    MajorVersion: UInt16;
    { Byte Order (2 bytes): This field MUST be set to 0xFFFE. This field is a byte order mark for all integer
      fields, specifying little-endian byte order. }
    ByteOrder: UInt16;
    { Sector Shift (2 bytes): This field MUST be set to 0x0009, or 0x000c, depending on the Major
      Version field. This field specifies the sector size of the compound file as a power of 2.
      If Major Version is 3, the Sector Shift MUST be 0x0009, specifying a sector size of 512 bytes.
      If Major Version is 4, the Sector Shift MUST be 0x000C, specifying a sector size of 4096 bytes. }
    SectorShift: UInt16;
    { Mini Sector Shift (2 bytes): This field MUST be set to 0x0006. This field specifies the sector size of
      the Mini Stream as a power of 2. The sector size of the Mini Stream MUST be 64 bytes. }
    MiniSectorShift: UInt16;
    { Reserved (6 bytes): This field MUST be set to all zeroes. }
    Reserved: array[0..5] of Byte;
    { Number of Directory Sectors (4 bytes): This integer field contains the count of the number of
      directory sectors in the compound file.
      If Major Version is 3, the Number of Directory Sectors MUST be zero. This field is not
      supported for version 3 compound files. }
    NumberOfDirectorySectors: UInt32;
    { Number of FAT Sectors (4 bytes): This integer field contains the count of the number of FAT
      sectors in the compound file. }
    NumberOfFATSectors: UInt32;
    { First Directory Sector Location (4 bytes): This integer field contains the starting sector number for
      the directory stream. }
    FirstDirectorySectorLocation: UInt32;
    { Transaction Signature Number (4 bytes): This integer field MAY contain a sequence number that
      is incremented every time the compound file is saved by an implementation that supports file
      transactions. This is the field that MUST be set to all zeroes if file transactions are not
      implemented. }
    TransactionSignatureNumber: UInt32;
    { Mini Stream Cutoff Size (4 bytes): This integer field MUST be set to 0x00001000. This field
      specifies the maximum size of a user-defined data stream that is allocated from the mini FAT
      and mini stream, and that cutoff is 4,096 bytes. Any user-defined data stream that is greater than
      or equal to this cutoff size must be allocated as normal sectors from the FAT. }
    MiniStreamCutoffSize: UInt32;
    { First Mini FAT Sector Location (4 bytes): This integer field contains the starting sector number for
      the mini FAT. }    FirstMiniFATSectorLocation: UInt32;    { Number of Mini FAT Sectors (4 bytes): This integer field contains the count of the number of mini      FAT sectors in the compound file. }    NumberOfMiniFATSectors: UInt32;    { First DIFAT Sector Location (4 bytes): This integer field contains the starting sector number for      the DIFAT. }    FirstDIFATSectorLocation: UInt32;    { Number of DIFAT Sectors (4 bytes): This integer field contains the count of the number of DIFAT      sectors in the compound file. }    NumberOfDIFATSectors: UInt32;    { DIFAT (436 bytes): This array of 32-bit integer fields contains the first 109 FAT sector locations of      the compound file. }    DIFAT: array[0..108] of UInt32;  end;

  TDirectoryEntryName = array[0..31] of Char;

  { Structure for Compound File Directory Entry, according to [MS-CFB] section 2.6 }
  TDirectoryEntry = record
    { Directory Entry Name (64 bytes): This field MUST contain a Unicode string for the storage or
      stream name encoded in UTF-16. The name MUST be terminated with a UTF-16 terminating null
      character. Thus, storage and stream names are limited to 32 UTF-16 code points, including the
      terminating null character. When locating an object in the compound file except for the root
      storage, the directory entry name is compared by using a special case-insensitive uppercase
      mapping, described in Red-Black Tree. The following characters are illegal and MUST NOT be part
      of the name: '/', '\', ':', '!'. }    DirectoryEntryName: TDirectoryEntryName;    { Directory Entry Name Length (2 bytes): This field MUST match the length of the Directory Entry      Name Unicode string in bytes. The length MUST be a multiple of 2 and include the terminating null
      character in the count. This length MUST NOT exceed 64, the maximum size of the Directory Entry
      Name field.}    DirectoryEntryNameLength: UInt16;    { Object Type (1 byte): This field MUST be 0x00, 0x01, 0x02, or 0x05, depending on the actual type      of object. All other values are not valid.      $00 Unknown or unallocated      $01 Storage Object      $02 Stream Object      $03 Root Storage Object }    ObjectType: Byte;    { Color Flag (1 byte): This field MUST be 0x00 (red) or 0x01 (black). All other values are not valid. }    ColorFlag: Byte;    { Left Sibling ID (4 bytes): This field contains the stream ID of the left sibling. If there is no left      sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF). }    LeftSiblingID: UInt32;    { Right Sibling ID (4 bytes): This field contains the stream ID of the right sibling. If there is no right      sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF). }    RightSiblingID: UInt32;    { Child ID (4 bytes): This field contains the stream ID of a child object. If there is no child object,      including all entries for stream objects, the field MUST be set to NOSTREAM (0xFFFFFFFF). }    ChildID: UInt32;    { CLSID (16 bytes): This field contains an object class GUID, if this entry is for a storage object or      root storage object. For a stream object, this field MUST be set to all zeroes. A value containing all
      zeroes in a storage or root storage directory entry is valid, and indicates that no object class is
      associated with the storage. If an implementation of the file format enables applications to create
      storage objects without explicitly setting an object class GUID, it MUST write all zeroes by default.
      If this value is not all zeroes, the object class GUID can be used as a parameter to start
      applications. }    CLSID: TGUID;    { State Bits (4 bytes): This field contains the user-defined flags if this entry is for a storage object or      root storage object. For a stream object, this field SHOULD be set to all zeroes because many
      implementations provide no way for applications to retrieve state bits from a stream object. If an
      implementation of the file format enables applications to create storage objects without explicitly
      setting state bits, it MUST write all zeroes by default. }    StateBits: UInt32;    { Creation Time (8 bytes): This field contains the creation time for a storage object, or all zeroes to      indicate that the creation time of the storage object was not recorded. The Windows FILETIME
      structure is used to represent this field in UTC. For a stream object, this field MUST be all zeroes.
      For a root storage object, this field MUST be all zeroes, and the creation time is retrieved or set on
      the compound file itself. }    CreationTime: UInt64;    { Modified Time (8 bytes): This field contains the modification time for a storage object, or all      zeroes to indicate that the modified time of the storage object was not recorded. The Windows
      FILETIME structure is used to represent this field in UTC. For a stream object, this field MUST be
      all zeroes. For a root storage object, this field MAY<2> be set to all zeroes, and the modified time
      is retrieved or set on the compound file itself. }    ModifiedTime: UInt64;    { Starting Sector Location (4 bytes): This field contains the first sector location if this is a stream      object. For a root storage object, this field MUST contain the first sector of the mini stream, if the
      mini stream exists. For a storage object, this field MUST be set to all zeroes. }    StartingSectorLocation: UInt32;    { Stream Size (8 bytes): This 64-bit integer field contains the size of the user-defined data if this is      a stream object. For a root storage object, this field contains the size of the mini stream. For a
      storage object, this field MUST be set to all zeroes.      For a version 3 compound file 512-byte sector size, the value of this field MUST be less than      or equal to 0x80000000. (Equivalently, this requirement can be stated: the size of a stream or
      of the mini stream in a version 3 compound file MUST be less than or equal to 2 gigabytes
      (GB).) Note that as a consequence of this requirement, the most significant 32 bits of this field
      MUST be zero in a version 3 compound file. However, implementers should be aware that
      some older implementations did not initialize the most significant 32 bits of this field, and
      these bits might therefore be nonzero in files that are otherwise valid version 3 compound
      files. Although this document does not normatively specify parser behavior, it is recommended
      that parsers ignore the most significant 32 bits of this field in version 3 compound files,
      treating it as if its value were zero, unless there is a specific reason to do otherwise (for
      example, a parser whose purpose is to verify the correctness of a compound file). }    StreamSize: UInt64;  end;

  { This object holds the content of a stream }
  TCFBStream = record
    Content  : TArray<Byte>;
    ContentName: string;
    FullName: string;
    StartingSector: UInt32;
    StorageName: string;
    StreamSize: UInt64;
  end;

  TCFBStreamArray = TArray<TCFBStream>;

{ function GetStreamList: returns the reference to the stream list objects }
function GetStreamList(): TCFBStreamArray;

{ function ProcessFileCFB: reads the specified file passed as argument
  It reads the sectors of the CFB and builds, but it doesn't parse the macros }
function ProcessFileCFB(const FileName: string): TProcessFileResult;

{ function ProcessFileZIP: reads the specified file passed as argument
  It reads the sectors of the CFB and builds, but it doesn't parse the macros }
function ProcessFileZIP(const FileName: string): TProcessFileResult;

implementation

uses
  System.SysUtils, System.Types, System.Zip, Vcl.Dialogs;

var
  CFBHeader        : TCFBHeader;
  CFBDiFATArray    : TArray<UInt32>;
  CFBDiFATLength   : UInt32;
  CFBFATArray      : TArray<UInt32>;
  CFBFATLength     : UInt32;
  CFBDirArray      : TArray<TDirectoryEntry>;
  CFBDirLength     : UInt32;
  CFBMiniFATArray  : TArray<UInt32>;
  CFBMiniFATLength : UInt32;
  CFBMiniStream    : TArray<Byte>;
  CFBStreamArray   : TCFBStreamArray;
  CFBStreamLength  : UInt32;
  CFBStreamCount   : UInt32;

procedure Reset();
begin
  Finalize(CFBHeader);
  CFBDiFATArray    := nil;
  CFBDiFATLength   := 0;
  CFBFATArray      := nil;
  CFBFATLength     := 0;
  CFBDirArray      := nil;
  CFBDirLength     := 0;
  CFBMiniFATArray  := nil;
  CFBMiniFATLength := 0;
  CFBMiniStream    := nil;
  CFBStreamArray   := nil;
  CFBStreamLength  := 0;
end;

function GetStreamList(): TCFBStreamArray;
begin
  Result := CFBStreamArray;
end;

{ function ReadSector: read a given sector, in an allocated buffer }
procedure ReadSector(Stream: TStream; SectorNumber: UInt32; Buffer: Pointer);
var
  SectorSize: Int32;
begin
  SectorSize := 1 shl CFBHeader.SectorShift;
  Inc(SectorNumber);
  Stream.Seek(SectorNumber shl CFBHeader.SectorShift, TSeekOrigin.soBeginning);
  Stream.ReadBuffer(Buffer^, SectorSize);
end;

{ function ReadMiniStream: reads a chained block of mini sectors }
function ReadMiniStream(Stream: TStream; StartingSectorLocation: UInt32; StreamSize: UInt32): TArray<Byte>;
var
  DataBuffer    : TArray<Byte>; // Result
  DataSize      : UInt32; // Calculated data size from FAT chain
  DataRemaining : UInt32; // Remaining bytes to read
  DataChunkSize : UInt32; // Bytes to read during the step
  DataOffset    : UInt32; // Pointer to the current reading location
  NextFATSector : UInt32; // Next chain mini sector
  SectorCount   : UInt32; // Calculated sector count from FAT chain
  SectorSize    : UInt32; // Calculated sector size in Byte
begin
  SectorSize  := 1 shl CFBHeader.MiniSectorShift;
  if StreamSize = 0 then
  begin
    SectorCount := 0;
    NextFATSector := StartingSectorLocation;
    while NextFATSector <> SectorENDOFCHAIN do
    begin
      Inc(SectorCount);
      NextFATSector := CFBMiniFATArray[NextFATSector];
    end;
    DataSize := SectorCount shl CFBHeader.MiniSectorShift;
  end
  else
    DataSize := StreamSize;
  SetLength(DataBuffer, DataSize);
  DataRemaining := DataSize;
  NextFATSector := StartingSectorLocation;
  DataOffset := 0;
  while NextFATSector <> SectorENDOFCHAIN do
  begin
    DataChunkSize := DataRemaining;
    if DataChunkSize > SectorSize then
      DataChunkSize := SectorSize;
    Move(CFBMiniStream[NextFATSector shl CFBHeader.MiniSectorShift], DataBuffer[DataOffset], DataChunkSize);
    Inc(DataOffset, DataChunkSize);
    Dec(DataRemaining, DataChunkSize);
    NextFATSector := CFBMiniFATArray[NextFATSector];
  end;
  Result := DataBuffer;
end;

{ function ReadStream: reads a chained block of sectors }
function ReadStream(Stream: TStream; StartingSectorLocation: UInt32; StreamSize: UInt32): TArray<Byte>;
var
  DataBuffer    : TArray<Byte>; // Result
  DataSize      : UInt32; // Calculated data size from FAT chain
  DataRemaining : UInt32; // Remaining bytes to read
  DataChunkSize : UInt32; // Bytes to read during the step
  DataOffset    : UInt32; // Pointer to the current reading location
  NextFATSector : UInt32; // Next chain sector
  SectorCount   : UInt32; // Calculated sector count from FAT chain
  SectorSize    : UInt32; // Calculated sector size in Byte
  SectorBuffer  : TArray<Byte>; // Reading buffer for last partial sector
begin
  SectorSize  := 1 shl CFBHeader.SectorShift;
  if StreamSize = 0 then
  begin
    SectorCount := 0;
    NextFATSector := StartingSectorLocation;
    while NextFATSector <> SectorENDOFCHAIN do
    begin
      Inc(SectorCount);
      NextFATSector := CFBFATArray[NextFATSector];
    end;
    DataSize := SectorCount shl CFBHeader.SectorShift;
  end
  else
    DataSize := StreamSize;
  SetLength(DataBuffer, DataSize);
  SetLength(SectorBuffer, SectorSize);
  DataRemaining := DataSize;
  NextFATSector := StartingSectorLocation;
  DataOffset := 0;
  while NextFATSector <> SectorENDOFCHAIN do
  begin
    DataChunkSize := DataRemaining;
    if DataChunkSize > SectorSize then
      DataChunkSize := SectorSize;
    ReadSector(Stream, NextFATSector, Pointer(SectorBuffer));
    Move(Pointer(SectorBuffer)^, DataBuffer[DataOffset], DataChunkSize);
    Inc(DataOffset, DataChunkSize);
    Dec(DataRemaining, DataChunkSize);
    NextFATSector := CFBFATArray[NextFATSector];
  end;
  Result := DataBuffer;
end;

function ReadCFB(Stream: TStream): TProcessFileResult;
var
  SectorData    : TArray<UInt32>;
  SectorSize    : UInt32; // Number of Bytes in a sector
  SectorLength  : UInt32; // Number of UInt32s in a sector
  DataSize      : Int32;
  NextFATSector : UInt32;
  I, N          : UInt32;
begin
  Stream.Position := 0;
  { Step 1: read and check the Compound File Header }
  if Stream.Read(CFBHeader, SizeOf(CFBHeader)) = SizeOf(CFBHeader) then
  begin
    if CFBHeader.Signature <> ValidSignature then
    begin
      Result := TProcessFileResult.pfHeaderSignatureError;
      Exit;
    end;
    if CFBHeader.HeaderCLSID <> GUID_NULL then
    begin
      Result := TProcessFileResult.pfHeaderGUIDError;
      Exit;
    end;
    SectorSize := 1 shl CFBHeader.SectorShift;
    SectorLength := SectorSize shr 2;
    SetLength(SectorData, SectorLength);
    { Step 2: read the DIFAT }
    N := 0;
    DataSize := CFBHeader.NumberOfDIFATSectors shl CFBHeader.SectorShift; // NumberOfDIFATSectors * SectorSize;
    CFBDiFATLength := DataSize shr 2 + 109;
    SetLength(CFBDiFATArray, CFBDiFATLength);
    for I := 0 to 108 do
    begin
      CFBDiFATArray[N] := CFBHeader.DIFAT[I];
      Inc(N);
    end;
    NextFATSector := CFBHeader.FirstDIFATSectorLocation;
    while NextFATSector <> SectorENDOFCHAIN do
    begin
      ReadSector(Stream, NextFATSector, SectorData);
      for I := 0 to SectorLength - 2 do
      begin
        CFBDiFATArray[N] := SectorData[I];
        Inc(N);
      end;
      NextFATSector := SectorData[SectorLength - 1];
    end;
    { Step 3: read the FAT }
    N := 0;
    DataSize := CFBHeader.NumberOfFATSectors shl CFBHeader.SectorShift; // NumberOfFATSectors * SectorSize;
    CFBFATLength := DataSize shr 2;
    SetLength(CFBFATArray, CFBFATLength);
    for I := 0 to CFBHeader.NumberOfFATSectors - 1 do
    begin
      ReadSector(Stream, CFBDiFATArray[I], @CFBFATArray[N]);
      Inc(N, SectorSize shr 2);
    end;
    { Step 4: read the Directory Entry Array }
    N := 0;
    NextFATSector := CFBHeader.FirstDirectorySectorLocation;
    while NextFATSector <> SectorENDOFCHAIN do
    begin
      Inc(N);
      NextFATSector := CFBFATArray[NextFATSector];
    end;
    if CFBHeader.MajorVersion = $0004 then
      if N <> CFBHeader.NumberOfDirectorySectors then
      begin
        Result := TProcessFileResult.pfHeaderDirectorySectorNumberError;
        Exit;
      end;
    DataSize := N shl CFBHeader.SectorShift; // NumberOfDirectorySectors * SectorSize;
    CFBDirLength := DataSize shr 7;
    CFBDirArray := TArray<TDirectoryEntry>(ReadStream(Stream, CFBHeader.FirstDirectorySectorLocation, DataSize));
    SetLength(CFBDirArray, CFBDirLength);
    { Step 5: read the mini FAT }
    N := 0;
    DataSize := CFBHeader.NumberOfMiniFATSectors shl CFBHeader.SectorShift; // NumberOfMiniFATSectors * SectorSize;
    CFBMiniFATLength := DataSize shr 2;
    SetLength(CFBMiniFATArray, CFBMiniFATLength);
    NextFATSector := CFBHeader.FirstMiniFATSectorLocation;
    while NextFATSector <> SectorENDOFCHAIN do
    begin
      ReadSector(Stream, NextFATSector, @CFBMiniFATArray[N]);
      Inc(N, SectorSize shr 2);
      NextFATSector := CFBFATArray[NextFATSector];
    end;
    { Step 6: read the mini stream }
    CFBMiniStream := TArray<Byte>(ReadStream(Stream, CFBDirArray[0].StartingSectorLocation, 0));
    Result := TProcessFileResult.pfOk;
  end
  else
    Result := TProcessFileResult.pfHeaderSizeError;
end;

procedure AppendStreamList(Stream: TStream; const StorageName: string; NewIndex: UInt32);
var
  ContentName : string;
begin
  if NewIndex <> SectorFREESECT then
  begin
    SetString(ContentName, CFBDirArray[NewIndex].DirectoryEntryName, CFBDirArray[NewIndex].DirectoryEntryNameLength shr 1 - 1);
    case CFBDirArray[NewIndex].ObjectType of
      $01:
      begin
        AppendStreamList(Stream, StorageName + '\' + ContentName, CFBDirArray[NewIndex].ChildID);
        { Recursively call AppendStreamList for the two siblings }
        AppendStreamList(Stream, StorageName, CFBDirArray[NewIndex].LeftSiblingID);
        AppendStreamList(Stream, StorageName, CFBDirArray[NewIndex].RightSiblingID);
      end;
      $02:
      begin
        CFBStreamArray[CFBStreamCount].ContentName := ContentName;
        CFBStreamArray[CFBStreamCount].StorageName := StorageName;
        CFBStreamArray[CFBStreamCount].FullName := StorageName + '\' + ContentName;
        CFBStreamArray[CFBStreamCount].StartingSector := CFBDirArray[NewIndex].StartingSectorLocation;
        CFBStreamArray[CFBStreamCount].StreamSize := CFBDirArray[NewIndex].StreamSize;
        Inc(CFBStreamCount);
        { Recursively call AppendStreamList for the two siblings }
        AppendStreamList(Stream, StorageName, CFBDirArray[NewIndex].LeftSiblingID);
        AppendStreamList(Stream, StorageName, CFBDirArray[NewIndex].RightSiblingID);
      end;
    end;
  end;
end;

procedure BuildStreamList(Stream: TStream);
var
  StreamIndex : UInt32;
begin
  { Initialize the array of streams }
  CFBStreamCount  := 0;
  CFBStreamLength := 0;
  for StreamIndex := 0 to CFBDirLength - 1 do
    if CFBDirArray[StreamIndex].ObjectType = $02 then
      Inc(CFBStreamLength);
  SetLength(CFBStreamArray, CFBStreamLength);
  { Start with the root entry, use AppendStreamList, to add the stream objects recursively }
  AppendStreamList(Stream, '', CFBDirArray[0].ChildID);
  { Read the streams }
  for StreamIndex := 0 to CFBStreamLength - 1 do
    if CFBStreamArray[StreamIndex].StorageName.IndexOf('VBA') >= 0 then
      if CFBStreamArray[StreamIndex].StreamSize <= CFBHeader.MiniStreamCutoffSize then
        CFBStreamArray[StreamIndex].Content := ReadMiniStream(Stream, CFBStreamArray[StreamIndex].StartingSector, CFBStreamArray[StreamIndex].StreamSize)
      else
        CFBStreamArray[StreamIndex].Content := ReadStream(Stream, CFBStreamArray[StreamIndex].StartingSector, CFBStreamArray[StreamIndex].StreamSize);
end;

function ProcessStream(Stream: TStream): TProcessFileResult;
begin
  Reset();
  Result := ReadCFB(Stream);
  if Result = TProcessFileResult.pfOk then
    BuildStreamList(Stream);
end;

function ProcessFileCFB(const FileName: string): TProcessFileResult;
var
  FileExtension : string;
  Stream        : TStream;
begin
  Result := TProcessFileResult.pfInvalidFileContent;
  FileExtension := ExtractFileExt(FileName).ToUpper();
  if (FileExtension = '.BIN')
  or (FileExtension = '.DOC')
  or (FileExtension = '.OTM')
  or (FileExtension = '.XLS') then
  begin
    Stream := TFileStream.Create(FileName, fmOpenRead);
    try
      ProcessFileCFB := ProcessStream(Stream);
    finally
      Stream.Free();
    end;
  end;
end;

function ProcessFileZIP(const FileName: string): TProcessFileResult;
var
  CompressedPath : string;
  FileExtension  : string;
  FileIndex      : Int32;
  MemoryStream   : TMemoryStream;
  ZipFile        : System.Zip.TZipFile;
  ZipHeader      : System.Zip.TZipHeader;
  ZipStream      : TStream;
begin
  Result := TProcessFileResult.pfInvalidFileContent;
  CompressedPath := '';
  FileExtension := ExtractFileExt(FileName).ToUpper();
  if (FileExtension = '.DOCM')
  or (FileExtension = '.DOCX')
  or (FileExtension = '.DOTM')
  or (FileExtension = '.DOTX') then
    CompressedPath := 'Word/vbaProject.bin'
  else
    if (FileExtension = '.XLAM')
    or (FileExtension = '.XLSB')
    or (FileExtension = '.XLSM')
    or (FileExtension = '.XLSX') then
      CompressedPath := 'xl/vbaProject.bin'
    else
      if (FileExtension = '.POTM')
      or (FileExtension = '.POTX')
      or (FileExtension = '.PPTM')
      or (FileExtension = '.PPTX') then
        CompressedPath := 'ppt/vbaProject.bin';
  if CompressedPath <> '' then
  begin
    MemoryStream := nil;
    ZipFile      := nil;
    ZipStream    := nil;
    try
      MemoryStream := TMemoryStream.Create();
      ZipFile      := TZipFile.Create();
      ZipFile.Open(FileName, TZipMode.zmRead);
      FileIndex := ZipFile.IndexOf(CompressedPath);
      if FileIndex >= 0 then
      begin
        ZipFile.Read(FileIndex, ZipStream, ZipHeader);
        MemoryStream.SetSize(ZipStream.Size);
        MemoryStream.CopyFrom(ZipStream, ZipStream.Size);
        Result := ProcessStream(MemoryStream);
      end;
      ZipFile.Close();
    finally
      MemoryStream.Free();
      ZipFile.Free();
      ZipStream.Free();
    end;
  end;
end;

initialization
  Reset();

finalization
  Reset();

end.
