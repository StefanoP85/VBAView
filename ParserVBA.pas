unit ParserVBA;
{$ALIGN 1}

interface

uses
  System.Generics.Collections, FilesCFB, Common;

function ParseFile(const FileName: string; var VBAProgram: TVBAProgram): TProcessFileResult;

implementation

uses
  Winapi.Windows, System.SysUtils, ParserPCode;

const
  MainStreamName = 'dir';
  ProjectStreamName = '_VBA_PROJECT';

procedure Reset(var VBAProgram: TVBAProgram);
begin
  Finalize(VBAProgram.Module);
  Finalize(VBAProgram.Reference);
  Finalize(VBAProgram);
end;

{ 2.4.1.3.1 Decompression Algorithm
  2.4.1.1.1 CompressedContainer

A CompressedContainer is an array of bytes holding the compressed data. The Decompression
algorithm (section 2.4.1.3.1) processes a CompressedContainer to populate a
DecompressedBuffer. The Compression algorithm (section 2.4.1.3.6) processes a
DecompressedBuffer to produce a CompressedContainer.
A CompressedContainer MUST be the last array of bytes in a stream. On read, the end of stream
indicator determines when the entire CompressedContainer has been read.
The CompressedContainer is a SignatureByte followed by array of CompressedChunk (section
2.4.1.1.4) structuresSignatureByte (1 byte): Specifies the beginning of the CompressedContainer. MUST be 0x01. TheDecompression algorithm (section 2.4.1.3.1) reads SignatureByte. The Compression
algorithm (section 2.4.1.3.6) writes SignatureByte.Chunks (variable): An array of CompressedChunk (section 2.4.1.1.4) records. Specifies thecompressed data. Read by the Decompression algorithm. Written by the Compression
algorithm. }function DecompressionAlgorithm(const CompressedContainer: TArray<Byte>; CompressedRecordEnd: UInt32): TArray<Byte>;
var
  { 2.4.1.2 State Variables
    The following state is maintained for the CompressedContainer (section 2.4.1.1.1). }
  { CompressedRecordEnd: The location of the byte after the last byte in the CompressedContainer
    (section 2.4.1.1.1). }
  { CompressedRecordEnd    : UInt32; }
  { CompressedCurrent: The location of the next byte in the CompressedContainer (section
    2.4.1.1.1) to be read by decompression or to be written by compression. }
  CompressedCurrent      : UInt32;
  { The following state is maintained for the current CompressedChunk (section 2.4.1.1.4). }
  { CompressedChunkStart: The location of the first byte of the CompressedChunk (section
    2.4.1.1.4) within the CompressedContainer (section 2.4.1.1.1). }
  CompressedChunkStart   : UInt32;
  { The following state is maintained for a DecompressedBuffer (section 2.4.1.1.2). }
  { DecompressedCurrent: The location of the next byte in the DecompressedBuffer (section
    2.4.1.1.2) to be written by decompression or to be read by compression. }
  DecompressedCurrent    : UInt32;
  { DecompressedBufferEnd: The location of the byte after the last byte in the DecompressedBuffer
    (section 2.4.1.1.2). }
  DecompressedBufferEnd  : UInt32;
  { The following state is maintained for the current DecompressedChunk (section 2.4.1.1.3). }
  { DecompressedChunkStart : The location of the first byte of the DecompressedChunk (section
    2.4.1.1.3) within the DecompressedBuffer (section 2.4.1.1.2). }
  DecompressedChunkStart : UInt32;
  { Other variables }
  DecompressedBuffer     : TArray<Byte>;
  CompressedChunkHeader  : UInt16;
  Size                   : UInt16;
  CompressedFlag         : UInt16;
  CompressedEnd          : UInt32;
  FlagByte               : Byte;
  Index                  : UInt16;
  Flag                   : Byte;
  CopyToken              : UInt16;
  Offset                 : UInt32;
  Length                 : UInt16;
  CopySource             : UInt32;

{ 2.4.1.3.11 Byte Copy }
procedure ByteCopy(CopySource: UInt32; DestinationSource: UInt32; ByteCount: UInt16);
var
  Counter    : UInt16;
  DstCurrent : UInt32;
  SrcCurrent : UInt32;
begin
  { SET SrcCurrent TO CopySource }
  SrcCurrent := CopySource;
  { SET DstCurrent TO DestinationSource }
  DstCurrent := DestinationSource;
  { FOR counter FROM 1 TO ByteCount INCLUSIVE }
  for Counter := 1 to ByteCount do
  begin
  { COPY the byte at SrcCurrent TO DstCurrent }
    DecompressedBuffer[DstCurrent] := DecompressedBuffer[SrcCurrent];
  { INCREMENT SrcCurrent }
    Inc(SrcCurrent);
  { INCREMENT DstCurrent }
    Inc(DstCurrent);
  { ENDFOR }
  end;
end;

{ 2.4.1.3.13 Pack CompressedChunkSize }
function ExtractCompressedChunkFlag(CompressedChunkHeader: UInt16): UInt16;
begin
  Result := CompressedChunkHeader and $8000 shr 15;
end;

{ 2.4.1.3.12 Extract CompressedChunkSize }
function ExtractCompressedChunkSize(CompressedChunkHeader: UInt16): UInt16;
begin
  Result := CompressedChunkHeader and $0FFF + 3;
end;

{ 2.4.1.3.17 Extract FlagBit }
function ExtractFlagBit(Index: UInt16; FlagByte: Byte): UInt16;
begin
  Result := (FlagByte shr Index) and $01;
end;

{ 2.4.1.3.19.1 CopyToken Help }
procedure CopyTokenHelp(out LengthMask: UInt16; out OffsetMask: UInt16; out BitCount: UInt16; out MaximumLength: UInt16);
var
  Difference : UInt32;
begin
  { SET difference TO DecompressedCurrent MINUS DecompressedChunkStart }
  Difference := DecompressedCurrent - DecompressedChunkStart;
  { SET BitCount TO the smallest integer that is GREATER THAN OR EQUAL TO LOGARITHM base 2 of
difference }
  BitCount   := 0;
  while 1 shl BitCount < Difference do
    Inc(BitCount);
  { SET BitCount TO the maximum of BitCount and 4 }
  if BitCount < 4 then
    BitCount := 4;
  { SET LengthMask TO 0xFFFF RIGHT SHIFT BY BitCount }
  LengthMask := $FFFF shr BitCount;
  { SET OffsetMask TO BITWISE NOT LengthMask }
  OffsetMask := not LengthMask;
  { SET MaximumLength TO (0xFFFF RIGHT SHIFT BY BitCount) PLUS 3 }
  MaximumLength := $FFFF shr BitCount + 3;
end;

{ 2.4.1.3.19.2 Unpack CopyToken }
procedure UnpackCopyToken(Token: UInt16; out Offset: UInt32; out Length: UInt16);
var
  BitCount      : UInt16;
  LengthMask    : UInt16;
  OffsetMask    : UInt16;
  MaximumLength : UInt16;
  Temp1, Temp2  : UInt16;
begin
  { CALL CopyToken Help (section 2.4.1.3.19.1) returning LengthMask, OffsetMask, and
    BitCount. }
  CopyTokenHelp(LengthMask, OffsetMask, BitCount, MaximumLength);
  { SET Length TO (Token BITWISE AND LengthMask) PLUS 3. }
  Length := (Token and LengthMask) + 3;
  { SET temp1 TO Token BITWISE AND OffsetMask. }
  Temp1 := Token and OffsetMask;
  { SET temp2 TO 16 MINUS BitCount. }
  Temp2 := 16 - BitCount;
  { SET Offset TO (temp1 RIGHT SHIFT BY temp2) PLUS 1. }
  Offset := (Temp1 shr Temp2) + 1;
end;

begin
  CompressedCurrent   := 0;
  DecompressedCurrent := 0;
  DecompressedBufferEnd := 0;
  { IF the byte located at CompressedCurrent EQUALS 0x01 THEN }
  if CompressedContainer[CompressedCurrent] = $01 then
  begin
  { INCREMENT CompressedCurrent }
    CompressedCurrent := CompressedCurrent + 1;
  { WHILE CompressedCurrent is LESS THAN CompressedRecordEnd }
    while CompressedCurrent < CompressedRecordEnd do
    begin
      DecompressedBufferEnd := DecompressedBufferEnd + 4096;
      SetLength(DecompressedBuffer, DecompressedBufferEnd);
  { SET CompressedChunkStart TO CompressedCurrent }
      CompressedChunkStart := CompressedCurrent;
  { CALL Decompressing a CompressedChunk }
  { SET Header TO the CompressedChunkHeader (section 2.4.1.1.5) located at CompressedChunkStart }
      Move(CompressedContainer[CompressedChunkStart], CompressedChunkHeader, SizeOf(CompressedChunkHeader));
  { CALL Extract CompressedChunkSize (section 2.4.1.3.12) with Header returning Size }
      Size := ExtractCompressedChunkSize(CompressedChunkHeader);
  { CALL Extract CompressedChunkFlag (section 2.4.1.3.15) with Header returning CompressedFlag }
      CompressedFlag := ExtractCompressedChunkFlag(CompressedChunkHeader);
  { SET DecompressedChunkStart TO DecompressedCurrent }
      DecompressedChunkStart := DecompressedCurrent;
  { SET CompressedEnd TO the minimum of CompressedRecordEnd and (CompressedChunkStart PLUS Size) }
      if CompressedRecordEnd < CompressedChunkStart + Size then
        CompressedEnd := CompressedRecordEnd
      else
        CompressedEnd := CompressedChunkStart + Size;
  { SET CompressedCurrent TO CompressedChunkStart PLUS 2 }
      CompressedCurrent := CompressedChunkStart + 2;
  { IF CompressedFlag EQUALS 1 THEN }
      if CompressedFlag = $01 then
      begin
  { WHILE CompressedCurrent is LESS THAN CompressedEnd }
        while CompressedCurrent < CompressedEnd do
        begin
  { CALL Decompressing a TokenSequence (section 2.4.1.3.4) with CompressedEnd }
  { SET Byte TO the FlagByte (section 2.4.1.1.7) located at CompressedCurrent }
          FlagByte := CompressedContainer[CompressedCurrent];
  { INCREMENT CompressedCurrent }
          CompressedCurrent := CompressedCurrent + 1;
  { IF CompressedCurrent is LESS THAN CompressedEnd THEN }
          if CompressedCurrent < CompressedEnd then
          begin
  { FOR index FROM 0 TO 7 INCLUSIVE }
            for Index := 0 to 7 do
            begin
  { IF CompressedCurrent is LESS THAN CompressedEnd THEN }
              if CompressedCurrent < CompressedEnd then
              begin
  { CALL Decompressing a Token (section 2.4.1.3.5) with index and Byte }
  { CALL Extract FlagBit (section 2.4.1.3.17) with index and Byte returning Flag }
                Flag := ExtractFlagBit(Index, FlagByte);
  { IF Flag EQUALS 0 THEN }
                if Flag = $00 then
                begin
  { COPY the byte at CompressedCurrent TO DecompressedCurrent }
                  DecompressedBuffer[DecompressedCurrent] := CompressedContainer[CompressedCurrent];
  { INCREMENT DecompressedCurrent }
                  DecompressedCurrent := DecompressedCurrent + 1;
  { INCREMENT CompressedCurrent }
                  CompressedCurrent := CompressedCurrent + 1;
  { ELSE }
                end
                else
                begin
  { SET Token TO the CopyToken (section 2.4.1.1.8) at CompressedCurrent }
                  Move(CompressedContainer[CompressedCurrent], CopyToken, SizeOf(CopyToken));
  { CALL Unpack CopyToken (section 2.4.1.3.19.2) with Token returning Offset and Length }
                  UnpackCopyToken(CopyToken, Offset, Length);
  { SET CopySource TO DecompressedCurrent MINUS Offset }
                  CopySource := DecompressedCurrent - Offset;
  { CALL Byte Copy (section 2.4.1.3.11) with CopySource, DecompressedCurrent, and Length }
                  ByteCopy(CopySource, DecompressedCurrent, Length);
  { INCREMENT DecompressedCurrent BY Length }
                  DecompressedCurrent := DecompressedCurrent + Length;
  { INCREMENT CompressedCurrent BY 2 }
                  CompressedCurrent := CompressedCurrent + 2;
  { ENDIF }                end;  { ENDIF }
              end;
  { ENDFOR }
            end;
  { ENDIF }
          end;
  { END WHILE }
        end;
  { ELSE }
      end
      else
      begin
  { CALL Decompressing a RawChunk (section 2.4.1.3.3) }
  { APPEND 4096 bytes from CompressedCurrent TO DecompressedCurrent }
        Move(CompressedContainer[CompressedCurrent], DecompressedBuffer[DecompressedCurrent], 4096);
  { INCREMENT DecompressedCurrent BY 4096 }
        DecompressedCurrent := DecompressedCurrent + 4096;
  { INCREMENT CompressedCurrent BY 4096 }
        CompressedCurrent := CompressedCurrent + 4096;
  { ENDIF }
      end;
    end;
    SetLength(DecompressedBuffer, DecompressedCurrent);
  end
  else
    DecompressedBuffer := nil;
  Result := DecompressedBuffer;
end;

function ParseDirStream(var DirStream: TVBAProgram): TProcessFileResult;
var
  { Stream management }
  StreamArray        : FilesCFB.TCFBStreamArray;
  StreamCount        : UInt32;
  StreamIndex        : UInt32;
  CompressedStream   : TArray<Byte>;
  DecompressedStream : TArray<Byte>;
  { Code page of MBCS }
  CodePage           : UInt32;
  { Offset of current parsing position }
  StreamOffset       : Int32;
  { Offset of last Byte of the decompressed stream }
  StreamEnd          : Int32;
  { Id is always the first field of a record }
  RecordId           : UInt16;
  { Size is always the second field of a record }
  RecordSize         : UInt32;
  { Reference structure }
  ReferenceCount     : UInt32;
  ReferencePointer   : PReference;
  ReferenceSize      : UInt32;
  ModulePointer      : PModule;

function ReadBYTE(): Byte;
begin
  Move(Result, DecompressedStream[StreamOffset], SizeOf(Result));
  Inc(StreamOffset, SizeOf(Byte));
end;

function ReadWORD(): UInt16;
begin
  Move(DecompressedStream[StreamOffset], Result, SizeOf(Result));
  Inc(StreamOffset, SizeOf(UInt16));
end;

function ReadDWORD(): UInt32;
begin
  Move(DecompressedStream[StreamOffset], Result, SizeOf(Result));
  Inc(StreamOffset, SizeOf(UInt32));
end;

function ReadQWORD(): UInt64;
begin
  Move(DecompressedStream[StreamOffset], Result, SizeOf(Result));
  Inc(StreamOffset, SizeOf(UInt64));
end;

function ReadString(CodePage: UInt32; NumberOfBytes: UInt32): string;
var
  StringBuffer: TArray<Byte>;
begin
  Result := '';
  if NumberOfBytes > 0 then
  begin
    StringBuffer := Copy(DecompressedStream, StreamOffset, NumberOfBytes);
    Inc(StreamOffset, NumberOfBytes);
    Result := AnsiString2UnicodeString(CodePage, StringBuffer, NumberOfBytes);
  end;
end;

begin
  { The dir stream is required }
  Result := TProcessFileResult.pfInvalidFileContent;
  StreamArray := FilesCFB.GetStreamList();
  if StreamArray <> nil then
  begin
    StreamCount := High(StreamArray);
    for StreamIndex := 0 to StreamCount do
      if (StreamArray[StreamIndex].StorageName.IndexOf('VBA') >= 0) and (StreamArray[StreamIndex].ContentName = MainStreamName) then
        if StreamArray[StreamIndex].Content <> nil then
        begin
          CompressedStream := StreamArray[StreamIndex].Content;
          DecompressedStream := DecompressionAlgorithm(CompressedStream, StreamArray[StreamIndex].StreamSize);
          { Process the dir stream }
          CodePage         := 0;
          StreamOffset     := 0;
          StreamEnd        := High(DecompressedStream);
          ReferenceCount   := 0;
          ReferencePointer := nil;
          ModulePointer    := nil;
          while StreamOffset <= StreamEnd do
          begin
            { Step 1: get the RecordId }
            RecordId := ReadWORD();
            { Step 2: get the RecordId }
            RecordSize := ReadDWORD();
            { Step 3: parse the RecordId field }
            case RecordId of
              $0001: { SysKindRecord }
                DirStream.SysKind := ReadDWORD();
              $0002: { LcidRecord }
                DirStream.Lcid := ReadDWORD();
              $0014: { LcidInvokeRecord }
                DirStream.LcidInvoke := ReadDWORD();
              $0003: { CodePageRecord }
              begin
                DirStream.CodePage := ReadWORD();
                CodePage := DirStream.CodePage;
              end;
              $0004: { NameRecord }
                DirStream.ProjectName := ReadString(CodePage, RecordSize);
              $0005: { DocStringRecord }
              begin
                DirStream.DocString := ReadString(CodePage, RecordSize);
                Inc(StreamOffset, 2); { Skip reserved space }
                ReferenceSize := ReadDWORD(); { Size of Unicode version }
                Inc(StreamOffset, ReferenceSize);
              end;
              $0006: { HelpFilePathRecord 1 }
                DirStream.HelpFile1 := ReadString(CodePage, RecordSize);
              $003D: { HelpFilePathRecord 2 }
                DirStream.HelpFile2 := ReadString(CodePage, RecordSize);
              $0007: { HelpContextRecord }
                DirStream.HelpContext := ReadDWORD();
              $0008: { LibFlagsRecord }
                DirStream.ProjectLibFlags := ReadDWORD();
              $0009: { VersionRecord }
              begin
                DirStream.VersionMajor := ReadDWORD();
                DirStream.VersionMinor := ReadWORD();
              end;
              $000C: { ConstantsRecord }
              begin
                DirStream.Constants := ReadString(CodePage, RecordSize);
                Inc(StreamOffset, 2); { Skip reserved space }
                ReferenceSize := ReadDWORD(); { Size of Unicode version }
                Inc(StreamOffset, ReferenceSize);
              end;
              $0016: { Reference name }
              begin
                SetLength(DirStream.Reference, ReferenceCount + 1);
                ReferencePointer := @DirStream.Reference[ReferenceCount];
                Inc(ReferenceCount);
                ReferencePointer^.ReferenceName := ReadString(CodePage, RecordSize);
                Inc(StreamOffset, 2); { Skip reserved space }
                ReferenceSize := ReadDWORD(); { Size of Unicode version }
                Inc(StreamOffset, ReferenceSize);
              end;
              $002F: { Reference Control }
              begin
                ReferenceSize := ReadDWORD();
                ReferencePointer^.ReferenceControl := ReadString(CodePage, ReferenceSize);
                Inc(StreamOffset, 6); { Skip reserved space }
                RecordId := ReadWORD();
                RecordSize := ReadDWORD();
                if RecordId = $0016 then { Reference name Extended }
                begin
                  ReferencePointer^.ExtendedName := ReadString(CodePage, RecordSize);
                  Inc(StreamOffset, 2); { Skip reserved space }
                  ReferenceSize := ReadDWORD(); { Size of Unicode version }
                  Inc(StreamOffset, ReferenceSize);
                end;
              end;
              $0030: { Reference extended }
              begin
                ReferenceSize := ReadDWORD();
                ReferencePointer^.ExtendedLibrary := ReadString(CodePage, ReferenceSize);
                Inc(StreamOffset, 6); { Skip reserved space }
                Move(ReferencePointer^.GUID, DecompressedStream[StreamOffset], SizeOf(ReferencePointer^.GUID));
                Inc(StreamOffset, SizeOf(ReferencePointer^.GUID));
                Inc(StreamOffset, 4); { Skip reserved space, Cookie }
              end;
              $0033: { Reference Original }
                ReferencePointer^.ReferenceOriginal := ReadString(CodePage, RecordSize);
              $000D: { Reference Registered }
              begin
                ReferenceSize := ReadDWORD();
                ReferencePointer^.ReferenceRegistered := ReadString(CodePage, ReferenceSize);
                Inc(StreamOffset, 6); { Skip reserved space }
              end;
              $000E: { Reference Project }
              begin
                ReferenceSize := ReadDWORD();
                ReferencePointer^.ReferenceProject := ReadString(CodePage, ReferenceSize);
                ReferenceSize := ReadDWORD();
                Inc(StreamOffset, ReferenceSize);
                ReferencePointer^.MajorVersion := ReadDWORD();
                ReferencePointer^.MinorVersion := ReadWORD();
              end;
              $000F: { Modules }
              begin
                DirStream.ModulesCount := ReadWORD();
                SetLength(DirStream.Module, DirStream.ModulesCount);
                ModulePointer := nil;
                Inc(StreamOffset, 8); { Skip reserved space, Cookie }
              end;
              $0019: { Module name }
              begin
                if ModulePointer = nil then
                  ModulePointer := @DirStream.Module[0]
                else
                  Inc(ModulePointer);
                ModulePointer^.ModuleName := ReadString(CodePage, RecordSize);
                ModulePointer^.ModuleReadOnly := False;
                ModulePointer^.ModulePrivate := False;
              end;
              $0047: { Module name in Unicode }
                Inc(StreamOffset, RecordSize);
              $001A: { Module stream name }
                ModulePointer^.ModuleStreamName := ReadString(CodePage, RecordSize);
              $0032: { Module stream name in Unicode }
                Inc(StreamOffset, RecordSize);
              $001C: { Module doc string }
                ModulePointer^.DocString := ReadString(CodePage, RecordSize);
              $0048: { Module doc string in Unicode }
                Inc(StreamOffset, RecordSize);
              $0031: { Module offset, this is VERY IMPORTANT!! }
                ModulePointer^.TextOffset := ReadDWORD();
              $001E: { Module help context }
                ModulePointer^.HelpContext := ReadDWORD();
              $002C: { Module cookie }
                Inc(StreamOffset, RecordSize);
              $0021: { Module type standard }
                ModulePointer^.ModuleType := TModuleType.ModuleStandard;
              $0022: { Module type class }
                ModulePointer^.ModuleType := TModuleType.ModuleClass;
              $0025: { Module read only }
                ModulePointer^.ModuleReadOnly := True;
              $0028: { Module private }
                ModulePointer^.ModulePrivate := True;
              $0010: { Terminator of dir stream }
                Continue;
              $002B: { Terminator of Module record }
                Continue;
              $004A: { Mystery field?? }
                Inc(StreamOffset, 4);
              else { Unknown field }
                raise EVBAParseError.Create('Unknown data found in the VBA main stream: ' + IntToStr(RecordId));
            end;
          end;
          { Only one stream can have the name 'dir' in the VBA storage, so quit here }
          Result := TProcessFileResult.pfOk;
          Break;
        end;
  end;
end;

function ParseSourceCode(const DecompressedText: string): string;
var
  StringBuilder : TStringBuilder;
  TextLineArray : TArray<string>;
  TextLine      : string;
  I             : UInt32;
begin
  Result := '';
  if DecompressedText.Length > 0 then
  begin
    StringBuilder := TStringBuilder.Create();
    try
      TextLineArray := DecompressedText.Split([#10]);
      for I := 0 to High(TextLineArray) do
      begin
        TextLine := TextLineArray[I];
        if not TextLine.Trim().StartsWith('Attribute') then
          StringBuilder.Append(TextLine);
      end;
      Result := StringBuilder.ToString();
    finally
      StringBuilder.Free();
    end;
  end;
end;

procedure ParseModuleStream(CodePage: UInt32; var Module: TModule);
var
  { Stream management }
  StreamArray        : FilesCFB.TCFBStreamArray;
  StreamCount        : UInt32;
  StreamIndex        : UInt32;
  CompressedSize     : UInt32;
  CompressedStream   : TArray<Byte>;
  DecompressedStream : TArray<Byte>;
  DecompressedText   : string;
begin
  Module.PerformanceCache := nil;
  Module.SourceCode := '';
  StreamArray := FilesCFB.GetStreamList();
  if StreamArray <> nil then
  begin
    StreamCount := High(StreamArray);
    for StreamIndex := 0 to StreamCount do
      if (StreamArray[StreamIndex].StorageName.IndexOf('VBA') >= 0) and (StreamArray[StreamIndex].ContentName = Module.ModuleStreamName) then
      begin
        if StreamArray[StreamIndex].Content <> nil then
        begin
          { This statement works even if Module.TextOffset is zero! }
          CompressedSize := StreamArray[StreamIndex].StreamSize - Module.TextOffset;
          CompressedStream := Copy(StreamArray[StreamIndex].Content, Module.TextOffset, CompressedSize);
          if Module.TextOffset > 0 then
            Module.PerformanceCache := Copy(StreamArray[StreamIndex].Content, 0, Module.TextOffset);
          DecompressedStream := DecompressionAlgorithm(CompressedStream, CompressedSize);
          SetString(DecompressedText, PAnsiChar(@DecompressedStream[0]), High(DecompressedStream));
          Module.SourceCode := ParseSourceCode(DecompressedText);
        end;
        Break;
      end;
  end;
end;

function ParseProjectStream(var VBAProgram: TVBAProgram): TProcessFileResult;
var
  StreamArray : FilesCFB.TCFBStreamArray;
  StreamCount : UInt32;
  StreamIndex : UInt32;
begin
  { The dir stream is required }
  Result := TProcessFileResult.pfInvalidFileContent;
  StreamArray := FilesCFB.GetStreamList();
  if StreamArray <> nil then
  begin
    StreamCount := High(StreamArray);
    for StreamIndex := 0 to StreamCount do
      if (StreamArray[StreamIndex].StorageName.IndexOf('VBA') >= 0) and (StreamArray[StreamIndex].ContentName = ProjectStreamName) then
        if StreamArray[StreamIndex].Content <> nil then
        begin
          VBAProgram.VbaProjectData := StreamArray[StreamIndex].Content;
          Move(VBAProgram.VbaProjectData[2], VBAProgram.Version, SizeOf(VBAProgram.Version));
          Result := TProcessFileResult.pfOk;
          Break;
        end;
  end;
end;

function ParseFile(const FileName: string; var VBAProgram: TVBAProgram): TProcessFileResult;
var
  FileExtension : string;
  I             : Integer;
begin
  FileExtension := ExtractFileExt(FileName).ToUpper();
  if (FileExtension = '.DOCM') or (FileExtension = '.DOCX')
  or (FileExtension = '.DOTM') or (FileExtension = '.DOTX')
  or (FileExtension = '.XLAM') or (FileExtension = '.XLSB')
  or (FileExtension = '.XLSM') or (FileExtension = '.XLSX')
  or (FileExtension = '.POTM') or (FileExtension = '.POTX')
  or (FileExtension = '.PPTM') or (FileExtension = '.PPTX') then
    Result := FilesCFB.ProcessFileZIP(FileName)
  else
    if (FileExtension = '.BIN') or (FileExtension = '.DOC')
    or (FileExtension = '.OTM') or (FileExtension = '.XLS') then
      Result := FilesCFB.ProcessFileCFB(FileName)
    else
      Result := TProcessFileResult.pfInvalidFileExtension;
  Reset(VBAProgram);
  if Result = TProcessFileResult.pfOk then
  begin
    { Step 1: parse the 'dir' stream }
    Result := ParseDirStream(VBAProgram);
    if Result = TProcessFileResult.pfOk then
    begin
      { Step 2: parse the '_VBA_PROJECT' stream }
      Result := ParseProjectStream(VBAProgram);
      if Result = TProcessFileResult.pfOk then
        for I := 0 to VBAProgram.ModulesCount - 1 do
        begin
          { Step 3: read the modules streams }
          ParseModuleStream(VBAProgram.CodePage, VBAProgram.Module[I]);
          { Step 4: parse the p-code }
          ParsePCode(VBAProgram, VBAProgram.Module[I]);
        end;
    end;
  end;
end;

end.
