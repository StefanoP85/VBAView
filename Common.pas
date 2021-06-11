unit Common;

interface

uses System.SysUtils;

type
  TProcessFileResult = (
    pfOk = 0,
    pfInvalidFileExtension,
    pfInvalidFileContent,
    pfHeaderSizeError,
    pfHeaderSignatureError,
    pfHeaderGUIDError,
    pfHeaderByteOrderError,
    pfHeaderDirectorySectorNumberError,
    pfNotFound
  );

type
  TModuleType           = (
    ModuleStandard      = $0021,
    ModuleClass         = $0022
  );
  TModule               = record
    ModuleName          : string; { Id = $0019 }
    ModuleStreamName    : string; { Id = $001A }
    DocString           : string; { Id = $001C }
    TextOffset          : UInt32; { Id = $0031 }
    HelpContext         : UInt32; { Id = $001E }
    ModuleType          : TModuleType; { Id = $0021 or $0022 }
    ModuleReadOnly      : Boolean; { Id = $0025 }
    ModulePrivate       : Boolean; { Id = $0028 }
    PerformanceCache    : TArray<Byte>;
    SourceCode          : string;
    ParsedPCode         : string;
  end;
  PModule               = ^TModule;
  TModuleArray          = TArray<TModule>;

  TReference            = record
    ReferenceName       : string; { Id = $0016 }
    ReferenceControl    : string;
    ReferenceOriginal   : string;
    ReferenceRegistered : string;
    ReferenceProject    : string;
    ExtendedName        : string;
    ExtendedLibrary     : string;
    GUID                : TGUID;
    MajorVersion        : UInt32;
    MinorVersion        : UInt16;
  end;
  PReference            = ^TReference;
  TReferenceArray       = TArray<TReference>;

  { The dir stream contains a series of bytes that specifies information for the VBA project, including
    project information, project references, and modules. The entire stream MUST be compressed as
    specified in Compression }
  TVBAProgram           = record
    { Office 2010 is $0097; Office 2013 is $00A3; Office 2016 32-bit is $00B2, 64-bit is $00D7 }
    Version             : UInt16;
    VbaProjectData      : TArray<Byte>;
    { SysKind (4 bytes): An unsigned integer that specifies the platform for which the VBA project is
      created. MUST have one of the following values:      0x00000000  For 16-bit Windows Platforms      0x00000001  For 32-bit Windows Platforms.      0x00000002  For Macintosh Platforms.      0x00000003  For 64-bit Windows Platforms. }    SysKind             : UInt32; { Id = $0001 }
    { Lcid (4 bytes): An unsigned integer that specifies the LCID value for the VBA project. MUST be
      0x00000409. }
    Lcid                : UInt32; { Id = $0002 }
    { LcidInvoke (4 bytes): An unsigned integer that specifies the LCID value used for Invoke calls.
      MUST be 0x00000409. }    LcidInvoke          : UInt32; { Id = $0014 }    { CodePage (2 bytes): An unsigned integer that specifies the code page for the VBA project. }    CodePage            : UInt16; { Id = $0003 }    { ProjectName (variable): An array of SizeOfProjectName bytes that specifies the VBA identifier      name for the VBA project. MUST contain MBCS characters encoded using the code page specified
      in PROJECTCODEPAGE (section 2.3.4.2.1.4). MUST NOT contain null characters. }    ProjectName         : string; { Id = $0004 }    { DocStringUnicode (variable): An array of SizeOfDocStringUnicode bytes that specifies the      description for the VBA project. MUST contain UTF-16 characters. MUST NOT contain null
      characters. MUST contain the UTF-16 encoding of DocString. }    DocString           : string; { Id = $0005 }    HelpFile1           : string; { Id = $0006 }    HelpFile2           : string; { Id = $0006 }    HelpContext         : UInt32; { Id = $0007 }    { ProjectLibFlags (4 bytes): An unsigned integer that specifies LIBFLAGS for the VBA project’s      Automation type library as specified in [MS-OAUT] section 2.2.20. MUST be 0x00000000. }    ProjectLibFlags     : UInt32; { Id = $0008 }    { VersionMajor (4 bytes): An unsigned integer specifying the major version of the VBA project. }    VersionMajor        : UInt32; { Id = $0009 }    { VersionMinor (2 bytes): An unsigned integer specifying the minor version of the VBA project. }    VersionMinor        : UInt16; { Id = $0009 }    Constants           : string; { Id = $000C }    { Specifies a reference to an Automation type library or VBA project. }    Reference           : TReferenceArray;    Module              : TModuleArray;    ModulesCount        : UInt16;  end;

  { The _VBA_PROJECT stream contains the version-dependent description of a VBA project.
  The first seven bytes of the stream are version-independent and therefore can be read by any version. }
  TVbaProjectStream     = record
    { Reserved1 (2 bytes): MUST be 0x61CC. MUST be ignored. }
    Reserved1           : UInt16;
    { Version (2 bytes): An unsigned integer that specifies the version of VBA used to create the VBA
      project. MUST be ignored on read. MUST be 0xFFFF on write. }    Version             : UInt16;
    { Reserved2 (1 byte): MUST be 0x00. MUST be ignored }
    Reserved2           : Byte;
    { Reserved3 (2 bytes): Undefined. MUST be ignored. }
    Reserved3           : UInt16;
    { PerformanceCache (variable): An array of bytes that forms an implementation-specific and
      version-dependent performance cache for the VBA project. The length of PerformanceCache
      MUST be seven bytes less than the size of _VBA_PROJECT Stream (section 2.3.4.1). MUST be
      ignored on read. MUST NOT be present on write }
    PerformanceCache    : TArray<Byte>;
  end;

  EVBAParseError = class(Exception)
  end;

{ Convert a MBCS to Unicode string; may raise an Error }
function AnsiString2UnicodeString(CodePage: UInt32; Bytes: TArray<Byte>; ByteCount: UInt32): string;

{ Convert a byte array to a HEX string }
function BinToStr(const Bin: TArray<Byte>): string;

function HexBYTE(Value: Byte): string;
function HexWORD(Value: UInt16): string;
function HexDWORD(Value: UInt32): string;

implementation

const
  HexSymbols = '0123456789ABCDEF';

function AnsiString2UnicodeString(CodePage: UInt32; Bytes: TArray<Byte>; ByteCount: UInt32): string;
var
  Encoding: TEncoding;
begin
  Encoding := TMBCSEncoding.Create(CodePage);
  try
    Result := Encoding.GetString(Bytes, 0, ByteCount);
  finally
    Encoding.Free();
  end;
end;

function HexBYTE(Value: Byte): string;
var
  Nibble0 : Byte;
  Nibble1 : Byte;
begin
  SetLength(Result, 2);
  Nibble0 := Value shr 4;
  Nibble1 := Value and $0F;
  Result := HexSymbols[1 + Nibble0] + HexSymbols[1 + Nibble1];
end;

function HexWORD(Value: UInt16): string;
var
  Byte0 : Byte;
  Byte1 : Byte;
begin
  Byte0 := Value shr 8;
  Byte1 := Value and $00FF;
  Result := HexBYTE(Byte0) + HexBYTE(Byte1);
end;

function HexDWORD(Value: UInt32): string;
var
  Word0 : Byte;
  Word1 : Byte;
begin
  Word0 := Value shr 16;
  Word1 := Value and $0000FFFF;
  Result := HexWORD(Word0) + HexWORD(Word1);
end;

function BinToStr(const Bin: TArray<Byte>): string;
var
  I: Integer;
begin
  SetLength(Result, Length(Bin) shl 1);
  for I :=  0 to Length(Bin)- 1 do
  begin
    Result[1 + I shl 1 + 0] := HexSymbols[1 + Bin[I] shr 4];
    Result[1 + I shl 1 + 1] := HexSymbols[1 + Bin[I] and $0F];
  end;
end;

end.
