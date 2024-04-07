unit ParserPCode;

interface

uses Common;

procedure ParsePCode(const VBAProgram: TVBAProgram; var Module: TModule);
procedure ReadIdentifiers(const VbaProjectData: TArray<Byte>; CodePage: UInt32);

implementation

uses
  System.SysUtils, System.Generics.Collections;

const
  MaxInternalNames = 254;
  MaxVariantType   = 38;
  InternalNames: array[0..MaxInternalNames] of string = (
    '<crash>', '0', 'Abs', 'Access', 'AddressOf', 'Alias', 'And', 'Any',
    'Append', 'Array', 'As', 'Assert', 'B', 'Base', 'BF', 'Binary',
    'Boolean', 'ByRef', 'Byte', 'ByVal', 'Call', 'Case', 'CBool', 'CByte',
    'CCur', 'CDate', 'CDec', 'CDbl', 'CDecl', 'ChDir', 'CInt', 'Circle',
    'CLng', 'Close', 'Compare', 'Const', 'CSng', 'CStr', 'CurDir', 'CurDir$',
    'CVar', 'CVDate', 'CVErr', 'Currency', 'Database', 'Date', 'Date$', 'Debug',
    'Decimal', 'Declare', 'DefBool', 'DefByte', 'DefCur', 'DefDate', 'DefDec', 'DefDbl',
    'DefInt', 'DefLng', 'DefObj', 'DefSng', 'DefStr', 'DefVar', 'Dim', 'Dir',
    'Dir$', 'Do', 'DoEvents', 'Double', 'Each', 'Else', 'ElseIf', 'Empty',
    'End', 'EndIf', 'Enum', 'Eqv', 'Erase', 'Error', 'Error$', 'Event',
    'WithEvents', 'Explicit', 'F', 'False', 'Fix', 'For', 'Format',
    'Format$', 'FreeFile', 'Friend', 'Function', 'Get', 'Global', 'Go', 'GoSub',
    'Goto', 'If', 'Imp', 'Implements', 'In', 'Input', 'Input$', 'InputB',
    'InputB', 'InStr', 'InputB$', 'Int', 'InStrB', 'Is', 'Integer', 'Left',
    'LBound', 'LenB', 'Len', 'Lib', 'Let', 'Line', 'Like', 'Load',
    'Local', 'Lock', 'Long', 'Loop', 'LSet', 'Me', 'Mid', 'Mid$',
    'MidB', 'MidB$', 'Mod', 'Module', 'Name', 'New', 'Next', 'Not',
    'Nothing', 'Null', 'Object', 'On', 'Open', 'Option', 'Optional', 'Or',
    'Output', 'ParamArray', 'Preserve', 'Print', 'Private', 'Property', 'PSet', 'Public',
    'Put', 'RaiseEvent', 'Random', 'Randomize', 'Read', 'ReDim', 'Rem', 'Resume',
    'Return', 'RGB', 'RSet', 'Scale', 'Seek', 'Select', 'Set', 'Sgn',
    'Shared', 'Single', 'Spc', 'Static', 'Step', 'Stop', 'StrComp', 'String',
    'String$', 'Sub', 'Tab', 'Text', 'Then', 'To', 'True', 'Type',
    'TypeOf', 'UBound', 'Unload', 'Unlock', 'Unknown', 'Until', 'Variant', 'WEnd',
    'While', 'Width', 'With', 'Write', 'Xor', '#Const', '#Else', '#ElseIf',
    '#End', '#If', 'Attribute', 'VB_Base', 'VB_Control', 'VB_Creatable', 'VB_Customizable', 'VB_Description',
    'VB_Exposed', 'VB_Ext_Key', 'VB_HelpID', 'VB_Invoke_Func', 'VB_Invoke_Property', 'VB_Invoke_PropertyPut', 'VB_Invoke_PropertyPutRef', 'VB_MemberFlags',
    'VB_Name', 'VB_PredeclaredID', 'VB_ProcData', 'VB_TemplateDerived', 'VB_VarDescription', 'VB_VarHelpID', 'VB_VarMemberFlags', 'VB_VarProcData',
    'VB_UserMemID', 'VB_VarUserMemID', 'VB_GlobalNameSpace', ',', '.', '"', '_', '!',
    '#', '&', '''', '(', ')', '*', '+', '-',
    ' /', ':', ';', '<', '<=', '<>', '=', '=<',
    '=>', '>', '><', '>=', '?', '\\', '^', ':='
  );

type
  TEndian       = (BigEndian, LittleEndian);
  TVariantType  = packed record
    VarType     : UInt16;
    TypeName    : string;
    Description : string;
  end;
  TVariantTypes = array[0..MaxVariantType] of TVariantType;

const
  VariantTypes : TVariantTypes = (
    (VarType: $00; TypeName: 'Empty'; Description: 'Empty'),
    (VarType: $01; TypeName: 'Null'; Description: 'Null'),
    (VarType: $02; TypeName: 'Integer'; Description: 'Integer (Int16)'),
    (VarType: $03; TypeName: 'Long'; Description: 'Long (Int32)'),
    (VarType: $04; TypeName: 'Single'; Description: 'Single (Float32)'),
    (VarType: $05; TypeName: 'Double'; Description: 'Double (Float64)'),
    (VarType: $06; TypeName: 'Currency'; Description: 'Currency (Fixed64)'),
    (VarType: $07; TypeName: 'Date'; Description: 'Date (Float64)'),
    (VarType: $08; TypeName: 'String'; Description: 'String (BSTR)'),
    (VarType: $09; TypeName: 'Object'; Description: 'Object (IDispatch)'),
    (VarType: $0A; TypeName: 'Error'; Description: 'Error'),
    (VarType: $0B; TypeName: 'Boolean'; Description: 'Boolean (Int16)'),
    (VarType: $0C; TypeName: 'Variant'; Description: 'Variant'),
    (VarType: $0D; TypeName: 'Nothing'; Description: 'Nothing (IUnknown)'),
    (VarType: $0E; TypeName: 'Decimal'; Description: 'Decimal (Fixed96)'),
    (VarType: $0F; TypeName: ''; Description: ''),
    (VarType: $10; TypeName: 'Char'; Description: 'BYTE (Int8)'),
    (VarType: $11; TypeName: 'Byte'; Description: 'BYTE (UInt8)'),
    (VarType: $12; TypeName: 'Word'; Description: 'WORD (UInt16)'),
    (VarType: $13; TypeName: 'DWord'; Description: 'DWORD (UInt32)'),
    (VarType: $14; TypeName: 'QWord'; Description: 'QWORD (Int64)'),
    (VarType: $15; TypeName: 'QWord'; Description: 'QWORD (UInt64)'),
    (VarType: $16; TypeName: 'Int'; Description: 'Int'),
    (VarType: $17; TypeName: 'UInt'; Description: 'UInt'),
    (VarType: $18; TypeName: 'Void'; Description: 'Void'),
    (VarType: $19; TypeName: 'HResult'; Description: 'HResult'),
    (VarType: $1A; TypeName: 'Ptr'; Description: 'Pointer'),
    (VarType: $1B; TypeName: 'SafeArray'; Description: 'SafeArray'),
    (VarType: $1C; TypeName: 'Array'; Description: 'Array'),
    (VarType: $1D; TypeName: 'User defined'; Description: 'User defined'),
    (VarType: $1E; TypeName: 'String'; Description: 'String (LPSTR)'),
    (VarType: $1F; TypeName: 'String'; Description: 'String (LPWSTR)'),
    (VarType: $20; TypeName: ''; Description: ''),
    (VarType: $21; TypeName: ''; Description: ''),
    (VarType: $22; TypeName: ''; Description: ''),
    (VarType: $23; TypeName: ''; Description: ''),
    (VarType: $24; TypeName: 'Record'; Description: 'Record'),
    (VarType: $25; TypeName: 'Int Ptr'; Description: 'Ptr Int'),
    (VarType: $26; TypeName: 'UInt Ptr'; Description: 'Ptr UInt')
  );

type
  TInstruction  = packed record
    Id          : UInt16;
    Mnemonic    : string;
    Arguments   : string;
    VarArg      : Boolean;
  end;
  TInstructions = array[0..263] of TInstruction;

const
  { Arguments:
    n = Name
    i = Imp
    f = Func
    v = Var
    r = Rec
    t = Type
    c = Context
    0 = UInt16
    1 = UInt32
    2 = UInt64
    3 = Float32
    4 = Float64
    5 = Date
  }
  Instructions : TInstructions = (
    (Id:   0; Mnemonic: 'Imp'; Arguments: ''; VarArg: False),
    (Id:   1; Mnemonic: 'Eqv'; Arguments: ''; VarArg: False),
    (Id:   2; Mnemonic: 'Xor'; Arguments: ''; VarArg: False),
    (Id:   3; Mnemonic: 'Or'; Arguments: ''; VarArg: False),
    (Id:   4; Mnemonic: 'And'; Arguments: ''; VarArg: False),
    (Id:   5; Mnemonic: 'Eq'; Arguments: ''; VarArg: False),
    (Id:   6; Mnemonic: 'Ne'; Arguments: ''; VarArg: False),
    (Id:   7; Mnemonic: 'Le'; Arguments: ''; VarArg: False),
    (Id:   8; Mnemonic: 'Ge'; Arguments: ''; VarArg: False),
    (Id:   9; Mnemonic: 'Lt'; Arguments: ''; VarArg: False),
    (Id:  10; Mnemonic: 'Gt'; Arguments: ''; VarArg: False),
    (Id:  11; Mnemonic: 'Add'; Arguments: ''; VarArg: False),
    (Id:  12; Mnemonic: 'Sub'; Arguments: ''; VarArg: False),
    (Id:  13; Mnemonic: 'Mod'; Arguments: ''; VarArg: False),
    (Id:  14; Mnemonic: 'IDiv'; Arguments: ''; VarArg: False),
    (Id:  15; Mnemonic: 'Mul'; Arguments: ''; VarArg: False),
    (Id:  16; Mnemonic: 'Div'; Arguments: ''; VarArg: False),
    (Id:  17; Mnemonic: 'Concat'; Arguments: ''; VarArg: False),
    (Id:  18; Mnemonic: 'Like'; Arguments: ''; VarArg: False),
    (Id:  19; Mnemonic: 'Pwr'; Arguments: ''; VarArg: False),
    (Id:  20; Mnemonic: 'Is'; Arguments: ''; VarArg: False),
    (Id:  21; Mnemonic: 'Not'; Arguments: ''; VarArg: False),
    (Id:  22; Mnemonic: 'UMi'; Arguments: ''; VarArg: False),
    (Id:  23; Mnemonic: 'FnAbs'; Arguments: ''; VarArg: False),
    (Id:  24; Mnemonic: 'FnFix'; Arguments: ''; VarArg: False),
    (Id:  25; Mnemonic: 'FnInt'; Arguments: ''; VarArg: False),
    (Id:  26; Mnemonic: 'FnSgn'; Arguments: ''; VarArg: False),
    (Id:  27; Mnemonic: 'FnLen'; Arguments: ''; VarArg: False),
    (Id:  28; Mnemonic: 'FnLenB'; Arguments: ''; VarArg: False),
    (Id:  29; Mnemonic: 'Paren'; Arguments: ''; VarArg: False),
    (Id:  30; Mnemonic: 'Sharp'; Arguments: ''; VarArg: False),
    (Id:  31; Mnemonic: 'LdLHS'; Arguments: 'n'; VarArg: False),
    (Id:  32; Mnemonic: 'Ld'; Arguments: 'n'; VarArg: False),
    (Id:  33; Mnemonic: 'MemLd'; Arguments: 'n'; VarArg: False),
    (Id:  34; Mnemonic: 'DictLd'; Arguments: 'n'; VarArg: False),
    (Id:  35; Mnemonic: 'IndexLd'; Arguments: 'n'; VarArg: False),
    (Id:  36; Mnemonic: 'ArgsLd'; Arguments: 'n0'; VarArg: False),
    (Id:  37; Mnemonic: 'ArgsMemLd'; Arguments: 'n0'; VarArg: False),
    (Id:  38; Mnemonic: 'ArgsDictLd'; Arguments: 'n0'; VarArg: False),
    (Id:  39; Mnemonic: 'St'; Arguments: 'n'; VarArg: False),
    (Id:  40; Mnemonic: 'MemSt'; Arguments: 'n'; VarArg: False),
    (Id:  41; Mnemonic: 'DictSt'; Arguments: 'n'; VarArg: False),
    (Id:  42; Mnemonic: 'IndexSt'; Arguments: 'n'; VarArg: False),
    (Id:  43; Mnemonic: 'ArgsSt'; Arguments: 'n0'; VarArg: False),
    (Id:  44; Mnemonic: 'ArgsMemSt'; Arguments: 'n0'; VarArg: False),
    (Id:  45; Mnemonic: 'ArgsDictSt'; Arguments: 'n0'; VarArg: False),
    (Id:  46; Mnemonic: 'Set'; Arguments: 'n'; VarArg: False),
    (Id:  47; Mnemonic: 'Memset'; Arguments: 'n'; VarArg: False),
    (Id:  48; Mnemonic: 'Dictset'; Arguments: 'n'; VarArg: False),
    (Id:  49; Mnemonic: 'Indexset'; Arguments: '0'; VarArg: False),
    (Id:  50; Mnemonic: 'ArgsSet'; Arguments: 'n0'; VarArg: False),
    (Id:  51; Mnemonic: 'ArgsMemSet'; Arguments: 'n0'; VarArg: False),
    (Id:  52; Mnemonic: 'ArgsDictSet'; Arguments: 'n0'; VarArg: False),
    (Id:  53; Mnemonic: 'MemLdWith'; Arguments: 'n'; VarArg: False),
    (Id:  54; Mnemonic: 'DictLdWith'; Arguments: 'n'; VarArg: False),
    (Id:  55; Mnemonic: 'ArgsMemLdWith'; Arguments: 'n0'; VarArg: False),
    (Id:  56; Mnemonic: 'ArgsDictLdWith'; Arguments: 'n0'; VarArg: False),
    (Id:  57; Mnemonic: 'MemStWith'; Arguments: 'n'; VarArg: False),
    (Id:  58; Mnemonic: 'DictStWith'; Arguments: 'n'; VarArg: False),
    (Id:  59; Mnemonic: 'ArgsMemStWith'; Arguments: 'n0'; VarArg: False),
    (Id:  60; Mnemonic: 'ArgsDictStWith'; Arguments: 'n0'; VarArg: False),
    (Id:  61; Mnemonic: 'MemSetWith'; Arguments: 'n'; VarArg: False),
    (Id:  62; Mnemonic: 'DictSetWith'; Arguments: 'n'; VarArg: False),
    (Id:  63; Mnemonic: 'ArgsMemSetWith'; Arguments: 'n0'; VarArg: False),
    (Id:  64; Mnemonic: 'ArgsDictSetWith'; Arguments: 'n0'; VarArg: False),
    (Id:  65; Mnemonic: 'ArgsCall'; Arguments: 'n0'; VarArg: False),
    (Id:  66; Mnemonic: 'ArgsMemCall'; Arguments: 'n0'; VarArg: False),
    (Id:  67; Mnemonic: 'ArgsMemCallWith'; Arguments: 'n0'; VarArg: False),
    (Id:  68; Mnemonic: 'ArgsArray'; Arguments: 'n0'; VarArg: False),
    (Id:  69; Mnemonic: 'Assert'; Arguments: ''; VarArg: False),
    (Id:  70; Mnemonic: 'BoS'; Arguments: '0'; VarArg: False),
    (Id:  71; Mnemonic: 'BoSImplicit'; Arguments: ''; VarArg: False),
    (Id:  72; Mnemonic: 'BoL'; Arguments: ''; VarArg: False),
    (Id:  73; Mnemonic: 'LdAddressOf'; Arguments: 'n'; VarArg: False),
    (Id:  74; Mnemonic: 'MemAddressOf'; Arguments: 'n'; VarArg: False),
    (Id:  75; Mnemonic: 'Case'; Arguments: ''; VarArg: False),
    (Id:  76; Mnemonic: 'CaseTo'; Arguments: ''; VarArg: False),
    (Id:  77; Mnemonic: 'CaseGt'; Arguments: ''; VarArg: False),
    (Id:  78; Mnemonic: 'CaseLt'; Arguments: ''; VarArg: False),
    (Id:  79; Mnemonic: 'CaseGe'; Arguments: ''; VarArg: False),
    (Id:  80; Mnemonic: 'CaseLe'; Arguments: ''; VarArg: False),
    (Id:  81; Mnemonic: 'CaseNe'; Arguments: ''; VarArg: False),
    (Id:  82; Mnemonic: 'CaseEq'; Arguments: ''; VarArg: False),
    (Id:  83; Mnemonic: 'CaseElse'; Arguments: ''; VarArg: False),
    (Id:  84; Mnemonic: 'CaseDone'; Arguments: ''; VarArg: False),
    (Id:  85; Mnemonic: 'Circle'; Arguments: '0'; VarArg: False),
    (Id:  86; Mnemonic: 'Close'; Arguments: '0'; VarArg: False),
    (Id:  87; Mnemonic: 'CloseAll'; Arguments: ''; VarArg: False),
    (Id:  88; Mnemonic: 'Coerce'; Arguments: ''; VarArg: False),
    (Id:  89; Mnemonic: 'CoerceVar'; Arguments: ''; VarArg: False),
    (Id:  90; Mnemonic: 'Context'; Arguments: 'c'; VarArg: False),
    (Id:  91; Mnemonic: 'Debug'; Arguments: ''; VarArg: False),
    (Id:  92; Mnemonic: 'DefType'; Arguments: '00'; VarArg: False),
    (Id:  93; Mnemonic: 'Dim'; Arguments: ''; VarArg: False),
    (Id:  94; Mnemonic: 'DimImplicit'; Arguments: ''; VarArg: False),
    (Id:  95; Mnemonic: 'Do'; Arguments: ''; VarArg: False),
    (Id:  96; Mnemonic: 'DoEvents'; Arguments: ''; VarArg: False),
    (Id:  97; Mnemonic: 'DoUnitil'; Arguments: ''; VarArg: False),
    (Id:  98; Mnemonic: 'DoWhile'; Arguments: ''; VarArg: False),
    (Id:  99; Mnemonic: 'Else'; Arguments: ''; VarArg: False),
    (Id: 100; Mnemonic: 'ElseBlock'; Arguments: ''; VarArg: False),
    (Id: 101; Mnemonic: 'ElseIfBlock'; Arguments: ''; VarArg: False),
    (Id: 102; Mnemonic: 'ElseIfTypeBlock'; Arguments: 'i'; VarArg: False),
    (Id: 103; Mnemonic: 'End'; Arguments: ''; VarArg: False),
    (Id: 104; Mnemonic: 'EndContext'; Arguments: ''; VarArg: False),
    (Id: 105; Mnemonic: 'EndFunc'; Arguments: ''; VarArg: False),
    (Id: 106; Mnemonic: 'EndIf'; Arguments: ''; VarArg: False),
    (Id: 107; Mnemonic: 'EndIfBlock'; Arguments: ''; VarArg: False),
    (Id: 108; Mnemonic: 'EndImmediate'; Arguments: ''; VarArg: False),
    (Id: 109; Mnemonic: 'EndProp'; Arguments: ''; VarArg: False),
    (Id: 110; Mnemonic: 'EndSelect'; Arguments: ''; VarArg: False),
    (Id: 111; Mnemonic: 'EndSub'; Arguments: ''; VarArg: False),
    (Id: 112; Mnemonic: 'EndType'; Arguments: ''; VarArg: False),
    (Id: 113; Mnemonic: 'EndWith'; Arguments: ''; VarArg: False),
    (Id: 114; Mnemonic: 'Erase'; Arguments: '0'; VarArg: False),
    (Id: 115; Mnemonic: 'Error'; Arguments: ''; VarArg: False),
    (Id: 116; Mnemonic: 'EventDecl'; Arguments: 'f'; VarArg: False),
    (Id: 117; Mnemonic: 'RaiseEvent'; Arguments: 'n0'; VarArg: False),
    (Id: 118; Mnemonic: 'ArgsMemRaiseEvent'; Arguments: 'n0'; VarArg: False),
    (Id: 119; Mnemonic: 'ArgsMemRaiseEventWith'; Arguments: 'n0'; VarArg: False),
    (Id: 120; Mnemonic: 'ExitDo'; Arguments: ''; VarArg: False),
    (Id: 121; Mnemonic: 'ExitFor'; Arguments: ''; VarArg: False),
    (Id: 122; Mnemonic: 'ExitFunc'; Arguments: ''; VarArg: False),
    (Id: 123; Mnemonic: 'ExitProp'; Arguments: ''; VarArg: False),
    (Id: 124; Mnemonic: 'ExitSub'; Arguments: ''; VarArg: False),
    (Id: 125; Mnemonic: 'FnCurDir'; Arguments: ''; VarArg: False),
    (Id: 126; Mnemonic: 'FnDir'; Arguments: ''; VarArg: False),
    (Id: 127; Mnemonic: 'Empty0'; Arguments: ''; VarArg: False),
    (Id: 128; Mnemonic: 'Empty1'; Arguments: ''; VarArg: False),
    (Id: 129; Mnemonic: 'FnError'; Arguments: ''; VarArg: False),
    (Id: 130; Mnemonic: 'FnFormat'; Arguments: ''; VarArg: False),
    (Id: 131; Mnemonic: 'FnFreeFile'; Arguments: ''; VarArg: False),
    (Id: 132; Mnemonic: 'FnInStr'; Arguments: ''; VarArg: False),
    (Id: 133; Mnemonic: 'FnInStr3'; Arguments: ''; VarArg: False),
    (Id: 134; Mnemonic: 'FnInStr4'; Arguments: ''; VarArg: False),
    (Id: 135; Mnemonic: 'FnInStrB'; Arguments: ''; VarArg: False),
    (Id: 136; Mnemonic: 'FnInStrB3'; Arguments: ''; VarArg: False),
    (Id: 137; Mnemonic: 'FnInStrB4'; Arguments: ''; VarArg: False),
    (Id: 138; Mnemonic: 'FnLBound'; Arguments: '0'; VarArg: False),
    (Id: 139; Mnemonic: 'FnMid'; Arguments: ''; VarArg: False),
    (Id: 140; Mnemonic: 'FnMidB'; Arguments: ''; VarArg: False),
    (Id: 141; Mnemonic: 'FnStrComp'; Arguments: ''; VarArg: False),
    (Id: 142; Mnemonic: 'FnStrComp3'; Arguments: ''; VarArg: False),
    (Id: 143; Mnemonic: 'FnStringVar'; Arguments: ''; VarArg: False),
    (Id: 144; Mnemonic: 'FnStringStr'; Arguments: ''; VarArg: False),
    (Id: 145; Mnemonic: 'FnUBound'; Arguments: '0'; VarArg: False),
    (Id: 146; Mnemonic: 'For'; Arguments: ''; VarArg: False),
    (Id: 147; Mnemonic: 'ForEach'; Arguments: ''; VarArg: False),
    (Id: 148; Mnemonic: 'ForEachAs'; Arguments: 'i'; VarArg: False),
    (Id: 149; Mnemonic: 'ForStep'; Arguments: ''; VarArg: False),
    (Id: 150; Mnemonic: 'FuncDefn'; Arguments: 'f'; VarArg: False),
    (Id: 151; Mnemonic: 'FuncDefnSave'; Arguments: 'f'; VarArg: False),
    (Id: 152; Mnemonic: 'GetRec'; Arguments: ''; VarArg: False),
    (Id: 153; Mnemonic: 'GoSub'; Arguments: 'n'; VarArg: False),
    (Id: 154; Mnemonic: 'GoTo'; Arguments: 'n'; VarArg: False),
    (Id: 155; Mnemonic: 'If'; Arguments: ''; VarArg: False),
    (Id: 156; Mnemonic: 'IfBlock'; Arguments: ''; VarArg: False),
    (Id: 157; Mnemonic: 'TypeOf'; Arguments: 'i'; VarArg: False),
    (Id: 158; Mnemonic: 'IfTypeBlock'; Arguments: 'i'; VarArg: False),
    (Id: 159; Mnemonic: 'Implements'; Arguments: '0000'; VarArg: False),
    (Id: 160; Mnemonic: 'Input'; Arguments: ''; VarArg: False),
    (Id: 161; Mnemonic: 'InputDone'; Arguments: ''; VarArg: False),
    (Id: 162; Mnemonic: 'InputItem'; Arguments: ''; VarArg: False),
    (Id: 163; Mnemonic: 'Label'; Arguments: 'n'; VarArg: False),
    (Id: 164; Mnemonic: 'Let'; Arguments: ''; VarArg: False),
    (Id: 165; Mnemonic: 'Line'; Arguments: '0'; VarArg: False),
    (Id: 166; Mnemonic: 'LineCont'; Arguments: ''; VarArg: True),
    (Id: 167; Mnemonic: 'LineInput'; Arguments: ''; VarArg: False),
    (Id: 168; Mnemonic: 'LineNum'; Arguments: 'n'; VarArg: False),
    (Id: 169; Mnemonic: 'LitCy'; Arguments: '0000'; VarArg: False),
    (Id: 170; Mnemonic: 'LitDate'; Arguments: '0000'; VarArg: False),
    (Id: 171; Mnemonic: 'LitDefault'; Arguments: ''; VarArg: False),
    (Id: 172; Mnemonic: 'LitDI2'; Arguments: '0'; VarArg: False),
    (Id: 173; Mnemonic: 'LitDI4'; Arguments: '1'; VarArg: False),
    (Id: 174; Mnemonic: 'LitDI8'; Arguments: '2'; VarArg: False),
    (Id: 175; Mnemonic: 'LitHI2'; Arguments: '0'; VarArg: False),
    (Id: 176; Mnemonic: 'LitHI4'; Arguments: '1'; VarArg: False),
    (Id: 177; Mnemonic: 'LitHI8'; Arguments: '2'; VarArg: False),
    (Id: 178; Mnemonic: 'LitNothing'; Arguments: ''; VarArg: False),
    (Id: 179; Mnemonic: 'LitOI2'; Arguments: '0'; VarArg: False),
    (Id: 180; Mnemonic: 'LitOI4'; Arguments: '1'; VarArg: False),
    (Id: 181; Mnemonic: 'LitOI8'; Arguments: '2'; VarArg: False),
    (Id: 182; Mnemonic: 'LitR4'; Arguments: '3'; VarArg: False),
    (Id: 183; Mnemonic: 'LitR8'; Arguments: '4'; VarArg: False),
    (Id: 184; Mnemonic: 'LitSmallI2'; Arguments: ''; VarArg: False),
    (Id: 185; Mnemonic: 'LitStr'; Arguments: ''; VarArg: True),
    (Id: 186; Mnemonic: 'LitVarSpecial'; Arguments: ''; VarArg: False),
    (Id: 187; Mnemonic: 'Lock'; Arguments: ''; VarArg: False),
    (Id: 188; Mnemonic: 'Loop'; Arguments: ''; VarArg: False),
    (Id: 189; Mnemonic: 'LoopUntil'; Arguments: ''; VarArg: False),
    (Id: 190; Mnemonic: 'LoopWhile'; Arguments: ''; VarArg: False),
    (Id: 191; Mnemonic: 'LSet'; Arguments: ''; VarArg: False),
    (Id: 192; Mnemonic: 'Me'; Arguments: ''; VarArg: False),
    (Id: 193; Mnemonic: 'MeImplicit'; Arguments: ''; VarArg: False),
    (Id: 194; Mnemonic: 'MemRedim'; Arguments: 'n0t'; VarArg: False),
    (Id: 195; Mnemonic: 'MemRedimWith'; Arguments: 'n0t'; VarArg: False),
    (Id: 196; Mnemonic: 'MemRedimAs'; Arguments: 'n0t'; VarArg: False),
    (Id: 197; Mnemonic: 'MemRedimAsWith'; Arguments: 'n0t'; VarArg: False),
    (Id: 198; Mnemonic: 'Mid'; Arguments: ''; VarArg: False),
    (Id: 199; Mnemonic: 'MidB'; Arguments: ''; VarArg: False),
    (Id: 200; Mnemonic: 'n'; Arguments: ''; VarArg: False),
    (Id: 201; Mnemonic: 'New'; Arguments: 'i'; VarArg: False),
    (Id: 202; Mnemonic: 'Next'; Arguments: ''; VarArg: False),
    (Id: 203; Mnemonic: 'NextVar'; Arguments: ''; VarArg: False),
    (Id: 204; Mnemonic: 'OnError'; Arguments: 'n'; VarArg: False),
    (Id: 205; Mnemonic: 'OnGosub'; Arguments: ''; VarArg: True),
    (Id: 206; Mnemonic: 'OnGoto'; Arguments: ''; VarArg: True),
    (Id: 207; Mnemonic: 'Open'; Arguments: '0'; VarArg: False),
    (Id: 208; Mnemonic: 'Option'; Arguments: ''; VarArg: False),
    (Id: 209; Mnemonic: 'OptionBase'; Arguments: ''; VarArg: False),
    (Id: 210; Mnemonic: 'ParamByVal'; Arguments: ''; VarArg: False),
    (Id: 211; Mnemonic: 'ParamOmitted'; Arguments: ''; VarArg: False),
    (Id: 212; Mnemonic: 'ParamNamed'; Arguments: 'n'; VarArg: False),
    (Id: 213; Mnemonic: 'PrintChan'; Arguments: ''; VarArg: False),
    (Id: 214; Mnemonic: 'PrintComma'; Arguments: ''; VarArg: False),
    (Id: 215; Mnemonic: 'PrintEoS'; Arguments: ''; VarArg: False),
    (Id: 216; Mnemonic: 'PrintItemComma'; Arguments: ''; VarArg: False),
    (Id: 217; Mnemonic: 'PrintItemNL'; Arguments: ''; VarArg: False),
    (Id: 218; Mnemonic: 'PrintItemSemi'; Arguments: ''; VarArg: False),
    (Id: 219; Mnemonic: 'PrintNL'; Arguments: ''; VarArg: False),
    (Id: 220; Mnemonic: 'PrintObj'; Arguments: ''; VarArg: False),
    (Id: 221; Mnemonic: 'PrintSemi'; Arguments: ''; VarArg: False),
    (Id: 222; Mnemonic: 'PrintSpc'; Arguments: ''; VarArg: False),
    (Id: 223; Mnemonic: 'PrintTab'; Arguments: ''; VarArg: False),
    (Id: 224; Mnemonic: 'PrintTabComma'; Arguments: ''; VarArg: False),
    (Id: 225; Mnemonic: 'PSet'; Arguments: '0'; VarArg: False),
    (Id: 226; Mnemonic: 'PutRec'; Arguments: ''; VarArg: False),
    (Id: 227; Mnemonic: 'QuoteRem'; Arguments: '0'; VarArg: True),
    (Id: 228; Mnemonic: 'Redim'; Arguments: 'n0t'; VarArg: False),
    (Id: 229; Mnemonic: 'RedimAs'; Arguments: 'n0t'; VarArg: False),
    (Id: 230; Mnemonic: 'Reparse'; Arguments: ''; VarArg: True),
    (Id: 231; Mnemonic: 'Rem'; Arguments: ''; VarArg: True),
    (Id: 232; Mnemonic: 'Resume'; Arguments: 'n'; VarArg: False),
    (Id: 233; Mnemonic: 'Return'; Arguments: ''; VarArg: False),
    (Id: 234; Mnemonic: 'RSet'; Arguments: ''; VarArg: False),
    (Id: 235; Mnemonic: 'Scale'; Arguments: '0'; VarArg: False),
    (Id: 236; Mnemonic: 'Seek'; Arguments: ''; VarArg: False),
    (Id: 237; Mnemonic: 'SelectCase'; Arguments: ''; VarArg: False),
    (Id: 238; Mnemonic: 'SelectIs'; Arguments: 'i'; VarArg: False),
    (Id: 239; Mnemonic: 'SelectType'; Arguments: ''; VarArg: False),
    (Id: 240; Mnemonic: 'SetStmt'; Arguments: ''; VarArg: False),
    (Id: 241; Mnemonic: 'Stack'; Arguments: '00'; VarArg: False),
    (Id: 242; Mnemonic: 'Stop'; Arguments: ''; VarArg: False),
    (Id: 243; Mnemonic: 'Type'; Arguments: 'r'; VarArg: False),
    (Id: 244; Mnemonic: 'Unlock'; Arguments: ''; VarArg: False),
    (Id: 245; Mnemonic: 'VarDefn'; Arguments: 'v'; VarArg: False),
    (Id: 246; Mnemonic: 'Wend'; Arguments: ''; VarArg: False),
    (Id: 247; Mnemonic: 'While'; Arguments: ''; VarArg: False),
    (Id: 248; Mnemonic: 'With'; Arguments: ''; VarArg: False),
    (Id: 249; Mnemonic: 'WriteChan'; Arguments: ''; VarArg: False),
    (Id: 250; Mnemonic: 'ConstFuncExpr'; Arguments: ''; VarArg: False),
    (Id: 251; Mnemonic: 'LbConst'; Arguments: 'n'; VarArg: False),
    (Id: 252; Mnemonic: 'LbIf'; Arguments: ''; VarArg: False),
    (Id: 253; Mnemonic: 'LbElse'; Arguments: ''; VarArg: False),
    (Id: 254; Mnemonic: 'LbElseIf'; Arguments: ''; VarArg: False),
    (Id: 255; Mnemonic: 'LbEndIf'; Arguments: ''; VarArg: False),
    (Id: 256; Mnemonic: 'LbMark'; Arguments: ''; VarArg: False),
    (Id: 257; Mnemonic: 'EndForVariable'; Arguments: ''; VarArg: False),
    (Id: 258; Mnemonic: 'StartForVariable'; Arguments: ''; VarArg: False),
    (Id: 259; Mnemonic: 'NewRedim'; Arguments: ''; VarArg: False),
    (Id: 260; Mnemonic: 'StartWithExpr'; Arguments: ''; VarArg: False),
    (Id: 261; Mnemonic: 'SetOrSt'; Arguments: 'n'; VarArg: False),
    (Id: 262; Mnemonic: 'EndEnum'; Arguments: ''; VarArg: False),
    (Id: 263; Mnemonic: 'Illegal'; Arguments: ''; VarArg: False)
  );

type
  TIdentifiers = TDictionary<UInt16, string>;

var
  Identifiers : TIdentifiers;

procedure Reset();
begin
  FreeAndNil(Identifiers);
end;

function ParseSimple(const CodeBytes: TArray<Byte>): string; forward;

{ Get one WORD from the stream; does not modify the offset }
function GetWORD(const Buffer: TArray<Byte>; Offset: UInt32; Endian: TEndian): UInt16;
var
  Half0 : UInt16;
  Half1 : UInt16;
begin
  Half0 := 0;
  Half1 := 0;
  Half0 := Buffer[Offset];
  Half1 := Buffer[Offset + 1];
  if Endian = TEndian.BigEndian then
    Result := Half0 shl 8 + Half1
  else
    Result := Half0 + Half1 shl 8;
end;

{ Get one DWORD from the stream; does not modify the offset }
function GetDWORD(const Buffer: TArray<Byte>; Offset: UInt32; Endian: TEndian): UInt32;
var
  Half0: UInt32;
  Half1: UInt32;
begin
  Half0 := 0;
  Half1 := 0;
  Half0 := GetWORD(Buffer, Offset, Endian);
  Half1 := GetWORD(Buffer, Offset + 2, Endian);
  if Endian = TEndian.BigEndian then
    Result := Half0 shl 16 + Half1
  else
    Result := Half0 + Half1 shl 16;
end;

{ Get one QWORD from the stream; does not modify the offset }
function GetQWORD(const Buffer: TArray<Byte>; Offset: UInt32; Endian: TEndian): UInt64;
var
  Half0: UInt64;
  Half1: UInt64;
begin
  Half0 := 0;
  Half1 := 0;
  Half0 := GetDWORD(Buffer, Offset, Endian);
  Half1 := GetDWORD(Buffer, Offset + 4, Endian);
  if Endian = TEndian.BigEndian then
    Result := Half0 shl 32 + Half1
  else
    Result := Half0 + Half1 shl 32
end;

{ Get one String of Char from the stream; does not modify the offset }
function GetString(const Buffer: TArray<Byte>; Offset: UInt32; CodePage: UInt32; NumberOfBytes: UInt32): string;
var
  StringBuffer: TArray<Byte>;
begin
  Result := '';
  if NumberOfBytes > 0 then
  begin
    StringBuffer := Copy(Buffer, Offset, NumberOfBytes);
    Result := AnsiString2UnicodeString(CodePage, StringBuffer, NumberOfBytes);
  end;
end;

{ Get one WORD from the stream; does modify the offset }
function ReadWORD(const Buffer: TArray<Byte>; var Offset: UInt32; Endian: TEndian): UInt16;
begin
  Result := GetWORD(Buffer, Offset, Endian);
  Inc(Offset, 2);
end;

{ Get one DWORD from the stream; does modify the offset }
function ReadDWORD(const Buffer: TArray<Byte>; var Offset: UInt32; Endian: TEndian): UInt32;
begin
  Result := GetDWORD(Buffer, Offset, Endian);
  Inc(Offset, 4);
end;

{ Get one QWORD from the stream; does modify the offset }
function ReadQWORD(const Buffer: TArray<Byte>; var Offset: UInt32; Endian: TEndian): UInt64;
begin
  Result := GetQWORD(Buffer, Offset, Endian);
  Inc(Offset, 8);
end;

function SkipStructure(
  const Buffer     : TArray<Byte>;
  Offset           : UInt32;
  Endian           : TEndian;
  IsLengthDWORD    : Boolean;
  ElementSize      : UInt32;
  CheckForMinusOne : Boolean
): UInt32;
var
  StructureLength : UInt32;
  DoSkip          : Boolean;
begin
  if IsLengthDWORD then
  begin
    StructureLength := ReadDWORD(Buffer, Offset, Endian);
    DoSkip := (CheckForMinusOne) and (StructureLength = $FFFFFFFF);
  end
  else
  begin
    StructureLength := ReadWORD(Buffer, Offset, Endian);
    DoSkip := (CheckForMinusOne) and (StructureLength = $FFFF);
  end;
  if not DoSkip then
    Offset := Offset + StructureLength * ElementSize;
  Result := Offset;
end;

procedure GetTypeAndLength(
  const Buffer  : TArray<Byte>;
  Offset        : UInt32;
  Endian        : TEndian;
  var VarType   : Byte;
  var VarLength : Byte
);
begin
  if Endian = TEndian.BigEndian then
  begin
    VarType   := Buffer[Offset];
    VarLength := Buffer[Offset + 1];
  end
  else
  begin
    VarType   := Buffer[Offset + 1];
    VarLength := Buffer[Offset];
  end;
end;

function TranslateOpCode(OpCode: UInt16; VBAVersion: UInt16; VBA64bit: Boolean): UInt16;
begin
  if VBAVersion = 3 then
    case OpCode of
      0..67:
        Result := OpCode;
      68..70:
        Result := OpCode + 2;
      71..111:
        Result := OpCode + 4;
      112..150:
        Result := OpCode + 8;
      151..164:
        Result := OpCode + 9;
      165..166:
        Result := OpCode + 10;
      167..169:
        Result := OpCode + 11;
      170..238:
        Result := OpCode + 12;
      else
        Result := OpCode + 24;
    end
  else
    if VBAVersion = 5 then
      case OpCode of
        0..68:
          Result := OpCode;
        69..71:
          Result := OpCode + 1;
        72..112:
          Result := OpCode + 3;
        113..151:
          Result := OpCode + 7;
        152..165:
          Result := OpCode + 8;
        166..167:
          Result := OpCode + 9;
        168..170:
          Result := OpCode + 10;
        else
          Result := OpCode + 11;
      end
    else
      if not VBA64bit then
        case OpCode of
          0..173:
            Result := OpCode;
          174..175:
            Result := OpCode + 1;
          176..178:
            Result := OpCode + 2;
          else
            Result := OpCode + 3;
        end
      else
        Result := OpCode;
end;

function GetIdentifier(Reference: UInt16; VbaVersion: UInt16; Vba64bit: Boolean): string;
begin
  Reference := Reference shr 1;
  if Reference >= $100 then
    if Identifiers.ContainsKey(Reference) then
      Result := Identifiers.Items[Reference]
    else
      Result := ''
  else
  begin
    if VBAVersion >= 7 then
      if Reference > $C3 then
        Reference := Reference - 1;
    Result := InternalNames[Reference];
  end;
end;

function GetName(
  const Buffer : TArray<Byte>;
  Offset       : UInt32;
  Endian       : TEndian;
  VbaVersion   : UInt16;
  Vba64bit     : Boolean
): string;
var
  ObjectID : UInt16;
begin
  ObjectID := GetWORD(Buffer, Offset, Endian);
  Result := GetIdentifier(ObjectID, VbaVersion, Vba64bit);
end;

function GetClassName(
  const ObjectTable : TArray<Byte>;
  ObjectID          : UInt32;
  Endian            : TEndian;
  VbaVersion        : UInt16;
  Vba64bit          : Boolean
): string;
var
  Offset : UInt32;
begin
  Offset := (ObjectID shr 3) * 10;
  if Length(ObjectTable) >= Offset + 10 then
  begin
    Offset := GetWORD(ObjectTable, Offset + 6, Endian);
    Result := GetIdentifier(Offset, VbaVersion, Vba64bit);
  end
  else
    Result := ''
end;

function GetTypeName(TypeID: UInt16): string;
var
  Flags : UInt16;
begin
  Flags := TypeID and $E0;
  TypeID := TypeID and (not $E0);
  if TypeID <= MaxVariantType then
    Result := VariantTypes[TypeID].Description
  else
    Result := 'Type 0x' + HexWORD(TypeID);
  if Flags and $80 <> 0 then
    Result := Result + ' Ptr';
end;

function DisasmInt16(Value: UInt16): string;
begin
  Result := IntToStr(Value);
end;

function DisasmInt32(Value: UInt32): string;
begin
  Result := IntToStr(Value);
end;

function DisasmInt64(Value: UInt64): string;
begin
  Result := IntToStr(Value);
end;

function DisasmFloat32(Value: UInt32): string;
var
  PValue  : Pointer;
  PResult : PSingle;
begin
  PValue  := @Value;
  PResult := PSingle(PValue);
  Result := FloatToStr(PResult^);
end;

function DisasmFloat64(Value: UInt64): string;
var
  PValue  : Pointer;
  PResult : PDouble;
begin
  PValue  := @Value;
  PResult := PDouble(PValue);
  Result := FloatToStr(PResult^);
end;

function DisasmName(
  ObjectID   : UInt16;
  Mnemonic   : string;
  OpType     : UInt16;
  VbaVersion : UInt16;
  Vba64bit   : Boolean
): string;
const
  VarTypes: array[0..13] of string = (
    '', '?', '%', '&', '!', '#', '@', '?', '$', '?', '?', '?', '?', '?'
  );
var
  StrType : string;
  VarName : string;
begin
  StrType := '';
  VarName := GetIdentifier(ObjectID, VbaVersion, Vba64bit);
  if OpType < High(VarTypes) then
    StrType := VarTypes[OpType]
  else
    if OpType = 32 then
      VarName := '[' + VarName + ']';
  if Mnemonic = 'OnError' then
  begin
    if OpType = 1 then
      VarName := ' Resume Next'
    else
      if OpType = 2 then
        VarName := ' GoTo 0'
  end
  else
    if Mnemonic = 'Resume' then
    begin
      if OpType = 1 then
        VarName := ' Next'
      else
        if OpType <> 0 then
          VarName := '';
    end;
  Result := VarName + StrType + ' ';
end;

function DisasmImp(
  IndirectTable : TArray<Byte>;
  ObjectTable : TArray<Byte>;
  Arg         : Char;
  WordChunk   : UInt16;
  Mnenomonic  : string;
  Endian      : TEndian;
  VbaVersion  : UInt16;
  Vba64bit    : Boolean
): string;
const
  AccessMode : array[0..2] of string = ('Read', 'Write', 'Read Write');
  LockMode   : array[0..2] of string = ('Read Write', 'Write', 'Read');
var
  ObjectName : string;
  Access     : UInt16;
  Lock       : UInt16;
  Mode       : UInt16;
begin
  if Mnenomonic <> 'Open' then
    if Length(ObjectTable) >= WordChunk + 8 then
      //ObjectName := GetName(ObjectTable, WordChunk, Endian, VbaVersion, Vba64bit)
      ObjectName := GetClassName(ObjectTable, WordChunk, Endian, VbaVersion, Vba64bit)
    else
      //ObjectName := '0x' + HexWORD(WordChunk)
  else
  begin
    Mode   := WordChunk and $00FF;
    Access := (WordChunk and $0F00) shr 8;
    Lock   := (WordChunk and $F000) shr 12;
    ObjectName := 'For ';
    if Mode and $01 <> 0 then
      ObjectName := ObjectName + 'Input '
    else
      if Mode and $02 <> 0 then
        ObjectName := ObjectName + 'Output '
      else
        if Mode and $04 <> 0 then
          ObjectName := ObjectName + 'Random '
        else
          if Mode and $08 <> 0 then
            ObjectName := ObjectName + 'Append '
          else
            if Mode = $20 then
              ObjectName := ObjectName + 'Binary ';
    if (Access <> 0) and (Access <= Length(AccessMode)) then
      ObjectName := ObjectName + 'Access ' + AccessMode[Access - 1] + ' ';
    if Lock <> 0 then
      if Lock and $04 <> 0 then
        ObjectName := ObjectName + 'Shared '
      else
        if Lock <= Length(AccessMode) then
          ObjectName := ObjectName + 'Lock ' + LockMode[Lock - 1];
  end;
  Result := ObjectName;
end;

function DisasmRec(
  IndirectTable : TArray<Byte>;
  DWordChunk    : UInt32;
  Endian        : TEndian;
  VbaVersion    : UInt16;
  Vba64bit      : Boolean
): string;
var
  ObjectName : string;
  Options    : UInt16;
begin
  ObjectName := GetName(IndirectTable, DWordChunk + 2, Endian, VbaVersion, Vba64bit);
  Options := GetWORD(IndirectTable, DWordChunk + 18, Endian);
  if (Options and 1) = 0 then
    ObjectName := 'Private ' + ObjectName
  else
    ObjectName := 'Public ' + ObjectName;
  Result := ObjectName;
end;

function DisasmType(
  IndirectTable : TArray<Byte>;
  DWordChunk    : UInt32
): string;
var
  ObjectID   : UInt16;
  ObjectName : string;
begin
  ObjectID := IndirectTable[DWordChunk + 6];
  if ObjectID <= MaxVariantType then
    ObjectName := VariantTypes[ObjectID].Description
  else
    ObjectName := 'type 0x' + HexWORD(ObjectID);
  Result := ObjectName;
end;

function DisasmObject(
  IndirectTable : TArray<Byte>;
  ObjectTable   : TArray<Byte>;
  Offset        : UInt32;
  Endian        : TEndian;
  VbaVersion    : UInt16;
  Vba64bit      : Boolean
): string;
var
  Flags          : UInt16;
  Offset2        : UInt16;
  TypeDescOffset : UInt32;
  WordChunk      : UInt16;
begin
  Result := '';
  TypeDescOffset := GetDWORD(IndirectTable, Offset, Endian);
  Flags := GetWORD(IndirectTable, TypeDescOffset, Endian);
  if Flags and $02 <> 0 then
    Result := DisasmType(IndirectTable, TypeDescOffset)
  else
  begin
    WordChunk := GetWORD(IndirectTable, TypeDescOffset + 2, Endian);
    Result := DisasmType(IndirectTable, WordChunk);
    if WordChunk = 0 then
      Result := ''
    else
    begin
      Offset2 := (WordChunk shr 3) * 10;
      if Length(ObjectTable) >= Offset2 + 10 then
      begin
        Flags      := GetWORD(ObjectTable, Offset2, Endian);
        WordChunk  := GetWORD(ObjectTable, Offset2 + 6, Endian);
        Result := GetIdentifier(WordChunk, VbaVersion, Vba64bit);
      end;
    end;
  end;
end;

function DisasmVar(
  IndirectTable : TArray<Byte>;
  ObjectTable   : TArray<Byte>;
  DWordChunk    : UInt32;
  Endian        : TEndian;
  VbaVersion    : UInt16;
  Vba64bit      : Boolean
): string;
var
  Flag1     : Byte;
  Flag2     : Byte;
  HasAs     : Boolean;
  HasNew    : Boolean;
  Offset2   : UInt32;
  ObjectID  : Byte;
  TypeName  : string;
  VarName   : string;
  VarType   : string;
  WordChunk : UInt16;
begin
  Flag1 := IndirectTable[DWordChunk];
  Flag2 := IndirectTable[DWordChunk + 1];
  HasAs := (Flag1 and $20) <> 0;
  HasNew := (Flag2 and $20) <> 0;
  VarName := GetName(IndirectTable, DWordChunk + 2, Endian, VbaVersion, Vba64bit);
  if (HasNew) or (HasAs) then
  begin
    VarType := '';
    if HasNew then
      VarType := VarType + 'New ';
    if HasAs then
    begin
      if Vba64bit then
        Offset2 := 16
      else
        Offset2 := 12;
      WordChunk := GetWORD(IndirectTable, DWordChunk + Offset2 + 2, Endian);
      if WordChunk = $FFFF then
      begin
        ObjectID := IndirectTable[DWordChunk + Offset2];
        TypeName := GetTypeName(ObjectID);
      end
      else
        TypeName := DisasmObject(IndirectTable, ObjectTable, DWordChunk + Offset2, Endian, VbaVersion, Vba64bit);
      if TypeName <> '' then
        VarType := VarType + TypeName;
    end;
  end;
  if VarType <> '' then
    VarName := VarName + ' As ' + VarType;
  Result := VarName;
end;

function DisasmArg(
  IndirectTable : TArray<Byte>;
  ArgOffset     : UInt32;
  Endian        : TEndian;
  VbaVersion    : UInt16;
  Vba64bit      : Boolean
): string;
var
  Flags       : UInt16;
  Offset2     : UInt32;
  ArgName     : string;
  ArgType     : UInt32;
  ArgOpts     : UInt16;
  ArgTypeID   : UInt32;
  ArgTypeName : string;
begin
  Flags := GetWORD(IndirectTable, ArgOffset, Endian);
  if Vba64bit then
    Offset2 := 4
  else
    Offset2 := 0;
  ArgName := GetName(IndirectTable, ArgOffset + 2, Endian, VbaVersion, Vba64bit);
  ArgType := GetDWORD(IndirectTable, ArgOffset + Offset2 + 12, Endian);
  ArgOpts := GetWORD(IndirectTable, ArgOffset + Offset2 + 24, Endian);
  if ArgOpts and $0004 <> 0 then
    ArgName := 'ByVal ' + ArgName;
  if ArgOpts and $0002 <> 0 then
    ArgName := 'ByRef ' + ArgName;
  if ArgOpts and $0200 <> 0 then
    ArgName := 'Optional ' + ArgName;
  if Flags and $0020 <> 0 then
  begin
    ArgTypeName := '';
    if ArgType and $FFFF0000 <> 0 then
    begin
      ArgTypeID := ArgType and $000000FF;
      ArgTypeName := GetTypeName(ArgTypeID);
    end;
    ArgName := ArgName + ' As ' + ArgTypeName;
  end;
  Result := ArgName;
end;

function DisasmFunc(
  ObjectTable      : TArray<Byte>;
  IndirectTable    : TArray<Byte>;
  DeclarationTable : TArray<Byte>;
  DWordChunk       : UInt32;
  OpType           : Byte;
  Endian           : TEndian;
  VbaVersion       : UInt16;
  Vba64bit         : Boolean
): string;
var
  Flags      : UInt16;
  Offset2    : UInt32;
  ArgOffset  : UInt32;
  RetType    : UInt32;
  DeclOffset : UInt16;
  FuncDecl   : string;
  SubName    : string;
  Options    : Byte;
  NewFlags   : Byte;
  HasDeclare : Boolean;
  HasAs      : Boolean;
  LibName    : string;
  ArgList    : string;
  ArgName    : string;
  TypeName   : string;
  TypeID     : UInt32;
begin
  FuncDecl := ' ';
  Flags := GetWORD(IndirectTable, DWordChunk, Endian);
  SubName := GetName(IndirectTable, DWordChunk + 2, Endian, VbaVersion, Vba64bit);
  if VbaVersion > 5 then
    Offset2 := 4
  else
    Offset2 := 0;
  if Vba64bit then
    Offset2 := Offset2 + 16;
  ArgOffset := GetDWORD(IndirectTable, DWordChunk + Offset2 + 36, Endian);
  RetType := GetDWORD(IndirectTable, DWordChunk + Offset2 + 40, Endian);
  DeclOffset := GetWORD(IndirectTable, DWordChunk + Offset2 + 44, Endian);
  Options := IndirectTable[DWordChunk + Offset2 + 54];
  NewFlags := IndirectTable[DWordChunk + Offset2 + 57];
  HasDeclare := False;
  if VbaVersion > 5 then
  begin
    if ((NewFlags and $0002) = 0) and (not Vba64bit) then
      FuncDecl := FuncDecl + 'Private ';
    if NewFlags and $0004 <> 0 then
      FuncDecl := FuncDecl + 'Friend ';
  end
  else
    if Flags and $0008 = 0 then
      FuncDecl := FuncDecl + 'Private ';
  if OpType and $04 <> 0 then
    FuncDecl := FuncDecl + 'Public ';
  if Flags and $0080 <> 0 then
    FuncDecl := FuncDecl + 'Static ';
  if ((Options and $90) = 0) and (DeclOffset <> $FFFF) and (not Vba64bit) then
  begin
    HasDeclare := True;
    FuncDecl := FuncDecl + 'Declare ';
  end;
  if VbaVersion > 5 then
    if NewFlags and $20 <> 0 then
      FuncDecl := FuncDecl + 'PtrSafe ';
  HasAs := (Flags and $0020) <> 0;
  if Flags and $1000 <> 0 then
    if (OpType = 2) or (OpType = 6) then
      FuncDecl := FuncDecl + 'Function '
    else
      FuncDecl := FuncDecl + 'Sub '
  else
    if Flags and $2000 <> 0 then
      FuncDecl := FuncDecl + 'Property Get '
    else
      if Flags and $4000 <> 0 then
        FuncDecl := FuncDecl + 'Property Let '
      else
        if Flags and $8000 <> 0 then
          FuncDecl := FuncDecl + 'Property Set ';
  FuncDecl := FuncDecl + subName;
  if HasDeclare then
  begin
    LibName := GetName(DeclarationTable, DeclOffset + 2, Endian, VbaVersion, Vba64bit);
    FuncDecl := FuncDecl + ' Lib "' + LibName + '" ';
  end;
  ArgList := '';
  while (ArgOffset <> $FFFFFFFF) and (ArgOffset <> 0) and (ArgOffset + 26 < Length(IndirectTable)) do
  begin
    ArgName := DisasmArg(IndirectTable, ArgOffset, Endian, VbaVersion, Vba64bit);
    if ArgList <> '' then
      ArgList := ArgList + ', ';
    ArgList := ArgList + ArgName;
    ArgOffset := GetDWORD(IndirectTable, ArgOffset + 20, Endian);
  end;
  FuncDecl := FuncDecl + '(' + ArgList + ')';
  if HasAs then
  begin
    FuncDecl := FuncDecl + ' As ';
    TypeName := '';
    if (RetType and $FFFF0000) = $FFFF0000 then
    begin
      TypeID := RetType and $000000FF;
      TypeName := GetTypeName(TypeID);
    end
    else
      TypeName := GetClassName(ObjectTable, GetWORD(IndirectTable, RetType + 2, Endian), Endian, VbaVersion, Vba64bit);
    FuncDecl := FuncDecl + TypeName;
  end;
  Result := FuncDecl;
end;

function DisasmVarArg(
  ModuleData       : TArray<Byte>;
  Offset           : UInt32;
  WLength          : UInt16;
  Mnemonic         : string;
  Endian           : TEndian;
  VbaVersion       : UInt16;
  Vba64bit         : Boolean
): string;
var
  SubString  : TArray<Byte>;
  VarArgName : string;
  Offset2    : UInt32;
  WordChunk  : UInt16;
  VarNames   : string;
  I          : Integer;
begin
  SubString := Copy(ModuleData, Offset, WLength);
  VarArgName := HexWORD(WLength);
  if (Mnemonic = 'LitStr') or (Mnemonic = 'QuoteRem') or (Mnemonic = 'Rem') or (Mnemonic = 'Reparse') then
    VarArgName := '0x' + VarArgName + ' "' + GetString(SubString, 0, 1252, WLength) + '"'
  else
    if (Mnemonic = 'OnGosub') or (Mnemonic = 'OnGoto') then
    begin
      Offset2 := Offset;
      VarNames := '';
      for I := 1 to WLength shr 1 do
      begin
        WordChunk := ReadWORD(ModuleData, Offset2, Endian);
        if VarNames <> '' then
          VarNames := VarNames + ', ';
        VarNames := VarNames + GetIdentifier(WordChunk, VbaVersion, Vba64bit);
      end;
      VarNames := VarNames + ' ';
    end
    else
      VarArgName := VarArgName + BinToStr(SubString);
  Result := VarArgName;
end;

function ParseLine(
  ModuleData       : TArray<Byte>;
  LineStart        : UInt32;
  LineLength       : UInt16;
  Endian           : TEndian;
  VbaVersion       : UInt16;
  Vba64bit         : Boolean;
  ObjectTable      : TArray<Byte>;
  IndirectTable    : TArray<Byte>;
  DeclarationTable : TArray<Byte>;
  LineNumber       : UInt16
): string;
const
  SpecialValues : array of string = ['False', 'True', 'Null', 'Empty'];
  OptionsValues : array of string = ['Base 0', 'Base 1', 'Compare Text', 'Compare Binary', 'Explicit', 'Private Module'];
var
  Buffer              : TStringBuilder;
  BufferParsed        : string;
  BufferPCode         : string;
  DimType             : string;
  DWordChunk          : UInt32;
  EndOfLineOffset     : UInt32;
  I                   : Integer;
  Instruction         : TInstruction;
  InstructionArgument : Char;
  Offset              : UInt32;
  OpCode              : UInt16;
  OpType              : UInt16;
  ParsedFunc          : string;
  ParsedImp           : string;
  ParsedName          : string;
  ParsedRec           : string;
  ParsedType          : string;
  ParsedVar           : string;
  ParsedVarArg        : string;
  StartOfLineOffset   : UInt32;
  TranslatedOpCode    : UInt16;
  WLength             : UInt16;
  WordChunk           : UInt16;
begin
  Result := '';
  if LineLength = 0 then
    Exit;
  Buffer := TStringBuilder.Create();
  try
    Buffer.AppendLine('                            Line #' + Format('%.5d', [LineNumber]) + '   ' + BinToStr(Copy(moduleData, LineStart, LineLength)));
    Offset := LineStart;
    EndOfLineOffset := LineStart + LineLength;
    while Offset < EndOfLineOffset do
    begin
      StartOfLineOffset := Offset;
      OpCode := ReadWORD(ModuleData, Offset, Endian);
      OpType := (OpCode and not $03FF) shr 10;
      OpCode := OpCode and $03FF;
      TranslatedOpCode := TranslateOpcode(OpCode, VbaVersion, Vba64bit);
      if TranslatedOpCode > 263 then
        raise EVBAParseError.Create('Unrecognized opcode ' + IntToStr(TranslatedOpCode) + ' at offset ' + HexDWORD(Offset));
      Instruction := Instructions[TranslatedOpCode];
      BufferParsed := HexWORD(OpCode) + ' ' + Instruction.Mnemonic;
      if Instruction.Mnemonic = 'Option' then
        BufferParsed := BufferParsed + ' ' + OptionsValues[OpType];
      if (Instruction.Mnemonic = 'Coerce')
      or (Instruction.Mnemonic = 'CoerceVar')
      or (Instruction.Mnemonic = 'DefType') then
        if OpType <= MaxVariantType then
          BufferParsed := BufferParsed + ' (' + VariantTypes[OpType].Description + ')'
        else
          BufferParsed := BufferParsed + ' (' + Format('%.2d', [OpType]) + ')';
      if (Instruction.Mnemonic = 'Dim')
      or (Instruction.Mnemonic = 'DimImplicit')
      or (Instruction.Mnemonic = 'Type') then
      begin
        DimType := '';
        if OpType and $04 <> 0 then
          DimType := DimType + 'Global '
        else
          if OpType and $08 <> 0 then
            DimType := DimType + 'Public '
          else
            if OpType and $10 <> 0 then
              DimType := DimType + 'Private '
            else
              if OpType and $20 <> 0 then
                DimType := DimType + 'Static ';
        if (OpType and $01 <> 0) and (Instruction.Mnemonic <> 'Type') then
          DimType := DimType + 'Const ';
        if DimType <> '' then
          BufferParsed := BufferParsed + ' ' + DimType;
      end;
      if Instruction.Mnemonic = 'LitVarSpecial' then
        BufferParsed := BufferParsed + ' ' + SpecialValues[OpType];
      if (Instruction.Mnemonic = 'Redim')
      or (Instruction.Mnemonic = 'RedimAs') then
        if OpType and 16 <> 0 then
          BufferParsed := BufferParsed + ' Preserve ';
      if (Instruction.Mnemonic = 'ArgsCall')
      or (Instruction.Mnemonic = 'ArgsMemCall')
      or (Instruction.Mnemonic = 'ArgsMemCallWith') then
        if OpType < 16 then
          BufferParsed := BufferParsed + ' (Call) '
        else
          OpType := OpType - 16;
      for I := 1 to Instruction.Arguments.Length do
      begin
        InstructionArgument := Instruction.Arguments[I];
        case InstructionArgument of
          '0': { Int16 }
            BufferParsed := BufferParsed + ' ' + DisasmInt16(ReadWORD(ModuleData, Offset, Endian));
          '1': { Int32 }
            BufferParsed := BufferParsed + ' ' + DisasmInt32(ReadDWORD(ModuleData, Offset, Endian));
          '2': { Int64 }
            BufferParsed := BufferParsed + ' ' + DisasmInt64(ReadQWORD(ModuleData, Offset, Endian));
          '3': { Float32 }
            BufferParsed := BufferParsed + ' ' + DisasmFloat32(ReadDWORD(ModuleData, Offset, Endian));
          '4': { Float64 }
            BufferParsed := BufferParsed + ' ' + DisasmFloat64(ReadQWORD(ModuleData, Offset, Endian));
          'n':
          begin
            WordChunk := ReadWORD(ModuleData, Offset, Endian);
            ParsedName := DisasmName(WordChunk, Instruction.Mnemonic, OpType, VbaVersion, Vba64bit);
            BufferParsed := BufferParsed + ' ' + ParsedName;
          end;
          'i':
          begin
            WordChunk := ReadWORD(ModuleData, Offset, Endian);
            ParsedImp := DisasmImp(IndirectTable, ObjectTable, InstructionArgument, WordChunk, Instruction.Mnemonic, Endian, VbaVersion, Vba64bit);
            BufferParsed := BufferParsed + ' ' + ParsedImp;
          end;
          'c':
          begin
            DWordChunk := ReadDWORD(ModuleData, Offset, Endian);
            BufferParsed := BufferParsed + ' ' + HexDWORD(DWordChunk);
          end;
          'f':
          begin
            DWordChunk := ReadDWORD(ModuleData, Offset, Endian);
            if Length(IndirectTable) >= DWordChunk + 61 then
            begin
              ParsedFunc := DisasmFunc(ObjectTable, IndirectTable, DeclarationTable, DWordChunk, OpType, Endian, VbaVersion, Vba64bit);
              BufferParsed := BufferParsed + ' ' + ParsedFunc;
            end;
          end;
          't':
          begin
            DWordChunk := ReadDWORD(ModuleData, Offset, Endian);
            if Length(IndirectTable) >= DWordChunk + 7 then
            begin
              ParsedType := DisasmType(IndirectTable, DWordChunk);
              BufferParsed := BufferParsed + ' As ' + ParsedRec;
            end;
          end;
          'r':
          begin
            DWordChunk := ReadDWORD(ModuleData, Offset, Endian);
            if Length(IndirectTable) >= DWordChunk + 20 then
            begin
              ParsedRec := DisasmRec(IndirectTable, DWordChunk, Endian, VbaVersion, Vba64bit);
              BufferParsed := BufferParsed + ' ' + ParsedRec;
            end;
          end;
          'v':
          begin
            DWordChunk := ReadDWORD(ModuleData, Offset, Endian);
            if Length(IndirectTable) >= DWordChunk + 16 then
            begin
              if OpType and $20 <> 0 then
                BufferParsed := BufferParsed + ' WithEvents';
              ParsedVar := DisasmVar(IndirectTable, ObjectTable, DWordChunk, Endian, VbaVersion, Vba64bit);
              BufferParsed := BufferParsed + ' ' + ParsedVar;
              if OpType and $10 <> 0 then
              begin
                WordChunk := ReadWORD(ModuleData, Offset, Endian);
                BufferParsed := BufferParsed + ' 0x' + HexWORD(WordChunk);
              end;
            end;
          end;
        end;
      end;
      if Instruction.VarArg then
      begin
        WLength := ReadWORD(ModuleData, Offset, Endian);
        ParsedVarArg := DisasmVarArg(ModuleData, Offset, wLength, Instruction.Mnemonic, Endian, VbaVersion, Vba64bit);
        BufferParsed := BufferParsed + ' ' + ParsedVarArg;
        Offset := Offset + WLength;
        if WLength and 1 <> 0 then
          Offset := Offset + 1
      end;
      BufferPCode := BinToStr(Copy(ModuleData, StartOfLineOffset, Offset - StartOfLineOffset));
      if Length(BufferPCode) > 40 then
        BufferPCode := Copy(BufferPCode, 1, 37) + '...';
      Buffer.AppendLine(Format('%-40s', [BufferPCode]) + '  ' + BufferParsed);
    end;
    Result := Buffer.ToString();
  finally
    Buffer.Free();
  end;
end;

function ParseModule(
  const ModuleData: TArray<Byte>;
  const VBAProgram: TVBAProgram
): string;
var
  Buffer           : TStringBuilder;
  DeclarationTable : TArray<Byte>;
  DWLength         : UInt32;
  DWLength2        : UInt32;
  Endian           : TEndian;
  IndirectTable    : TArray<Byte>;
  Line             : UInt16;
  LineLength       : UInt16;
  LineOffset       : UInt32;
  MagicSignature   : UInt16;
  NumberOfLines    : UInt16;
  ObjectTable      : TArray<Byte>;
  Offset           : UInt32;
  Offset2          : UInt32;
  PCodeStart       : UInt32;
  TableStart       : UInt32;
  Vba64bit         : Boolean;
  VbaVersion       : Byte;
  Version          : UInt16;
begin
  Buffer := TStringBuilder.Create();
  try
    // Determine endinanness: PC (little-endian) or Mac (big-endian)
    if GetWord(ModuleData, 2, TEndian.LittleEndian) > $FF then
      Endian := TEndian.BigEndian
    else
      Endian := TEndian.LittleEndian;
    Vba64bit := VBAProgram.SysKind = $00000003;
    VbaVersion := 3;
    Version := GetWORD(VBAProgram.VbaProjectData, 2, TEndian.LittleEndian);
    if Version >= $6B then
    begin
      if Version >= $97 then
        VbaVersion := 7
      else
        VbaVersion := 6;
        if Vba64bit then
        begin
          DWLength := GetDWord(ModuleData, $0043, endian);
          DeclarationTable := Copy(ModuleData, $0047, DWLength);
          DWLength := GetDWord(ModuleData, $0011, endian);
          TableStart := DWLength + 12;
        end
        else
        begin
          DWLength := GetDWord(ModuleData, $003F, endian);
          DeclarationTable := Copy(ModuleData, $0043, DWLength);
          DWLength := GetDWord(ModuleData, $0011, endian);
          TableStart := DWLength + 10;
        end;
        DWLength := getDWord(moduleData, TableStart, endian);
        TableStart := TableStart + 4;
        IndirectTable := Copy(ModuleData, TableStart, DWLength);
        DWLength := GetDWord(ModuleData, $0005, endian);
        DWLength2 := DWLength + $8A;
        DWLength := GetDWord(ModuleData, DWLength2, endian);
        DWLength2 := DWLength2 + 4;
        ObjectTable := Copy(ModuleData, DWLength2, DWLength);
        Offset := $0019;
    end
    else
    begin
      VbaVersion := 5;
      Offset := 11;
      DWLength := GetDWord(ModuleData, Offset, endian);
      Offset2 := Offset + 4;
      DeclarationTable := Copy(ModuleData, Offset2, DWLength);
      Offset := SkipStructure(ModuleData, Offset, endian,  True,  1, False);
      Offset := Offset + 64;
      Offset := SkipStructure(ModuleData, Offset, endian, False, 16, False);
      Offset := SkipStructure(ModuleData, Offset, endian,  True,  1, False);
      Offset := Offset + 6;
      Offset := SkipStructure(ModuleData, Offset, endian,  True,  1, False);
      Offset2 := Offset + 8;
      DWLength := GetDWord(moduleData, Offset2, endian);
      TableStart := DWLength + 14;
      Offset2 := DWLength + 10;
      DWLength := GetDWord(moduleData, Offset2, endian);
      IndirectTable := Copy(ModuleData, TableStart, DWLength);
      DWLength := GetDWord(moduleData, Offset, endian);
      Offset2 := DWLength + $008A;
      DWLength := GetDWord(moduleData, Offset2, endian);
      Offset2 := Offset2 + 4;
      ObjectTable := Copy(ModuleData, Offset2, DWLength);
      Offset := Offset + 77;
    end;
    DWLength := getDWord(ModuleData, Offset, endian);
    Offset := DWLength + $003C;
    MagicSignature := ReadWORD(ModuleData, Offset, endian);
    if MagicSignature <> $CAFE then
      Exit;
    Offset := Offset + 2;
    NumberOfLines := ReadWORD(ModuleData, Offset, Endian);
    PCodeStart := Offset + NumberOfLines * 12 + 10;
    for Line := 1 to NumberOfLines do
    begin
      Offset := Offset + 4;
      LineLength := ReadWORD(ModuleData, Offset, Endian);
      Offset := Offset + 2;
      LineOffset := ReadDWORD(ModuleData, Offset, Endian);
      Buffer.Append(ParseLine(ModuleData, PCodeStart + LineOffset, LineLength, Endian, VbaVersion, Vba64bit, ObjectTable, IndirectTable, DeclarationTable, line));
    end;
    Buffer.AppendLine();
    Buffer.AppendLine('Declaration table: ').AppendLine(ParseSimple(DeclarationTable)).AppendLine();
    Buffer.AppendLine('Indirect table: ').AppendLine(ParseSimple(IndirectTable)).AppendLine();
    Buffer.AppendLine('Object table: ').AppendLine(ParseSimple(ObjectTable)).AppendLine();
    Result := Buffer.ToString();
  finally
    Buffer.Free();
  end;
end;

procedure ReadIdentifiers(const VbaProjectData: TArray<Byte>; CodePage: UInt32);
var
  ByteChunk           : byte;
  Endian              : TEndian;
  I                   : UInt32;
  Identifier          : string;
  IdentifierNumber    : UInt16;
  IdLength            : Byte;
  IdType              : Byte;
  IsKeyword           : Boolean;
  MagicSignature      : UInt16;
  NonUnicodeName      : Boolean;
  NumberOfIdentifiers : UInt16;
  NumberOfProjects    : UInt16;
  NumberOfReferences  : UInt16;
  Offset              : UInt32;
  ReferenceLength     : UInt16;
  UnicodeName         : Boolean;
  UnicodeRef          : Boolean;
  Version             : UInt16;
  W0, W1              : UInt16;
  WLength             : UInt16;
  WordChunk           : UInt16;
begin
  Reset();
  MagicSignature := GetWord(VbaProjectData, 0, TEndian.LittleEndian);
  if MagicSignature <> $61CC then
    Exit;
  Version := getWord(VbaProjectData, 2, TEndian.LittleEndian);
  UnicodeRef := (Version >= $5B) and (not (Version in [$60, $62, $63])) or (Version = $4E);
  UnicodeName := (Version >= $59) and (not (Version in [$60, $62, $63])) or (Version = $4E);
  NonUnicodeName := ((Version <= $59) and (Version <> $4E)) or ((Version < $5F) and (Version > $6B));
  WordChunk := getWord(VbaProjectData, 5, TEndian.LittleEndian);
  if WordChunk = $000E then
    Endian := TEndian.BigEndian
  else
    Endian := TEndian.LittleEndian;
  Offset := $1E;
  NumberOfReferences := ReadWORD(VbaProjectData, Offset, Endian);
  Offset := Offset + 2;
  for I := 1 to NumberOfReferences do
  begin
    ReferenceLength := ReadWORD(VbaProjectData, Offset, Endian);
    if ReferenceLength = 0 then
      Offset := Offset + 6
    else
      if (((UnicodeRef) and (ReferenceLength < 5)) or ((not UnicodeRef) and (ReferenceLength < 3))) then
        Offset := Offset + ReferenceLength
      else
      begin
        if UnicodeRef then
          ByteChunk := VbaProjectData[Offset + 4]
        else
          ByteChunk := VbaProjectData[Offset + 2];
        Offset := Offset + ReferenceLength;
        if (ByteChunk = 67) or (ByteChunk = 68) then
          Offset := SkipStructure(VbaProjectData, Offset, Endian, False, 1, False);
      end;
    Offset := Offset + 10;
    WordChunk := ReadWORD(VbaProjectData, Offset, Endian);
    if WordChunk <> 0 then
    begin
      Offset := SkipStructure(VbaProjectData, Offset, Endian, False, 1, False);
      WLength := ReadWORD(VbaProjectData, Offset, Endian);
      if WLength <> 0 then
        Offset := Offset + 2;
      Offset := Offset + WLength + 30
    end;
  end;
  // Number of entries in the class/user forms table
  Offset := SkipStructure(VbaProjectData, Offset, Endian, False, 2, False);
  // Number of compile-time identifier-value pairs
  Offset := SkipStructure(VbaProjectData, Offset, Endian, False, 4, False);
  Offset := Offset + 2;
  // Typeinfo typeID
  Offset := SkipStructure(VbaProjectData, Offset, Endian, False, 1, True);
  // Project description
  Offset := SkipStructure(VbaProjectData, Offset, Endian, False, 1, True);
  // Project help file name
  Offset := skipStructure(VbaProjectData, Offset, Endian, False, 1, True);
  Offset := Offset + $64;
  // Skip the module descriptors
  NumberOfProjects := ReadWORD(VbaProjectData, Offset, Endian);
  for I := 1 to NumberOfProjects do
  begin
    WLength := ReadWORD(VbaProjectData, Offset, Endian);
    // Code module name
    if UnicodeName then
      Offset := Offset + WLength;
    if NonUnicodeName then
    begin
      if WLength <> 0 then
        WLength := ReadWORD(VbaProjectData, Offset, Endian);
      Offset := Offset + WLength;
    end;
    // Stream time
    Offset := SkipStructure(VbaProjectData, Offset, Endian, False, 1, False);
    Offset := SkipStructure(VbaProjectData, Offset, Endian, False, 1, True);
    ReadWORD(VbaProjectData, Offset, Endian);
    if Version >= $6B then
      Offset := SkipStructure(VbaProjectData, Offset, Endian, False, 1, True);
    Offset := SkipStructure(VbaProjectData, Offset, Endian, False, 1, True);
    Offset := Offset + 2;
    if Version <> $51 then
      Offset := Offset + 4;
    Offset := SkipStructure(VbaProjectData, Offset, Endian, False, 8, False);
    Offset := Offset + 11;
  end;
  Offset := Offset + 6;
  Offset := SkipStructure(VbaProjectData, Offset, Endian, True, 1, False);
  Offset := Offset + 6;
  W0 := ReadWORD(VbaProjectData, Offset, Endian);
  NumberOfIdentifiers := ReadWORD(VbaProjectData, Offset, Endian);
  W1 := ReadWORD(VbaProjectData, Offset, Endian);
  Offset := Offset + 4;
  IdentifierNumber := W0 - NumberOfIdentifiers;
  Identifiers := TDictionary<UInt16, string>.Create(NumberOfIdentifiers);
  for I := 1 to NumberOfIdentifiers do
  begin
    IsKeyword := False;
    Identifier := '';
    GetTypeAndLength(VbaProjectData, Offset, Endian, IdType, IdLength);
    Offset := Offset + 2;
    if (IdLength = 0) and (IdType = 0) then
    begin
      Offset := Offset + 2;
      GetTypeAndLength(VbaProjectData, Offset, Endian, IdType, IdLength);
      Offset := Offset + 2;
      IsKeyword := True;
    end;
    if IdType and $80 <> 0 then
      Offset := Offset + 6;
    if IdLength > 0 then
    begin
      Identifier := GetString(VbaProjectData, Offset, CodePage, IdLength);
      Inc(IdentifierNumber);
      Identifiers.Add(IdentifierNumber, Identifier);
      Offset := Offset + IdLength;
    end;
    if not IsKeyword then
      Offset := Offset + 4;
  end;
end;

function ParseSimple(const CodeBytes: TArray<Byte>): string;
var
  CodeLength    : UInt32;
  CodeText      : string;
  I             : UInt32;
  StringBuilder : TStringBuilder;
  TextChars     : string;
  TextLine      : string;
begin
  Result := '';
  if Length(CodeBytes) > 0 then
  begin
    StringBuilder := TStringBuilder.Create();
    try
      CodeLength := High(CodeBytes);
      CodeText := BinToStr(CodeBytes);
      TextChars := '';
      TextLine := '';
      I := 0;
      while I <= CodeLength do
      begin
        TextLine := TextLine + CodeText.Chars[I shl 1];
        TextLine := TextLine + CodeText.Chars[I shl 1 + 1];
        if CodeBytes[I] >= 32 then
          TextChars := TextChars + Chr(CodeBytes[I])
        else
          TextChars := TextChars + '.';
        Inc(I);
        if I mod 20 = 0 then
        begin
          StringBuilder.AppendLine(TextLine + '        ' + TextChars);
          TextLine := '';
          TextChars := '';
        end;
      end;
      for I := TextLine.Length to 39 do
        TextLine := TextLine + ' ';
      StringBuilder.AppendLine(TextLine + '        ' + TextChars);
      Result := StringBuilder.ToString();
    finally
      StringBuilder.Free();
    end;
  end;
end;

procedure ParsePCode(const VBAProgram: TVBAProgram; var Module: TModule);
var
  ParseResult: string;
begin
  try
    if Identifiers <> nil then
      ParseResult := ParseModule(Module.PerformanceCache, VBAProgram)
    else
      ParseResult := ParseSimple(Module.PerformanceCache);
  except
    ParseResult := ParseSimple(Module.PerformanceCache);
  end;
  Module.ParsedPCode := ParseResult;
end;

initialization
  Reset();

finalization
  Reset();

end.
