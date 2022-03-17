unit UnitFormVBAMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons,
  Vcl.ActnList, System.Actions, Vcl.StdActns, System.ImageList, Vcl.ImgList,
  Vcl.Grids, System.Classes, Vcl.Menus, SynEdit, SynEditHighlighter,
  SynHighlighterVB, SynHighlighterST, SynEditSearch, SynEditMiscClasses,
  UnitFormVBAProperties, UnitFormModuleSearch, UnitFormVBACheck,
  UnitFormAbout, UnitFormThemeSettings, UnitFormVBASettings, ParserVBA, Common;

type
  TFormVBAView = class(TForm)
    ActionCheck: TAction;
    ActionExportAll: TAction;
    ActionExportThis: TAction;
    ActionGlobalSearch: TAction;
    ActionHelp1031: TAction;
    ActionHelp1033: TAction;
    ActionHelp1040: TAction;
    ActionList: TActionList;
    ActionProjectProperties: TAction;
    EditFileSpec: TEdit;
    Exportthismodule1: TMenuItem;
    FileExit: TFileExit;
    FileOpen: TFileOpen;
    FileSaveAll: TFileSaveAs;
    FileSaveThis: TFileSaveAs;
    GridPanelMainLayout: TGridPanel;
    GridPanelModule: TGridPanel;
    ImageList: TImageList;
    LabelModules: TLabel;
    LabelPCode: TLabel;
    LabelVB: TLabel;
    MainMenu: TMainMenu;
    MenuItemCheckAutoruns: TMenuItem;
    MenuItemExit: TMenuItem;
    MenuItemFile: TMenuItem;
    MenuItemFileExportAll: TMenuItem;
    MenuItemFileExportThis: TMenuItem;
    MenuItemFileOpen: TMenuItem;
    MenuItemHelp1031: TMenuItem;
    MenuItemHelp1033: TMenuItem;
    MenuItemHelp1040: TMenuItem;
    MenuItemSearch: TMenuItem;
    MenuItemSearchFind: TMenuItem;
    MenuItemSearchFindNext: TMenuItem;
    MenuItemSearchGlobalSearch: TMenuItem;
    MenuItemView: TMenuItem;
    MenuItemViewInformation: TMenuItem;
    PanelModulePCode: TPanel;
    PanelModules: TPanel;
    PanelModuleVB: TPanel;
    SaveAll1: TMenuItem;
    SearchFind: TSearchFind;
    SearchFindNext: TSearchFindNext;
    StringGridModules: TStringGrid;
    SynEditPCode: TSynEdit;
    SynEditSearch: TSynEditSearch;
    SynEditVB: TSynEdit;
    SynVBSyn: TSynVBSyn;
    MenuItemHelp: TMenuItem;
    MenuItemHelpAbout: TMenuItem;
    ActionHelpAbout: TAction;
    ActionThemeSettings: TAction;
    ActionVBASettings: TAction;
    MenuItemThemeSettings: TMenuItem;
    MenuItemVBASettings: TMenuItem;
    procedure ActionCheckExecute(Sender: TObject);
    procedure ActionExportAllExecute(Sender: TObject);
    procedure ActionExportThisExecute(Sender: TObject);
    procedure ActionGlobalSearchExecute(Sender: TObject);
    procedure ActionHelp1031Execute(Sender: TObject);
    procedure ActionHelp1033Execute(Sender: TObject);
    procedure ActionHelp1040Execute(Sender: TObject);
    procedure ActionProjectPropertiesExecute(Sender: TObject);
    procedure FileOpenAccept(Sender: TObject);
    procedure FileSaveAllAccept(Sender: TObject);
    procedure FileSaveThisAccept(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure GridPanelMainLayoutResize(Sender: TObject);
    procedure StringGridModulesSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
    procedure ActionHelpAboutExecute(Sender: TObject);
    procedure ActionVBASettingsExecute(Sender: TObject);
    procedure ActionThemeSettingsExecute(Sender: TObject);
  private
    { Private declarations }
    var VBAProgram : TVBAProgram;
    procedure CreateWnd(); override;
    procedure DestroyWnd(); override;
    procedure DropFiles(var Msg: TWMDropFiles); message WM_DROPFILES;
    procedure ReadFile(const FileName: string);
    procedure ResetGrid();
    procedure ShowErrorMessage(const E: Exception);
    procedure ShowHintMessage(const ProcessFileResult: TProcessFileResult);
    procedure ShowWarningMessage(const E: EVBAParseError);
    procedure UpdateGrid();
  public
    { Public declarations }
  end;

var
  FormVBAView: TFormVBAView;

implementation

{$R *.dfm}

uses System.Generics.Collections, System.UITypes, Winapi.ShellAPI, Vcl.Clipbrd;

const
  MessageFileNotFound = 'File not found!';
  MessageInvalidFileExtension = 'Unsupported file extension! This program ' +
    'can parse only Microsoft Excel and Microsoft Word files, with the ' +
    'following extensions: .BIN .DOC .DOCM .DOCX .DOT .DOTM .DOTX .OTM ' +
    '.POTM .POTX .PPTM .PPTX .XLAM .XLS .XLSB .XLSM .XLSX';
  MessageInvalidFileContent = 'Could not find VBA macros in the selected filed.';
  MessageHeaderSizeError = 'The file has a wrong header size: maybe the file ' +
    'file is damaged or it is not a valid Microsoft Excel or Microsoft Word file.';
  MessageHeaderSignatureError = 'The file has a wrong header signature: maybe the ' +
    'file is damaged or it is not a valid Microsoft Excel or Microsoft Word file.';
  MessageHeaderGUIDError = 'The file has a wrong header GUID: maybe the ' +
    'file is damaged or it is not a valid Microsoft Excel or Microsoft Word file.';
  MessageByteOrderError = 'The file has a wrong byte order sequence: maybe the ' +
    'file is damaged or it is not a valid Microsoft Excel or Microsoft Word file.';
  MessageDirectorySectorNumberError = 'The file has an internal problema: maybe the ' +
    'file is damaged or it has been modified in a wrong way.';
  MessageParseError = 'An unexpected error has occured during parsing the file.'#13#10 +
    'Please note that this application is an BETA version, and some information ' +
    'are not well documented by Microsoft. If you want report the error to the author, ' +
    'please include the following information:'#13#10;
  MessageError = 'An enexpected error has occured, causing the failure of the '#13#10 +
    'parsing and probably this application needs to be restarted. Please note that ' +
    'this application is an BETA version. If you want report the error to the author, ' +
    'please include the following information:'#13#10;

procedure TFormVBAView.ActionCheckExecute(Sender: TObject);
begin
  UnitFormVBACheck.FormVBACheck.SetReference(VBAProgram);
end;

procedure TFormVBAView.ActionExportAllExecute(Sender: TObject);
var
  I             : Int32;
  StringBuilder : TStringBuilder;
begin
  StringBuilder := TStringBuilder.Create();
  try
    for I := 0 to VBAProgram.ModulesCount - 1 do
    begin
      StringBuilder.AppendFormat(''' Module ', []).AppendFormat(VBAProgram.Module[I].ModuleName, []).AppendLine();
      StringBuilder.AppendFormat(VBAProgram.Module[I].SourceCode, []);
      StringBuilder.AppendLine();
    end;
    Clipboard().AsText := StringBuilder.ToString();
  finally
    FreeAndNil(StringBuilder);
  end;
end;

procedure TFormVBAView.ActionExportThisExecute(Sender: TObject);
begin
  Clipboard().AsText := SynEditVB.Text;
end;

procedure TFormVBAView.ActionGlobalSearchExecute(Sender: TObject);
begin
  UnitFormModuleSearch.FormModuleSearch.SetReference(VBAProgram);
end;

procedure TFormVBAView.ActionHelp1031Execute(Sender: TObject);
begin
  ShellExecute(Handle, 'open', '1031\VBUI6.CHM', nil, nil, SW_SHOW);
end;

procedure TFormVBAView.ActionHelp1033Execute(Sender: TObject);
begin
  ShellExecute(Handle, 'open', '1033\VBUI6.CHM', nil, nil, SW_SHOW);
end;

procedure TFormVBAView.ActionHelp1040Execute(Sender: TObject);
begin
  ShellExecute(Handle, 'open', '1040\VBUI6.CHM', nil, nil, SW_SHOW);
end;

procedure TFormVBAView.ActionHelpAboutExecute(Sender: TObject);
begin
  UnitFormAbout.FormAbout.Show();
end;

procedure TFormVBAView.ActionProjectPropertiesExecute(Sender: TObject);
begin
  UnitFormVBAProperties.FormVBAProperties.SetReference(VBAProgram);
end;

procedure TFormVBAView.ActionThemeSettingsExecute(Sender: TObject);
begin
  UnitFormThemeSettings.FormThemeSettings.Show();
end;

procedure TFormVBAView.ActionVBASettingsExecute(Sender: TObject);
begin
  UnitFormVBASettings.FormVBASettings.Show();
end;

procedure TFormVBAView.CreateWnd();
begin
  inherited;
  DragAcceptFiles(Handle, True);
end;

procedure TFormVBAView.DestroyWnd();
begin
  inherited;
  DragAcceptFiles(Handle, False);
end;

procedure TFormVBAView.DropFiles(var Msg: TWMDropFiles);
const
  Index = 0;
var
  FileName : string;
  Size     : UINT;
begin
  Size := DragQueryFile(Msg.Drop, Index, nil, 0);
  FileName := '';
  SetLength(FileName, Size);
  DragQueryFile(Msg.Drop, Index, PChar(FileName), Size + 1);
  ReadFile(FileName);
end;

procedure TFormVBAView.FileOpenAccept(Sender: TObject);
var
  FileName    : string;
begin
  FileName := FileOpen.Dialog.FileName;
  ReadFile(FileName);
end;

procedure TFormVBAView.FileSaveAllAccept(Sender: TObject);
var
  Encoding   : TEncoding;
  FileData   : TBytes;
  FileExt    : string;
  FileName   : string;
  FilePath   : string;
  FileStream : TFileStream;
  I: Integer;
begin
  FilePath := ExtractFilePath(FileSaveAll.Dialog.FileName);
  if FileSaveAll.Dialog.FilterIndex = 1 then
    FileExt := '.bas'
  else
    FileExt := '.txt';
  Encoding := nil;
  FileStream := nil;
  try
    case FileSaveThis.Dialog.FilterIndex of
      1:
        Encoding := TMBCSEncoding.Create();
      2:
        Encoding := TMBCSEncoding.Create();
      3:
        Encoding := TUTF8Encoding.Create();
      4:
        Encoding := TUnicodeEncoding.Create();
    end;
    for I := 0 to VBAProgram.ModulesCount - 1 do
    begin
      FileName := FilePath + VBAProgram.Module[I].ModuleName + FileExt;
      FileData := Encoding.GetBytes(VBAProgram.Module[I].SourceCode);
      FileStream := TFileStream.Create(FileName, fmCreate);
      try
        FileStream.WriteBuffer(FileData, Length(FileData));
      finally
        FreeAndNil(FileStream);
      end;
    end;
  finally
    FreeAndNil(Encoding);
  end;
end;

procedure TFormVBAView.FileSaveThisAccept(Sender: TObject);
var
  Encoding   : TEncoding;
  FileData   : TBytes;
  FileName   : string;
  FileStream : TFileStream;
begin
  FileName := FileSaveThis.Dialog.FileName;
  if ExtractFileExt(FileName) = '' then
    if FileSaveThis.Dialog.FilterIndex = 1 then
      FileName := FileName + '.bas'
    else
      FileName := FileName + '.txt';
  Encoding := nil;
  FileStream := nil;
  try
    case FileSaveThis.Dialog.FilterIndex of
      1:
        Encoding := TMBCSEncoding.Create();
      2:
        Encoding := TMBCSEncoding.Create();
      3:
        Encoding := TUTF8Encoding.Create();
      4:
        Encoding := TUnicodeEncoding.Create();
    end;
    FileData := Encoding.GetBytes(SynEditVB.Text);
    FileStream := TFileStream.Create(FileName, fmCreate);
    FileStream.WriteBuffer(FileData, Length(FileData));
  finally
    FreeAndNil(FileStream);
    FreeAndNil(Encoding);
  end;
end;

procedure TFormVBAView.FormActivate(Sender: TObject);
begin
  ActionHelp1031.Enabled := FileExists('1031\VBUI6.CHM');
  ActionHelp1033.Enabled := FileExists('1033\VBUI6.CHM');
  ActionHelp1040.Enabled := FileExists('1040\VBUI6.CHM');
end;

procedure TFormVBAView.FormCreate(Sender: TObject);
begin
  StringGridModules.Cells[0, 0] := 'Module';
  StringGridModules.Cells[1, 0] := 'Source';
  StringGridModules.Cells[2, 0] := 'P-Code';
  StringGridModules.ColWidths[0] := 240;
  StringGridModules.ColWidths[1] := 60;
  StringGridModules.ColWidths[2] := 60;
end;

procedure TFormVBAView.GridPanelMainLayoutResize(Sender: TObject);
begin
  PanelModules.Height := GridPanelMainLayout.CellSize[0,1].Y - 8;
end;

procedure TFormVBAView.ReadFile(const FileName: string);
var
  ParseResult : TProcessFileResult;
begin
  ActionCheck.Enabled := False;
  ActionExportAll.Enabled := False;
  ActionExportThis.Enabled := False;
  ActionGlobalSearch.Enabled := False;
  ActionProjectProperties.Enabled := False;
  FileSaveAll.Enabled := False;
  FileSaveThis.Enabled := False;
  SearchFind.Enabled := False;
  SearchFindNext.Enabled := False;
  ResetGrid();
  try
    ParseResult := ParserVBA.ParseFile(FileName, VBAProgram);
    if ParseResult = TProcessFileResult.pfOk then
    begin
      ActionCheck.Enabled := True;
      ActionExportAll.Enabled := True;
      ActionExportThis.Enabled := True;
      ActionGlobalSearch.Enabled := True;
      ActionProjectProperties.Enabled := True;
      FileSaveAll.Enabled := True;
      FileSaveThis.Enabled := True;
      SearchFind.Enabled := True;
      SearchFindNext.Enabled := True;
      EditFileSpec.Text := FileName;
      UpdateGrid();
    end
    else
      ShowHintMessage(ParseResult);
    except
      on E: EVBAParseError do
        ShowWarningMessage(E);
      on E: Exception do
        ShowErrorMessage(E);
  end;
end;

procedure TFormVBAView.ResetGrid();
var
  I: UInt32;
begin
  EditFileSpec.Text := '';
  StringGridModules.RowCount := 2;
  for I := 0 to StringGridModules.ColCount - 1 do
    StringGridModules.Cells[I, 1] := '';
  SynEditPCode.Text := '';
  SynEditVB.Text := '';
end;

procedure TFormVBAView.ShowErrorMessage(const E: Exception);
var
  ReportMessage: string;
begin
  ReportMessage := MessageError + E.Message;
  MessageDlg(ReportMessage, TMsgDlgType.mtWarning, [mbOk], 0);
end;

procedure TFormVBAView.ShowHintMessage(const ProcessFileResult: TProcessFileResult);
begin
  case ProcessFileResult of
    TProcessFileResult.pfInvalidFileExtension:
      MessageDlg(MessageInvalidFileExtension, mtInformation, [mbOK], 0);
    TProcessFileResult.pfInvalidFileContent:
      MessageDlg(MessageInvalidFileContent, mtInformation, [mbOK], 0);
    TProcessFileResult.pfHeaderSizeError:
      MessageDlg(MessageHeaderSizeError, mtInformation, [mbOK], 0);
    TProcessFileResult.pfHeaderSignatureError:
      MessageDlg(MessageHeaderSignatureError, mtInformation, [mbOK], 0);
    TProcessFileResult.pfHeaderGUIDError:
      MessageDlg(MessageHeaderGUIDError, mtInformation, [mbOK], 0);
    TProcessFileResult.pfHeaderByteOrderError:
      MessageDlg(MessageInvalidFileContent, mtInformation, [mbOK], 0);
    TProcessFileResult.pfHeaderDirectorySectorNumberError:
      MessageDlg(MessageByteOrderError, mtInformation, [mbOK], 0);
    TProcessFileResult.pfNotFound:
      MessageDlg(MessageDirectorySectorNumberError, mtInformation, [mbOK], 0);
  end;
end;

procedure TFormVBAView.ShowWarningMessage(const E: EVBAParseError);
var
  ReportMessage: string;
begin
  ReportMessage := MessageParseError + E.Message;
  MessageDlg(ReportMessage, TMsgDlgType.mtWarning, [mbOk], 0);
end;

procedure TFormVBAView.StringGridModulesSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
  if VBAProgram.ModulesCount > 0 then
  begin
    SynEditVB.ClearAll();
    SynEditVB.Text := VBAProgram.Module[ARow - 1].SourceCode;
    SynEditPCode.ClearAll();
    SynEditPCode.Text := VBAProgram.Module[ARow - 1].ParsedPCode;
  end;
end;

procedure TFormVBAView.UpdateGrid();
var
  I : Int32;
begin
  if VBAProgram.ModulesCount > 0 then
  begin
    StringGridModules.RowCount := VBAProgram.ModulesCount + 1;
    for I := 0 to VBAProgram.ModulesCount - 1 do
    begin
      StringGridModules.Cells[0, I + 1] := VBAProgram.Module[I].ModuleName;
      StringGridModules.Cells[1, I + 1] := IntToStr(Length(VBAProgram.Module[I].SourceCode));
      StringGridModules.Cells[2, I + 1] := IntToStr(Length(VBAProgram.Module[I].PerformanceCache));
    end;
  end;
end;

end.
