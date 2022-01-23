unit UnitFormVBAMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons,
  Vcl.ActnList, System.Actions, Vcl.StdActns, System.ImageList, Vcl.ImgList,
  Vcl.Grids, System.Classes, Vcl.Menus, SynEdit, SynEditHighlighter,
  SynHighlighterVB, SynHighlighterST, SynEditSearch, SynEditMiscClasses,
  UnitFormVBAProperties, ParserVBA, Common, Clipbrd;

type
  TFormVBAView = class(TForm)
    ActionExportAll: TAction;
    ActionExportThis: TAction;
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
    MenuItemExit: TMenuItem;
    MenuItemFile: TMenuItem;
    MenuItemFileExportAll: TMenuItem;
    MenuItemFileExportThis: TMenuItem;
    MenuItemFileOpen: TMenuItem;
    MenuItemSearch: TMenuItem;
    MenuItemSearchFind: TMenuItem;
    MenuItemSearchFindFirst: TMenuItem;
    MenuItemSearchFindNext: TMenuItem;
    MenuItemView: TMenuItem;
    MenuItemViewInformation: TMenuItem;
    PanelModulePCode: TPanel;
    PanelModules: TPanel;
    PanelModuleVB: TPanel;
    SaveAll1: TMenuItem;
    SearchFind: TSearchFind;
    SearchFindFirst: TSearchFindFirst;
    SearchFindNext: TSearchFindNext;
    StringGridModules: TStringGrid;
    SynEditPCode: TSynEdit;
    SynEditSearch: TSynEditSearch;
    SynEditVB: TSynEdit;
    SynVBSyn: TSynVBSyn;
    procedure ActionExportAllExecute(Sender: TObject);
    procedure ActionExportThisExecute(Sender: TObject);
    procedure ActionProjectPropertiesExecute(Sender: TObject);
    procedure FileOpenAccept(Sender: TObject);
    procedure FileSaveAllAccept(Sender: TObject);
    procedure FileSaveThisAccept(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure GridPanelMainLayoutResize(Sender: TObject);
    procedure SearchFindFindDialogFind(Sender: TObject);
    procedure SearchFindFirstFindDialogFind(Sender: TObject);
    procedure StringGridModulesSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
  private
    { Private declarations }
    var VBAProgram : TVBAProgram;
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

uses System.Generics.Collections, System.UITypes;

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

procedure TFormVBAView.ActionProjectPropertiesExecute(Sender: TObject);
begin
  UnitFormVBAProperties.FormVBAProperties.SetReference(VBAProgram);
end;

procedure TFormVBAView.FileOpenAccept(Sender: TObject);
var
  FileName    : string;
  ParseResult : TProcessFileResult;
begin
  ActionExportAll.Enabled := False;
  ActionExportThis.Enabled := False;
  ActionProjectProperties.Enabled := False;
  FileSaveAll.Enabled := False;
  FileSaveThis.Enabled := False;
  SearchFind.Enabled := False;
  SearchFindFirst.Enabled := False;
  SearchFindNext.Enabled := False;
  FileName := FileOpen.Dialog.FileName;
  ResetGrid();
  try
    ParseResult := ParserVBA.ParseFile(FileName, VBAProgram);
    if ParseResult = TProcessFileResult.pfOk then
    begin
      ActionExportAll.Enabled := True;
      ActionExportThis.Enabled := True;
      ActionProjectProperties.Enabled := True;
      FileSaveAll.Enabled := True;
      FileSaveThis.Enabled := True;
      SearchFind.Enabled := True;
      SearchFindFirst.Enabled := True;
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

procedure TFormVBAView.SearchFindFindDialogFind(Sender: TObject);
var
  SearchText : string;
begin
  SearchText := SearchFind.Dialog.FindText;
  SynEditSearch.FindFirst(SearchText);
end;

procedure TFormVBAView.SearchFindFirstFindDialogFind(Sender: TObject);
var
  I          : Int32;
  SearchText : string;
begin
  SearchText := SearchFindFirst.Dialog.FindText;
  if VBAProgram.ModulesCount > 0 then
    for I := 0 to VBAProgram.ModulesCount - 1 do
      if VBAProgram.Module[I].SourceCode.ToUpper().Contains(SearchText.ToUpper()) then
      begin
        StringGridModules.Selection := TGridRect(Rect(0, I + 1, 2, I + 1));
        SynEditVB.Text := VBAProgram.Module[I].SourceCode;
        SynEditPCode.Text := VBAProgram.Module[I].ParsedPCode;
      end;
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
