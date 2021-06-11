unit UnitFormVBAMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons,
  Vcl.ActnList, System.Actions, Vcl.StdActns, System.ImageList, Vcl.ImgList,
  Vcl.Grids, System.Classes, SynEdit, SynEditHighlighter,
  SynHighlighterVB, SynHighlighterST,
  UnitFormVBAProperties, ParserVBA, Common;

type
  TFormVBAView = class(TForm)
    ActionList: TActionList;
    ActionProjectProperties: TAction;
    FileOpen: TFileOpen;
    ImageList: TImageList;
    LabeledEditFileSpec: TLabeledEdit;
    LabelModules: TLabel;
    LabelPCode: TLabel;
    LabelVB: TLabel;
    SpeedButtonFileOpen: TSpeedButton;
    SpeedButtonInformation: TSpeedButton;
    StringGridModules: TStringGrid;
    SynEditPCode: TSynEdit;
    SynEditVB: TSynEdit;
    SynSTSyn: TSynSTSyn;
    SynVBSyn: TSynVBSyn;
    procedure ActionProjectPropertiesExecute(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FileOpenAccept(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure StringGridModulesSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
  private
    { Private declarations }
    VBAProgram : TVBAProgram;
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
    'Please note that this application is an ALPHA version, and some information ' +
    'are not well documented by Microsoft. If you want report the error to the author, ' +
    'please include the following information:'#13#10;
  MessageError = 'An enexpected error has occured, causing the failure of the '#13#10 +
    'parsing and probably this application needs to be restarted. Please note that ' +
    'this application is an ALPHA version. If you want report the error to the author, ' +
    'please include the following information:'#13#10;

procedure TFormVBAView.ActionProjectPropertiesExecute(Sender: TObject);
begin
  UnitFormVBAProperties.FormVBAProperties.SetReference(VBAProgram);
end;

procedure TFormVBAView.FileOpenAccept(Sender: TObject);
var
  FileName    : string;
  ParseResult : TProcessFileResult;
begin
  ActionProjectProperties.Enabled := False;
  FileName := FileOpen.Dialog.FileName;
  ResetGrid();
  try
    ParseResult := ParserVBA.ParseFile(FileName, VBAProgram);
    if ParseResult = TProcessFileResult.pfOk then
    begin
      ActionProjectProperties.Enabled := True;
      LabeledEditFileSpec.Text := FileName;
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

procedure TFormVBAView.FormCreate(Sender: TObject);
begin
  StringGridModules.Cells[0, 0] := 'Module';
  StringGridModules.Cells[1, 0] := 'Source';
  StringGridModules.Cells[2, 0] := 'P-Code';
  StringGridModules.ColWidths[0] := 240;
  StringGridModules.ColWidths[1] := 60;
  StringGridModules.ColWidths[2] := 60;
end;

procedure TFormVBAView.FormResize(Sender: TObject);
var
  ScreenHeight : Int32;
begin
  ScreenHeight := Height - StringGridModules.Top - 64;
  LabeledEditFileSpec.Width := Width - LabeledEditFileSpec.Left - 32;
  StringGridModules.Height  := ScreenHeight;
  SynEditVB.Top             := StringGridModules.Top;
  SynEditVB.Height          := ScreenHeight div 2 - 16;
  SynEditVB.Width           := Width - 508;
  LabelPCode.Top            := SynEditVB.Top + ScreenHeight div 2 - 8;
  SynEditPCode.Top          := SynEditVB.Top + ScreenHeight div 2 + 16;
  SynEditPCode.Height       := ScreenHeight div 2 - 16;
  SynEditPCode.Width        := Width - 508;
end;

procedure TFormVBAView.ResetGrid();
var
  I: UInt32;
begin
  LabeledEditFileSpec.Text := '';
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
  I : Integer;
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
