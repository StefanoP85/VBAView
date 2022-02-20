unit UnitFormModuleSearch;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  System.Actions, Vcl.ActnList, System.ImageList, Vcl.ImgList, Vcl.Buttons,
  Common;

type
  TFormModuleSearch = class(TForm)
    ActionList: TActionList;
    ActionSearch: TAction;
    GridPanelMainLayout: TGridPanel;
    ImageList: TImageList;
    LabeledEditSearchText: TLabeledEdit;
    MemoSearchResult: TMemo;
    SpeedButtonSearch: TSpeedButton;
    procedure ActionSearchExecute(Sender: TObject);
  private
    { Private declarations }
    var VBAProgram : TVBAProgram;
  public
    { Public declarations }
    procedure SetReference(const AVBAProgram: TVBAProgram);
  end;

var
  FormModuleSearch: TFormModuleSearch;

implementation

{$R *.dfm}

procedure TFormModuleSearch.ActionSearchExecute(Sender: TObject);
const
  NewLine = #13#10;
var
  LineNumber    : Int32;
  ModulePrinted : Boolean;
  ModuleIndex   : Int32;
  ModuleText    : string;
  ModuleLines   : TArray<string>;
  SearchProcess : TStrings;
  SearchText    : string;
begin
  SearchText := LabeledEditSearchText.Text;
  SearchText := SearchText.ToUpper();
  MemoSearchResult.Lines.Clear();
  SearchProcess := TStringList.Create();
  try
    for ModuleIndex := 0 to VBAProgram.ModulesCount - 1 do
    begin
      ModulePrinted := False;
      ModuleText := VBAProgram.Module[ModuleIndex].SourceCode.ToUpper();
      ModuleLines := ModuleText.Split([#13]);
      for LineNumber := Low(ModuleLines) to High(ModuleLines) - 1 do
        if ModuleLines[LineNumber].Contains(SearchText) then
        begin
          if ModulePrinted then
            SearchProcess.Append('Line ' + LineNumber.ToString())
          else
          begin
            ModulePrinted := True;
            SearchProcess.Append('Module ' + VBAProgram.Module[ModuleIndex].ModuleName
              + ' Line ' + LineNumber.ToString());
          end;
          SearchProcess.Append(ModuleLines[LineNumber]);
          SearchProcess.Append(NewLine);
        end;
    end;
    Finalize(ModuleLines);
    MemoSearchResult.Lines.SetStrings(SearchProcess);
  finally
    FreeAndNil(SearchProcess);
  end;
end;

procedure TFormModuleSearch.SetReference(const AVBAProgram: TVBAProgram);
begin
  VBAProgram := AVBAProgram;
  Show();
end;

end.
