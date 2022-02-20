unit UnitFormVBACheck;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  Common;

type
  TFormVBACheck = class(TForm)
    GridPanelMainLayout: TGridPanel;
    MemoSearchResult: TMemo;
  private
    { Private declarations }
    var VBAProgram : TVBAProgram;
  public
    { Public declarations }
    procedure SetReference(const AVBAProgram: TVBAProgram);
  end;

var
  FormVBACheck: TFormVBACheck;

implementation

{$R *.dfm}

{ TFormVBACheck }

procedure TFormVBACheck.SetReference(const AVBAProgram: TVBAProgram);
const
  NewLine   = #13#10;
  AutoNames : array[0..4] of string = ('AutoExec', 'AutoNew', 'AutoOpen', 'AutoClose', 'AutoExit');
  Events    : array[0..2] of string = ('_Calculate', '_Change', '_Open');
var
  ModuleIndex   : Int32;
  ModuleText    : string;
  SearchProcess : TStrings;
  I             : Integer;
begin
  VBAProgram := AVBAProgram;
  MemoSearchResult.Lines.Clear();
  SearchProcess := TStringList.Create();
  try
    for ModuleIndex := 0 to VBAProgram.ModulesCount - 1 do
    begin
      ModuleText := VBAProgram.Module[ModuleIndex].SourceCode.ToUpper();
      for I := 0 to 4 do
      begin
        if ModuleText.Contains(AutoNames[I].ToUpper()) then
        begin
          SearchProcess.Append('Module ' + VBAProgram.Module[ModuleIndex].ModuleName);
          SearchProcess.Append('Contains an autorun identifier: ' + AutoNames[I]);
          SearchProcess.Append(NewLine);
        end;
        if ( VBAProgram.Module[ModuleIndex].ModuleName = AutoNames[I].ToUpper() )
        and ( ModuleText.Contains('MAIN') ) then
        begin
          SearchProcess.Append('Module ' + VBAProgram.Module[ModuleIndex].ModuleName);
          SearchProcess.Append('Has the name of an autorun identifier: ' + AutoNames[I]);
          SearchProcess.Append('Contains the autorun identifier for module: Main');
          SearchProcess.Append(NewLine);
        end;
      end;
      for I := 0 to 2 do
      begin
        if ModuleText.Contains(Events[I].ToUpper()) then
        begin
          SearchProcess.Append('Module ' + VBAProgram.Module[ModuleIndex].ModuleName);
          SearchProcess.Append('Contains an autorun event: ' + Events[I]);
          SearchProcess.Append(NewLine);
        end;
      end;
    end;
    MemoSearchResult.Lines.SetStrings(SearchProcess);
  finally
    FreeAndNil(SearchProcess);
  end;
  Show();
end;

end.
