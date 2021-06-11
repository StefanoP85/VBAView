unit UnitFormVBAProperties;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, ParserVBA, Vcl.StdCtrls, Common;

type
  TFormVBAProperties = class(TForm)
    LabelProperties: TLabel;
    LabelReferences: TLabel;
    StringGridProperties: TStringGrid;
    StringGridReferences: TStringGrid;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure SetReference(const VBAProgram: TVBAProgram);
  end;

var
  FormVBAProperties: TFormVBAProperties;

implementation

{$R *.dfm}

procedure TFormVBAProperties.FormCreate(Sender: TObject);
begin
  StringGridProperties.Cells[0,  0] := 'Property name';
  StringGridProperties.Cells[1,  0] := 'Property value';
  StringGridProperties.Cells[0,  1] := 'Project sys kind';
  StringGridProperties.Cells[0,  2] := 'Locale identifier';
  StringGridProperties.Cells[0,  3] := 'Locale identifier Invoke';
  StringGridProperties.Cells[0,  4] := 'Code page';
  StringGridProperties.Cells[0,  5] := 'Project name';
  StringGridProperties.Cells[0,  6] := 'Doc string';
  StringGridProperties.Cells[0,  7] := 'Help file 1';
  StringGridProperties.Cells[0,  8] := 'Help file 2';
  StringGridProperties.Cells[0,  9] := 'Help context';
  StringGridProperties.Cells[0, 10] := 'Project library flags';
  StringGridProperties.Cells[0, 11] := 'Major version';
  StringGridProperties.Cells[0, 12] := 'Minor version';  
  StringGridReferences.Cells[0,  0] := 'Name';
  StringGridReferences.Cells[1,  0] := 'Control';
  StringGridReferences.Cells[2,  0] := 'Original';
  StringGridReferences.Cells[3,  0] := 'Registered';
  StringGridReferences.Cells[4,  0] := 'Project';
end;

procedure TFormVBAProperties.SetReference(const VBAProgram: TVBAProgram);
var
  I: UInt32;
begin
  case VBAProgram.SysKind of
    0:
      StringGridProperties.Cells[1, 1] := '16-bit Windows platform';
    1:
      StringGridProperties.Cells[1, 1] := '32-bit Windows platform';
    2:
      StringGridProperties.Cells[1, 1] := 'Macintosh platform';
    3:  
      StringGridProperties.Cells[1, 1] := '64-bit Windows platform';
    else
      StringGridProperties.Cells[1, 1] := 'Unknown platform';
  end;
  StringGridProperties.Cells[1,  2] := IntToStr(VBAProgram.Lcid);
  StringGridProperties.Cells[1,  3] := IntToStr(VBAProgram.LcidInvoke);
  StringGridProperties.Cells[1,  4] := IntToStr(VBAProgram.CodePage);
  StringGridProperties.Cells[1,  5] := VBAProgram.ProjectName;
  StringGridProperties.Cells[1,  6] := VBAProgram.ProjectName;
  StringGridProperties.Cells[1,  7] := VBAProgram.HelpFile1;
  StringGridProperties.Cells[1,  8] := VBAProgram.HelpFile2;
  StringGridProperties.Cells[1,  9] := IntToStr(VBAProgram.HelpContext);
  StringGridProperties.Cells[1, 10] := IntToStr(VBAProgram.ProjectLibFlags);
  StringGridProperties.Cells[1, 11] := IntToStr(VBAProgram.VersionMajor);
  StringGridProperties.Cells[1, 12] := IntToStr(VBAProgram.VersionMinor);
  StringGridReferences.RowCount := High(VBAProgram.Reference);
  for I := 0 to High(VBAProgram.Reference) do
  begin
    StringGridReferences.Cells[0, I + 1] := VBAProgram.Reference[I].ReferenceName;
    StringGridReferences.Cells[1, I + 1] := VBAProgram.Reference[I].ReferenceControl;
    StringGridReferences.Cells[2, I + 1] := VBAProgram.Reference[I].ReferenceOriginal;
    StringGridReferences.Cells[3, I + 1] := VBAProgram.Reference[I].ReferenceRegistered;
    StringGridReferences.Cells[4, I + 1] := VBAProgram.Reference[I].ReferenceProject;
  end;
  Show();
end;

end.
