program VBAView;

uses
  Vcl.Forms,
  UnitFormVBAMain in 'UnitFormVBAMain.pas' {FormVBAView},
  UnitFormVBAProperties in 'UnitFormVBAProperties.pas' {FormVBAProperties},
  ParserVBA in 'ParserVBA.pas',
  FilesCFB in 'FilesCFB.pas',
  ParserPCode in 'ParserPCode.pas',
  Common in 'Common.pas';

{$R *.res}

begin
  System.ReportMemoryLeaksOnShutdown := True;
  Application.Initialize();
  Application.MainFormOnTaskbar := True;
  Application.Title := 'VBA Preview';
  Application.CreateForm(TFormVBAView, FormVBAView);
  Application.CreateForm(TFormVBAProperties, FormVBAProperties);
  Application.Run();
end.
