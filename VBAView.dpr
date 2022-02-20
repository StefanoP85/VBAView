program VBAView;

uses
  Vcl.Forms,
  UnitFormVBAMain in 'UnitFormVBAMain.pas' {FormVBAView},
  UnitFormVBAProperties in 'UnitFormVBAProperties.pas' {FormVBAProperties},
  ParserVBA in 'ParserVBA.pas',
  FilesCFB in 'FilesCFB.pas',
  ParserPCode in 'ParserPCode.pas',
  Common in 'Common.pas',
  UnitFormModuleSearch in 'UnitFormModuleSearch.pas' {FormModuleSearch},
  UnitFormVBACheck in 'UnitFormVBACheck.pas' {FormVBACheck};

{$R *.res}

begin
  System.ReportMemoryLeaksOnShutdown := True;
  Application.Initialize();
  Application.MainFormOnTaskbar := True;
  Application.Title := 'VBA Preview';
  Application.CreateForm(TFormVBAView, FormVBAView);
  Application.CreateForm(TFormVBAProperties, FormVBAProperties);
  Application.CreateForm(TFormModuleSearch, FormModuleSearch);
  Application.CreateForm(TFormVBACheck, FormVBACheck);
  Application.Run();
end.
