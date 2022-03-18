program VBAView;

{$R *.dres}

uses
  Vcl.Forms,
  UnitFormVBAMain in 'UnitFormVBAMain.pas' {FormVBAView},
  UnitFormVBAProperties in 'UnitFormVBAProperties.pas' {FormVBAProperties},
  UnitFormModuleSearch in 'UnitFormModuleSearch.pas' {FormModuleSearch},
  UnitFormVBACheck in 'UnitFormVBACheck.pas' {FormVBACheck},
  UnitFormThemeSettings in 'UnitFormThemeSettings.pas' {FormThemeSettings},
  UnitFormVBASettings in 'UnitFormVBASettings.pas' {FormVBASettings},
  UnitFormAbout in 'UnitFormAbout.pas' {FormAbout},
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
  Application.CreateForm(TFormModuleSearch, FormModuleSearch);
  Application.CreateForm(TFormVBACheck, FormVBACheck);
  Application.CreateForm(TFormThemeSettings, FormThemeSettings);
  Application.CreateForm(TFormAbout, FormAbout);
  Application.CreateForm(TFormVBASettings, FormVBASettings);
  Application.Run();
end.
