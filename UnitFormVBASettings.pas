unit UnitFormVBASettings;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, SynEdit, Vcl.StdCtrls,
  SynEditHighlighter, SynHighlighterVB;

type
  TFormVBASettings = class(TForm)
    ButtonApply: TButton;
    ComboBoxFont: TComboBox;
    ComboBoxTheme: TComboBox;
    EditFontSize: TEdit;
    LabelFont: TLabel;
    LabelSize: TLabel;
    LabelSynEdit: TLabel;
    SynEditVB: TSynEdit;
    SynVBSyn: TSynVBSyn;
    UpDownSize: TUpDown;
    procedure EditFontSizeChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormVBASettings: TFormVBASettings;

implementation

{$R *.dfm}

procedure TFormVBASettings.EditFontSizeChange(Sender: TObject);
begin
  SynEditVB.Font.Size := StrToInt(EditFontSize.Text);
end;

end.
