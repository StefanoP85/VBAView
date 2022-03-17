unit UnitFormThemeSettings;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.Themes, Vcl.Styles, Vcl.Styles.Ext, Vcl.Buttons;

type
  TFormThemeSettings = class(TForm)
    CheckBoxDisableVClStylesNC: TCheckBox;
    ComboBoxVCLStyle: TComboBox;
    ImageVCLStyle: TImage;
    LabelVCLStyle: TLabel;
    SpeedButtonApply: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure ComboBoxVCLStyleChange(Sender: TObject);
    procedure SpeedButtonApplyClick(Sender: TObject);
  private
    { Private declarations }
    procedure DrawSeletedVCLStyle();
    procedure LoadThemes();
  public
    { Public declarations }
  end;

var
  FormThemeSettings: TFormThemeSettings;

implementation

{$R *.dfm}

uses Vcl.Imaging.PngImage, PngFunctions;

procedure TFormThemeSettings.ComboBoxVCLStyleChange(Sender: TObject);
begin
  DrawSeletedVCLStyle();
end;

procedure TFormThemeSettings.DrawSeletedVCLStyle();
var
  StyleName : string;
  LBitmap   : TBitmap;
  LStyle    : TCustomStyleExt;
  SourceInfo: TSourceInfo;
  LPng      : TPngImage;
begin
   ImageVCLStyle.Picture := nil;
   StyleName := ComboBoxVCLStyle.Text;
   if (StyleName <> '') and (CompareText('Windows', StyleName) <> 0) then
   begin
    LBitmap := TBitmap.Create();
    try
       LBitmap.PixelFormat := TPixelFormat.pf32bit;
       LBitmap.Width := ImageVCLStyle.ClientRect.Width;
       LBitmap.Height := ImageVCLStyle.ClientRect.Height;
       SourceInfo := TStyleManager.StyleSourceInfo[StyleName];
       LStyle := TCustomStyleExt.Create(TStream(SourceInfo.Data));
       try
         DrawSampleWindow(LStyle, LBitmap.Canvas, ImageVCLStyle.ClientRect, StyleName);
         ConvertToPNG(LBitmap, LPng);
         try
           ImageVCLStyle.Picture.Assign(LPng);
           //LPng.SaveToFile(ChangeFileExt(ParamStr(0),'.png'));
         finally
           LPng.Free();
         end;
         //ImageVCLStyle.Picture.Assign(LBitmap);
       finally
         LStyle.Free();
       end;
    finally
      LBitmap.Free();
    end;
   end;
end;

procedure TFormThemeSettings.FormCreate(Sender: TObject);
begin
  LoadThemes();
end;

procedure TFormThemeSettings.LoadThemes();
var
  Style : string;
begin
  try
    ComboBoxVCLStyle.Items.BeginUpdate();
    ComboBoxVCLStyle.Items.Clear();
    // ComboBoxVCLStyle.Items.Add('Windows');
    for Style in TStyleManager.StyleNames do
      ComboBoxVCLStyle.Items.Add(Style);
  finally
    ComboBoxVCLStyle.Items.EndUpdate();
  end;
end;

procedure TFormThemeSettings.SpeedButtonApplyClick(Sender: TObject);
var
  StyleName : string;
begin
  StyleName := ComboBoxVCLStyle.Text;
  if StyleName <> '' then
    TStyleManager.SetStyle(StyleName)
  else
    TStyleManager.SetStyle(TStyleManager.SystemStyle.Name);
end;

end.
