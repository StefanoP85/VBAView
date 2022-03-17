object FormThemeSettings: TFormThemeSettings
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'Theme settings'
  ClientHeight = 354
  ClientWidth = 433
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object LabelVCLStyle: TLabel
    Left = 5
    Top = 41
    Width = 45
    Height = 13
    Caption = 'VCL Style'
  end
  object ImageVCLStyle: TImage
    Left = 5
    Top = 87
    Width = 422
    Height = 258
    Transparent = True
  end
  object SpeedButtonApply: TSpeedButton
    Left = 312
    Top = 54
    Width = 113
    Height = 27
    Caption = 'Apply'
    OnClick = SpeedButtonApplyClick
  end
  object ComboBoxVCLStyle: TComboBox
    Left = 5
    Top = 60
    Width = 301
    Height = 21
    Style = csDropDownList
    TabOrder = 1
    OnChange = ComboBoxVCLStyleChange
  end
  object CheckBoxDisableVClStylesNC: TCheckBox
    Left = 5
    Top = 10
    Width = 396
    Height = 17
    Caption = 
      'Disable VCL Styles in Non client Area (Only valid when Vcl Style' +
      's are activated)'
    TabOrder = 0
  end
end
