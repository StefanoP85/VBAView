object FormVBAProperties: TFormVBAProperties
  Left = 0
  Top = 0
  Caption = 'VBA Project properties'
  ClientHeight = 441
  ClientWidth = 624
  Color = clBtnFace
  Constraints.MinHeight = 480
  Constraints.MinWidth = 640
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object LabelProperties: TLabel
    Left = 8
    Top = 8
    Width = 108
    Height = 13
    Caption = 'VBA project properties'
    FocusControl = StringGridProperties
  end
  object LabelReferences: TLabel
    Left = 8
    Top = 224
    Width = 89
    Height = 13
    Caption = 'Project references'
    FocusControl = StringGridReferences
  end
  object StringGridProperties: TStringGrid
    Left = 8
    Top = 27
    Width = 608
    Height = 190
    ColCount = 2
    DefaultColWidth = 256
    RowCount = 13
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    TabOrder = 0
  end
  object StringGridReferences: TStringGrid
    Left = 8
    Top = 243
    Width = 608
    Height = 190
    DefaultColWidth = 256
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    TabOrder = 1
  end
end
