object FormVBACheck: TFormVBACheck
  Left = 0
  Top = 0
  Caption = 'Check for autoruns'
  ClientHeight = 441
  ClientWidth = 624
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object GridPanelMainLayout: TGridPanel
    Left = 0
    Top = 0
    Width = 624
    Height = 441
    Align = alClient
    ColumnCollection = <
      item
        Value = 100.000000000000000000
      end>
    ControlCollection = <
      item
        Column = 0
        Control = MemoSearchResult
        Row = 0
      end>
    RowCollection = <
      item
        Value = 100.000000000000000000
      end>
    TabOrder = 0
    ExplicitLeft = -1
    ExplicitTop = 32
    ExplicitWidth = 625
    ExplicitHeight = 409
    object MemoSearchResult: TMemo
      Left = 1
      Top = 1
      Width = 622
      Height = 439
      Align = alClient
      ReadOnly = True
      ScrollBars = ssBoth
      TabOrder = 0
      ExplicitWidth = 623
      ExplicitHeight = 407
    end
  end
end
