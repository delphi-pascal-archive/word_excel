object Form1: TForm1
  Left = 252
  Top = 126
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Work with Microsoft Word and Excel'
  ClientHeight = 545
  ClientWidth = 414
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Icon.Data = {
    0000010001002020100000000000E80200001600000028000000200000004000
    0000010004000000000080020000000000000000000000000000000000000000
    0000000080000080000000808000800000008000800080800000C0C0C0008080
    80000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00CCC0
    000CCCC0000000000CCCC7777CCCCCCC0000CCCC00000000CCCC7777CCCCCCCC
    C0000CCCCCCCCCCCCCC7777CCCCC0CCCCC0000CCCCCCCCCCCC7777CCCCC700CC
    C00CCCC0000000000CCCC77CCC77000C0000CCCC00000000CCCC7777C7770000
    00000CCCC000000CCCC777777777C000C00000CCCC0000CCCC77777C777CCC00
    CC00000CCCCCCCCCC77777CC77CCCCC0CCC000CCCCC00CCCCC777CCC7CCCCCCC
    CCCC0CCCCCCCCCCCCCC7CCCCCCCCCCCC0CCCCCCCCCCCCCCCCCCCCCC7CCC70CCC
    00CCCCCCCC0CC0CCCCCCCC77CC7700CC000CCCCCC000000CCCCCC777CC7700CC
    0000CCCC00000000CCCC7777CC7700CC0000C0CCC000000CCC7C7777CC7700CC
    0000C0CCC000000CCC7C7777CC7700CC0000CCCC00000000CCCC7777CC7700CC
    000CCCCCC000000CCCCCC777CC7700CC00CCCCCCCC0CC0CCCCCCCC77CC770CCC
    0CCCCCCCCCCCCCCCCCCCCCC7CCC7CCCCCCCC0CCCCCCCCCCCCCC7CCCCCCCCCCC0
    CCC000CCCCC00CCCCC777CCC7CCCCC00CC00000CCCCCCCCCC77777CC77CCC000
    C00000CCCC0000CCCC77777C777C000000000CCCC000000CCCC777777777000C
    0000CCCC00000000CCCC7777C77700CCC00CCCC0000000000CCCC77CCC770CCC
    CC0000CCCCCCCCCCCC7777CCCCC7CCCCC0000CCCCCCCCCCCCCC7777CCCCCCCCC
    0000CCCC00000000CCCC7777CCCCCCC0000CCCC0000000000CCCC7777CCC0000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    000000000000000000000000000000000000000000000000000000000000}
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  PixelsPerInch = 120
  TextHeight = 16
  object Bevel2: TBevel
    Left = 32
    Top = 184
    Width = 353
    Height = 9
    Shape = bsTopLine
  end
  object Image1: TImage
    Left = 8
    Top = 256
    Width = 65
    Height = 33
    Visible = False
  end
  object Bevel3: TBevel
    Left = 8
    Top = 352
    Width = 401
    Height = 9
    Shape = bsBottomLine
  end
  object Bevel4: TBevel
    Left = 40
    Top = 48
    Width = 353
    Height = 9
    Shape = bsTopLine
  end
  object Button1: TButton
    Left = 16
    Top = 56
    Width = 385
    Height = 25
    Caption = 'Open document (Word Application)'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 16
    Top = 88
    Width = 385
    Height = 25
    Caption = 'New doc + table + save doc (OleVariant)'
    TabOrder = 1
    OnClick = Button2Click
  end
  object Button4: TButton
    Left = 16
    Top = 120
    Width = 385
    Height = 25
    Caption = 'Search and Replace (X-Man to M_A_T_R_I_X)'
    TabOrder = 2
    OnClick = Button4Click
  end
  object Button5: TButton
    Left = 16
    Top = 152
    Width = 385
    Height = 25
    Caption = 'Page Setup (PageWidth, PageHeight, Orientation, Margin)'
    TabOrder = 3
    OnClick = Button5Click
  end
  object Button6: TButton
    Left = 16
    Top = 192
    Width = 249
    Height = 25
    Caption = 'Text (Range(a,b), between (a,b))'
    TabOrder = 4
    OnClick = Button6Click
  end
  object Button7: TButton
    Left = 16
    Top = 224
    Width = 385
    Height = 25
    Caption = 'Search and select word "Picture"'
    TabOrder = 5
    OnClick = Button7Click
  end
  object Button8: TButton
    Left = 16
    Top = 256
    Width = 185
    Height = 25
    Caption = 'Pictures 1'
    TabOrder = 6
    OnClick = Button8Click
  end
  object Button9: TButton
    Left = 16
    Top = 288
    Width = 385
    Height = 25
    Caption = 'Statistics of document'
    TabOrder = 7
    OnClick = Button9Click
  end
  object Button10: TButton
    Left = 16
    Top = 320
    Width = 385
    Height = 25
    Caption = 'Table (Rows and Columns)'
    TabOrder = 8
    OnClick = Button10Click
  end
  object Edit1: TEdit
    Left = 280
    Top = 192
    Width = 121
    Height = 24
    TabOrder = 9
    Text = 'text between (a,b)'
  end
  object Button14: TButton
    Left = 216
    Top = 256
    Width = 185
    Height = 25
    Caption = 'Pictures 2 (Frames)'
    TabOrder = 10
    OnClick = Button14Click
  end
  object Button11: TButton
    Left = 64
    Top = 16
    Width = 305
    Height = 25
    Caption = 'Launch Word'
    TabOrder = 11
    OnClick = Button11Click
  end
  object StringGrid1: TStringGrid
    Left = 16
    Top = 376
    Width = 385
    Height = 129
    RowCount = 3
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
    TabOrder = 12
    ColWidths = (
      64
      64
      64
      64
      64)
    RowHeights = (
      24
      24
      24)
  end
  object Button13: TButton
    Left = 16
    Top = 512
    Width = 185
    Height = 25
    Caption = 'To Excel'
    TabOrder = 13
    OnClick = Button13Click
  end
  object Button15: TButton
    Left = 216
    Top = 512
    Width = 185
    Height = 25
    Caption = 'From Excel'
    TabOrder = 14
    OnClick = Button15Click
  end
  object WordApplication1: TWordApplication
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    AutoQuit = False
    Left = 32
    Top = 64
  end
  object WordDocument1: TWordDocument
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    Left = 360
    Top = 64
  end
  object XLApp: TExcelApplication
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    AutoQuit = False
    Left = 184
    Top = 464
  end
end
