unit UnitFormVBASettings;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  System.Generics.Collections, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.ComCtrls, SynEdit, Vcl.StdCtrls,
  SynEditHighlighter, SynHighlighterVB;

type
  TIDEHighlightElements = (
    AdditionalSearchMatchHighlight,
    &Assembler,
    AttributeNames,
    AttributeValues,
    BraceHighlight,
    Character,
    CodeFoldingTree,
    Comment,
    DiffAddition,
    DiffDeletion,
    DiffMove,
    DisabledBreak,
    EnabledBreak,
    ErrorLine,
    ExecutionPoint,
    Float,
    FoldedCode,
    Hex,
    HotLink,
    Identifier,
    IllegalChar,
    InvalidBreak,
    LineHighlight,
    LineNumber,
    MarkedBlock,
    ModifiedLine,
    Number,
    Octal,
    PlainText,
    Preprocessor,
    ReservedWord,
    RightMargin,
    Scripts,
    SearchMatch,
    &String,
    Symbol,
    SyncEditBackground,
    SyncEditHighlight,
    Tags,
    Whitespace
  );
  TIDEHighlightElementsAttributes = (
    Bold,
    Italic,
    Underline,
    DefaultForeground,
    DefaultBackground,
    ForegroundColorNew,
    BackgroundColorNew
  );
  TItemIDEHighlightElementsAttributes = record
    Bold               : Boolean;
    Italic             : Boolean;
    Underline          : Boolean;
    DefaultForeground  : Boolean;
    DefaultBackground  : Boolean;
    ForegroundColorNew : string;
    BackgroundColorNew : string;
  end;
  TIDETheme = array[TIDEHighlightElements] of TItemIDEHighlightElementsAttributes;

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
    procedure ButtonApplyClick(Sender: TObject);
    procedure ComboBoxFontChange(Sender: TObject);
    procedure ComboBoxThemeChange(Sender: TObject);
    procedure EditFontSizeChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    var IDETheme: TIDETheme;
    var MainSynEdit: TSynEdit;
    var MainVBSyn: TSynVBSyn;
    var Themes : TDictionary<string, string>;
    procedure LoadFixedSizeFonts();
    procedure LoadThemes();
  public
    { Public declarations }
    procedure SetReferences(SynEdit: TSynEdit; SynVBSyn: TSynVBSyn);
  end;

var
  FormVBASettings: TFormVBASettings;

implementation

uses System.StrUtils, System.TypInfo, Vcl.GraphUtil, Xml.XMLDoc, Xml.XMLDom, Xml.XMLIntf;

{$R *.dfm}

function EnumFontsProc(
  var LogFont: TLogFont;
  var TextMetric: TTextMetric;
  FontType: UInt32;
  Data: Pointer): Integer; stdcall;
begin
  if (LogFont.lfPitchAndFamily and FIXED_PITCH) <> 0 then
    if (FormVBASettings.ComboBoxFont.Items.IndexOf(LogFont.lfFaceName) < 0)
    and (not StartsText('@', LogFont.lfFaceName)) then
      FormVBASettings.ComboBoxFont.Items.Add(LogFont.lfFaceName);
  Result := 1;
end;

procedure SetSynAttr(IDETheme: TIDETheme; Element: TIDEHighlightElements; SynAttr: TSynHighlighterAttributes);
begin
  SynAttr.Background := StringToColor(IDETheme[Element].BackgroundColorNew);
  SynAttr.Foreground := StringToColor(IDETheme[Element].ForegroundColorNew);
  SynAttr.Style      := [];
  if IDETheme[Element].Bold then
    SynAttr.Style := SynAttr.Style + [fsBold];
  if IDETheme[Element].Italic then
    SynAttr.Style := SynAttr.Style + [fsItalic];
  if IDETheme[Element].Underline then
    SynAttr.Style := SynAttr.Style + [fsUnderline];
end;

procedure RefreshSynEdit(IDETheme: TIDETheme; SynEdit: TSynEdit);
var
  Element : TIDEHighlightElements;
begin
  Element := TIDEHighlightElements.RightMargin;
  SynEdit.RightEdgeColor := StringToColor(IDETheme[Element].ForegroundColorNew);
  Element := TIDEHighlightElements.MarkedBlock;
  SynEdit.SelectedColor.Foreground := StringToColor(IDETheme[Element].ForegroundColorNew);
  SynEdit.SelectedColor.Background := StringToColor(IDETheme[Element].BackgroundColorNew);
  Element := TIDEHighlightElements.LineNumber;
  SynEdit.Gutter.Color := StringToColor(IDETheme[Element].BackgroundColorNew);
  SynEdit.Gutter.Font.Color := StringToColor(IDETheme[Element].ForegroundColorNew);
  Element := TIDEHighlightElements.LineHighlight;
  SynEdit.ActiveLineColor := StringToColor(IDETheme[Element].BackgroundColorNew);
  Element := TIDEHighlightElements.PlainText;
  SynEdit.Gutter.BorderColor := GetHighLightColor(StringToColor(IDETheme[Element].BackgroundColorNew));
end;

procedure RefreshSynHighlighter(IDETheme: TIDETheme; SynEdit: TSynEdit; SynVBSyn: TSynVBSyn);
begin
  RefreshSynEdit(IDETheme, SynEdit);
  SetSynAttr(IDETheme, TIDEHighlightElements.Comment, SynVBSyn.CommentAttri);
  SetSynAttr(IDETheme, TIDEHighlightElements.Identifier, SynVBSyn.IdentifierAttri);
  SetSynAttr(IDETheme, TIDEHighlightElements.ReservedWord, SynVBSyn.KeyAttri);
  SetSynAttr(IDETheme, TIDEHighlightElements.Number, SynVBSyn.NumberAttri);
  SetSynAttr(IDETheme, TIDEHighlightElements.Whitespace, SynVBSyn.SpaceAttri);
  SetSynAttr(IDETheme, TIDEHighlightElements.String, SynVBSyn.StringAttri);
  SetSynAttr(IDETheme, TIDEHighlightElements.Symbol, SynVBSyn.SymbolAttri);
end;

procedure TFormVBASettings.ButtonApplyClick(Sender: TObject);
begin
  MainSynEdit.Font.Name := ComboBoxFont.Text;
  MainSynEdit.Font.Size := StrToInt(EditFontSize.Text);
  RefreshSynHighlighter(IDETheme, MainSynEdit, MainVBSyn);
end;

procedure TFormVBASettings.ComboBoxFontChange(Sender: TObject);
begin
  SynEditVB.Font.Name := ComboBoxFont.Text;
end;

procedure TFormVBASettings.ComboBoxThemeChange(Sender: TObject);
var
  XmlDocIDETheme : IXmlDocument;
  XPathElement   : string;
  XPathSelect    : IDomNodeSelect;
  Element        : TIDEHighlightElements;
  ElementName    : string;
begin
  XmlDocIDETheme := TXmlDocument.Create(Self);
  try
    XmlDocIDETheme.LoadFromXML(Themes.Items[ComboBoxTheme.Text]);
    for Element in [Low(TIDEHighlightElements)..High(TIDEHighlightElements)] do
    begin
      ElementName := GetEnumName(TypeInfo(TIDEHighlightElements), Integer(Element));
      XPathElement := Format('//DelphiIDETheme/%s/', [ElementName]);
      if Supports(XmlDocIDETheme.DocumentElement.DOMNode, IDomNodeSelect, XPathSelect) then
      begin
        IDETheme[Element].Bold      := CompareText(XPathSelect.selectNode(Format('%s%s', [XPathElement, 'Bold'])).firstChild.nodeValue, 'True') = 0;
        IDETheme[Element].Italic    := CompareText(XPathSelect.selectNode(Format('%s%s', [XPathElement, 'Italic'])).firstChild.nodeValue, 'True') = 0;
        IDETheme[Element].Underline := CompareText(XPathSelect.selectNode(Format('%s%s', [XPathElement, 'Underline'])).firstChild.nodeValue, 'True') = 0;
        IDETheme[Element].DefaultForeground := CompareText(XPathSelect.selectNode(Format('%s%s', [XPathElement, 'DefaultForeground'])).firstChild.nodeValue, 'True') = 0;
        IDETheme[Element].DefaultBackground := CompareText(XPathSelect.selectNode(Format('%s%s', [XPathElement, 'DefaultBackground'])).firstChild.nodeValue, 'True') = 0;
        IDETheme[Element].ForegroundColorNew := XPathSelect.selectNode(Format('%s%s', [XPathElement, 'ForegroundColorNew'])).firstChild.nodeValue;
        IDETheme[Element].BackgroundColorNew := XPathSelect.selectNode(Format('%s%s', [XPathElement, 'BackgroundColorNew'])).firstChild.nodeValue;
      end;
    end;
  finally
    ;
  end;
  RefreshSynHighlighter(IDETheme, SynEditVB, SynVBSyn);
end;

procedure TFormVBASettings.EditFontSizeChange(Sender: TObject);
begin
  SynEditVB.Font.Size := StrToInt(EditFontSize.Text);
end;

procedure TFormVBASettings.FormActivate(Sender: TObject);
begin
  ComboBoxFont.ItemIndex := ComboBoxFont.Items.IndexOf(SynEditVB.Font.Name);
  EditFontSize.Text := IntToStr(SynEditVB.Font.Size);
end;

procedure TFormVBASettings.FormCreate(Sender: TObject);
begin
  LoadThemes();
  LoadFixedSizeFonts();
end;

procedure TFormVBASettings.FormDestroy(Sender: TObject);
begin
  FreeAndNil(Themes);
end;

procedure TFormVBASettings.LoadFixedSizeFonts();
var
  TheDC   : HDC;
  LogFont : TLogFont;
begin
  ComboBoxFont.Items.Clear();
  TheDC := GetDC(0);
  try
    ZeroMemory(@LogFont, SizeOf(LogFont));
    LogFont.lfCharSet := DEFAULT_CHARSET;
    EnumFontFamiliesEx(TheDC, LogFont, @EnumFontsProc, 0, 0);
  finally
    ReleaseDC(0, TheDC);
  end;
end;

procedure TFormVBASettings.LoadThemes();
const
  ResourcePrefix = 'Resource_';
var
  ResourceIndex  : Int32;
  ResourceName   : string;
  ResourceReader : TStreamReader;
  ResourceStream : TResourceStream;
  ThemeList      : TStrings;
begin
  Themes := TDictionary<string, string>.Create();
  ThemeList := TStringList.Create();
  try
    try
      ResourceStream := TResourceStream.Create(HInstance, 'Resource_0', RT_RCDATA);
      ThemeList.LoadFromStream(ResourceStream);
    finally
      FreeAndNil(ResourceStream);
    end;
    for ResourceIndex := 0 to ThemeList.Count - 1 do
    begin
      ResourceName := ResourcePrefix + IntToStr(ResourceIndex + 1);
      ResourceStream := TResourceStream.Create(HInstance, ResourceName, RT_RCDATA);
      try
        ResourceReader := TStreamReader.Create(ResourceStream);
        try
          Themes.Add(
            ThemeList.Strings[ResourceIndex], ResourceReader.ReadToEnd()
          );
        finally
          FreeAndNil(ResourceStream);
        end;
      finally
        FreeAndNil(ResourceReader);
      end;
    end;
  finally
    FreeAndNil(ThemeList);
  end;
  ComboBoxTheme.Items.Clear();
  for ResourceIndex := 0 to Themes.Count - 1 do
    ComboBoxTheme.Items.Add(Themes.Keys.ToArray[ResourceIndex]);
end;

procedure TFormVBASettings.SetReferences(SynEdit: TSynEdit; SynVBSyn: TSynVBSyn);
begin
  MainSynEdit := SynEdit;
  MainVBSyn := SynVBSyn;
end;

end.
