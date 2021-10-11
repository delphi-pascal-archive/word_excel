unit Wdof;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Word97, OleServer, StdCtrls, Word2000, ExtCtrls, ComObj, DdeMan, Clipbrd,
  WordXP, ExcelXP, Grids, Excel2000, Variants, JPEG;

type
  TForm1 = class(TForm)
    Button1: TButton;
    WordApplication1: TWordApplication;
    WordDocument1: TWordDocument;
    Button2: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Button9: TButton;
    Button10: TButton;
    Bevel2: TBevel;
    Edit1: TEdit;
    Image1: TImage;
    Button14: TButton;
    Bevel3: TBevel;
    Bevel4: TBevel;
    Button11: TButton;
    StringGrid1: TStringGrid;
    Button13: TButton;
    Button15: TButton;
    XLApp: TExcelApplication;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button14Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure Button15Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.Button1Click(Sender: TObject);
var
 FileName,ConfirmConversions,ReadOnly,AddToRecentFiles,
  PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam: OleVariant;
begin
 WordApplication1.Connect;
 WordApplication1.Visible:=true;
 // открываем шаблон документа
 FileName:=ExtractFilePath(Application.ExeName)+'3.doc';
 ConfirmConversions:=False;
 ReadOnly:=False;
 AddToRecentFiles:=False;
 PasswordDocument:='';
 PasswordTemplate:='';
 Revert:=False;
 WritePasswordDocument:='';
 WritePasswordTemplate:='';
 Format:=0;
 WordApplication1.Documents.Open(FileName,ConfirmConversions,ReadOnly,
  AddToRecentFiles,PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 // связываем компонент с существующим интерфейсом
 WordDocument1.ConnectKind:=ckAttachToInterface;
 WordDocument1.ConnectTo(WordApplication1.ActiveDocument);
 // сохранение и выход
 // WordDocument1.SaveAs(FileName);
 // WordDocument1.Close;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 WordDocument1.Disconnect;
 WordApplication1.Disconnect; 
 Action:=caFree;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
 MsWord,EmptyParam: OleVariant;
 i: integer;
begin
 try
  MsWord:=GetActiveOleObject('Word.Application.10');
 except
  try
   MsWord:=CreateOleObject('Word.Application.10');
   MsWord.Visible:=true;
  except
   Exception.Create('Error');
  end;
 end;
 MsWord.Documents.Add;
 MsWord.Selection.Font.Size:=12;
 MsWord.Selection.TypeText('Текст:');
 MsWord.Selection.Font.Bold:=true;
 MsWord.Selection.TypeText(#13#10'New string...');
 MsWord.ActiveDocument.Tables.Add(Range:=MsWord.Selection.Range,NumRows:=5,NumColumns:=3);
 for i:=0 to 3 do
  begin
   MsWord.ActiveDocument.Bookmarks.Add(Range:=MsWord.Selection.Range, Name:='klop');
   MsWord.ActiveDocument.Bookmarks.DefaultSorting:=0;
   MsWord.ActiveDocument.Bookmarks.ShowHidden:=false;
   if i<3
   then MsWord.Selection.MoveRight(Unit:=12);
  end;
 for i:=0 to 3 do
  begin
   MsWord.Selection.Goto(What:=-1, Name:='klop');
   MsWord.Selection.TypeText('_X_X_X_');
  end;
 MsWord.ActiveDocument.SaveAs(ExtractFilePath(Application.ExeName)+'sample.doc');
end;

procedure TForm1.Button4Click(Sender: TObject);
var
 FileName,ConfirmConversions,ReadOnly,AddToRecentFiles,
  PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam: OleVariant;
 oldStr,newStr,replace: OleVariant;
begin
 WordApplication1.Connect;
 WordApplication1.Visible:=true;
 // открываем шаблон документа
 FileName:=ExtractFilePath(Application.ExeName)+'3.doc';
 ConfirmConversions:=False;
 ReadOnly:=False;
 AddToRecentFiles:=False;
 PasswordDocument:='';
 PasswordTemplate:='';
 Revert:=False;
 WritePasswordDocument:='';
 WritePasswordTemplate:='';
 Format:=0;
 WordApplication1.Documents.Open(FileName,ConfirmConversions,ReadOnly,
  AddToRecentFiles,PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 // связываем компонент с существующим интерфейсом
 WordDocument1.ConnectKind:=ckAttachToInterface;
 WordDocument1.ConnectTo(WordApplication1.ActiveDocument);
 oldStr:='X-Man';
 newStr:='M_A_T_R_I_X';
 replace:=1;
 // поиск X-Man и замена на M_A_T_R_I_X
 WordDocument1.Range.Find.Execute(oldStr,EmptyParam,EmptyParam,EmptyParam,
 EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,newStr,replace,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 WordApplication1.Visible:=true;
 // выход wdSaveChanges, wdDoNotSaveChanges, wdPromptToSaveChanges
 // vsave:=wdDoNotSaveChanges;
 // WordDocument1.Close(vsave);
 // или
 // WordDocument1.Save;
 // WordDocument1.Close;
end;

procedure TForm1.Button5Click(Sender: TObject);
var
 FileName,ConfirmConversions,ReadOnly,AddToRecentFiles,
  PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam: OleVariant;
begin
 WordApplication1.Connect;
 WordApplication1.Visible:=true;
 // открываем шаблон документа
 FileName:=ExtractFilePath(Application.ExeName)+'3.doc';
 ConfirmConversions:=False;
 ReadOnly:=False;
 AddToRecentFiles:=False;
 PasswordDocument:='';
 PasswordTemplate:='';
 Revert:=False;
 WritePasswordDocument:='';
 WritePasswordTemplate:='';
 Format:=0;
 WordApplication1.Documents.Open(FileName,ConfirmConversions,ReadOnly,
  AddToRecentFiles,PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 // связываем компонент с существующим интерфейсом
 WordDocument1.ConnectKind:=ckAttachToInterface;
 WordDocument1.ConnectTo(WordApplication1.ActiveDocument);
 // Page Setup
 WordApplication1.ActiveDocument.PageSetup.PageWidth:=WordApplication1.CentimetersToPoints(10);
 WordApplication1.ActiveDocument.PageSetup.PageHeight:=WordApplication1.CentimetersToPoints(10);
 WordApplication1.ActiveDocument.PageSetup.Orientation:=1;
 WordApplication1.ActiveDocument.PageSetup.TopMargin:=WordApplication1.CentimetersToPoints(2);
 WordApplication1.ActiveDocument.PageSetup.BottomMargin:=WordApplication1.CentimetersToPoints(2);
 WordApplication1.ActiveDocument.PageSetup.LeftMargin:=WordApplication1.CentimetersToPoints(2.5);
 WordApplication1.ActiveDocument.PageSetup.RightMargin:=WordApplication1.CentimetersToPoints(2);
 // орфография
 WordApplication1.Options.CheckSpellingAsYouType:=False;
 WordApplication1.Options.CheckGrammarAsYouType:=False;
 // сохранение или распечатка документа
 //WordDocument1.PrintOut;
 //WordDocument1.SaveAs(FileName);
 WordApplication1.Visible:=true;
end;

procedure TForm1.Button6Click(Sender: TObject);
var
 FileName,ConfirmConversions,ReadOnly,AddToRecentFiles,
  PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam: OleVariant;
 Range1,Range2,Range3,a,b: OleVariant;
begin
 WordApplication1.Connect;
 WordApplication1.Visible:=true;
 // открываем шаблон документа
 FileName:=ExtractFilePath(Application.ExeName)+'opit.doc';
 ConfirmConversions:=False;
 ReadOnly:=False;
 AddToRecentFiles:=False;
 PasswordDocument:='';
 PasswordTemplate:='';
 Revert:=False;
 WritePasswordDocument:='';
 WritePasswordTemplate:='';
 Format:=0;
 WordApplication1.Documents.Open(FileName,ConfirmConversions,ReadOnly,
  AddToRecentFiles,PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 // связываем компонент с существующим интерфейсом
 WordDocument1.ConnectKind:=ckAttachToInterface;
 WordDocument1.ConnectTo(WordApplication1.ActiveDocument);
 WordApplication1.Visible:=true;
 // вставка, шрифт
 Range1:=WordDocument1.Range;
 a:=15; // непонятно
 b:=35; // непонятно
 Range2:=WordDocument1.Range(a,b);
 Range3:=WordDocument1.Range(a);
 a:=25; // непонятно
 b:=35; // непонятно
 WordDocument1.Range(a,b).Font.Bold:=1;
 WordDocument1.Range(a,b).Font.Size:=14;
 WordDocument1.Range(a,b).Font.Color:=clRed;
 // вставка после/перед (перед - InsertBefore('-- INSERTED --'));
 Range2.InsertAfter('-- INSERTED --'); // после
 // текст, который содержится между (a,b)
 Edit1.Text:=WordDocument1.Range(a,b).Text;
 // команды
 {
 WordDocument1.Range(a,b).Select;
 WordDocument1.Range(a,b).Cut;
 WordDocument1.Range(a,b).Copy;
 WordDocument1.Range(a,b).Paste;
 }
end;

procedure TForm1.Button7Click(Sender: TObject);
var
 FileName,ConfirmConversions,ReadOnly,AddToRecentFiles,
  PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam: OleVariant;
 a,b,vstart,vend: OleVariant;
 j,ilengy: integer;
begin
 WordApplication1.Connect;
 WordApplication1.Visible:=true;
 // открываем шаблон документа
 FileName:=ExtractFilePath(Application.ExeName)+'opit.doc';
 ConfirmConversions:=False;
 ReadOnly:=False;
 AddToRecentFiles:=False;
 PasswordDocument:='';
 PasswordTemplate:='';
 Revert:=False;
 WritePasswordDocument:='';
 WritePasswordTemplate:='';
 Format:=0;
 WordApplication1.Documents.Open(FileName,ConfirmConversions,ReadOnly,
  AddToRecentFiles,PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 // связываем компонент с существующим интерфейсом
 WordDocument1.ConnectKind:=ckAttachToInterface;
 WordDocument1.ConnectTo(WordApplication1.ActiveDocument);
 WordApplication1.Visible:=true;
 // поиск и выделение
 ilengy:=Length(WordDocument1.Range.Text);
 for j:=0 to ilengy-8 do
  begin
   a:=j;
   b:=j+7;
   if WordDocument1.Range(a,b).Text='Picture'
   then
    begin
     vstart:=j;
     vend:=j+7;
    end;
  end;
 WordDocument1.Range(vstart,vend).Select;
{
 // вставка перед/после выделенного текста
 WordApplication1.Selection.InsertAfter("text1");
 WordApplication1.Selection.InsertBefore("text2");
 // Форматирование выделенного текста
 WordApplication1.Selection.Font.Bold:=1;
 WordApplication1.Selection.Font.Size:=16;
 WordApplication1.Selection.Font.Color:=clGreen;
 // Для выравнивания проще воспользоваться компонентом WordParagraphFormat.
 // Сначала только нужно "подключить" его к выделенному фрагменту текста:
 WordParagraphFormat1.ConnectTo(WordApplication1.Selection.ParagraphFormat);
 WordParagraphFormat1.Alignment:=wdAlignParagraphCenter;
 // значения его свойства Alignment может принимать значения
 // wdAlignParagraphCenter, wdAlignParagraphLeft, wdAlignParagraphRight
 // Имеются и методы Cut, Copy и Paste
 WordApplication1.Selection.Cut;
 WordApplication1.Selection.Copy;
 WordApplication1.Selection.Paste;
 // убираем выделение с помощью метода Collapse. При этом необходимо указать, в какую сторону сместится курсор, будет ли он до ранее выделенного фрагмента или после:
 var vcol: OleVariant;
 ...
 vcol:=wdCollapseStart;
 WordApplication1.Selection.Collapse(vcol);
 // при этом выделение пропадет, а курсор займет позицию перед
 // фрагментом текста. Если присвоить переменной значение wdCollapseEnd,
 // то курсор переместится назад. Можно просто поставить в скобках "пустышку":
 WordApplication1.Selection.Collapse(EmptyParam);
 // Тогда свертывание выделения производится по умолчанию,
 // к началу выделенного текста.
}
end;

procedure TForm1.Button8Click(Sender: TObject);
var
 FileName,ConfirmConversions,ReadOnly,AddToRecentFiles,
  PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam: OleVariant;
 a,b,vstart,vend: OleVariant;
 j,ilengy: integer;
begin
 Image1.Picture.LoadFromFile('AtExpl.jpg');
 ClipBoard.Assign(Image1.Picture);
//////////
 WordApplication1.Connect;
 WordApplication1.Visible:=true;
 // открываем шаблон документа
 FileName:=ExtractFilePath(Application.ExeName)+'opit.doc';
 ConfirmConversions:=False;
 ReadOnly:=False;
 AddToRecentFiles:=False;
 PasswordDocument:='';
 PasswordTemplate:='';
 Revert:=False;
 WritePasswordDocument:='';
 WritePasswordTemplate:='';
 Format:=0;
 WordApplication1.Documents.Open(FileName,ConfirmConversions,ReadOnly,
  AddToRecentFiles,PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 // связываем компонент с существующим интерфейсом
 WordDocument1.ConnectKind:=ckAttachToInterface;
 WordDocument1.ConnectTo(WordApplication1.ActiveDocument);
 WordApplication1.Visible:=true;
 // поиск и выделение
 ilengy:=Length(WordDocument1.Range.Text);
 for j:=0 to ilengy-8 do
  begin
   a:=j;
   b:=j+7;
   if WordDocument1.Range(a,b).Text='Picture'
   then
    begin
     vstart:=j;
     vend:=j+7;
    end;
  end;
 WordDocument1.Range(vstart,vend).Select;
 // вставка рисунка (или WordDocument1.Range(a,b).Paste)
 WordApplication1.Selection.Paste;
end;

procedure TForm1.Button14Click(Sender: TObject);
var
 FileName,ConfirmConversions,ReadOnly,AddToRecentFiles,
  PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam: OleVariant;
 vstart,vend: OleVariant;
begin
 Image1.Picture.LoadFromFile('AtExpl.jpg');
 ClipBoard.Assign(Image1.Picture);
//////////
 WordApplication1.Connect;
 WordApplication1.Visible:=true;
 // открываем шаблон документа
 FileName:=ExtractFilePath(Application.ExeName)+'opit.doc';
 ConfirmConversions:=False;
 ReadOnly:=False;
 AddToRecentFiles:=False;
 PasswordDocument:='';
 PasswordTemplate:='';
 Revert:=False;
 WritePasswordDocument:='';
 WritePasswordTemplate:='';
 Format:=0;
 WordApplication1.Documents.Open(FileName,ConfirmConversions,ReadOnly,
  AddToRecentFiles,PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 // связываем компонент с существующим интерфейсом
 WordDocument1.ConnectKind:=ckAttachToInterface;
 WordDocument1.ConnectTo(WordApplication1.ActiveDocument);
 WordApplication1.Visible:=true;
 // вставка кадра
 vstart:=1;
 vend:=2;
 WordDocument1.Frames.Add(WordDocument1.Range(vstart,vend));
 WordDocument1.Frames.Item(1).Height:=Image1.Height;
 WordDocument1.Frames.Item(1).Width:=Image1.Width;
 WordDocument1.Frames.Item(1).Select;
 WordApplication1.Selection.Paste;
 // положение
 WordDocument1.Frames.Item(1).VerticalPosition:=30;
 WordDocument1.Frames.Item(1).HorizontalPosition:=50;
 // отступ между краями рамки и текстом
 WordDocument1.Frames.Item(1).HorizontalDistanceFromText:=10;
 WordDocument1.Frames.Item(1).VerticalDistanceFromText:=10;
 // масштабирование
 WordDocument1.Frames.Item(1).Height:=Image1.Height*4;
 WordDocument1.Frames.Item(1).Width:=Image1.Width*2;
 // удаление рамки
 // WordDocument1.Frames.Item(1).Delete;
end;

procedure TForm1.Button9Click(Sender: TObject);
var
 FileName,ConfirmConversions,ReadOnly,AddToRecentFiles,
  PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam: OleVariant;
 WCount, CCount, SCount, PCount: integer;
begin
 WordApplication1.Connect;
 WordApplication1.Visible:=true;
 // открываем шаблон документа
 FileName:=ExtractFilePath(Application.ExeName)+'opit.doc';
 ConfirmConversions:=False;
 ReadOnly:=False;
 AddToRecentFiles:=False;
 PasswordDocument:='';
 PasswordTemplate:='';
 Revert:=False;
 WritePasswordDocument:='';
 WritePasswordTemplate:='';
 Format:=0;
 WordApplication1.Documents.Open(FileName,ConfirmConversions,ReadOnly,
  AddToRecentFiles,PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 // связываем компонент с существующим интерфейсом
 WordDocument1.ConnectKind:=ckAttachToInterface;
 WordDocument1.ConnectTo(WordApplication1.ActiveDocument);
 WordApplication1.Visible:=true;
 // статистика
{
 $00000000  wdStatisticWords Количество слов
 $00000001  wdStatisticLines Количество строк
 $00000002  wdStatisticPages Количество страниц
 $00000003  wdStatisticCharacters Знаки без пробелов
 $00000004  wdStatisticParagraphs Количество разделов
 $00000005  wdStatisticCharactersWithSpaces Знаки с пробелами
}
 WCount:=WordDocument1.ComputeStatistics($00000000);
 CCount:=WordDocument1.ComputeStatistics($00000003);
 SCount:=WordDocument1.ComputeStatistics($00000005);
 PCount:=WordDocument1.ComputeStatistics($00000002);
 Button9.Caption:='Words: '+IntToStr(WCount)+', Znaki: '+IntToStr(CCount)+', Znaki with _: '+IntToStr(SCount)+', Pages: '+IntToStr(PCount);
end;

procedure TForm1.Button10Click(Sender: TObject);
var
 FileName,ConfirmConversions,ReadOnly,AddToRecentFiles,
  PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam: OleVariant;
 tcount,i,j: integer;
 a,b: OleVariant;
begin
 WordApplication1.Connect;
 WordApplication1.Visible:=true;
 // открываем шаблон документа
 FileName:=ExtractFilePath(Application.ExeName)+'tables.doc';
 ConfirmConversions:=False;
 ReadOnly:=False;
 AddToRecentFiles:=False;
 PasswordDocument:='';
 PasswordTemplate:='';
 Revert:=False;
 WritePasswordDocument:='';
 WritePasswordTemplate:='';
 Format:=0;
 WordApplication1.Documents.Open(FileName,ConfirmConversions,ReadOnly,
  AddToRecentFiles,PasswordDocument,PasswordTemplate,Revert,WritePasswordDocument,
   WritePasswordTemplate,Format,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 // связываем компонент с существующим интерфейсом
 WordDocument1.ConnectKind:=ckAttachToInterface;
 WordDocument1.ConnectTo(WordApplication1.ActiveDocument);
 WordApplication1.Visible:=true;
 // статистика кол-ва таблиц
 // tcount:=WordDocument1.Tables.Count;
 // к отдельной таблице обращаемся по ее номеру
 // WordDocument1.Tables.Item(i)..., // где i - целое число.
 // вместо всего того, что на странице - таблица
 i:=10; // кол-во строк
 j:=6; // кол-во столбцов
 //WordDocument1.Tables.Add(WordDocument1.Range, i, j, EmptyParam, EmptyParam);
 // вместо диапазона (a,b) - таблица
 a:=290;
 b:=292;
 WordDocument1.Tables.Add(WordDocument1.Range(a,b), i, j, EmptyParam, EmptyParam);
 // кол-во строк и столбцов первой таблицы
 i:=WordDocument1.Tables.Item(1).Columns.Count;
 j:=WordDocument1.Tables.Item(1).Rows.Count;
 Button10.Caption:=Button10.Caption+' - Rows: '+IntToStr(i)+', Columns: '+IntToStr(j);
{{
 // ширина столбцов или высота строк
 WordDocument1.Tables.Item(i).Columns.Width:=90;
 WordDocument1.Tables.Item(i).Rows.Height:=45;
 // размеры отдельных строк и столбцов
 WordDocument1.Tables.Item(i).Columns.Item(j).Width:=90;
 WordDocument1.Tables.Item(i).Rows.Item(j).Height:=45;
}
 // вставка текста в ячейку (номерстроки, номер столбца)
 WordDocument1.Tables.Item(1).Cell(2,2).Range.Text:='MaTriX';
{
 // отступы от края ячеек, как для всей таблицы сразу,
 // так и для отдельной ячейки
 WordDocument1.Tables.Item(i).TopPadding:=10;
 WordDocument1.Tables.Item(i).BottomPadding:=10;
 WordDocument1.Tables.Item(i).RightPadding:=10;
 WordDocument1.Tables.Item(i).LeftPadding:=10;
}
{
 // выделить нужную ячейку, столбец или строку
 WordDocument1.Tables.Item(i).Cell(j,k).Select;
 WordDocument1.Tables.Item(i).Columns.Item(j).Select;
 WordDocument1.Tables.Item(i).Rows.Item(j).Select;
 // подогон размера ячеек по содержимому
 WordDocument1.Tables.Item(i).Columns.AutoFit;
 // добавление строк и столбцов
 WordDocument1.Tables.Item(1).Rows.Add(EmptyParam);
 WordDocument1.Tables.Item(1).Columns.Add(EmptyParam);
}
{
 // вставка столбца в определенном месте таблицы:
 var
  i, j: шnteger;
  varcol: OleVariant;
 begin
 j:=2;
 varcol:=WordDocument1.Tables.Item(1).Columns.Item(1);
 WordDocument1.Tables.Item(1).Columns.Add(varcol);
 // объединение ячеек
 WordDocument1.Tables.Item(i).Cell(j,k).Merge(WordDocument1.Tables.Item(i).Cell(j,k+1));
 // мы объединили две соседние по горизонтали ячейки (j,k) и (j,k+1).
 // при этом получается, что большая ячейка как бы имеет два "адреса".
 varrow:=1;
 varcol:=2;
 WordDocument1.Tables.Item(i).Cell(j,k).Split(varrow, varcol);
}
{
 // удаление из таблицы второго столбца или третьей строки
 WordDocument1.Tables.Item(1).Columns.Item(2).Delete;
 WordDocument1.Tables.Item(1).Rows.Item(3).Delete;
}
 // фон
{
 wdTextureNone
 wdTexture2Pt5Percent
 wdTexture7Pt5Percent
 wdTexture35Percent
 wdTexture62Pt5Percent
 wdTextureSolid
 wdTextureDarkHorizontal
 wdTextureCross
}
 i:=10;
 j:=6;
 // WordDocument1.Tables.Item(1).Cell(i,j).Shading.Texture:=wdTexture20Percent;
 WordDocument1.Tables.Item(1).Columns.Item(j).Shading.Texture:=wdTexture20Percent;
 WordDocument1.Tables.Item(1).Rows.Item(j).Shading.Texture:=wdTexture20Percent;
 // формат
 WordDocument1.Tables.Item(1).Cell(1,2).Select;
 WordApplication1.Selection.Text:='xXx';
 WordApplication1.Selection.Font.Color:=clRed;
 WordApplication1.Selection.Font.Italic:=1;
 WordApplication1.Selection.Font.Size:=16;
end;

procedure TForm1.Button11Click(Sender: TObject);
begin
 WordApplication1.Connect;
 WordApplication1.Visible:=true;
end;

procedure TForm1.Button13Click(Sender: TObject);
var
 Excel, WorkBook, Sheet: Variant;
begin
 Excel:=CreateOleObject('Excel.Application.10'); // для Office XP
// Excel:=CreateOleObject('Excel.Application'); // для остальных
 Excel.SheetsInNewWorkbook:=1;
 WorkBook:=Excel.WorkBooks.Add;
 Sheet:=WorkBook.WorkSheets[1];
 Sheet.Cells.VerticalAlignment:=xlCenter;
 Sheet.Cells[1, 1]:='XX___1___XX';
 Sheet.Cells[5, 5]:='XX___1___XX'; 
 Sheet.Cells.Columns.AutoFit;
 Excel.Visible:=True;
end;

procedure TForm1.Button15Click(Sender: TObject);
var
 WorkBk: _WorkBook;
 WorkSheet: _WorkSheet;
 K,R,X,Y: integer;
 IIndex: OleVariant;
 RangeMatrix: Variant;
 NomFich: WideString;
begin
 NomFich:=ExtractFilePath(ParamStr(0))+'xl.xls';
 IIndex:=1;
 XLApp.Connect;
 // Открываем файл Excel
 XLApp.WorkBooks.Open(NomFich,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
       EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                                                EmptyParam,EmptyParam,0);
 WorkBk:=XLApp.WorkBooks.Item[IIndex];
 WorkSheet:=WorkBk.WorkSheets.Get_Item(1) as _WorkSheet;
 // Чтобы знать размер листа (WorkSheet), т.е. количество строк и количество
 // столбцов, мы активируем его последнюю непустую ячейку
 WorkSheet.Cells.SpecialCells(xlCellTypeLastCell,EmptyParam).Activate;
 // Получаем значение последней строки
 X:=XLApp.ActiveCell.Row;
 // Получаем значение последней колонки
 Y:=XLApp.ActiveCell.Column;
 // Определяем количество колонок в TStringGrid
 StringGrid1.ColCount:=Y+1;
 // Сопоставляем матрицу WorkSheet с нашей Delphi матрицей
 RangeMatrix:=XLApp.Range['A1',XLApp.Cells.Item[X,Y]].Value2;
 // Выходим из Excel и отсоединяемся от сервера
 XLApp.Quit;
 XLApp.Disconnect;
 //  Определяем цикл для заполнения TStringGrid
 K:=1;
  repeat
   for R:=1 to Y do
    StringGrid1.Cells[(R),(K)]:=RangeMatrix[K,R];
    StringGrid1.RowCount:=K+1;
    inc(K,1);
  until K>X;
 // Un assign the Delphi Variant Matrix
 RangeMatrix:=Unassigned;
end;

end.
