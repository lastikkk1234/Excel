unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Excel_TLB, VBIDE_TLB,
  Math, Graph_TLB, Vcl.Imaging.pngimage, Vcl.ExtCtrls, Vcl.StdCtrls,ComObj;

type
  TForm1 = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label6: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit6: TEdit;
    Button2: TButton;
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  AChart: _Chart;
  mchart: ExcelChart;
  mshape: Shape;
  CellName: String;
  oChart: ExcelChart;
  Col: Char;
  defaultLCID: Cardinal;
  Row: Integer;
  mAxis:Axis;
  GridPrevFile: string;
  MyDisp: IDispatch;
  ExcelApp: ExcelApplication;
  v:variant;
  Sheet: ExcelWorksheet;
  y1, y2, y3, x, xb, xe, a1, a2, a3, st: Extended;

implementation

{$R *.dfm}

procedure TForm1.Button2Click(Sender: TObject);
var
 e1,e2,e3: Integer;
 b: boolean;
begin
  ExcelApp := CreateOleObject('Excel.Application') as ExcelApplication;
  ExcelApp.Visible[0] := True;
  ExcelApp.Workbooks.Add(xlWBatWorkSheet, 0);

  Sheet := ExcelApp.Workbooks[1].WorkSheets[1] as ExcelWorksheet;
  ExcelApp.Application.ReferenceStyle[0] := xlA1;

  Val(Edit1.Text,xb,e1);
  Val(Edit2.Text,xe,e2);
  Val(Edit6.Text,st,e3);
  if (e1 <> 0) or (e2 <> 0) or  (e3 <> 0) then
     raise Exception.Create('Ошибка данных'+#13+'Пожалуйста вводите только числа');
  if xb = 0 then b:= true else b:= false;

  col:='A';
  x:=xb;
  Sheet.Range[col+'1', col+'1'].Value[xlRangeValueDefault]:='X';
  row:=2;

  while ((x<=xe) and (x>=xb)) do    //Заполнения х
    begin
      if x <> 0 then begin
        Sheet.Range[col+IntToStr(row), col+IntToStr(row)].Value[xlRangeValueDefault]:=x;
        row:=row+1;
      end;
      x:=x+st;
    end;

  col:='B';                        //Заполнения Y
  Sheet.Range[col+'1', col+'1'].Value[xlRangeValueDefault]:='Y';
  row:=2;
  x:=xb;
  while (x<=0) and (x>=xb) and (X<=xe) do
    begin
     if x <> 0 then begin
      y1:= 2/(x*x* + 2*x);
      Sheet.Range[col+IntToStr(row), col+IntToStr(row)].Value[xlRangeValueDefault]:=y1;
      row:=row+1;
     end;
     x:=x+st;
    end;
  while (x>=1) and (x>=xb) and (X<=xe) do
    begin
      y1:=2/(x*x* + 2*x);
      Sheet.Range[col+IntToStr(row), col+IntToStr(row)].Value[xlRangeValueDefault]:=y1;
      x:=x+st;
      row:=row+1;
    end;
  while (x>=xb) and (X<=xe) and not ((x>=1) and (x<=0)) do
    begin
      y1:=2/(x*x* + 2*x);
      Sheet.Range[col+IntToStr(row), col+IntToStr(row)].Value[xlRangeValueDefault]:=y1;
      x:=x+st;
      row:=row+1;
    end;
  sheet.Range['A2','B'+inttostr(row)].Select;
  mshape:=Sheet.Shapes.AddChart(xlXYScatterSmoothNoMarkers,250,1,800,800);
  mchart:=(mshape.Chart as ExcelChart).Location(xlLocationAsNewSheet,EmptyParam);
  ExcelApp.Application.ActiveWorkbook.ActiveChart.SetElement(1);
  ExcelApp.Application.ActiveWorkbook.ActiveChart.ChartTitle[0].Text:='График функции';
  MyDisp:=mchart.Axes(xlValue, xlPrimary, 0);
  mAxis:=Axis(MyDisp);
  mAxis.HasTitle:=True;
  mAxis.AxisTitle.Caption:='X';

  MyDisp:=mchart.Axes(xlCategory, xlPrimary, 0);
  mAxis:=Axis(MyDisp);
  mAxis.HasTitle:=True;
  mAxis.AxisTitle.Caption:='Y';

  ExcelApp.Application.ActiveWorkbook.ActiveChart.SetElement(328);
  mchart.HasLegend[0] := False;
  sheet.name := 'Данные';
  mchart.name := 'График функции';
end;

end.
