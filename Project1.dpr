program Project1;

uses
  Vcl.Forms,
  Graph_TLB in 'Graph_TLB.pas',
  Unit1 in 'Unit1.pas' {Form1},
  Office_TLB in 'Office_TLB.pas',
  Excel_TLB in 'Excel_TLB.pas';

{$R *.res}

begin
  Vcl.Forms.Application.Initialize;
  Vcl.Forms.Application.MainFormOnTaskbar := True;
  Vcl.Forms.Application.CreateForm(TForm1, Form1);
  Vcl.Forms.Application.Run;
end.
