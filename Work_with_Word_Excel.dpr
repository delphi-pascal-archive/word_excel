program Work_with_Word_Excel;

uses
  Forms,
  Wdof in 'Wdof.pas' {Form1};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
