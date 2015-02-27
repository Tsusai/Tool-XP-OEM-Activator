program Activator;

uses
  Forms,
  Main in 'src\Main.pas' {Form1},
  WindowsAPI in 'src\WindowsAPI.pas';

{$R *.res}

begin
	Application.Initialize;
	Application.Title := 'Windows OEM Activator';
	Application.CreateForm(TForm1, Form1);
	Application.Run;
end.
