program SADPF;
//library SADPF;

uses
  //ISAPIApp, {ISAPI}
  //IWInitISAPI, {ISAPI}
  Forms,
  IWMain,
  untEntrada in 'untEntrada.pas' {tfrmEntrada: TIWAppForm},
  ServerController in 'ServerController.pas' {IWServerController: TIWServerControllerBase},
  UserSessionUnit in 'UserSessionUnit.pas' {IWUserSession: TIWUserSessionBase},
  untResultados in 'untResultados.pas' {tfrmResultados: TIWAppForm},
  uCode in 'uCode.pas';

  //exports GetExtensionVersion, HttpExtensionProc, TerminateExtension;

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TformIWMain, formIWMain);
  Application.Run;
  //IwRun; {ISAPI}
end.
