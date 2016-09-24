library ieSnake;

uses
  ComServ,
  ieSnake_TLB,
  main,
  ieFilter_TLB in 'ieFilter_TLB.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.

