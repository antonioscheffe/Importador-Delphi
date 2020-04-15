program ImportadorRotas;

uses
  Forms,
  UImportarRotas in 'UImportarRotas.pas' {FrmPrincipal};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFrmPrincipal, FrmPrincipal);
  Application.Run;
end.
