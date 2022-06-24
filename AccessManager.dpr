program AccessManager;

uses
  Vcl.Forms,
  uMain in 'uMain.pas' {frmMain},
  uAccessUtil in 'uAccessUtil.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
