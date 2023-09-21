program Browser;

uses
  (*ExceptionLog,*)
  Forms,
  uBDEInfo in 'uBDEInfo.pas' {Form_BDESystemInfo},
  uBrowser in 'uBrowser.pas' {Form_Browser},
  uBrowseMemo in 'uBrowseMemo.pas' {frmBrowseMemo},
  GotoRecNo in 'GotoRecNo.pas' {frmGotoRecNo};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TForm_Browser, Form_Browser);
  Application.CreateForm(TfrmGotoRecNo, frmGotoRecNo);
  Application.Run;
end.
