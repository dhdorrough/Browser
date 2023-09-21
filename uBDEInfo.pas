unit uBDEInfo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, BDE, DB, DBTables;

type
  TForm_BDESystemInfo = class(TForm)
    Button_Close: TButton;
    ListBox1: TListBox;
    procedure FormCreate(Sender: TObject);
    procedure Button_CloseClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

(*
var
  Form_BDESystemInfo: TForm_BDESystemInfo;
*)  

implementation

{$R *.DFM}

function fDbiGetSysVersion(SysVerList: TStringList): SYSVersion;
  var
    Month, Day, iHour, iMin, iSec: Word;
    Year: SmallInt;

  procedure ListDriverInfo;
    var
      DriverList: TStringList;
      i: integer;
      s: string;

    function fDbiOpenTableTypesList(Driver: string): string;

    var
      hTypeCur: hDBICur;
      TblTypes: TBLType;
      BufStr: string;
    begin
      hTypeCur:= nil;
      FillChar(TblTypes, sizeof(TblTypes), 0);
      Check(DbiOpenTableTypesList(PChar(Driver), hTypeCur));
      while (DbiGetNextRecord(hTypeCur, dbiNOLOCK, @TblTypes, nil) = DBIERR_NONE) do
        begin
          BufStr := format('    Name: %s, TableLevel: %d',[Tbltypes.szName, Tbltypes.iTblLevel]);
          result := BufStr;
        end;
    end;

    procedure List1Driver(Driver: string);
    begin { List1Driver }
      with SysVerList do
        try
          begin
            Add(Driver);
            Add(fDbiOpenTableTypesList(Driver));
          end;
        except
          on e: exception do
            Add(Format('    Error on %s [%s]', [Driver, E.message]));
      end;
    end;  { List1Driver }

  begin { ListDriverInfo }
    DriverList := TStringList.Create;
    with SysVerList do
      try
        Session.GetDriverNames(DriverList);
        if DriverList.Count > 0 then
          begin
            // cfmVirtual, cfmPersistent, cfmSession
            s := '';
            if cfmVirtual in Session.ConfigMode then
              s := s + 'Virtual ';
            if cfmPersistent in Session.ConfigMode then
              s := s + 'Persistant ';
            if cfmSession in Session.ConfigMode then
              s := s + 'Session ';
            Add(Format('Available drivers: ConfigMode=%s', [s]));
            for i := 0 to DriverList.Count-1 do
              List1Driver(DriverList[i]);
            List1Driver('DBASE');
            List1Driver('FOXPRO');
            List1Driver('PARADOX');
            List1Driver('MSACCESS');
          end;
      finally
        DriverList.Free;
      end;
  end;  { ListDriverInfo }

begin { fDbiGetSysVersion }
  Check(DbiGetSysVersion(Result));
  if SysVerList <> nil then
  begin
    with SysVerList do
    begin
      Clear;
      Add(Format('ENGINE VERSION=%d', [Result.iVersion]));
      Add(Format('INTERFACE LEVEL=%d', [Result.iIntfLevel]));
      Check(DbiDateDecode(Result.dateVer, Month, Day, Year));
      Add(Format('VERSION DATE=%s', [DateToStr(EncodeDate(Year, Month,Day))]));
      Check(DbiTimeDecode(Result.timeVer, iHour, iMin, iSec));
      Add(Format('VERSION TIME=%s', [TimeToStr(EncodeTime(iHour, iMin,
        iSec div 1000, iSec div 100))]));
      ListDriverInfo;
    end;
  end;
end;  { fDbiGetSysVersion }

procedure TForm_BDESystemInfo.FormCreate(Sender: TObject);
  var
    List: TStringList;
    i: integer;
begin
  List := TStringList.Create;
  Check(DbiInit(nil));
  fDbiGetSysVersion(List);
  for i := 0 to List.Count-1 do
    ListBox1.Items.Add(List[i]);
  List.Free;
end;

procedure TForm_BDESystemInfo.Button_CloseClick(Sender: TObject);
begin
  Close;
end;

end.



