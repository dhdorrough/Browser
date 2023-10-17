// Program was compiled with Delphi7.0 (Build 8.1)
// Program uses Turbo Power Orpheus library (available on SourceForge)
// Program uses Turbo Power Systools library (available on SourceForge)
{$Define pos}
unit uBrowser;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, DBGrids, Db, DBTables, ExtCtrls, Buttons, ComCtrls,
  DBCtrls, SoftReg, uBrowseMemo, DBInpReq, Menus; 

type
  TForm_Browser = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    DBGrid1: TDBGrid;
    OpenDialog1: TOpenDialog;
    Panel4: TPanel;
    DBNavigator1: TDBNavigator;
    Edit2: TLabel;
    btnGotoRecNo: TButton;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Exit1: TMenuItem;
    N2: TMenuItem;
    Open1: TMenuItem;
    Close1: TMenuItem;
    Navigate1: TMenuItem;
    FindRecord1: TMenuItem;
    AddRecord1: TMenuItem;
    TopRecord1: TMenuItem;
    PreviousRecord1: TMenuItem;
    NextRepord1: TMenuItem;
    BottomRecord1: TMenuItem;
    HideRecord1: TMenuItem;
    N1: TMenuItem;
    Order1: TMenuItem;
    N3: TMenuItem;
    RebuildIndexes1: TMenuItem;
    Functions1: TMenuItem;
    ShowBDEInfo2: TMenuItem;
    N4: TMenuItem;
    RecentDBFs1: TMenuItem;
    Label_DBFName: TLabel;
    FindDialog1: TFindDialog;
    PackTable1: TMenuItem;
    PopupMenu1: TPopupMenu;
    DisplayMemoasText1: TMenuItem;
    DisplayMemoasObjects1: TMenuItem;
    MovetoFirstColumn1: TMenuItem;
    ShowDeletedRecords1: TMenuItem;
    lblIsNullField: TLabel;
    procedure Table1AfterScroll(DataSet: TDataSet);
    procedure FormCreate(Sender: TObject);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure Table1AfterOpen(DataSet: TDataSet);
    procedure btnRebuildIndexesClick(Sender: TObject);
    procedure ShowBDEInfo1Click(Sender: TObject);
    procedure Open1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Navigate1Click(Sender: TObject);
    procedure Close1Click(Sender: TObject);
    procedure AddRecord1Click(Sender: TObject);
    procedure TopRecord1Click(Sender: TObject);
    procedure PreviousRecord1Click(Sender: TObject);
    procedure NextRepord1Click(Sender: TObject);
    procedure BottomRecord1Click(Sender: TObject);
    procedure HideRecord1Click(Sender: TObject);
    procedure RebuildIndexes1Click(Sender: TObject);
    procedure ShowBDEInfo2Click(Sender: TObject);
    procedure ModifyStructure1Click(Sender: TObject);
    procedure FindDialog1Find(Sender: TObject);
    procedure FindRecord1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure PackTable1Click(Sender: TObject);
    procedure DisplayMemoasText1Click(Sender: TObject);
    procedure DisplayMemoasObjects1Click(Sender: TObject);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure MovetoFirstColumn1Click(Sender: TObject);
    procedure btnGotoRecNoClick(Sender: TObject);
    procedure ShowDeletedRecords1Click(Sender: TObject);
  private
    { Private declarations }
    fDefaultPath: string;
    RecentDBFs: TStringList;
    BrowseMemoList: TList;
    fRecordCount: integer;
    SoftReg1: TSoftReg;
    procedure OrderClick(Sender: TObject);
    procedure Enable_Stuff;
    procedure EnableNavigateMenu;
    procedure ReBuildIndexes;
    procedure UpdateTagNames;
    procedure CloseDBF;
    procedure UpdateRecentDBFsMenuItem;
    procedure DisplayMemoAsObjects(BrowseHow: TBrowseHow);
    procedure ShowBDEInfo;
    procedure Table1AfterInsert(DataSet: TDataSet);
    procedure UpdateOpenMemoFields;
  public
    { Public declarations }
    Table1: TTable;
    DataSource1: TDataSource;
    procedure OpenDBF(lfn: string);
    procedure AddHistoryItem(lfn: string);
    procedure SaveHistoryList;
    procedure OpenRecentDBF(Sender: TObject);
  end;

var
  Form_Browser: TForm_Browser;
  
implementation

{$R *.DFM}

uses
  BDE, uBDEInfo, StStrL, StStrZ, StBase, MyUtils,
  GotoRecNo;

CONST
  MAXHISTORY = 10;
  FILE_HISTORY = 'File History';
  FILE_PATHS   = 'File Paths';
  DEFAULT_PATH = 'Default Path';

procedure TForm_Browser.SaveHistoryList;
  var
    Key  : HKEY;
    i: integer;
begin { TForm_Browser.SaveHistoryList }
  with SoftReg1 do
    begin
      ProductName := 'Browser';
      OpenRootKey(rkUser);
      Key := OpenKey(0, FILE_HISTORY, true);
      with RecentDBFs do
        for i := 0 to Count-1 do
          WriteString(Key, 'File'+IntToStr(i), RecentDBFs[i]);

      Key := OpenKey(0, FILE_PATHS, true);
      WriteString(Key, DEFAULT_PATH, fDefaultPath);

      CloseRootKey;
    end;
end;  { TForm_Browser.SaveHistoryList }


procedure TForm_Browser.Table1AfterScroll(DataSet: TDataSet);
begin
  with DataSet as TTable do
    if not ControlsDisabled then
      begin
        if Active and (fRecordCount > 0) then
          begin
            btnGotoRecNo.Caption := IntToStr(RecNo);
            Edit2.Caption := IntToStr(fRecordCount);
            lblIsNullField.Visible := DBGrid1.SelectedField.IsNull;
          end
        else
          begin
            btnGotoRecNo.Caption := '';
            Edit2.Caption := 'No records';
            lblIsNullField.Visible := false;
          end;

        EnableNavigateMenu;
        UpdateOpenMemoFields;
      end;
end;

procedure TForm_Browser.Table1AfterInsert(DataSet: TDataSet);
begin
  DBGrid1.SelectedField := DataSet.Fields[0];
end;


procedure TForm_Browser.CloseDBF;
  var
    i: integer;
begin
  with Table1 do
    if Active then
      begin
        if State in [dsEdit, dsInsert] then
          Post;
        Active := false;
        for i := 0 to BrowseMemoList.Count-1 do
          TForm(BrowseMemoList[i]).Free;
        BrowseMemoList.Clear;
        Label_DBFName.Caption := '';
        btnGotoRecNo.Caption  := '';
        Table1AfterScroll(Table1);
        SaveHistoryList;
      end;
  Enable_Stuff;
end;

procedure TForm_Browser.OpenDBF(lfn: string);
  var
    ext: string[4];
begin
  with Table1 do
    begin
      CloseDBF;
      DataBaseName := ExtractFilePath(lfn);
      TableName    := ExtractFileName(lfn);
      Ext          := UpperCase(ExtractFileExt(lfn));
      if ext = '.DB' then
        TableType  := ttParadox
      else
        TableType  := ttDBase;
      IndexName    := '';
      try
        Active       := true;
        fRecordCount := RecordCount;
        Label_DBFName.Caption := Format('%s (TableLevel=%d)', [lfn, TableLevel]);
        AddHistoryItem(lfn);
        fDefaultPath := ExtractFilePath(lfn);
        Check(DbiSetProp(hDBIObj(Table1.Handle), curSOFTDELETEON, LongInt(ShowDeletedRecords1.Checked)));
        Table1AfterScroll(Table1);
      except
        on e:Exception do
          Error(Format('Unable to open %s [%s]', [lfn, e.message]));
      end;
    end;
  Enable_Stuff;
end;

procedure TForm_Browser.Enable_Stuff;
begin
  with Table1 do
    begin
      Open1.Enabled             := not Active;
      Close1.Enabled            := Active;
      RebuildIndexes1.Enabled   := Active;
      PackTable1.Enabled        := Active;
      if not Active then
        begin
          btnGotoRecNo.Caption  := '';
          Edit2.Caption         := '';
        end;
    end;
  EnableNavigateMenu;
end;

procedure TForm_Browser.FormCreate(Sender: TObject);
  var
    Key  : HKEY;
    b    : boolean;
    lfn  : string;
    MailingListDir: string;
begin { TForm_Browser.FormCreate }
  Label_DBFName.Caption := '';
  Table1           := TTable.Create(self);
  with Table1 do
    begin
      AfterOpen   := Table1AfterOpen;
      AfterScroll := Table1AfterScroll;
      AfterInsert := Table1AfterInsert;
    end;
  DataSource1         := TDataSource.Create(self);
  DataSource1.DataSet := Table1;
  DBGrid1.DataSource  := DataSource1;
  DBNavigator1.DataSource := DataSource1;

  RecentDBFs          := TStringList.Create;
  BrowseMemoList      := TList.Create;

  SoftReg1 := TSoftReg.Create;
  with SoftReg1 do
    begin
      CompanyName := 'BCC Software';
      ProductName := 'Mail Manager 2010';

      OpenRootKey(rkLocal);
      Key  := OpenKey(0, 'Paths', false);
      MailingListDir := ReadString(Key, 'BCC_MailLists_Path', 'C:\BCC\MM2010\LISTS');
      CloseRootKey;

      ProductName := 'Browser';
      OpenRootKey(rkUser);
      Key := OpenKey(0, FILE_HISTORY, true);
      lfn := ReadString(Key, 'File'+IntToStr(RecentDBFs.Count), '');
      b   := lfn <> '';
      while b do
        begin
          RecentDBFs.Add(Lfn);
          if RecentDBFs.Count < MAXHISTORY then
            begin
              lfn := ReadString(Key, 'File'+IntToStr(RecentDBFs.Count), '');
              b   := lfn <> '';
            end
          else
            b := false;
        end;

      Key := OpenKey(0, FILE_PATHS, true);
      fDefaultPath := ReadString(Key, DEFAULT_PATH, MailingListDir);
      CloseRootKey;

//    if RecentDBFs.Count = 0 then
//      begin
//        if fMailingListDir <> '' then
//          AddHistoryItem(fMailingListDir + 'MAILLIST.DBF');
//
//        CloseRootKey;
//      end;
      UpdateRecentDBFsMenuItem;
    end;

  if ParamStr(1) <> '' then { if passed a parameter, try to use it as lfn }
    OpenDBF(ParamStr(1));

  Enable_Stuff;
end;  { TForm_Browser.FormCreate }

procedure TForm_Browser.AddHistoryItem(lfn: string);
  var
    i: integer;
begin { TForm_Browser.AddHistoryItem }
  with RecentDBFs do
    begin
      i := IndexOf(lfn);
      if i >= 0 then { already in list }
        begin
          if i > 0 then  { not already at top }
            begin
              Delete(i);
              Insert(0, lfn);        { move to the top }
            end;
        end else
      if Count = MAXHISTORY then { history list is full }
        begin
          Delete(MAXHISTORY-1);  { delete the oldest }
          Insert(0, lfn);        { add newset at the top }
        end
      else
        Insert(0, lfn);          { add at the top }
    end;
  UpdateRecentDBFsMenuItem;
end;  { TForm_Browser.AddHistoryItem }

procedure TForm_Browser.UpdateRecentDBFsMenuItem;
  var
    i: integer;
    aMenuItem: TMenuItem;
begin
  with RecentDBFs1 do
    begin
      { empty previous 'Recent DBFs' sub-menu }
      for i := Count-1 downto 0 do
        Delete(i);

      for i := 0 to RecentDBFs.Count-1 do
        begin
          aMenuItem := TMenuItem.Create(self);
          with aMenuItem do
            begin
              Caption := RecentDBFs[i];
              OnClick := OpenRecentDBF;
              AutoHotkeys := maManual;
            end;
          Add(aMenuItem);
        end;
    end;
end;



procedure TForm_Browser.DBGrid1DblClick(Sender: TObject);
  var
    aBrowseMemo: TForm;
begin
  aBrowseMemo := TfrmBrowseMemo.Create(self, DataSource1, DBGrid1.SelectedField, bhAsText);
  BrowseMemoList.Add(aBrowseMemo);
  aBrowseMemo.Show;
end;

procedure TForm_Browser.Table1AfterOpen(DataSet: TDataSet);
begin
  with Table1 do
    begin

    end;
  UpdateTagNames;
end;

procedure TForm_Browser.UpdateTagNames;
  var
    aMenuItem: TMenuItem;
    i: integer;
    Items: TStringList;
begin
  with Table1 do
    begin
      Items := TStringList.Create;
      try
        Items.Clear;
        GetIndexNames(Items);
        Items.Add('(recno order)');

        { empty old 'Order' sub-menu }

        with Order1 do
          for i := Count-1 downto 0 do
            Delete(i);

        { Add items to 'Order1' sub-menu }
        for i := Items.Count-1 downto 0 do
          begin
            aMenuItem := TMenuItem.Create(self);
            aMenuItem.Caption := Items[i];
            aMenuItem.OnClick := OrderClick;
            Order1.Add(aMenuItem);
          end;
      finally
        Items.Free;
      end;
    end;
end;


procedure TForm_Browser.btnRebuildIndexesClick(Sender: TObject);
begin
  RebuildIndexes;
end;


procedure TForm_Browser.ShowBDEInfo1Click(Sender: TObject);
begin
  ShowBDEInfo;
end;

procedure TForm_Browser.ShowBDEInfo;
begin
  with TForm_BDESystemInfo.Create(self) do
    begin
      ShowModal;
      Free;
    end;
end;


procedure TForm_Browser.Open1Click(Sender: TObject);
begin
  with OpenDialog1 do
    begin
      InitialDir := fDefaultPath;
      if Execute then
        OpenDBF(FileName);
    end;
end;

procedure TForm_Browser.Exit1Click(Sender: TObject);
begin
  Close;
end;

procedure TForm_Browser.Navigate1Click(Sender: TObject);
begin
  EnableNavigateMenu;
end;

procedure TForm_Browser.EnableNavigateMenu;
begin
  with Table1 do
    begin
      Open1.Enabled            := not Active;
      Close1.Enabled           := Active;
      Navigate1.Enabled        := Active;
      FindRecord1.Enabled      := not EOF;
      AddRecord1.Enabled       := Active;
      TopRecord1.Enabled       := not BOF;
      PreviousRecord1.Enabled  := not BOF;
      NextRepord1.Enabled      := not EOF;
      BottomRecord1.Enabled    := not EOF;
      Order1.Enabled           := Active;
    end;
end;


procedure TForm_Browser.Close1Click(Sender: TObject);
begin
  CloseDBF;
end;

procedure TForm_Browser.AddRecord1Click(Sender: TObject);
begin
  Table1.Append;
end;

procedure TForm_Browser.TopRecord1Click(Sender: TObject);
begin
  Table1.First;
end;

procedure TForm_Browser.PreviousRecord1Click(Sender: TObject);
begin
  Table1.Prior;
end;

procedure TForm_Browser.NextRepord1Click(Sender: TObject);
begin
  Table1.Next;
end;

procedure TForm_Browser.BottomRecord1Click(Sender: TObject);
begin
  Table1.Last;
end;

procedure TForm_Browser.HideRecord1Click(Sender: TObject);
begin
  Table1.Delete;
end;

procedure TForm_Browser.RebuildIndexes1Click(Sender: TObject);
begin
  RebuildIndexes;
end;

procedure TForm_Browser.ShowBDEInfo2Click(Sender: TObject);
begin
  ShowBDEInfo
end;

procedure TForm_Browser.OrderClick(Sender: TObject);
  var
    idx: integer;
    aMenuItem: TMenuItem;
begin
  with Table1 do
    begin
      aMenuItem := TMenuItem(Sender);
      idx := Order1.IndexOf(aMenuItem);
      if idx > 0 then { index selected }
        IndexName := aMenuItem.Caption
      else
        IndexName := '';   // recno order selected
    end;
end;

procedure TForm_Browser.ModifyStructure1Click(Sender: TObject);
begin
(*
  with TModify_Table_Structure.Create(self, Table1) do
    begin
      ShowModal;
      Free;
    end;
*)
end;

procedure TForm_Browser.OpenRecentDBF(Sender: TObject);
begin
  with Sender as TMenuItem do
    OpenDBF(Caption);
end;

procedure TForm_Browser.FindDialog1Find(Sender: TObject);
  label
    FOUND_IT;
  var
    i: integer;
    Saved_Recno: integer;
    FoundIt: boolean;
    MatchString: string;
    FldNr: integer;
{$IfDef pos}
    temp2: string;
{$EndIf}
begin { TForm_Browser.FindDialog1Find }
  with FindDialog1 do
    begin
{$IfDef Pos}
      MatchString := UpperCase(FindText);
{$EndIf}
      with Table1 do
        begin
          DisableControls;
          Saved_Recno := Recno;
          FoundIt     := false;
          FldNr       := -1;   // just to keep compiler happy
          Next;                // start searching at next record
          while not EOF do
            begin
              for i := 0 to Fields.Count-1 do
                with Fields[i] do
                  begin
{$IfDef Pos}
                    Temp2 := UpperCase(AsString);
                    FoundIt := Pos(MatchString, Temp2) > 0;
{$EndIf}
                    if FoundIt then
                      begin
                        FldNr := i;
                        goto FOUND_IT;
                      end;
                  end;
              Next;
            end;
FOUND_IT:
          EnableControls;
          if FoundIt then
            begin
              DBGrid1.SelectedField := Fields[FldNr];
              FindDialog1.CloseDialog;
            end
          else
            begin
              Recno := Saved_Recno;
              Error(Format('Search string "%s" could not be found', [FindText]));
            end;
          Table1AfterScroll(Table1);  // do it with Controls Enabled
        end;
  end;
end;  { TForm_Browser.FindDialog1Find }

procedure TForm_Browser.FindRecord1Click(Sender: TObject);
begin
  FindDialog1.Execute;
end;

procedure TForm_Browser.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  SaveHistoryList;
end;

procedure TForm_Browser.PackTable1Click(Sender: TObject);
begin
  with Table1 do
    begin
      close;
      Exclusive := true;
      open;
      Check(DbiPackTable(DBHandle, Handle, NIL, szDBASE, TRUE));
      close;
      Exclusive := false;
      Open;
    end;
end;

procedure TForm_Browser.ReBuildIndexes;
begin
  with Table1 do
    begin
      close;
      Exclusive := true;
      Check(DbiSetProp(hDBIObj(Table1.Handle), curSOFTDELETEON, LongInt(true)));  // include hidden records
      open;
      Check(DbiRegenIndexes(handle));
      close;
      Exclusive := false;
      Open;
    end;
end;

procedure TForm_Browser.DisplayMemoasText1Click(Sender: TObject);
  var
    aBrowseMemo: TForm;
begin
  aBrowseMemo := TfrmBrowseMemo.Create(self, DataSource1, DBGrid1.SelectedField, bhAsText);
  BrowseMemoList.Add(aBrowseMemo);
  aBrowseMemo.Show;
end;

procedure TForm_Browser.DisplayMemoAsObjects(BrowseHow: TBrowseHow);
  var
    aBrowseMemo: TForm;
begin
  aBrowseMemo := TfrmBrowseMemo.Create(self, DataSource1, DBGrid1.SelectedField, BrowseHow);
  BrowseMemoList.Add(aBrowseMemo);
  aBrowseMemo.Show;
end;

procedure TForm_Browser.DisplayMemoasObjects1Click(Sender: TObject);
begin
  DisplayMemoAsObjects(bhAsObject);
end;

procedure TForm_Browser.PopupMenu1Popup(Sender: TObject);
  var
    b1, b2: boolean;
begin
  b1 := DBGrid1.SelectedField is TBlobField;
  DisplayMemoasText1.Enabled := b1;
  if b1 then
    begin
      b2 := MemoContainsObject(DBGrid1.SelectedField);
      DisplayMemoasObjects1.Enabled := b2;
    end
  else
    DisplayMemoasObjects1.Enabled := false;
end;

procedure TForm_Browser.MovetoFirstColumn1Click(Sender: TObject);
begin
  DBGrid1.SelectedField := Table1.Fields[0];
end;

procedure TForm_Browser.UpdateOpenMemoFields;
  var
    i: integer;
begin
  for i := 0 to BrowseMemoList.Count-1 do
    with TfrmBrowseMemo(BrowseMemoList[i]) do
      begin
        if BrowseHow = bhAsObject then
          UpdateObjectMemo;
      end;
end;

procedure TForm_Browser.btnGotoRecNoClick(Sender: TObject);
begin
  with frmGotoRecNo do
     if ShowModal = mrOk then
       begin
         Check(DbiSetToRecordNo(Table1.Handle, OvcNumericField1.AsInteger));
         Table1.Resync([rmCenter]);
         Table1AfterScroll(Table1);
       end;
end;

procedure TForm_Browser.ShowDeletedRecords1Click(Sender: TObject);
begin
  ShowDeletedRecords1.Checked := not ShowDeletedRecords1.Checked;
  Check(DbiSetProp(hDBIObj(Table1.Handle), curSOFTDELETEON, LongInt(ShowDeletedRecords1.Checked)));
  Table1.Refresh;
  Table1AfterScroll(Table1);
end;

end.
