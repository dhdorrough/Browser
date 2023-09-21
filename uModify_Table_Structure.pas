unit uModify_Table_Structure;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, OvcTCHdr, OvcTCBEF, OvcTCNum, OvcTCmmn, OvcTCell, OvcTCStr,
  OvcTCEdt, OvcBase, OvcTable, ExtCtrls, OvcTCCBx, DB, DBTables, BCCTables,
  OvcTCSim, ImportTypes;

type
  TMyFieldDef = class
                 FieldDescRecord: TFieldDescRecord;
                 Old_Name  : string[16];
                 Type_Index: longint;  { Contains the index into the col_fieldType.Items array. }
                 Field_Name: string[16];
                 DBFldType : word;
                 Field_Width: integer;
                 Field_Dec: integer;
                end;

  TMyFieldDefs = TList;

  TModify_Table_Structure = class(TForm)
    OvcController1: TOvcController;
    Col_FieldWidth: TOvcTCNumericField;
    Button_Ok: TButton;
    Button_Cancel: TButton;
    Panel1: TPanel;
    Button_Insert: TButton;
    Button_Delete: TButton;
    OvcTCColHead1: TOvcTCColHead;
    col_FieldType: TOvcTCComboBox;
    BatchMove1: TBatchMove;
    Label_TableName: TLabel;
    Panel_BehindGrid: TPanel;
    OvcTable1: TOvcTable;
    Col_FieldDec: TOvcTCNumericField;
    col_FieldName: TOvcTCSimpleField;
    procedure FormDestroy(Sender: TObject);
    procedure OvcTable1GetCellData(Sender: TObject; RowNum: Longint;
      ColNum: Integer; var Data: Pointer; Purpose: TOvcCellDataPurpose);
    procedure Button_InsertClick(Sender: TObject);
    procedure Button_DeleteClick(Sender: TObject);
    procedure Button_OkClick(Sender: TObject);
    procedure Col_FieldNameExit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure Panel_BehindGridResize(Sender: TObject);
    procedure Col_FieldDecUserValidation(Sender: TObject;
      var ErrorCode: Word);
    procedure OvcController1Error(Sender: TObject; ErrorCode: Word;
      const ErrorMsg: String);
    procedure OvcTable1GetCellAttributes(Sender: TObject; RowNum,
      ColNum: Integer; var CellAttr: TOvcCellAttributes);
    procedure OvcTable1DoneEdit(Sender: TObject; RowNum, ColNum: Integer);
    procedure col_FieldNameUserValidation(Sender: TObject;
      var ErrorCode: Word);
    procedure Col_FieldWidthUserValidation(Sender: TObject;
      var ErrorCode: Word);
  private
    { Private declarations }
    Table_MailingList: TBCCTable;
  public
    { Public declarations }
    MyFieldDefs : TMyFieldDefs;
    CellData: string[255];
    CellNumber: shortint;
    CellCombo: TCellComboBoxInfo;
    function Any_Duplicate_Field_Names: boolean;
    constructor Create(AOwner: TComponent; aTable: TBCCTable);
                reintroduce;
  end;

  ECantDelete = class(exception);
  ECantRename = class(exception);

var
  Modify_Table_Structure: TModify_Table_Structure;

implementation

uses MW_Decl, uMW_Codes, MWStrings, BccStrings, BDE, OvcData;

const
  oeOutOfRange = oeCustomError + 1;
  oeINVALIDFIELDNAME = oeCustomError + 2;

  COLNR_FIELDNAME  = 0;
  COLNR_FIELDTYPE  = 1;
  COLNR_FIELDWIDTH = 2;
  COLNR_FIELDDEC   = 3;

{$R *.DFM}

procedure TModify_Table_Structure.FormDestroy(Sender: TObject);
  var
    i : integer;
begin
  if Assigned(MyFieldDefs) then
    for i := MyFieldDefs.Count-1 downto 0 do
      begin
        TMyFieldDef(MyFieldDefs.Items[i]).Free;
        MyFieldDefs.Delete(i);
      end;

  MyFieldDefs.Free;
end;

procedure TModify_Table_Structure.OvcTable1GetCellData(Sender: TObject; RowNum: Longint;
      ColNum: Integer; var Data: Pointer; Purpose: TOvcCellDataPurpose);
begin { TForm_Table_Structure.OvcTable1GetCellData }
  if RowNum > 0 then
    with TMyFieldDef(MyFieldDefs[RowNum-1]) do
      case colnum of
        COLNR_FIELDNAME: { the field name column }
            data := @Field_Name;

        COLNR_FIELDTYPE: { the field type column }
            data := @Type_Index;

        COLNR_FIELDWIDTH: { the field width column }
            data := @Field_Width;

        COLNR_FIELDDEC: { the field decimal column }
            data := @Field_Dec;
      end;
end;  { TForm_Table_Structure.OvcTable1GetCellData }

procedure TModify_Table_Structure.Button_InsertClick(Sender: TObject);
  var
    aMyFieldDef : TMyFieldDef;
begin
  aMyFieldDef := TMyFieldDef.Create;
  with aMyFieldDef do
    begin
      Old_Name    := '';
      Type_Index  := 0;
      Field_Name  := 'Untitled' + IntToStr(MyFieldDefs.Count+1);
      DBFldType   := fldDBCHAR;
      Field_Width := 10;
      Field_Dec   := 0;
    end;
  MyFieldDefs.Insert(OvcTable1.ActiveRow-1, aMyFieldDef);
  OvcTable1.RowLimit := MyFieldDefs.Count+1;
  OvcTable1.Refresh;
  ActiveControl := OvcTable1;
end;

procedure TModify_Table_Structure.Button_DeleteClick(Sender: TObject);
begin
  MyFieldDefs.Delete(OvcTable1.ActiveRow-1);
  OvcTable1.RowLimit := MyFieldDefs.Count+1;
  OvcTable1.Refresh;
  ActiveControl := OvcTable1;
end;

function TModify_Table_Structure.Any_Duplicate_Field_Names: boolean;
  var
    i, j: integer;
begin { TModify_Table_Structure.Any_Duplicate_Field_Names }
  for i := 0 to MyFieldDefs.Count-2 do
    for j := i+1 to MyFieldDefs.Count-1 do
      if UpperCase(TMyFieldDef(MyFieldDefs[i]).Field_Name) =
         UpperCase(TMyFieldDef(MyFieldDefs[j]).Field_Name) then
        begin
          result := true;
          exit;
        end;
  result := false;
end;  { TModify_Table_Structure.Any_Duplicate_Field_Names }

procedure TModify_Table_Structure.Button_OkClick(Sender: TObject);
  var
    oldlfn          : string;
    newlfn          : string;
    temp            : string;
    aMyFieldDef     : TMyFieldDef;
    NewBCCFieldDescList: TBCCFieldDescList;
    Table_OutFile   : TBCCTable;

  procedure CreateFieldDescList;
    var
      i: integer;
  begin { CreateFieldDescList }
    NewBCCFieldDescList := TBCCFieldDescList.Create;
    for i := 0 to MyFieldDefs.Count-1 do
      begin
        aMyFieldDef := TMyFieldDef(MyFieldDefs.Items[i]);

        { Create the field definitions for the new table }

        with aMyFieldDef do
          begin
            StrPCopy(FieldDescRecord.Name, Field_Name);
            FieldDescRecord.FldType   := DBFldType;
            FieldDescRecord.FldNum    := i+1;
            FieldDescRecord.Len       := Field_Width;
            FieldDescRecord.Units1    := Field_Width;
            FieldDescRecord.Units2    := Field_Dec;

            NewBCCFieldDescList.Add(FieldDescRecord);

            { Create the field mappings }

            if not empty(Old_Name) then
              begin
                if (Old_Name = Field_Name) then
                  temp := Field_Name
                else
                  temp := Field_Name+'='+Old_Name;

                BatchMove1.Mappings.Add(temp)
              end;
          end;
      end;
  end;  { CreateFieldDescList }

  procedure AddIndexes;
    var
      i: integer;
  begin { AddIndexes }
    { Create index definitions for the new table }

    { Note: we should probably be trying to determine if any
            of the fields needed for this index have been
            deleted or significantly changed during the
            table modification process }

    for i := 0 to Table_MailingList.IndexDefs.count - 1 do
      with Table_MailingList.IndexDefs.Items[i] do
        Table_Outfile.AddIndex(Name, Fields, Options);
  end;  { AddIndexes }

  procedure InitializeOutTable;
    var
      temp_table_name : string;
      temp_table_path : string;
  begin { InitializeOutTable }
    { initialize the output file }

    Temp_Table_Name := RandomStr(8) + ExtractFileExt(Table_MailingList.TableName);
    Table_OutFile   := TBCCTable.Create(self);

    with Table_OutFile do
      begin
{$ifdef NoNovellKludge}
        DataBaseName := Table_MailingList.DataBaseName;
        TableName    := Temp_Table_Name;
{$endif}
{$ifNdef NoNovellKludge}
        DataBaseName := '';
        with Table_MailingList do
          temp_table_path := ExtractFilePath(DataBaseName+TableName);
        TableName    := Temp_Table_Path + Temp_Table_Name;
{$endif}
        TableType    := Table_MailingList.TableType;
      end;
  end;  { InitializeOutTable }

  procedure ConvertTable;
  begin { ConvertTable }
    { now actually do the conversion }

    Table_MailingList.active  := false;

    { Should probably verify that enough space exists here
      before doing the copy }

    with BatchMove1 do
      begin
        Source         := Table_MailingList;
        Destination    := Table_Outfile;
        AbortOnProblem := true;
        mode           := batAppend;
        Execute;
      end;
  end;  { ConvertTable }

  procedure CleanUp;
  begin { CleanUp }
    { get rid of the old file }

    with Table_MailingList do
      oldlfn := DatabaseName + TableName;

    with Table_OutFile do
      NewLfn := DatabaseName + TableName;

    if not DeleteFile(oldlfn) then
      ECantDelete.CreateFmt('Unable to delete file "%s"', [oldlfn]);

    { rename temporary file to have permanent name }

    if not RenameFile(NewLfn, OldLfn) then
      ECantRename.CreateFmt( 'Unable to rename %s to %s',
                             [NewLfn, OldLfn]);

    Table_MailingList.Active := true;  { open the restructured table }
  end;  { CleanUp }

begin { TModify_Table_Structure.Button_OkClick }
  if Any_Duplicate_Field_Names then
    begin
      alert('There are duplicate field names. Cannot update list structure');
      ModalResult := mrNone;
    end else
  if Yes('Do you want to permanently update the table structure? ') then
    begin
      if not OvcTable1.SaveEditedData then
        alert('OvcTable1.SaveEditedData failed');;
      OvcTable1.Refresh;

      try
        InitializeOutTable;
        CreateFieldDescList;
        try
          Table_Outfile.BCCCreateDBaseTable(4, NewBCCFieldDescList);   { Create the new table }
          ConvertTable;
          AddIndexes;
          CleanUp;
        except
          on e:Exception do
            begin
              Error(Format('Unable to create new table [%s]', [e.message]));
              ModalResult := mrNone;
            end;
        end;
      finally
        Table_Outfile.Free;
      end;
    end;
end;  { TModify_Table_Structure.Button_OkClick }

procedure TModify_Table_Structure.Col_FieldNameExit(Sender: TObject);
begin
  OvcTable1.SaveEditedData;
end;

procedure TModify_Table_Structure.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
   if ORD(Key) = VK_RETURN then begin
     {change the focus to the next control in the tab order}

     SelectNext(ActiveControl as TWinControl, True, True);
     {clear "Key" to skip default handling}
     Key := #0;
   end;
end;

constructor TModify_Table_Structure.Create(AOwner: TComponent;
  aTable: TBCCTable);
  var
    i                 : integer;
    aMyFieldDef       : TMyFieldDef;
    aFieldDescRecordP : PFieldDescRecord;
    FieldDescList     : TBCCFieldDescList;

begin { TModify_Table_Structure.Create }
  inherited Create(AOwner);
  Table_MailingList := aTable;
  with Table_MailingList do
    begin
      Label_TableName.Caption := ExtractFileName(UpperCase(DataBaseName+TableName));

      { set up the drop-down list of types }
      with col_FieldType.Items do
        begin
          AddObject('CHAR',           TObject(fldDBCHAR));     { Char string }
          AddObject('NUMBER',         TObject(fldDBNUM));      { Number }
          AddObject('MEMO',           TObject(fldDBMEMO));     { Memo          (blob) }
          AddObject('LOGICAL',        TObject(fldDBBOOL));     { Logical }
          AddObject('DATE',           TObject(fldDBDATE));     { Date }
          AddObject('FLOAT',          TObject(fldDBFLOAT));    { Float }
          AddObject('BLOB',           TObject(fldDBBINARY));   { Binary data   (blob) }
          AddObject('LONGINT',        TObject(fldDBLONG));     { Long (Integer) }
          AddObject('DATETIME',       TObject(fldDBDATETIME)); { DateTime }
          AddObject('DOUBLE',         TObject(fldDBDOUBLE));   { Double }
          AddObject('AUTOINC',        TObject(fldDBAUTOINC));  { Auto increment (long) }
        end;

      MyFieldDefs        := TMyFieldDefs.Create;
      OvcTable1.RowLimit := FieldDefs.count+1; { save room for title row }

      FieldDescList      := TBCCFieldDescList.Create;
      try
        GetFieldDescriptions(FieldDescList);
        for i := 1 to FieldDescList.count do
          begin
            aFieldDescRecordP := FieldDescList.Items[i];
            aMyFieldDef       := TMyFieldDef.Create;
            with aMYFieldDef, aFieldDescRecordP^ do
              begin
                aMyFieldDef.FieldDescRecord := aFieldDescRecordP^;
                Field_Name  := Name;
                Old_Name    := Name;
                Type_Index  := col_FieldType.Items.IndexOfObject(TObject(FldType));
                DBFldType   := FldType;
                Field_Width := Len;
                Field_Dec   := Units2;
              end;
            MyFieldDefs.Add(aMyFieldDef);
          end;
      finally
        FieldDescList.Free;
      end;
    end;
end;  { TModify_Table_Structure.Create }

procedure TModify_Table_Structure.Panel_BehindGridResize(Sender: TObject);
begin
  OvcTable1.Width  := Panel_BehindGrid.Width;
  OvcTable1.Height := Panel_BehindGrid.Height;
end;

procedure TModify_Table_Structure.Col_FieldDecUserValidation(
  Sender: TObject; var ErrorCode: Word);
  var
    aMyFieldDef: TMyFieldDef;
    FldNr: integer;
begin
  with OvcTable1 do
    if ActiveRow > 0 then
      begin
        FldNr := ActiveRow - 1;
        aMyFieldDef := TMyFieldDef(MyFieldDefs[FldNr]);
        if (aMyFieldDef.DBFldType = fldDBNUM) OR
           (aMyFieldDef.DBFldType = fldDBFLOAT) then
          begin
            with sender as TOvcTCNumericFieldEdit do
              if (AsInteger < 0) or (AsInteger > aMyFieldDef.Field_Width) then
                begin
                  ErrorCode := oeOutOfRange;
                  OvcController1.ErrorText := Format('DEC must be >= 0 and <= %d',
                                                     [aMyFieldDef.Field_Width]);
                end
          end
        else
          with sender as TOvcTCNumericFieldEdit do
            if (AsInteger <> 0) then
              begin
                ErrorCode := oeOutOfRange;
                OvcController1.ErrorText := 'DEC must be 0';
                AsInteger := 0;
              end;
      end;
end;

procedure TModify_Table_Structure.OvcController1Error(Sender: TObject;
  ErrorCode: Word; const ErrorMsg: String);
begin
  case ErrorCode of
    oeOutOfRange,
    oeINVALIDFIELDNAME: Error(OvcController1.ErrorText);
  else
    Error(ErrorMsg);
  end;
end;

procedure TModify_Table_Structure.OvcTable1GetCellAttributes(
  Sender: TObject; RowNum, ColNum: Integer;
  var CellAttr: TOvcCellAttributes);
  var
    FldNr: integer;
    aMyFieldDef: TMyFieldDef;
begin { TModify_Table_Structure.OvcTable1GetCellAttributes }
  if RowNum > 0 then
    begin
      FldNr := RowNum - 1;
      aMyFieldDef := TMyFieldDef(MyFieldDefs[FldNr]);
      case ColNum of
        COLNR_FIELDNAME: ;
        COLNR_FIELDTYPE: ;
        COLNR_FIELDWIDTH:
          begin
            with CellAttr do
            case aMyFieldDef.DBFldType of
              fldDBCHAR,                           { Char string }
              fldDBNUM,                            { Number }
              fldDBFLOAT:                          { Float }
                begin
                  caAccess := otxNormal;
                  caFontHiColor := clHighLight;
                end;
            else
              begin
                caAccess := otxInvisible;
                caFontHiColor := clBtnFace;
              end;
            end;
          end;
        COLNR_FIELDDEC:
          begin
            with CellAttr do
              if (aMyFieldDef.DBFldType = fldDBNUM) OR
                 (aMyFieldDef.DBFldType = fldDBFLOAT) then
                  caAccess := otxNormal
              else
                caAccess := otxInvisible;
          end;
      end;
    end;
end;  { TModify_Table_Structure.OvcTable1GetCellAttributes }

procedure TModify_Table_Structure.OvcTable1DoneEdit(Sender: TObject;
  RowNum, ColNum: Integer);
  var
    aMyFieldDef: TMyFieldDef;
begin
  if RowNum > 0 then
    begin
      aMyFieldDef := TMyFieldDef(MyFieldDefs[RowNum-1]);
      with aMyFieldDef do
        case ColNum of
          COLNR_FIELDTYPE:
            DBFldType := Word(col_FieldType.Items.Objects[Type_Index]);
        end;
    end;
end;

procedure TModify_Table_Structure.col_FieldNameUserValidation(
  Sender: TObject; var ErrorCode: Word);
  var
    mode : tSEARCH_TYPE;
    i    : integer;
    f1   : string;
    f2   : string;
    ErrorMsg: string;
    AllowIt: boolean;
    RowNum: integer;
begin
  f2 := (Sender as TOvcTcSimpleFieldEdit).AsString;
  RowNum := OvcTable1.ActiveRow;
  AllowIt := ValidFieldName(f2, ift_dbase, 3, ErrorMsg);
  if not AllowIt then
    begin
      ErrorCode := oeINVALIDFIELDNAME;
      OvcController1.ErrorText := ErrorMsg;
    end
  else
    begin
      mode := SEARCHING;
      i := 0;
      repeat
        if i >= MyFieldDefs.Count then
          mode := NOT_FOUND else
        if i <> (RowNum-1) then  { ignore ourself }
          begin
            f1 := TMyFieldDef(MyFieldDefs.Items[i]).Field_Name;
            if CompareText(f1, f2) = 0 then
              mode := SEARCH_FOUND
            else
              inc(i)
          end
        else
          inc(i);
      until mode <> SEARCHING;
      AllowIt := mode <> SEARCH_FOUND;
      if not AllowIt then
        begin
          ErrorCode := oeINVALIDFIELDNAME;
          OvcController1.ErrorText := Format('Duplicate field name: %s', [f2]);
        end;
    end;
end;

procedure TModify_Table_Structure.Col_FieldWidthUserValidation(
  Sender: TObject; var ErrorCode: Word);
  var
    FldNr: integer;
    MaxWidth: integer;
    aMyFieldDef: TMyFieldDef;
begin
  with OvcTable1 do
    if ActiveRow > 0 then
      begin
        FldNr := ActiveRow - 1;
        aMyFieldDef := TMyFieldDef(MyFieldDefs[FldNr]);
        if (aMyFieldDef.DBFldType = fldDBCHAR) or
           (aMyFieldDef.DBFldType = fldDBNUM) then
          begin
            MaxWidth := 20;
            case aMyFieldDef.DBFldType of
              fldDBCHAR: MaxWidth := 254;
              fldDBNUM:  MaxWidth := 20;
            end;
            with sender as TOvcTCNumericFieldEdit do
              if (AsInteger < 1) or (AsInteger > MaxWidth) then
                begin
                  ErrorCode := oeOutOfRange;
                  OvcController1.ErrorText := Format('Width must be > 0 and <= %d',
                                                     [MaxWidth]);
                end
          end;
      end;
end;

end.
