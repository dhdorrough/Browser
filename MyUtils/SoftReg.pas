unit SoftReg; {PB}

{$B-}

interface

uses
  Classes, ComCtrls, Controls, ExtCtrls, Forms, StdCtrls, SysUtils, Windows;

const
  SViewKey = 'View';

type
  ERegistryError = class(EWin32Error);
  
  TRootKey = (rkNone, rkLocal, rkUser);

  TKeyInfo = record
    NumSubKeys: Integer;
    MaxSubKeyLen: Integer;
    NumValues: Integer;
    MaxValueLen: Integer;
    MaxDataLen: Integer;
    FileTime: TFileTime;
  end;

  TOpenRootKeyEvent = procedure(Sender: TObject; ARootKey: TRootKey;
    var Key: HKEY; var DoDefault: Boolean) of object;

  TSoftReg = class
  private
    fCompanyName: string;
    fOnOpenRootKey: TOpenRootKeyEvent;
    fProductName: string;
    fVersion: string;
    fRootKey: TRootKey;
    fpValue: pByte;
    fValueSize: DWORD;
    fValueType: DWORD;
    procedure SetCompanyName(Value: string);
    procedure SetProductName(Value: string);
    procedure SetVersion(Value: string);
    procedure CheckKey(var Key: HKEY);
    function CreateKey(BaseKey: HKEY; Name: string): HKEY;
//  procedure CreateValues(SourceKey : HKEY; Subkey : string; DestinationKey : HKEY);
//  function TempKeyName : string;
  protected
    fHKeyRoot: HKEY;
    function DoOpenRootKey(ARootKey: TRootKey): HKEY; virtual;
  public
    constructor Create; virtual;
    destructor Destroy; override;
    procedure CloseKey(Key: HKEY);
    procedure CloseRootKey;
    function ComponentValueName(Component: TComponent): string;
    function DeleteKey(Key: HKEY; SubKey: string): Boolean;
    function DeleteValue(Key: HKEY; Name: string): Boolean;
    procedure GetKeyNames(Key: HKEY; Strings: TStrings);
    procedure GetValueNames(Key: HKEY; Strings: TStrings);
    procedure ClearKey(Key: HKEY);
//  procedure SaveKey(Key: HKEY; Filename: string);
//  procedure RestoreKey(Key: HKEY; Filename: string);
//  procedure CopyKey(SourceKey, DestinationKey: HKEY);
    function KeyInfo(Key: HKEY): TKeyInfo;
    function OpenKey(BaseKey: HKEY; Name: string; CanCreate: Boolean): HKEY;
    procedure OpenRootKey(ARootKey: TRootKey); virtual;
    function ReadBinary(Key: HKEY; Name: string; var Value; Size: DWORD): Boolean;
    function ReadBinaryEx(Key: HKEY; Name: string; var Value; Size: DWORD; Initialize: Boolean): Boolean;
    procedure WriteBinary(Key: HKEY; Name: string; const Value; Size: DWORD);
//  function ReadComponent(Key: HKEY; Name: string; Component: TComponent; Clear: Boolean): Boolean;
//  procedure WriteComponent(Key: HKEY; Name: string; Component: TComponent);
    function ReadHandle(Key: HKEY; Name: string): THandle;
    procedure WriteHandle(Key: HKEY; Name: string; Value: THandle);
    function ReadInteger(Key: HKEY; Name: string; Default: Integer): Integer;
    procedure WriteInteger(Key: HKEY; Name: string; Value: Integer);
    function ReadBoolean(Key: HKEY; Name: string; Default: Boolean): Boolean;
    procedure WriteBoolean(Key: HKEY; Name: string; Value: Boolean);
//  function ReadPath(Key: HKEY; Name: string; Default: string; CanCreate: Boolean): string;
    procedure WritePath(Key: HKEY; Name: string; Value: string);
    function ReadString(Key: HKEY; Name: string; Default: string): string;
    procedure WriteString(Key: HKEY; Name: string; Value: string);
    function ReadStrings(Key: HKEY; Name: string; Strings: TStrings): Boolean;
    procedure WriteStrings(Key: HKEY; Name: string; Strings: TStrings);
//  function ReadWindowText(Key: HKEY; Control: TWinControl): string;
//  procedure WriteWindowText(Key: HKEY; Control: TWinControl);
//  function ReadForm(Key: HKEY; Form: TForm): Boolean; overload;
//  function ReadFormEx(Key: HKEY; Form: TForm; Maximized, NoMaximize: Boolean): Boolean;
//  function ReadForm(Key: HKEY; Form: TForm; Options: TRegWindowOptions): Boolean; overload; {CR 9829 - PB}
//  procedure WriteForm(Key: HKEY; Form: TForm);
//  function ReadListViewWidths(Key: HKEY; Wnd: HWND; ValueName: string): Boolean; overload; {CR 9706 - PB}
//  function ReadListViewWidths(Key: HKEY; ListView: TListView; ValueName: string): Boolean; overload; {CR 9631 - PB}
//  function ReadListViewWidths(Key: HKEY; ListView: TListView): Boolean; overload;
//  procedure WriteListViewWidths(Key: HKEY; Wnd: HWND; ValueName: string); overload; {CR 9706 - PB}
//  procedure WriteListViewWidths(Key: HKEY; ListView: TListView; ValueName: string); overload; {CR 9631 - PB}
//  procedure WriteListViewWidths(Key: HKEY; ListView: TListView); overload;
    procedure WriteVariant(Key: HKEY; Name: string; Value: Variant);
    function ReadVariant(Key: HKEY; Name: string; VariantType : Integer; Default: Variant): Variant;
    property RootKey: TRootKey read fRootKey;
    property OnOpenRootKey: TOpenRootKeyEvent read fOnOpenRootKey write fOnOpenRootKey;
  published
    property CompanyName: string read fCompanyName write SetCompanyName;
    property ProductName: string read fProductName write SetProductName;
    property Version: string read fVersion write SetVersion;
  end;

implementation

uses
  CommCtrl, Registry, {BccSoftwareDeclX, PathUtilsX,
  BccControlUtils, BccExceptions, BCCPersistObject, BccRegistryComCtrls,
  BccSoftwareConstX, BccUtils, Files, FileUtilsX,} Variants;

const
  REGSTR_MAX_VALUE_LENGTH  = 256;
  HEAP_GENERATE_EXCEPTIONS = 4;
  HEAP_ZERO_MEMORY         = 8;

  SErrorSavingKey          = 'Error saving registry key to the file, %s';
  SErrorLoadingKey         = 'Error loading registry key from the file, %s';
  SErrorCopyingKey         = 'Error copying registry key';
  SCreateValuesError       = 'Error creating registry values';

constructor TSoftReg.Create;
begin
//fCompanyName:= BCCCompanyNameKey;
//fProductName:= BCCProductNameKey; // Application.Title;
end;

destructor TSoftReg.Destroy;
begin
  CloseRootKey;
end;

procedure TSoftReg.CloseKey(Key: HKEY);
begin
  if Key<>0 then
    RegCloseKey(Key);
end;

procedure TSoftReg.CloseRootKey;
begin
  if fHKeyRoot<>0 then
    RegCloseKey(fHKeyRoot);
  fHKeyRoot:= 0;
  fRootKey:= rkNone;
end;

function TSoftReg.ComponentValueName(Component: TComponent): string;

  function ComponentName(Component: TComponent; IsOwner: Boolean): string;
  var
    ClassName: string;
  begin
    if not Assigned(Component) then
      Result:= ''
    else
    begin
      ClassName:= Component.ClassName;
      if Copy(ClassName, 1, 1)='T' then
        ClassName:= Copy(ClassName, 2, Length(ClassName));
      if IsOwner or (Component is TCustomForm) then
        Result:= ClassName
      else
      begin
        Result:= Component.Name;
        if Result='' then
          Result:= ClassName
        else
          if (Pos(ClassName, Result)=1) and (Length(ClassName)<Length(Result)) then
            Result:= Copy(Result, Succ(Length(ClassName)), Length(Result));
      end;
    end;
  end;

begin
  Result:= ComponentName(Component, False);
  if ComponentName(Component.Owner, True)<>'' then
    Result:= ComponentName(Component.Owner, True)+'.'+Result;
end;

function TSoftReg.DeleteKey(Key: HKEY; SubKey: string): Boolean;
var
  Index: Integer;
  HSubKey: HKEY;
  KeyNames: TStrings;
begin
  Result:= False;
  CheckKey(Key);
  HSubKey:= OpenKey(Key, SubKey, False);
  if HSubKey<>0 then
  try
    KeyNames:= TStringList.Create;
    try
      Result:= True;
      GetKeyNames(HSubKey, KeyNames);
      with KeyNames do
        for Index:= 0 to Pred(Count) do
          Result:= Result and DeleteKey(HSubKey, Strings[Index]);
    finally
      KeyNames.Free;
    end;
  finally
    CloseKey(HSubKey);
  end;
  Result:= Result and (RegDeleteKey(Key, pChar(SubKey))=ERROR_SUCCESS);
end;

function TSoftReg.DeleteValue(Key: HKEY; Name: string): Boolean;
begin
  CheckKey(Key);
  Result:= RegDeleteValue(Key, pChar(Name))=ERROR_SUCCESS;
end;

procedure TSoftReg.GetKeyNames(Key: HKEY; Strings: TStrings);
var
  Index: Integer;
  Len: DWORD;
  S: string;
begin
  with Strings do
  begin
    BeginUpdate;
    try
      Clear;
      CheckKey(Key);
      with KeyInfo(Key) do
      begin
        SetString(S, nil, Succ(MaxSubKeyLen));
        for Index:= 0 to Pred(NumSubKeys) do
        begin
          Len:= Succ(MaxSubKeyLen);
          RegEnumKeyEx(Key, Index, pChar(S), Len, nil, nil, nil, nil);
          Add(pChar(S));
        end;
      end;
    finally
      EndUpdate;
    end;
  end;
end;

procedure TSoftReg.GetValueNames(Key: HKEY; Strings: TStrings);
var
  Index: Integer;
  Len: DWORD;
  S: string;
begin
  with Strings do
  begin
    Clear;
    BeginUpdate;
    try
      CheckKey(Key);
      with KeyInfo(Key) do
      begin
        SetString(S, nil, Succ(MaxValueLen));
        for Index:= 0 to Pred(NumValues) do
        begin
          Len:= Succ(MaxValueLen);
          RegEnumValue(Key, Index, pChar(S), Len, nil, nil, nil, nil);
          Add(pChar(S));
        end;
      end;
    finally
      EndUpdate;
    end;
  end;
end;

function TSoftReg.KeyInfo(Key: HKEY): TKeyInfo;
begin
  FillChar(Result, SizeOf(Result), 0);
  CheckKey(Key);
  with Result do
    RegQueryInfoKey(Key, nil, nil, nil, @NumSubKeys, @MaxSubKeyLen, nil, @NumValues, @MaxValueLen, @MaxDataLen, nil, @FileTime);
end;

function TSoftReg.OpenKey(BaseKey: HKEY; Name: string; CanCreate: Boolean): HKEY;
begin
  Result:= 0;
  if CanCreate then
    Result:= CreateKey(BaseKey, Name)
  else
  begin
    CheckKey(BaseKey);
    SetLastError(RegOpenKey(BaseKey, pChar(Name), Result));
  end;
end;

procedure TSoftReg.OpenRootKey(ARootKey: TRootKey);
begin
  if fRootKey<>ARootKey then
  begin
    CloseRootKey;
    fHKeyRoot:= DoOpenRootKey(ARootKey);
    fRootKey:= ARootKey;
  end;
end;

function TSoftReg.DoOpenRootKey(ARootKey: TRootKey): HKEY;
var
  BaseKey: HKEY;
  RootName: string;
  DoDefault: Boolean;
begin
  Result:= 0;
  DoDefault:= True;
  if Assigned(OnOpenRootKey) then
    OnOpenRootKey(Self, ARootKey, Result, DoDefault);
  if DoDefault then
  begin
    case ARootKey of
      rkLocal: BaseKey:= HKEY_LOCAL_MACHINE;
      rkUser: BaseKey:= HKEY_CURRENT_USER;
    else
      BaseKey:= 0;
    end;
    if BaseKey<>0 then
    begin
      RootName:= 'Software\'+CompanyName;
      if ProductName<>'' then
        RootName:= RootName+'\'+ProductName;
      if Version<>'' then
        RootName:= RootName+'\'+Version;
      Result:= CreateKey(BaseKey, RootName);
    end;
  end;
end;

procedure TSoftReg.SetCompanyName(Value: string);
begin
  if fCompanyName<>Value then
  begin
    CloseRootKey;
    fCompanyName:= Value;
  end;
end;

procedure TSoftReg.SetProductName(Value: string);
begin
  if fProductName<>Value then
  begin
    CloseRootKey;
    fProductName:= Value;
  end;
end;

procedure TSoftReg.SetVersion(Value: string);
begin
  if fVersion<>Value then
  begin
    CloseRootKey;
    fVersion:= Value;
  end;
end;

procedure TSoftReg.CheckKey(var Key: HKEY);
begin
  if Key=0 then
    Key:= fHKeyRoot;
end;

function TSoftReg.CreateKey(BaseKey: HKEY; Name: string): HKEY;
begin
  Result:= 0;
  CheckKey(BaseKey);
  SetLastError(RegCreateKey(BaseKey, pChar(Name), Result));
end;

function TSoftReg.ReadBinary(Key: HKEY; Name: string; var Value;
  Size: DWORD): Boolean;
begin
  Result:= ReadBinaryEx(Key, Name, Value, Size, True);
end;

function TSoftReg.ReadBinaryEx(Key: HKEY; Name: string; var Value; Size: DWORD;
  Initialize: Boolean): Boolean;
var
  Handle: THandle;
begin
  Result:= False;
  if Initialize then
    FillChar(Value, Size, 0);
  Handle:= ReadHandle(Key, Name);
  if Handle<>0 then
  try
    fpValue:= GlobalLock(Handle);
    if Assigned(fpValue) then
    try
      if GlobalSize(Handle)>Size then
        MoveMemory(@Value, fpValue, Size)
      else
        MoveMemory(@Value, fpValue, GlobalSize(Handle));
      Result:= True;
    finally
      GlobalUnlock(Handle);
    end;
  finally
    GlobalFree(Handle);
  end;
end;

procedure TSoftReg.WriteBinary(Key: HKEY; Name: string; const Value;
  Size: DWORD);
begin
  CheckKey(Key);
  RegSetValueEx(Key, pChar(Name), 0, REG_BINARY, @Value, Size);
end;

(*
function TSoftReg.ReadComponent(Key: HKEY; Name: string;
  Component: TComponent; Clear: Boolean): Boolean;
var
  Stream: TMemoryStream;
begin
  Result:= False;
  CheckKey(Key);
  fValueSize:= 0;
  if (RegQueryValueEx(Key, pChar(Name), nil,
    @fValueType, nil, @fValueSize)=ERROR_SUCCESS) and (fValueType=REG_BINARY) then
  begin
    Stream:= TMemoryStream.Create;
    try
      Stream.Size:= fValueSize;
      Result:= RegQueryValueEx(Key, pChar(Name), nil, @fValueType, Stream.Memory, @fValueSize)=ERROR_SUCCESS;
      if Result then
      try
        BCCReadComponent(Stream, Component, Clear);
      except
        on EStreamError do;
      else
        raise;
      end;
    finally
      Stream.Free;
    end;
  end;
end;
*)

(*
procedure TSoftReg.WriteComponent(Key: HKEY; Name: string; Component: TComponent);
begin
  with TMemoryStream.Create do
  try
    WriteComponent(Component);
    WriteBinary(Key, Name, pChar(Memory)^, Size);
  finally
    Free;
  end;
end;
*)

function TSoftReg.ReadHandle(Key: HKEY; Name: string): THandle;
begin
  Result:= 0;
  CheckKey(Key);
  if (RegQueryValueEx(Key, pChar(Name), nil,
     @fValueType, nil, @fValueSize)=ERROR_SUCCESS) and (fValueType=REG_BINARY) then
  begin
    Result:= GlobalAlloc(GHND, fValueSize);
    if Result<>0 then
    try
      fpValue:= GlobalLock(Result);
      if Assigned(fpValue) then
      try
        RegQueryValueEx(Key, pChar(Name), nil, @fValueType, fpValue, @fValueSize);
      finally
        GlobalUnlock(Result);
      end;
    except
      GlobalFree(Result);
      raise;
    end;
  end;
end;

procedure TSoftReg.WriteHandle(Key: HKEY; Name: string; Value: THandle);
begin
  fpValue:= GlobalLock(Value);
  if Assigned(fpValue) then
  try
    WriteBinary(Key, Name, fpValue, GlobalSize(Value));
  finally
    GlobalUnlock(Value);
  end;
end;

function TSoftReg.ReadInteger(Key: HKEY; Name: string; Default: Integer): Integer;
begin
  CheckKey(Key);
  fValueSize:= SizeOf(Result);
  if not((RegQueryValueEx(Key, pChar(Name), nil, @fValueType, @Result, @fValueSize)=ERROR_SUCCESS) and (fValueType=REG_DWORD)) then
    Result:= Default;
end;

procedure TSoftReg.WriteInteger(Key: HKEY; Name: string; Value: Integer);
begin
  CheckKey(Key);
  RegSetValueEx(Key, pChar(Name), 0, REG_DWORD, @Value, SizeOf(DWORD));
end;

function TSoftReg.ReadBoolean(Key: HKEY; Name: string; Default: Boolean): Boolean;
begin
  Result:= Boolean(ReadInteger(Key, Name, Ord(Default)));
end;

procedure TSoftReg.WriteBoolean(Key: HKEY; Name: string; Value: Boolean);
begin
  WriteInteger(Key, Name, Integer(Value));
end;

(*
function TSoftReg.ReadPath(Key: HKEY; Name: string; Default: string;
  CanCreate: Boolean): string;
begin
  Result := LongPathName(ReadString(Key, Name, Default));
  if Result<>'' then
  begin
    Result:= BccIncludeTrailingPathDelimiter(Result);
    if CanCreate then
      BccCreateDirectories(Result);
    WritePath(Key, Name, Result);
  end;
end;
*)

procedure TSoftReg.WritePath(Key: HKEY; Name: string; Value: string);
begin
  WriteString(Key, Name, Value);
end;

function TSoftReg.ReadString(Key: HKEY; Name: string; Default: string): string;
begin
  Result:= Default;
  CheckKey(Key);
  fValueSize:= 0;
  if (RegQueryValueEx(Key, pChar(Name), nil,
    @fValueType, nil, @fValueSize)=ERROR_SUCCESS) and (fValueType=REG_SZ) then
  begin
    GetMem(fpValue, fValueSize);
    try
      if RegQueryValueEx(Key, pChar(Name), nil,
        @fValueType, fpValue, @fValueSize)=ERROR_SUCCESS then
          Result:= pChar(fpValue)
      else
        Result:= Default;
    finally
      FreeMem(fpValue);
    end;
  end;
end;

procedure TSoftReg.WriteString(Key: HKEY; Name: string; Value: string);
begin
  CheckKey(Key);
  RegSetValueEx(Key, pChar(Name), 0, REG_SZ, pChar(Value), Succ(Length(Value)));
end;

procedure TSoftReg.WriteVariant(Key: HKEY; Name: string; Value: Variant);
var
  aDouble : Double;
  aDate   : TDateTime;
begin
  case VarType(Value) of
    varEmpty,
    varNull      : DeleteValue(Key, Name);
    varSmallInt,
    varInteger,
    varByte      : WriteInteger(Key, Name, Value);
    varSingle,
    varDouble,
    varCurrency  : begin
                     aDouble := Value;
                     WriteBinary(Key, Name, aDouble, SizeOf(aDouble));
                   end;
    varDate      : begin
                     aDate := Value;
                     WriteBinary(Key, Name, aDate, SizeOf(aDate));
                   end;
    varOleStr,
    varString    : WriteString(Key, Name, Value);
    varBoolean   : WriteBoolean(Key, Name, Value);
  end;
end;

function TSoftReg.ReadVariant(Key: HKEY; Name: string; VariantType : Integer; Default: Variant): Variant;
var
  aDouble : Double;
  aDate   : TDateTime;
begin
  try
    Result                        := Default;
    CheckKey(Key);
    fValueSize                    := 0;
    if (RegQueryValueEx(Key, pChar(Name), nil, @fValueType, nil, @fValueSize) = ERROR_SUCCESS) then
      begin
        case VariantType of
          varString  : if (fValueType = REG_SZ) then
                         Result   := ReadString(Key, Name, Default)
                       else
                         DeleteValue(Key, Name);
          varBoolean : if (fValueType = REG_DWORD) then
                         Result   := ReadBoolean(Key, Name, Default)
                       else
                         DeleteValue(Key, Name);
          varInteger : if (fValueType = REG_DWORD) then
                         Result   := ReadInteger(Key, Name, Default)
                       else
                         DeleteValue(Key, Name);
          varDouble  : if (fValueType = REG_BINARY) then
                         if ReadBinary(Key, Name, aDouble, SizeOf(aDouble)) then
                           Result := aDouble
                         else
                           Result := Default
                       else
                         DeleteValue(Key, Name);
          varDate    : if (fValueType = REG_BINARY) then
                         if ReadBinary(Key, Name, aDate, SizeOf(aDate)) then
                           Result := aDate
                         else
                           Result := Default
                       else
                         DeleteValue(Key, Name);
        end;
      end;
  except
    Result                        := Default;
  end;
end;

function TSoftReg.ReadStrings(Key: HKEY; Name: string; Strings: TStrings): Boolean;
var
  P, pLine: pChar;
  Text: string;
begin
  Result:= False;
  CheckKey(Key);
  fValueSize:= 0;
  if (RegQueryValueEx(Key, pChar(Name), nil,
    @fValueType, nil, @fValueSize)=ERROR_SUCCESS) and (fValueType=REG_MULTI_SZ) then
  begin
    GetMem(fpValue, fValueSize);
    try
      if RegQueryValueEx(Key, pChar(Name), nil,
        @fValueType, fpValue, @fValueSize)=ERROR_SUCCESS then
      begin
        Strings.Clear;
        P:= pChar(fpValue);
        while P^<>#0 do
        begin
          pLine:= P;
          while p^<>#0 do Inc(P);
          SetString(Text, pLine, P-pLine);
          Strings.Add(Text);
          Inc(P);
        end;
        Result:= True;
      end;
    finally
      FreeMem(fpValue);
    end;
  end;
end;

procedure TSoftReg.WriteStrings(Key: HKEY; Name: string; Strings: TStrings);
var
  Text: string;
  Index: Integer;
begin
  CheckKey(Key);
  fValueSize:= 1;
  for Index:= 0 to Pred(Strings.Count) do
  begin
    Text:= Text+Strings[Index]+#0;
    Inc(fValueSize, Succ(Length(Strings[Index])));
  end;
  Text:= Text+#0;
  RegSetValueEx(Key, pChar(Name), 0,
    REG_MULTI_SZ, pChar(Text), fValueSize);
end;

(*
function TSoftReg.ReadWindowText(Key: HKEY; Control: TWinControl): string;
var
  i             : integer;
  IndexStrList  : TStringList;
begin
  if Control is TListBox then
  begin
    IndexStrList := TStringList.Create;
    try
      ReadStrings(Key, ComponentValueName(Control), IndexStrList);
      with TListBox(Control) do
      begin
        for i:=0 to (IndexStrList.Count - 1) do
          if (Items.IndexOf(IndexStrList[i]) <> -1) then
            Selected[Items.IndexOf(IndexStrList[i])] := TRUE;
      end;
    finally
      IndexStrList.Free;
    end;
  end
  else
  begin
    Result:= ReadString(Key,
      ComponentValueName(Control), WindowText(Control.Handle));
    if Result <> '' then
    begin
      if Control is TComboBox then
      begin
        with TComboBox(Control) do
        begin
          if Style=csDropDownList then
          begin
            if Items.IndexOf(Result)>=0 then
              ItemIndex:= Items.IndexOf(Result);
          end
          else
            Text:= Result;
        end;
      end
      else
        SetWindowText(Control.Handle, pChar(Result));
    end;
  end;
end;
*)

(*
procedure TSoftReg.WriteWindowText(Key: HKEY; Control: TWinControl);
var
  i             : integer;
  IndexStrList  : TStringList;
begin
  if Control is TListBox then
  begin
    with TListBox(Control) do
    begin
      IndexStrList := TStringList.Create;
      try
        for i:=0 to (Items.Count - 1) do
        begin
          if Selected[i] then
            IndexStrList.Add(Items[i]);
        end;
        WriteStrings(Key, ComponentValueName(Control), IndexStrList{.Text});
      finally
        IndexStrList.Free;
      end;
    end;
  end
  else
    WriteString(Key, ComponentValueName(Control), WindowText(Control.Handle));
end;
*)

(*
function TSoftReg.ReadForm(Key: HKEY; Form: TForm): Boolean;
begin
  Result:= ReadFormEx(Key, Form, False, False);
end;
*)

(*
function TSoftReg.ReadFormEx(Key: HKEY; Form: TForm;
  Maximized, NoMaximize: Boolean): Boolean; {CR 9829 - PB}
var
  Options: TRegWindowOptions;
begin
  Options:= [];
  if Maximized then
    Include(Options, rwoDefaultMaximize);
  if NoMaximize then
    Include(Options, rwoNoMaximize);
  Result:= ReadForm(Key, Form, Options);
end;
*)

(*
function TSoftReg.ReadForm(Key: HKEY; Form: TForm; Options: TRegWindowOptions): Boolean;
begin
  Result:= RegReadWindow(Key, ComponentValueName(Form), Form.Handle, Options);
end;

procedure TSoftReg.WriteForm(Key: HKEY; Form: TForm);
begin
  RegWriteWindow(Key, ComponentValueName(Form), Form.Handle);
end;
*)

(*
function TSoftReg.ReadListViewWidths(Key: HKEY; Wnd: HWND;
  ValueName: string): Boolean;
begin
  Result:= RegReadListView(Key, ValueName, Wnd);
end;

function TSoftReg.ReadListViewWidths(Key: HKEY; ListView: TListView): Boolean;
begin
  Result:= ReadListViewWidths(Key, ListView, ComponentValueName(ListView));
end;

function TSoftReg.ReadListViewWidths(Key: HKEY; ListView: TListView;
  ValueName: string): Boolean;
begin
  Result:= ReadListViewWidths(Key, ListView.Handle, ValueName);
end;

procedure TSoftReg.WriteListViewWidths(Key: HKEY; Wnd: HWND;
  ValueName: string);
begin
  RegWriteListView(Key, ValueName, Wnd);
end;

procedure TSoftReg.WriteListViewWidths(Key: HKEY; ListView: TListView;
  ValueName: string); {CR 9631 - PB}
begin
  WriteListViewWidths(Key, ListView.Handle, ValueName);
end;

procedure TSoftReg.WriteListViewWidths(Key: HKEY; ListView: TListView);
begin
  WriteListViewWidths(Key, ListView, ComponentValueName(ListView));
end;
*)

(*
procedure TSoftReg.SaveKey(Key: HKEY; Filename: string); {spb}
var
  Err : DWORD;
begin               {converts long path name to short path name but requires short filename}
  CheckKey(Key);
  Err := RegSaveKey(Key, PChar(ShortPathName(Filename)), nil);
  if (Err <> ERROR_SUCCESS) then
  begin
    SetLastError(Err);
    BccRaiseLastWin32Error(ERegistryError, Format(SErrorSavingKey, [FileName]));
  end;
end;
*)

(*
function TSoftReg.TempKeyName : string;
begin
  Result := TempFilename('', '');
  SysUtils.DeleteFile(Result);
  Result := ExtractFilename(Result);
  Result := ChangeFileExt(Result, '');
end;
*)

procedure TSoftReg.ClearKey(Key: HKEY);
var
  List  : TStringList;
  index : Integer;
begin
  List := TStringList.Create;
  try
    GetKeyNames(Key, List);
    for Index := 0 to Pred(List.Count) do
      DeleteKey(Key, List[Index]);
    List.Clear;
    GetValueNames(Key, List);
    for Index := 0 to Pred(List.Count) do
      DeleteValue(Key, List[Index]);
  finally
    List.Free;
  end;
end;

(*
procedure TSoftReg.RestoreKey(Key: HKEY; Filename: string);
var
  ErrLoad    : DWORD;
  TempKey    : HKEY;
  TempName   : string;
  BackupName : string;
  BackupKey  : HKEY;
begin               {converts long path name to short path name but requires short filename}
  TempName          := TempKeyName;
  ErrLoad           := RegLoadKey(HKEY_USERS, PChar(TempName), PChar(ShortPathName(Filename))); {load saved hive}
  if (ErrLoad <> ERROR_SUCCESS) then
  begin
    SetLastError(ErrLoad);
    BccRaiseLastWin32Error(ERegistryError, Format(SErrorLoadingKey, [FileName]));
  end;
  try
    TempKey         := OpenKey(HKEY_USERS, PChar(TempName), false);  {open saved hive}
    if (TempKey <> 0) then
      try
        CheckKey(Key);
        BackupName  := TempKeyName;
        BackupKey   := OpenKey(HKEY_USERS, PChar(BackupName), true); {create backup key}
        try
          CopyKey(Key, BackupKey);                                   {backup current key}
          try
            ClearKey(Key);                                           {clear current key}
            CopyKey(TempKey, Key);                                   {copy saved hive to current key}
          except
            ClearKey(Key);                                           {clear current key}
            CopyKey(BackupKey, Key);                                 {Restore backedup hive}
            raise;
          end;
        finally
          CloseKey(BackupKey);
          BackupKey := OpenKey(HKEY_USERS, '', false);
          if (BackupKey <> 0) then
            try
              DeleteKey(BackupKey, BackupName);                      {delete backup hive}
            finally
              CloseKey(Backupkey);
            end;
        end;
      finally
        CloseKey(TempKey);
      end;
  finally
    RegUnloadKey(HKEY_USERS, PChar(TempName));                       {unload saved hive}
  end;
end;
*)

(*
procedure TSoftReg.CopyKey(SourceKey, DestinationKey: HKEY);
var
 dwSubKeyLength : DWORD;
  dwKeyIndex     : DWORD;
  szSubKey       : PChar;
  Err            : DWORD;
  hNewKey        : HKEY;
  hRtnKey        : HKEY;
begin
  dwKeyIndex             := 0;
  szSubKey               := StrAlloc(REGSTR_MAX_VALUE_LENGTH);
  try
    Err                := ERROR_SUCCESS;
    while true do
      begin
        dwSubKeyLength := REGSTR_MAX_VALUE_LENGTH;
        Err            := RegEnumKey(SourceKey, dwKeyIndex, szSubKey, dwSubKeyLength);
        case Err of
          ERROR_NO_MORE_ITEMS : begin
                                  Err := ERROR_SUCCESS;
                                  Break;
                                end;
          ERROR_SUCCESS       : begin
                                  Err := RegCreateKey(DestinationKey, szSubKey, hNewKey);
                                  case Err of
                                    ERROR_SUCCESS : try
                                                      CreateValues(SourceKey, szSubKey, hNewKey);
                                                      Err := RegOpenKey(SourceKey, szSubKey, hRtnKey);
                                                      case Err of
                                                        ERROR_SUCCESS : try
                                                                          CopyKey(hRtnKey, hNewKey);
                                                                        finally
                                                                          CloseHandle(hRtnKey);
                                                                        end;
                                                        else            Break;
                                                      end;
                                                    finally
                                                      CloseHandle(hNewKey);
                                                    end;
                                    else            Break;
                                  end;
                                end;
          else                  Break;
        end;
        Inc(dwKeyIndex);
      end;
    if (Err <> ERROR_SUCCESS) then
    begin
      SetLastError(Err);
      BccRaiseLastWin32Error(ERegistryError, SErrorCopyingKey);
    end;
  finally
    StrDispose(szSubKey);
  end;
end;
*)

(*
procedure TSoftReg.CreateValues(SourceKey : HKEY; Subkey : string; DestinationKey : HKEY);
var
  cbValue         : DWORD;
  dwType          : DWORD;
  cdwBuf          : DWORD;
  i               : DWORD;
  szValue         : PChar;
  TempKey         : HKEY;
  lRet            : DWORD;
  pBuf            : PByte;
  TempKeyInfo     : TKeyInfo;
begin
  lRet            := RegOpenKey(SourceKey, PChar(Subkey), Tempkey);
  if (lRet <> ERROR_SUCCESS) then
  begin
    SetLastError(lRet);
    BccRaiseLastWin32Error(ERegistryError, SCreateValuesError);
  end
  else
    try
      TempKeyInfo := KeyInfo(TempKey);
      if (TempKeyInfo.NumValues > 0) then
        begin
          szValue          := StrAlloc(Succ(TempKeyInfo.MaxValueLen));
          try
            pBuf           := HeapAlloc(GetProcessHeap, HEAP_GENERATE_EXCEPTIONS or HEAP_ZERO_MEMORY, Succ(TempKeyInfo.MaxDataLen));
            try
              for i := 0 to Pred(TempKeyinfo.NumValues) do
                begin
                  cbValue  := Succ(TempKeyInfo.MaxValueLen);
                  cdwBuf   := Succ(TempKeyInfo.MaxDataLen);
                  lRet     := RegEnumValue(Tempkey, i, szValue, cbValue, nil, @dwType, pBuf, @cdwBuf);
                  if (lRet <> ERROR_SUCCESS) then
                  begin
                    SetLastError(lRet);
                    BccRaiseLastWin32Error(ERegistryError, SCreateValuesError);
                  end
                  else
                    begin
                      lRet := RegSetValueEx(DestinationKey, szValue, 0, dwType, Pointer(pBuf), cdwBuf);
                      if (lRet <> ERROR_SUCCESS) then
                      begin
                        SetLastError(lRet);
                        BccRaiseLastWin32Error(ERegistryError, SCreateValuesError);
                      end;
                    end;
                end;
            finally
              HeapFree(GetProcessHeap, 0 , pBuf);
            end;
          finally
            StrDispose(szValue);
          end;
        end;
    finally
      CloseHandle(TempKey);
    end;
end;
*)

end.
