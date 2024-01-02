{ ----------------------------------------------------- }
{ @Unit Description:  CSV Parser to DataSource          }
{ @Author:            Lukáš Vlček [xnukes@gmail.com]    }
{ @Licence:           GNU General Public Licence 3.0    }
{ ----------------------------------------------------- }
{ Contributor  : Christopohe Fontaine                   }
{ Improvements :                                        }
{ - Every object created wasn't freed                   }
{ - Replacememt of GetLocaleFormatSettings, wich is     }
{   deprecated in Delphi and not portable in FreePascal }
{ - Multiple load and save methods                      }
{ - Access fields by name or by index                   }
{ - Browse the csv in both directions                   }
{ - Add or delete a line, or clear all data             }
{ - New method to create and load at the same time      }
{ Developments planned soon :                           }
{ - Adding a TFormatSettings to configure all formats   }
{ - Adding a IndexFieldNames to sort the lines          }
{ ----------------------------------------------------- }

unit uCSVDataSource;

interface

uses
  Classes, SysUtils, Variants;

type

  { TCSVDataSource }

  TCSVDataSource = class(TObject)
    protected
      _columns: TStringList;
      _rows: TStringList;
      _delimiter: Char;
      _index: Integer;
      _feof: Boolean;
      _fbof: Boolean;
      _date_separator: Char;
    public
      constructor Create;
      constructor Create(FileName: ShortString; Delimiter: Char = ';');
      constructor Create(List: TStringList; Delimiter: Char = ';');
      destructor Destroy; override;

      procedure LoadFromFile(const FileName: ShortString);
      procedure LoadFromString(const Chaine: AnsiString);
      procedure LoadFromList(const List: TStringList);
      procedure LoadFromStream(const Stream: TStream);

      procedure SaveToFile(const FileName: ShortString);
      procedure SaveToList(const List: TStringList);
      procedure SaveToStream(const Stream: TStream);

      procedure SetDelimiter(const Character: Char);
      procedure SetDateSeparator(const Separator: Char);

      procedure First;
      procedure Last;
      procedure Next;
      procedure Previous;

      function Add(Row: String): Integer;
      function Add(Fields : array of Variant): Integer;
      procedure Delete;
      procedure Clear;

      function GetTotal: Integer;
      function GetRowString: String;

      function FieldAsString(Column: ShortString): String;
      function FieldAsString(Idx: Integer): String;
      function FieldAsInteger(Column: ShortString): Integer;
      function FieldAsInteger(Idx: Integer): Integer;
      function FieldAsFloat(Column: ShortString): Extended;
      function FieldAsFloat(Idx: Integer): Extended;
      function FieldAsDate(Column: ShortString): TDate;
      function FieldAsDate(Idx: Integer): TDate;
      function FieldAsTime(Column: ShortString): TTime;
      function FieldAsTime(Idx: Integer): TTime;
      function FieldAsDateTime(Column: ShortString): TDateTime;
      function FieldAsDateTime(Idx: Integer): TDateTime;

      function FieldByNameAsString(Column: ShortString): String;
      function FieldByNameAsInteger(Column: ShortString): Integer;
      function FieldByNameAsFloat(Column: ShortString): Extended;
      function FieldByNameAsDate(Column: ShortString): TDate;
      function FieldByNameAsTime(Column: ShortString): TTime;
      function FieldByNameAsDateTime(Column: ShortString): TDateTime;

      property Eof: Boolean read _feof;
      property Bof: Boolean read _fbof;
      property Count: Integer read GetTotal;
    private
      function GetColumnIndex(Column: ShortString): Integer;
      function GetHeaderRow: AnsiString;
      procedure Load(const loadedFile: TStringList);
      procedure TestLimits;
    published
  end;

implementation

constructor TCSVDataSource.Create;
begin
  Self._columns := TStringList.Create;
  Self._rows := TStringList.Create;
  Self._columns.Clear;
  Self._rows.Clear;
  Self._index := -1;
  Self._date_separator := DefaultFormatSettings.DateSeparator;
end;

constructor TCSVDataSource.Create(FileName: ShortString; Delimiter: Char);
begin
  Self.Create;
  Self.SetDelimiter(Delimiter);
  Self.LoadFromFile(FileName);
end;

destructor TCSVDataSource.Destroy;
begin
  Self._rows.Free;
  Self._columns.Free;
  inherited Destroy;
end;

constructor TCSVDataSource.Create(List: TStringList; Delimiter: Char);
begin
  Self.Create;
  Self.SetDelimiter(Delimiter);
  Self.LoadFromList(List);
end;

procedure TCSVDataSource.LoadFromFile(const FileName: ShortString);
var
  loadedFile: TStringList;
begin
  loadedFile := TStringList.Create;
  loadedFile.LoadFromFile(FileName);

  Self.Load(loadedFile);

  loadedFile.Free;
end;

procedure TCSVDataSource.LoadFromString(const Chaine: AnsiString);
var
  loadedFile: TStringList;
begin
  loadedFile := TStringList.Create;
  loadedFile.Text := Chaine;

  Self.Load(loadedFile);

  loadedFile.Free;
end;

procedure TCSVDataSource.LoadFromList(const List: TStringList);
begin
  Self.LoadFromString(List.Text);
end;

procedure TCSVDataSource.LoadFromStream(const Stream: TStream);
var
  Lst : TStringList;
begin
  Lst := TStringList.Create;
  Lst.LoadFromStream(Stream);
  Self.LoadFromList(lst);
  Lst.Free;
end;

procedure TCSVDataSource.SaveToFile(const FileName: ShortString);
var
  loadedFile : TStringList;
begin
  loadedFile := TStringList.Create;
  loadedFile.Text := Self._rows.Text;
  loadedFile.Insert(0, Self.GetHeaderRow);
  loadedFile.SaveToFile(FileName);
  loadedFile.Free;
end;

procedure TCSVDataSource.SaveToList(const List: TStringList);
begin
  List.Clear;
  List.Text := Self._rows.Text;
  List.Insert(0, Self.GetHeaderRow);
end;

procedure TCSVDataSource.SaveToStream(const Stream: TStream);
var
  List : TStringList;
begin
  List := TStringList.Create;
  Self.SaveToList(List);
  List.SaveToStream(Stream);
  List.Free;
end;

procedure TCSVDataSource.SetDelimiter(const Character: Char);
begin
  Self._delimiter := Character;
end;

procedure TCSVDataSource.SetDateSeparator(const Separator: Char);
begin
  Self._date_separator := Separator;
end;

procedure TCSVDataSource.First;
begin
  Self._index := 0;
  Self.TestLimits;
end;

procedure TCSVDataSource.Last;
begin
  Self._index := Self.GetTotal - 1;
  Self.TestLimits;
end;

function TCSVDataSource.GetTotal: Integer;
begin
  Result := Self._rows.Count;
end;

function TCSVDataSource.GetRowString: String;
begin
  Result := Self._rows.Strings[Self._index];
end;

function TCSVDataSource.FieldAsString(Column: ShortString): String;
var
  ColumnIndex: Integer;
  Row: TStringList;
begin
  ColumnIndex := Self.GetColumnIndex(Column);
  Row := TStringList.Create;
  Row.StrictDelimiter := True;
  Row.Delimiter := Self._delimiter;
  Row.DelimitedText := Self._rows.Strings[Self._index];
  Result := Row.Strings[ColumnIndex];
  Row.Free;
end;

function TCSVDataSource.FieldAsString(Idx: Integer): String;
var
  Row: TStringList;
begin
  Row := TStringList.Create;
  Row.StrictDelimiter := True;
  Row.Delimiter := Self._delimiter;
  Row.DelimitedText := Self._rows.Strings[Self._index];
  Result := Row.Strings[Idx];
  Row.Free;
end;

function TCSVDataSource.FieldAsInteger(Column: ShortString): Integer;
begin
  Result := StrToIntDef(Self.FieldAsString(Column), 0);
end;

function TCSVDataSource.FieldAsInteger(Idx: Integer): Integer;
begin
  Result := StrToIntDef(Self.FieldAsString(Idx), 0);
end;

function TCSVDataSource.FieldAsFloat(Column: ShortString): Extended;
begin
  Result := StrToFloatDef(Self.FieldAsString(Column), 0);
end;

function TCSVDataSource.FieldAsFloat(Idx: Integer): Extended;
begin
  Result := StrToFloatDef(Self.FieldAsString(Idx), 0);
end;

function TCSVDataSource.FieldAsDate(Column: ShortString): TDate;
var
  MySettings: TFormatSettings;
begin
  MySettings := DefaultFormatSettings;
  MySettings.DateSeparator := Self._date_separator;
  Result := StrToDate(Self.FieldAsString(Column), MySettings);
end;

function TCSVDataSource.FieldAsDate(Idx: Integer): TDate;
var
  MySettings: TFormatSettings;
begin
  MySettings := DefaultFormatSettings;
  MySettings.DateSeparator := Self._date_separator;
  Result := StrToDate(Self.FieldAsString(Idx), MySettings);
end;

function TCSVDataSource.FieldAsTime(Column: ShortString): TTime;
begin
  Result := StrToTime(Self.FieldAsString(Column));
end;

function TCSVDataSource.FieldAsTime(Idx: Integer): TTime;
begin
  Result := StrToTime(Self.FieldAsString(Idx));
end;

function TCSVDataSource.FieldAsDateTime(Column: ShortString): TDateTime;
var
  MySettings: TFormatSettings;
begin
  MySettings := DefaultFormatSettings;
  MySettings.DateSeparator := Self._date_separator;
  Result := StrToDateTime(Self.FieldAsString(Column), MySettings);
end;

function TCSVDataSource.FieldAsDateTime(Idx: Integer): TDateTime;
var
  MySettings: TFormatSettings;
begin
  MySettings := DefaultFormatSettings;
  MySettings.DateSeparator := Self._date_separator;
  Result := StrToDateTime(Self.FieldAsString(Idx), MySettings);
end;

function TCSVDataSource.GetColumnIndex(Column: ShortString): Integer;
var
  ColumnIndex: Integer;
begin
  ColumnIndex := Self._columns.IndexOf(Column);
  if ColumnIndex <> -1 then
    Result := ColumnIndex
  else
    raise Exception.Create('Error: Column "' + Column + '" not found !');
end;

function TCSVDataSource.GetHeaderRow: AnsiString;
var
  c : Integer;
begin
  Result := '';
  for c := 0 to Self._columns.Count-1 do
  begin
    if c > 0 then
      Result := Result + Self._delimiter;
    Result := Result + Self._columns[c];
  end;
end;

procedure TCSVDataSource.Load(const loadedFile: TStringList);
var
  Row: TStringList;
  I: Integer;
begin
  // Empty list
  Self.Clear;

  // load columns
  Row := TStringList.Create;
  Row.StrictDelimiter := True;
  Row.Delimiter := Self._delimiter;
  Row.DelimitedText := loadedFile.Strings[0]; // first row is column names

  for I := 0 to Row.Count -1 do
  begin
    Self._columns.Add(Row.Strings[I]);
  end;
  Row.Free;

  // load rows
  for I := 1 to loadedFile.Count - 1 do
  begin
    Self._rows.Add(loadedFile.Strings[I]);
  end;

  Self._index := 0;
  Self.TestLimits;
end;

procedure TCSVDataSource.TestLimits;
begin
  Self._feof := (Self._index = Self._rows.Count);
  Self._fbof := (Self._index < 0);
end;

function TCSVDataSource.FieldByNameAsString(Column: ShortString): String;
begin
  Result := Self.FieldAsString(Column);
end;

function TCSVDataSource.FieldByNameAsInteger(Column: ShortString): Integer;
begin
  Result := Self.FieldAsInteger(Column);
end;

function TCSVDataSource.FieldByNameAsFloat(Column: ShortString): Extended;
begin
  Result := Self.FieldAsFloat(Column);
end;

function TCSVDataSource.FieldByNameAsDate(Column: ShortString): TDate;
begin
  Result := Self.FieldAsDate(Column);
end;

function TCSVDataSource.FieldByNameAsTime(Column: ShortString): TTime;
begin
  Result := Self.FieldAsTime(Column);
end;

function TCSVDataSource.FieldByNameAsDateTime(Column: ShortString): TDateTime;
begin
  Result := Self.FieldAsDateTime(Column);
end;

procedure TCSVDataSource.Next;
begin
  Inc(Self._index);
  if Self._index = Self._rows.Count then
    Self._feof := True;
  Self.TestLimits;
end;

procedure TCSVDataSource.Previous;
begin
  Dec(Self._index);
  Self.TestLimits;
end;

function TCSVDataSource.Add(Row: String): Integer;
begin
  Self._index := Self._rows.Add(Row);
  Self.TestLimits;
  Result := Self._index;
end;

function TCSVDataSource.Add(Fields: array of Variant): Integer;
var
  F   : Variant;
  Row : String;
  i   : Integer;
begin
  Row := '';
  i := 0;
  for F in Fields do
  begin
    if i > 0 then
      Row := Row + Self._delimiter;
    Row := Row + VarToStr(F);
    Inc(i);
  end;
  Result := Self.Add(Row);
end;

procedure TCSVDataSource.Delete;
begin
  Self._rows.Delete(Self._index);
  Self.TestLimits;
end;

procedure TCSVDataSource.Clear;
begin
  Self._rows.Clear;
  Self._columns.Clear;
  Self.TestLimits;
end;

end.
