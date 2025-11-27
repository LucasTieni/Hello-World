unit MockHelpers;

interface

uses
  OracleData, System.JSON, System.SysUtils, System.Classes, JvDBLookup, Vcl.StdCtrls;
procedure CreateSQLUtil(aUserName, aPassword, aAlias: string);
procedure FreeSQLUtil;
procedure PopulateDataSet(aDataset: TOracleDataSet; aFieldsAndData: string);
procedure OpenDataSet(aDataset: TOracleDataSet; aCustomSQL: string);
procedure ReplaceComponentWithTEdit(var aComponent: TComponent; aText: string ='');
function FindStringInFile(aFileName, aSearchString: string): Boolean;
implementation
uses
  LoginUnit, DFMCommon, DFMSession, UELSqlUtils;


procedure CreateSQLUtil(aUserName, aPassword, aAlias: string);
begin
  if (not Assigned(gSqlUtil)) and (gSqlUtil = nil) then
  begin
    FreeAndNil(gSqlUtil);
  end;

  if (not Assigned(FRM_Login)) and (FRM_Login = nil) then
  begin
    FRM_Login := TFRM_Login.Create(nil);
  end
  else
    FRM_Login.MainSession.Connected := False;

  FRM_Login.MainSession.LogonUsername := aUserName;
  FRM_Login.MainSession.LogonPassword := aPassword;
  FRM_Login.MainSession.LogonDatabase := aAlias;
  USERID := aUserName;
  FRM_Login.MainSession.Connected := True;
  gSqlUtil := TDfmSqlUtil.Create(Frm_Login.MainSession);
end;

procedure FreeSQLUtil;
begin
  if (Assigned(gSqlUtil)) or (gSqlUtil <> nil) then
    FreeAndNil(gSqlUtil);

  if (Assigned(FRM_Login)) or (FRM_Login <> nil) then
    FreeAndNil(FRM_Login);

end;

procedure PopulateDataSet(aDataset: TOracleDataSet; aFieldsAndData: string);
var
  jsonObject: TJSONObject;
  fieldsArray, dataArray: TJSONArray;
  fieldNames, fieldTypes: TStringList;
  i, j: Integer;
  rowArray: TJSONArray;
  rowValues: TStringList;
  jsonValue: TJSONValue;
  finalSQL, RowSQL: string;
begin
  {$REGION 'Sample JSON'}
  (*
    {
      "fields": [
        { "name": "ID", "type": "ftInteger" },
        { "name": "AgentName", "type": "ftString", "size": 50 },
        { "name": "IsActive", "type": "ftBoolean" },
        { "name": "Rating", "type": "ftFloat" }
      ],
      "data": [
        [ 1, "Agent Smith", true, 8.5 ],
        [ 2, "Agent Jones", true, 7.0 ],
        [ 3, "Agent Brown", false, 6.2 ]
      ]
    }
  *)

  {$ENDREGION 'Sample JSON'}

  finalSQL := '';
  jsonObject := TJSONObject.ParseJSONValue(aFieldsAndData) as TJSONObject;
  if not Assigned(jsonObject) then
    raise Exception.Create('Invalid JSON');

  try
    fieldsArray := jsonObject.GetValue('fields') as TJSONArray;
    dataArray := jsonObject.GetValue('data') as TJSONArray;

    fieldNames := TStringList.Create;
    fieldTypes := TStringList.Create;
    rowValues := TStringList.Create;
    try
      // Collect field names and types
      for i := 0 to FieldsArray.Count - 1 do
      begin
        FieldNames.Add((FieldsArray.Items[i] as TJSONObject).GetValue('name').Value);
        fieldTypes.Add((FieldsArray.Items[i] as TJSONObject).GetValue('type').Value);
      end;

      for i := 0 to DataArray.Count - 1 do
      begin
        rowArray := DataArray.Items[i] as TJSONArray;
        rowValues.Clear;

        for j := 0 to rowArray.Count - 1 do
        begin
          jsonValue := rowArray.Items[j];
          if FieldTypes[j] = 'ftString' then
            rowValues.Add(QuotedStr(jsonValue.Value))
          else if FieldTypes[j] = 'ftBoolean' then
          begin
            if SameText(jsonValue.Value, 'true') then
              rowValues.Add('1')
            else
              rowValues.Add('0');
          end
          else
            rowValues.Add(jsonValue.Value);
        end;

        if i = 0 then
        begin
          // First row includes AS field names
          RowSQL := 'SELECT ';
          for j := 0 to RowValues.Count - 1 do
          begin
            RowSQL := RowSQL + RowValues[j] + ' AS "' + FieldNames[j] + '"';
            if j < RowValues.Count - 1 then
              RowSQL := RowSQL + ', '
            else
              RowSQL := RowSQL + ' FROM DUAL';
          end;
        end
        else
        begin
          // Subsequent rows just values
          RowSQL := 'UNION ALL SELECT ' + RowValues.CommaText;
        end;

        finalSQL := finalSQL + RowSQL + sLineBreak;
      end;
      OpenDataSet(aDataset,finalSQL);
    finally
      FieldNames.Free;
      FieldTypes.Free;
      RowValues.Free;
    end;
  finally
    JSONObject.Free;
  end;
end;

procedure OpenDataSet(aDataset: TOracleDataSet; aCustomSQL: string);
begin
  aDataset.Close;
  aDataset.DeleteVariables;
  aDataset.SQL.Clear;
  aDataset.SQL.Add(aCustomSQL);
  aDataset.Open;
end;

procedure ReplaceComponentWithTEdit(var aComponent: TComponent; aText: string = '');
var
  lOwner: TComponent;
  lName: string;
  lEdit: TEdit;
begin
  lOwner := aComponent.Owner;
  lName := aComponent.Name;
  FreeAndNil(aComponent);
  lEdit := TEdit.Create(lOwner);
  lEdit.Name := lName;
  lEdit.Text := aText;
end;

function FindStringInFile(aFileName, aSearchString: string): Boolean;
var
  textFile: TStringList;
begin
  if FileExists(aFileName) then
  begin
    textfile := TStringList.Create;
    try
      textfile.LoadFromFile(aFileName);
      Result := (Pos(aSearchString,textFile.Text) >=0);
    finally
      textfile.Free;
    end;
  end;

end;
end.
