unit uFriendlyCreditConfig;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, OracleData, Oracle, System.JSON,
  REST.JSON, Vcl.Grids, Data.DB, Datasnap.DBClient, AdvUtil, AdvObj, BaseGrid,
  AdvGrid, DBAdvGrid, Math, Vcl.ComCtrls, Vcl.StdCtrls, Vcl.Buttons;

type
  TFRM_Friendly_Credit_Config = class(TForm)
    GridFCPeriod: TStringGrid;
    btnSaveChanges: TBitBtn;
    btnClose: TBitBtn;
    lblRequierdDistributionTotal: TLabel;
    lblDistributionTotal: TLabel;
    procedure GridFCPeriodDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure GridFCPeriodSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure GridFCPeriodSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: string);
    procedure GridFCPeriodKeyPress(Sender: TObject; var Key: Char);
    procedure btnSaveChangesClick(Sender: TObject);
    procedure GridFCPeriodExit(Sender: TObject);
  {$IFDEF CRMTEST}
  protected
  {$ELSE}
  private
  {$ENDIF}
    FCPeriodIDs: array of Integer;
    ModifiedRows: TStringList;
    FPrevCol, FPrevRow: Integer;
    function Refreshdata: boolean;
    procedure AutoSizeStringGridColumns(GridFCPeriod: TStringGrid);
    procedure AdjustGridWidthToColumns(Grid: TStringGrid);
    procedure UpdateSaveButtonState;
    procedure SaveChanges;
    procedure UpdateDistributionTotalLabel;
    function CalculateRequiredDistributionTotal: Integer;
  public
    destructor Destroy; override;
    class procedure ShowFCConfig(aOwner: TComponent);
  end;

var
  FRM_Friendly_Credit_Config: TFRM_Friendly_Credit_Config;

implementation

uses
  CrmCommon, UelSqlUtils, LoginUnit;

{$R *.dfm}

function TFRM_Friendly_Credit_Config.Refreshdata: boolean;
var
  sqlText, jsonString, status, message: string;
  response: variant;
  i, j: Integer;
  jsonResponse: TJSONObject;
  jsonArray: TJSONArray;
  LFieldName: string;
  LFieldValue: string;

begin
  Result := true;

  if not Assigned(ModifiedRows) then
    ModifiedRows := TStringList.Create;
  ModifiedRows.Clear;

  sqlText := 'ods.pk_crmui_ndc.get_fc_periods(:p_response)';

  gSqlUtil.ExecProc(sqlText, TRANSACTION_NO, ['p_response', otString, pdOutput,
    @response]);

  jsonString := VarToStr(response);
  jsonResponse := TJSONObject.ParseJSONValue(jsonString) as TJSONObject;

  status := LowerCase(jsonResponse.GetValue<string>('status'));
  message := jsonResponse.GetValue<string>('message');

  if not Assigned(jsonResponse) then
  begin
    MessageDlg('Unable to load Friendly Credit Period. Please contact support.',
      mtError, [mbOk], 0);
    Result := false;
  end;

  if jsonResponse.TryGetValue('data', jsonArray) and (jsonArray.Count > 0) then
  begin
    GridFCPeriod.Cells[0, 0] := 'End Time';
    GridFCPeriod.Cells[1, 0] := 'Required Distribution %';
    GridFCPeriod.Cells[2, 0] := 'Latest Distribution %';

    GridFCPeriod.Options := GridFCPeriod.Options + [goEditing];
    GridFCPeriod.Options := GridFCPeriod.Options - [goRangeSelect, goRowSelect];

    GridFCPeriod.FixedCols := 1;

    GridFCPeriod.OnSelectCell := GridFCPeriodSelectCell;
    GridFCPeriod.OnKeyPress := GridFCPeriodKeyPress;
    GridFCPeriod.OnExit := GridFCPeriodExit;

    GridFCPeriod.Font.Style := [fsBold];
    GridFCPeriod.RowCount := jsonArray.Size + 1;
    SetLength(FCPeriodIDs, jsonArray.Size);

    for i := 0 to jsonArray.Size - 1 do
    begin
      jsonResponse := jsonArray.Items[i] as TJSONObject;

      FCPeriodIDs[i] := StrToIntDef(jsonResponse.Pairs[0].JsonValue.Value, 0);

      for j := 0 to jsonResponse.Count - 1 do
      begin
        if j > 0 then
        begin
          LFieldName := jsonResponse.Pairs[j].jsonString.Value;
          LFieldValue := jsonResponse.Pairs[j].JsonValue.Value;
          if LowerCase(LFieldValue) = 'null' then
            GridFCPeriod.Cells[j - 1, i + 1] := ''
          else
            GridFCPeriod.Cells[j - 1, i + 1] := LFieldValue;
        end;
      end;
    end;

    UpdateDistributionTotalLabel;
  end;
  FreeAndNil(jsonResponse);
  
  FPrevCol := -1;
  FPrevRow := -1;

  if status = 'error' then
  begin
    MessageDlg(message, mtError, [mbOk], 0);
    Result := false;
  end;

  AutoSizeStringGridColumns(GridFCPeriod);
  AdjustGridWidthToColumns(GridFCPeriod);
  UpdateSaveButtonState;
end;

procedure TFRM_Friendly_Credit_Config.AutoSizeStringGridColumns(GridFCPeriod: TStringGrid);
var
  Col, Row: Integer;
  MaxWidth: Integer;

begin
  GridFCPeriod.DefaultDrawing := false;
  GridFCPeriod.RowHeights[0] := 20;

  for Col := 0 to GridFCPeriod.ColCount - 1 do
  begin
    MaxWidth := 0;

    if GridFCPeriod.FixedCols > Col then
      MaxWidth := Max(MaxWidth, GridFCPeriod.Canvas.TextWidth(GridFCPeriod.Cells
        [Col, 0]) + 10);

    for Row := 0 to GridFCPeriod.RowCount - 1 do
    begin
      MaxWidth := Max(MaxWidth, GridFCPeriod.Canvas.TextWidth(GridFCPeriod.Cells
        [Col, Row]) + 10);
    end;

    GridFCPeriod.ColWidths[Col] := MaxWidth + 35;
  end;
end;

procedure TFRM_Friendly_Credit_Config.btnSaveChangesClick(Sender: TObject);
begin
  SaveChanges;
end;

procedure TFRM_Friendly_Credit_Config.AdjustGridWidthToColumns(Grid: TStringGrid);
var
  TotalColWidth: Integer;
  i: Integer;

begin
  TotalColWidth := 0;

  for i := 0 to Grid.ColCount - 1 do
  begin
    TotalColWidth := TotalColWidth + Grid.ColWidths[i];
  end;

  Grid.Width := TotalColWidth + 25;
end;

procedure TFRM_Friendly_Credit_Config.UpdateSaveButtonState;
var
  Total: Integer;
begin
  Total := CalculateRequiredDistributionTotal;
  btnSaveChanges.Enabled := Assigned(ModifiedRows) and (ModifiedRows.Count > 0) and (Total = 100);
end;

procedure TFRM_Friendly_Credit_Config.GridFCPeriodDrawCell(Sender: TObject;
  ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  TextToDraw: string;
  TextWidth: Integer;
  CellWidth: Integer;
  LeftOffset: Integer;

begin

  GridFCPeriod.Canvas.Font.Size := 9;
  if ARow = 0 then
  begin
    GridFCPeriod.Canvas.Font.Style := [fsBold];
    GridFCPeriod.Canvas.Brush.Color := clMenu;
  end
  else
  begin
    GridFCPeriod.Canvas.Font.Style := GridFCPeriod.Canvas.Font.Style - [fsBold];
    GridFCPeriod.Canvas.Brush.Color := clWhite;
  end;

  TextToDraw := GridFCPeriod.Cells[ACol, ARow];

  if ACol > 0 then
  begin
    TextWidth := GridFCPeriod.Canvas.TextWidth(TextToDraw);
    CellWidth := GridFCPeriod.ColWidths[ACol];

    LeftOffset := CellWidth - TextWidth;

    GridFCPeriod.Canvas.TextRect(Rect, Rect.Left + (LeftOffset - 5), Rect.Top,
      TextToDraw);
  end
  else
  begin
    GridFCPeriod.Canvas.TextRect(Rect, Rect.Left + 15, Rect.Top, TextToDraw);
  end;

end;

procedure TFRM_Friendly_Credit_Config.GridFCPeriodSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
begin
  if (GridFCPeriod.Col = 1) and (GridFCPeriod.Row > 0) and 
     ((ACol <> GridFCPeriod.Col) or (ARow <> GridFCPeriod.Row)) then
  begin
    if Trim(GridFCPeriod.Cells[GridFCPeriod.Col, GridFCPeriod.Row]) = '' then
    begin
      GridFCPeriod.Cells[GridFCPeriod.Col, GridFCPeriod.Row] := '0';
      if ModifiedRows.IndexOf(IntToStr(GridFCPeriod.Row)) = -1 then
        ModifiedRows.Add(IntToStr(GridFCPeriod.Row));
      UpdateSaveButtonState;
      UpdateDistributionTotalLabel;
    end;
  end;
  
  CanSelect := (ACol = 1) and (ARow > 0);
end;

procedure TFRM_Friendly_Credit_Config.GridFCPeriodKeyPress(Sender: TObject; var Key: Char);
begin
  if (GridFCPeriod.Col = 1) and (GridFCPeriod.Row > 0) then
  begin
    if not (Key in ['0'..'9', #8]) then
      Key := #0;
  end;
end;

procedure TFRM_Friendly_Credit_Config.GridFCPeriodSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: string);
begin
  if (ACol = 1) and (ARow > 0) then
  begin
    if ModifiedRows.IndexOf(IntToStr(ARow)) = -1 then
      ModifiedRows.Add(IntToStr(ARow));
    UpdateSaveButtonState;
    UpdateDistributionTotalLabel;
  end;
end;

procedure TFRM_Friendly_Credit_Config.SaveChanges;
var
  i, RowIndex, FCPeriodID, k: Integer;
  RequiredDist: string;
  JsonPayload: TJSONArray;
  JsonItem: TJSONObject;
  PayloadStr: string;
  sqlText: string;
  vResponse: variant;
  jsonResponse: TJSONObject;
  status, message, failedIds: string;
  failedArray: TJSONArray;
begin
  JsonPayload := TJSONArray.Create;
  try
    for i := 0 to ModifiedRows.Count - 1 do
    begin
      RowIndex := StrToInt(ModifiedRows[i]);
      if (RowIndex > 0) and (RowIndex - 1 < Length(FCPeriodIDs)) then
      begin
        FCPeriodID := FCPeriodIDs[RowIndex - 1];
        RequiredDist := GridFCPeriod.Cells[1, RowIndex];

        JsonItem := TJSONObject.Create;
        JsonItem.AddPair('fc_period_id', TJSONNumber.Create(FCPeriodID));
        JsonItem.AddPair('required_distribution', TJSONNumber.Create(StrToFloatDef(RequiredDist, 0)));

        JsonPayload.AddElement(JsonItem);
      end;
    end;

    PayloadStr := JsonPayload.ToString;

    sqlText := 'ods.pk_crmui_ndc.friendly_credit_req_dist_upd(:p_payload, :p_response)';

    gSqlUtil.ExecProc(sqlText, TRANSACTION_NO, [
      'p_payload', otString, pdInput, PayloadStr,
      'p_response', otString, pdOutput, @vResponse
    ]);

    jsonResponse := TJSONObject.ParseJSONValue(VarToStr(vResponse)) as TJSONObject;
    if Assigned(jsonResponse) then
    begin
      status := LowerCase(jsonResponse.GetValue<string>('status'));
      message := jsonResponse.GetValue<string>('message');

      if status = 'success' then
      begin
        gSqlUtil.Commit;
        ShowMessage('Changes saved successfully: ' + message);
        ModifiedRows.Clear;
        UpdateSaveButtonState;
      end
      else if status = 'partial_success' then
      begin
        failedIds := '';
        if jsonResponse.TryGetValue('failed_fc_period_ids', failedArray) then
        begin
          for k := 0 to failedArray.Count - 1 do
          begin
            if failedIds <> '' then failedIds := failedIds + ', ';
            failedIds := failedIds + failedArray.Items[k].Value;
          end;
        end;

        gSqlUtil.Commit;
        ShowMessage('Partial success: ' + message + #13#10 +
                   'Failed FC Period IDs: ' + failedIds + #13#10 +
                   'Please verify the failed entries and try again.');

        ModifiedRows.Clear;
        UpdateSaveButtonState;
      end
      else if status = 'error' then
      begin
        failedIds := '';
        if jsonResponse.TryGetValue('failed_fc_period_ids', failedArray) then
        begin
          for k := 0 to failedArray.Count - 1 do
          begin
            if failedIds <> '' then failedIds := failedIds + ', ';
            failedIds := failedIds + failedArray.Items[k].Value;
          end;
        end;

        MessageDlg('No rows were updated: ' + message + #13#10 +
                  'Invalid FC Period IDs: ' + failedIds + #13#10 +
                  'Please check your data and try again.',
                  mtError, [mbOK], 0);
      end
      else
      begin
        MessageDlg('Update failed: ' + message, mtError, [mbOK], 0);
      end;

      FreeAndNil(jsonResponse);
    end
    else
    begin
      MessageDlg('Error: Invalid response from server.', mtError, [mbOK], 0);
    end;

  finally
    JsonPayload.Free;
  end;
end;

destructor TFRM_Friendly_Credit_Config.Destroy;
begin
  if Assigned(ModifiedRows) then
    FreeAndNil(ModifiedRows);
  inherited Destroy;
end;

class procedure TFRM_Friendly_Credit_Config.ShowFCConfig(aOwner: TComponent);
var
  frm: TFRM_Friendly_Credit_Config;

begin
  try
    frm := TFRM_Friendly_Credit_Config.Create(aOwner);

    if frm.Refreshdata then
    begin
      frm.BringToFront;
      frm.Position := poDesktopCenter;
      frm.Activate;
      frm.ShowModal;
    end;
  finally
    FreeAndNil(frm);
  end;

end;

function TFRM_Friendly_Credit_Config.CalculateRequiredDistributionTotal: Integer;
var
  i: Integer;
  CellValue: string;
begin
  Result := 0;
  for i := 1 to GridFCPeriod.RowCount - 1 do
  begin
    CellValue := GridFCPeriod.Cells[1, i];
    if CellValue <> '' then
      Result := Result + StrToIntDef(CellValue, 0);
  end;
end;

procedure TFRM_Friendly_Credit_Config.UpdateDistributionTotalLabel;
var
  Total, Difference: Integer;
  DiffText: string;
begin
  Total := CalculateRequiredDistributionTotal;

  if Total = 100 then
  begin
    lblDistributionTotal.Caption := IntToStr(Total) + '%';
    lblDistributionTotal.Font.Color := clGreen;
  end
  else
  begin
    Difference := 100 - Total;
    if Difference > 0 then
      DiffText := ' (need +' + IntToStr(Difference) + '%)'
    else
      DiffText := ' (excess ' + IntToStr(Abs(Difference)) + '%)';

    lblDistributionTotal.Caption := IntToStr(Total) + '%' + DiffText;
    lblDistributionTotal.Font.Color := clRed;
  end;
end;

procedure TFRM_Friendly_Credit_Config.GridFCPeriodExit(Sender: TObject);
begin
  if (GridFCPeriod.Col = 1) and (GridFCPeriod.Row > 0) then
  begin
    if Trim(GridFCPeriod.Cells[GridFCPeriod.Col, GridFCPeriod.Row]) = '' then
    begin
      GridFCPeriod.Cells[GridFCPeriod.Col, GridFCPeriod.Row] := '0';
      if ModifiedRows.IndexOf(IntToStr(GridFCPeriod.Row)) = -1 then
        ModifiedRows.Add(IntToStr(GridFCPeriod.Row));
      UpdateSaveButtonState;
      UpdateDistributionTotalLabel;
    end;
  end;
end;

end.