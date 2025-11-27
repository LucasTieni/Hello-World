unit SMETS_READINGS_DCC;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, OracleData, StdCtrls, Buttons, ExtCtrls, Grids, DBGrids,
  DBCtrls, Oracle, AdvUtil, AdvObj, BaseGrid, AdvGrid, DBAdvGrid, JvExControls,
  JvDBLookup, Vcl.ComCtrls;

type
  TFrm_Smets_Readings_Dcc = class(TForm)
    RegistersQuery: TOracleDataSet;
    ReadingsSrce: TDataSource;
    Panel1: TPanel;
    BitBtn1: TBitBtn;
    GroupBox1: TGroupBox;
    MetersQuery: TOracleDataSet;
    Label1: TLabel;
    MeterSrce: TDataSource;
    oqRefresh: TOracleQuery;
    Meter_Lookup: TJvDBLookupCombo;
    tabRegisters: TTabControl;
    gdReadings: TDBAdvGrid;
    procedure BitBtn1Click(Sender: TObject);
    procedure Meter_LookupChange(Sender: TObject);
    procedure TabRegistersChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

  {$IFDEF CRMTEST}
  protected
  {$ELSE}
  private
  {$ENDIF}
    fSpan : string;
    fHasE7AndE10Reads: boolean;
    procedure SetSpanAndDevice;
    procedure GetMeters;
    procedure GetReadings;
    procedure BuildTabs;
    procedure BuildGrid;

    procedure Refreshdata;
  public
    FormId : integer;

    constructor Create(aOwner: TComponent; aSpan : string); reintroduce;
    class procedure Start(aOwner: TComponent; aSpan : string);
  end;

var
  Frm_Smets_Readings_Dcc: TFrm_Smets_Readings_Dcc;

implementation

uses smets, LoginUnit, MAIN, smets_updates, CrmCommon, UELSqlUtils;
{$R *.dfm}

{==============================================================================}
{$region 'Class: TFrm_Smets_Readings_Dcc'}
{------------------------------------------------------------------------------}
constructor TFrm_Smets_Readings_Dcc.Create(aOwner: TComponent; aSpan : string);
begin
  inherited Create(aOwner);

  fSpan := aSpan;
end;

{------------------------------------------------------------------------------}
class procedure TFrm_Smets_Readings_Dcc.Start(aOwner: TComponent; aSpan : string);
var
  frm : TFrm_Smets_Readings_Dcc;
begin
  frm := TFrm_Smets_Readings_Dcc.Create(aOwner, aSpan);

  if not frm.MetersQuery.IsEmpty then
  begin
    Frm_Main.CrmFormList.Add(frm, frm.FormId);
    frm.Show;
  end
  else
  begin
    MessageDlg('There are no SMETS DCC readings for this Supply', mtError, [mbOk], 0);
    FreeAndNil(frm);
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Readings_Dcc.FormCreate(Sender: TObject);
begin
  RefreshData;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Readings_Dcc.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Frm_Main.CrmFormList.ReleaseForm(FormId);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Readings_Dcc.Refreshdata;
begin
  tabRegisters.tabs.clear;

  GetMeters;

  if MetersQuery.recordcount > 0 then
  begin
    Meter_Lookup.keyvalue := MetersQuery.FieldByName('meter').text;
  end;
end;

procedure TFrm_Smets_Readings_Dcc.BitBtn1Click(Sender: TObject);
begin
  close;
end;

procedure TFrm_Smets_Readings_Dcc.Meter_LookupChange(Sender: TObject);
begin
  fHasE7AndE10Reads := false;

  if MetersQuery.active or (MetersQuery.recordcount > 0) then
  begin
    SetSpanAndDevice;
    GetReadings;
    BuildTabs;
    BuildGrid;
  end;
end;

procedure TFrm_Smets_Readings_Dcc.SetSpanAndDevice;
var
  sql: string;
begin
  try
    sql := 'liberty100.pk_tier_readings.pr_set_span_and_device@stagingdb(:p_span_in, :p_meter_in)';

    gSqlUtil.ExecProc(sql, TRANSACTION_NO,
      ['p_span_in',  otString, pdInput, fSpan,
       'p_meter_in', otString, pdInput, MetersQuery.FieldByName('meter').text]);

  except
    on E: exception do
      MessageDlg('Unable to load readings. ' + E.Message, mtError, [mbOk],0);
  end;
end;

procedure TFrm_Smets_Readings_Dcc.GetMeters;
var
  sql: string;
begin
  try
    // Fields: METER / FULLDESC / SPAN
    sql := 'ods.pk_dcc_metering.pr_get_mpxn_meters(:p_span_in, :po_result_out)';

    gSqlUtil.CreateCursor(MetersQuery, sql, TRANSACTION_NO,
           ['p_span_in',     otString, fSpan,
            'po_result_out', otCursor, null]);

  except
    on E: exception do
      MessageDlg('Unable to get meters. ' + E.Message, mtError, [mbOk],0);
  end;
end;

procedure TFrm_Smets_Readings_Dcc.GetReadings;
var
  sql: string;
begin
  RegistersQuery.Filtered := False;
  RegistersQuery.Filter := EmptyStr;
  gdReadings.DataSource := nil;

  sql := 'select tpr_name, flow, readdate, kwh ' +
         'from liberty100.vw_tier_readings@stagingdb %s ' +
         'order by readdate desc';
  try
    if IsSmetsEA(fSpan) then
    begin
      sql := Format(sql, [' where readdate >= :readdate ']);
      gSqlUtil.SelectQuery(RegistersQuery, sql,
        ['readdate', otDate, GetSmetsEAEffectiveDate(fSpan)]);
    end
    else
    begin
      sql := Format(sql, ['']);
      gSqlUtil.SelectQuery(RegistersQuery, sql);
    end;

  except
    on E: exception do
      MessageDlg('Unable to get readings.' + E.Message, mtError, [mbOk],0);
  end;
end;

procedure TFrm_Smets_Readings_Dcc.BuildTabs;
var
  slTabNames: TStringList;
  index: Integer;
begin
  tabRegisters.Tabs.Clear;
  slTabNames := TStringList.Create;

  try
    RegistersQuery.First;

    while not RegistersQuery.EOF do
    begin
      if slTabNames.IndexOf(RegistersQuery.FieldByName('tpr_name').AsString) = -1 then
      begin
        slTabNames.Add(RegistersQuery.FieldByName('tpr_name').AsString);
      end;

      RegistersQuery.Next;
    end;

    if slTabNames.IndexOf('Peak') <> -1 then
    begin
      slTabNames.Delete(slTabNames.IndexOf('Peak'));
      slTabNames.Insert(0, 'Peak');

      if slTabNames.IndexOf('Day') <> -1 then
      begin
        slTabNames.Delete(slTabNames.IndexOf('Day'));
        fHasE7AndE10Reads := true;
      end;
    end
    else if slTabNames.IndexOf('Day') <> -1 then
    begin
      slTabNames.Delete(slTabNames.IndexOf('Day'));
      slTabNames.Insert(0, 'Day');
    end;

    if slTabNames.IndexOf('OffPeak') <> -1 then
    begin
      slTabNames.Delete(slTabNames.IndexOf('OffPeak'));
      slTabNames.Insert(1, 'OffPeak');

      if slTabNames.IndexOf('Night') <> -1 then
      begin
        slTabNames.Delete(slTabNames.IndexOf('Night'));
        fHasE7AndE10Reads := true;
      end;
    end
    else if slTabNames.IndexOf('Night') <> -1 then
    begin
      slTabNames.Delete(slTabNames.IndexOf('Night'));
      slTabNames.Insert(1, 'Night');
    end;

    For index := 0 to slTabNames.Count - 1 do
    begin
      tabRegisters.Tabs.Add(slTabNames[index]);
    end;

    if tabRegisters.Tabs.Count > 0 then
    begin
      tabRegisters.Tabindex := 0;
    end;

  finally
    FreeAndNil(slTabNames);
  end;
end;

procedure TFrm_Smets_Readings_Dcc.tabRegistersChange(Sender: TObject);
begin
  BuildGrid;
 end;

procedure TFrm_Smets_Readings_Dcc.BuildGrid;
var
  selectedTprName: string;
  filterSql: string;
begin
  selectedTprName := EmptyStr;

  if tabRegisters.Tabs.Count > 0 then
  begin
    selectedTprName := tabRegisters.tabs[tabRegisters.Tabindex];
    tabRegisters.Enabled := True;
  end
  else
  begin
    selectedTprName := '-1';
    tabRegisters.Enabled := False;
  end;

  filterSql := '(tpr_name = ' + QuotedStr(selectedTprName) + ')';
  if fHasE7AndE10Reads then
  begin
    if selectedTprName = 'Peak' then
    begin
      filterSql := filterSql + ' or (tpr_name = ' + QuotedStr('Day') + ')';
    end
    else if selectedTprName = 'OffPeak' then
    begin
      filterSql := filterSql + ' or (tpr_name = ' + QuotedStr('Night') + ')';
    end;
  end;

  try
    with RegistersQuery do
    begin
      Filtered := False;
      Filter := filterSql;
      Filtered := True;
    end;

    RegistersQuery.First;

    gdReadings.DataSource := ReadingsSrce;
  except
    on E: exception do
      ShowMessage(E.Message);
  end;
end;
{------------------------------------------------------------------------------}
{$endregion TFrm_Smets_Readings_Dcc}
{==============================================================================}

end.