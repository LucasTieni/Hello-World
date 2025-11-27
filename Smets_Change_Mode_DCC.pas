unit Smets_Change_Mode_DCC;
interface
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, ExtCtrls, Buttons,
  DB, OracleData, DBCtrls, RxToolEdit, RxCurrEdit,SYSTEM.UITypes,
  Vcl.Samples.Spin, Oracle, JvExControls, JvDBLookup, JvExMask, JvToolEdit;

type
  TFrm_Smets_Change_Mode_Dcc = class(TForm)
    Group_Tariff: TGroupBox;
    TT_LOOKUP: TDBLookupComboBox;
    TT_L: TLabel;
    Label22: TLabel;
    ModeLookup: TDBLookupComboBox;
    Label27: TLabel;
    Vat_Lookup: TDBLookupComboBox;
    Label23: TLabel;
    Label6: TLabel;
    DeviceLookup: TDBLookupComboBox;
    Group_Meter: TGroupBox;
    Label1: TLabel;
    MeterNo: TEdit;
    Group_Credit: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label15: TLabel;
    IC: TCurrencyEdit;
    ID: TCurrencyEdit;
    Group_Customer: TGroupBox;
    CustomerName: TEdit;
    Label9: TLabel;
    Group_Config: TGroupBox;
    GAS_L: TLabel;
    GAS_LOOKUP: TDBLookupComboBox;
    FCD_L: TLabel;
    FCW_L: TLabel;
    Label8: TLabel;
    Group_SMS: TGroupBox;
    Label_SMS: TLabel;
    Group_Card: TGroupBox;
    Label12: TLabel;
    CardNO: TEdit;
    Panel_SPAN: TPanel;
    Prepay_Query: TOracleDataSet;
    PREPAY_SRCE: TDataSource;
    Mm_Query: TOracleDataSet;
    MM_SRCE: TDataSource;
    Tt_Query: TOracleDataSet;
    TT_SRCE: TDataSource;
    Vat_Query: TOracleDataSet;
    VAT_SRCE: TDataSource;
    Ct_Query: TOracleDataSet;
    CT_SRCE: TDataSource;
    Rs_Query: TOracleDataSet;
    RS_SRCE: TDataSource;
    Evt_Query: TOracleDataSet;
    EVT_SRCE: TDataSource;
    Gas_Query: TOracleDataSet;
    GAS_SRCE: TDataSource;
    Co2_Query: TOracleDataSet;
    CO2_SRCE: TDataSource;
    LatestConfig: TOracleDataSet;
    DeviceQuery: TOracleDataSet;
    DeviceSrce: TDataSource;
    DataConfig: TOracleDataSet;
    DataConfigSrce: TDataSource;
    VendMode: TOracleDataSet;
    VendModeSrce: TDataSource;
    DebtMode: TOracleDataSet;
    DebtModeSrce: TDataSource;
    TEmpQuery: TOracleDataSet;
    SupplyImage: TImage;
    OldCard: TLabel;
    olddr: TLabel;
    panel_bottom: TPanel;
    SendBtn: TBitBtn;
    BitBtn2: TBitBtn;
    DR: TSpinEdit;
    grbMeterChange: TGroupBox;
    lcbMeterChange: TDBLookupComboBox;
    lblMeterChange: TLabel;
    qrMeterChange: TOracleDataSet;
    dsMeterChange: TDataSource;
    qrMeterChangeID: TFloatField;
    qrMeterChangeMODE_CHANGE_REASON: TStringField;
    qrMeterChangeMETER_MODE: TStringField;
    qrMeterChangeCAN_USER_SELECT: TStringField;
    qrInsMeterChLog: TOracleQuery;
    qrMeterChangeNEEDS_VALIDATION: TStringField;
    Tariff_Query_Com: TOracleDataSet;
    TARIFF_SRCE_CoM: TDataSource;
    TariffLookupCoM: TJvDBLookupCombo;
    DBGuid: TDBEdit;
    ReadSchedule: TEdit;
    FCW_Edit: TEdit;
    FCD_Edit: TEdit;
    Label_IHD: TLabel;
    IHD: TMemo;
    SMS: TMemo;
    lbWeekLimit: TLabel;
    chkWeekLimit: TCheckBox;
    edtEffectiveDate: TDateEdit;
    lblEffectiveDate: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SendBtnClick(Sender: TObject);
    procedure ModeLookupCloseUp(Sender: TObject);
    procedure MM_SRCEDataChange(Sender: TObject; Field: TField);
    procedure HandleEdtEffectiveDateVisibility();
  strict private
    fService     : integer;
    fSpan        : string;
    fMeter       : string;
    fCustomerId  : Int64;
    fAgreementId : Int64;
    fCanChangeEffectiveDate: boolean;
    fCustomerCurrentTariff: integer;
    fEfsdmsMtd   : TDate;

    function GetReadCycle: string;
    procedure GetFriendlyCredit;
    procedure DoLookupTariff;
    procedure DoModeChange;
    procedure SetIDHTxt(aMeterMode: variant);
    function CheckNewPeakTimes(out oSendSms: boolean): boolean;
    function IsRightTariff: boolean;
    function IsChangeAllowed: boolean;
    function IsTariffChangeAllowed: boolean;
    function IsChangeModeAllowed: boolean;
  public
    CustomerId: String;

    constructor Create(aOwner: TComponent; aService: integer; aSpan, aMeter: string; aCustomerId, aAgreementId: Int64;  aEfsdmsMtd: TDate); reintroduce;
    class function StartModal(aOwner: TComponent; aService: integer; aSpan, aMeter: string; aCustomerId, aAgreementId: Int64;  aEfsdmsMtd: TDate): boolean;
  end;

var
  Frm_Smets_Change_Mode_Dcc: TFrm_Smets_Change_Mode_Dcc;

implementation
uses
  Smets_Updates, Main, DataModule, Common, Smets, DMImages, LoginUnit, Smets_Configuration_DCC,
  Smets_DCC, CrmCommon, UelSqlUtils, Math, StrUtils, SmetsCommon;
{$R *.dfm}

const
  WEEKLY_LIMIT = 'CreditRiskWeekly';

{==============================================================================}
{$region 'Class: TFrm_Smets_Change_Mode_Dcc'}
{------------------------------------------------------------------------------}
constructor TFrm_Smets_Change_Mode_Dcc.Create(aOwner: TComponent; aService: integer; aSpan, aMeter: string; aCustomerId, aAgreementId: Int64; aEfsdmsMtd: TDate);
begin
  inherited Create(aOwner);

  fService     := aService;
  fSpan        := aSpan;
  fMeter       := aMeter;
  fCustomerId  := aCustomerId;
  fAgreementId := aAgreementId;
  fEfsdmsMtd   := aEfsdmsMtd;
end;

{------------------------------------------------------------------------------}
class function TFrm_Smets_Change_Mode_Dcc.StartModal(aOwner: TComponent; aService: integer; aSpan, aMeter: string; aCustomerId, aAgreementId: Int64; aEfsdmsMtd: TDate): boolean;
var
  frm : TFrm_Smets_Change_Mode_Dcc;
begin
  frm := TFrm_Smets_Change_Mode_Dcc.Create(aOwner, aService, aSpan, aMeter, aCustomerId, aAgreementId, aEfsdmsMtd);
  try
    if frm.LatestConfig.IsEmpty then
    begin
      MessageDlg('Data is not available for this service!', mtError, [mbOk], 0);
      exit;
    end;

    Result := frm.ShowModal = mrOk;
  finally
    FreeAndNil(frm);
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Change_Mode_Dcc.FormCreate(Sender: TObject);
var
  paymentCardNo : variant;
  cv            : string;
begin
  Panel_Span.Caption := 'SPAN - ' + fSpan;

  gSqlUtil.SelectQuery(LatestConfig,
    'select suppliercode, jobrequestno, supplytype, supplytype_description, servicepointno, deviceno, '+    // 5
           'devicetype, devicetype_description, deviceinstallationdate, devicemfgno, metertype, '+          // 10
           'metertype_description, metermode, metermode_description, metermode_efd, meter_firmware, '+      // 15
           'chargingtype, chargingtype_description, chargingtype_efd, tariffrefno, tariffrefno_description, '+ // 20
           'tariffrefno_efd, vatgroupid, vatgroupid_description, vatgroupid_efd, gasconfigid, '+  // 25
           'gasconfigid_description, gas_config_efd, prepayconfigid, prepayconfigid_description, '+  // 29
           'prepayconfig_efd, fcsdayconfigid, fcsdayconfigid_description, fcsdayconfig_efd, '+       // 33
           'fcweekconfigid, fcweekconfigid_description, fcweekconfig_efd, rdgscheduleid, '+          // 37
           'rdgscheduleid_description, rdgscheduleid_efd, evtnotifycfgid, evtnotifycfgid_description, '+   // 41
           'evtnotifycfgid_efd, tarifftype, tarifftype_description, tarifftype_efd, co2configid, '+      // 46
           'co2configid_description, co2configid_efd, paymentcardid, paymentcardid_efd, customername, '+   // 51
           'customername_efd, gatewayno, accountbalance, accountbalance_as_of, outstandingdebt, '+         // 56
           'outstandingdebt_as_of, debtrecoveryrate, debtrecoveryrate_efd, currencycode, tarifflabel, '+   // 61
           'meter_removed, meter_removed_date, gateway_status, gateway_status_date, supply_status, '+      // 66
           'supply_status_date, smiff_id, tariff_line_1, tariff_line_2, last_refreshed, deviceguid, customer_id, '+  // 73
           'registration_status, registration_date, manufact_make_type, commission_in_progress, off_peak_hours '+ // 78
      'from ods.vw_sn_current_values '+
      'where servicepointno = :servicepointno',
    ['servicepointno', otString, fSpan]);

  if LatestConfig.IsEmpty then
    exit;

  gSqlUtil.SelectQuery(Ct_Query,
    'select value, description '+
      'from staging.v_lk_charging_type@stagingdb '+
      'order by 2');

  gSqlUtil.SelectQuery(Vat_Query,
    'select value, description, suppliercode '+
      'from staging.v_lk_vat@stagingdb '+
      'order by 2');

  gSqlUtil.SelectQuery(Prepay_Query,
    'select value, description, supplytype, is_default, suppliercode '+
      'from staging.v_lk_prepay_config@stagingdb '+
      'where supplytype = :supplytype '+
      'order by 4 nulls last, 2',
    ['supplytype', otInteger, fService]);

  gSqlUtil.SelectQuery(DeviceQuery,
    'select value, description '+
      'from staging.v_lk_device_type@stagingdb '+
      'order by 2');

  gSqlUtil.SelectQuery(DataConfig,
    'select value, description '+
      'from staging.v_lk_meter_data@stagingdb '+
      'order by 2');

  gSqlUtil.SelectQuery(VendMode,
    'select value, description '+
      'from staging.v_lk_vend_mode@stagingdb '+
      'where value = 2 '+
      'order by 2');

  gSqlUtil.SelectQuery(DebtMode,
    'select value, description '+
      'from staging.v_lk_debt_mode@stagingdb '+
      'where value = 2 '+
      'order by 2');

  gSqlUtil.SelectQuery(Tt_Query,
    'select value, description '+
      'from staging.v_lk_tariff_type@stagingdb '+
      'order by 2');

  gSqlUtil.SelectQuery(Rs_Query,
    'select value, description, use_as_default, suppliercode '+
      'from staging.v_lk_read_schedule_config@stagingdb '+
      'order by 2');

  gSqlUtil.SelectQuery(Evt_Query,
    'select value, description, suppliercode, use_as_default '+
      'from staging.v_lk_events@stagingdb '+
      'order by 2');

  gSqlUtil.SelectQuery(Gas_Query,
    'select value, description, use_as_default, suppliercode '+
      'from staging.v_lk_gas_config@stagingdb '+
      'order by 2');

  gSqlUtil.SelectQuery(Co2_Query,
    'select value, description, supplytype, is_default, suppliercode '+
      'from staging.v_lk_co2@stagingdb '+
      'where supplytype = :supplytype '+
      'order by 2',
    ['supplytype', otInteger, fService]);

  CustomerName.Text := LatestConfig.FieldByName('customername').AsString;
  MeterNo.Text      := fMeter;

  DeviceLookup.KeyValue := LatestConfig.FieldByName('devicetype').Value;
  Vat_Lookup.KeyValue   := LatestConfig.FieldByName('vatgroupid').Value;
  Tt_Lookup.KeyValue    := LatestConfig.FieldByName('tarifftype').Value;
  ReadSchedule.Text     := GetReadCycle;
  Gas_Lookup.KeyValue   := LatestConfig.FieldByName('gasconfigid').Value;
  Ic.Value              := StrToFloat(Frm_Common.GetValue('S2_S1EA_COM_CREDIT_AMOUNT'));
  Id.Value              := 0;

  gSqlUtil.ExecProc('ods.pk_crmui_metering.get_pan(:p_span_in, :p_supply_type_in, :p_pan_out, :p_reserved_for_process_in)', TRANSACTION_YES,
    ['p_span_in',                 otLong,    pdInput,  StrToInt64(fSpan),
     'p_supply_type_in',          otInteger, pdInput,  fService,
     'p_pan_out',                 otString,  pdOutput, @paymentCardNo,
     'p_reserved_for_process_in', otString,  pdInput,  'COS_WIN']);

  CardNo.Text := VarToStr(paymentCardNo);

  SetSupplyTypeIcon(LatestConfig.FieldByName('supplytype').AsInteger, SupplyImage);

  case LatestConfig.FieldByName('supplytype').AsInteger of
    SERVICE_ELECTRICITY:
    begin
      Gas_L.Visible      := false;
      Gas_Lookup.Visible := false;
      edtEffectiveDate.Visible := true;
      edtEffectiveDate.Date := Now;
      HandleEdtEffectiveDateVisibility;
    end;
    SERVICE_GAS:
    begin
      Gas_L.Visible      := true;
      Gas_Lookup.Visible := true;
      edtEffectiveDate.Visible := false;
      lblEffectiveDate.Visible := false;
      Group_Tariff.Height := Group_Tariff.Height - 25;

      // for gas get the default Calorific Value that should be in force at the time.
      cv := GetDefaultCalorificValue(fSpan);
      if cv <> '' then
        Gas_Lookup.KeyValue := StrToInt(cv);
    end;
  end;

  // Now For Meter Mode
  ModeLookup.Enabled  := true;
  ModeLookup.KeyValue := LatestConfig.FieldByName('metermode').Value;

  gSqlUtil.SelectQuery(Mm_Query,
    'select value, description '+
      'from staging.v_lk_meter_mode@stagingdb '+
      'order by 2');

  DBGuid.Text    := LatestConfig.FieldByName('deviceguid').AsString;
  DBGuid.Enabled := false;

  case TMeterMode(LatestConfig.FieldByName('metermode').AsInteger) of
    mmCredit:
    begin
      Modelookup.KeyValue := Ord(mmPrepayment);

      try
        Dr.Value := StrToInt(gSqlUtil.SelectQueryString('select item_value from crm.standing_data where item_name = ''DEFAULT_DEBT_RECOVERY_RATE'''));
      except
        on e:Exception do
        begin
          MessageDlg(e.Message, mtError, [mbOk], 0);
        end;
      end;
    end;
    mmPrepayment:
    begin
      ModeLookup.KeyValue := Ord(mmCredit);
      Dr.Value            := 0;
      Dr.Enabled          := false;
      Dr.Visible          := false;
      Label5.Visible      := false;
      Label15.Visible     := false;
    end;
  end;

  DoLookupTariff;
  DoModeChange;
  GetFriendlyCredit;

  if GetMeterEA(fSpan, fMeter) then
  begin
    IHD.Visible       := false;
    Label_IHD.Visible := false;
    Group_SMS.Height  := Group_SMS.Height - 35;
    SMS.Top           := SMS.Top - 35;
    Label_SMS.Top     := Label_SMS.Top - 35;
    Self.Height       := Self.Height - 35;
  end;

  chkWeekLimit.Checked := false;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Change_Mode_Dcc.DoLookupTariff;
  {----------------------------------------------------------------------------}
  procedure LoadElecSimplifiedTariff;
  var
    simplifiedTariffId : variant;
  begin
    try
      gSqlUtil.CreateCursor(Tariff_Query_Com, 'ods.pk_crmui_tariffs.get_simplified_tariffnames(:p_res_out)', TRANSACTION_NO, ['p_res_out', otCursor, null]);

      TariffLookupCom.LookupDisplay := 'simplified_tariffname';
      TariffLookupCom.LookupField   := 'simplified_tariffid';

      gSqlUtil.ExecProc('ods.pk_crmui_tariffs.preselect_simplified_tariffname(:p_mpxn_in, :p_customer_id_in, :p_aid_in, :p_simplified_tariffid_out, :p_simplified_tariffname_out)', TRANSACTION_NO,
        ['p_mpxn_in',                   otLong,    pdInput,  StrToInt64(fSpan),
         'p_customer_id_in',            otLong,    pdInput,  fCustomerId,
         'p_aid_in',                    otLong,    pdInput,  fAgreementId,
         'p_simplified_tariffid_out',   otInteger, pdOutput, @simplifiedTariffId,
         'p_simplified_tariffname_out', otString,  pdOutput, null]);

      fCustomerCurrentTariff := simplifiedTariffId;
      TariffLookupCom.KeyValue := simplifiedTariffId;
    except
      on e:Exception do
      begin
        MessageDlg('Unable to load Elec Tariffs. Please contact support.' + #13#13 + e.Message, mtError, [mbOk], 0);
        Close;
      end;
    end;
  end;
  {----------------------------------------------------------------------------}
  procedure LoadGasTariff;
  var
    tariffRefNo : variant;
  begin
    tariffRefNo := '';

    TariffLookupCom.LookupDisplay := 'tariff_name';
    TariffLookupCom.LookupField   := 'tariffrefno';
    TariffLookupCom.KeyValue      := GetLatestTariff(fSpan);

    try
      gSqlUtil.CreateCursorExt(Tariff_Query_Com, 'ods.pk_crmui_tariffs.get_tariffs_com(:p_mpxn_in, :p_res_out, :p_preselected_out)', TRANSACTION_NO,
        ['p_mpxn_in',         otLong,   pdInput,  StrToInt64(fSpan),
         'p_res_out',         otCursor, pdOutput, null,
         'p_preselected_out', otString, pdOutput, @tariffRefNo]);

      if VarToStr(tariffRefNo) <> '' then
        TariffLookupCom.KeyValue := VarToStr(tariffRefNo);
    except
      on e:Exception do
      begin
        raise Exception.Create('The tariff list could not be loaded.' + #13#13 + e.Message);
      end;
    end;
  end;
  {----------------------------------------------------------------------------}
begin
  case fService of
    SERVICE_ELECTRICITY: LoadElecSimplifiedTariff;
    SERVICE_GAS:         LoadGasTariff;
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Change_Mode_Dcc.DoModeChange;
  {----------------------------------------------------------------------------}
  function GetCalculatedHeight(aControls: TArray<TWinControl>): integer;
  var
    control : TWinControl;
  begin
    Result := 36; // base

    for control in aControls do
      if control.Visible then
        Result := Result + control.Height;
  end;
  {----------------------------------------------------------------------------}
begin
  CustomerName.Enabled := false;
  Tt_Lookup.Enabled    := false;
  ModeLookup.Enabled   := false;
  ReadSchedule.Enabled := false;
  Group_Config.Height  := 125;
  Vat_Lookup.Enabled   := false;
  Fcd_Edit.Enabled     := false;
  Fcw_Edit.Enabled     := false;
  Gas_Lookup.Enabled   := false;
  DeviceLookup.Enabled := false;

  if TMeterMode(ModeLookup.KeyValue) = mmCredit then
  begin
    Group_Credit.Visible := false;
    Group_Card.Visible   := false;
    Fcd_Edit.Visible     := false;
    Fcw_Edit.Visible     := false;
    Fcd_L.Visible        := false;
    Fcw_L.Visible        := false;
    Dr.Value             := 0;
    Ic.Value             := 0;
    Id.Value             := 0;
  end
  else
  begin
    Group_Credit.Visible := true;
    Group_Card.Visible   := true;
    Fcd_Edit.Visible     := true;
    Fcw_Edit.Visible     := true;
    Fcd_L.Visible        := false;
    Fcw_L.Visible        := false;
  end;

  if Trim(Vat_Lookup.Text) = '' then
  begin
    Vat_Lookup.KeyValue := gSqlUtil.SelectQueryVariant(
      'select item_value '+
        'from liberty100.vw_lib100_sn_vat_group_id@stagingdb '+
        'where servicepointno = :servicepointno and item_value is not null '+
        'order by item_value_date desc',
      ['servicepointno', otString, fSpan]);
  end;

  if Trim(Fcd_Edit.Text) = '' then
  begin
    Fcd_Edit.Text := gSqlUtil.SelectQueryString(
      'select description '+
        'from staging.v_lk_fc_special_day@stagingdb '+
        'where use_as_default = ''Y''');
  end;

  if Trim(Fcw_Edit.text) = '' then
  begin
    Fcw_Edit.Text := gSqlUtil.SelectQueryString(
      'select fcwcdescription '+
        'from staging.tb_std_fcweekdef@stagingdb '+
        'where is_default = ''Y''');
  end;

  if Trim(ReadSchedule.Text) = '' then
  begin
     ReadSchedule.Text := GetReadCycle;
  end;

  if (Gas_Lookup.Text='') and (fService = SERVICE_GAS) then
  begin
    Gas_Lookup.KeyValue := gSqlUtil.SelectQueryVariant(
      'select gasconfigid '+
        'from staging.tb_std_gasconfig@stagingdb '+
        'where is_default = ''Y''');
  end;

  if Trim(CardNo.Text) = '' then
  begin
    CardNo.Text := gSqlUtil.SelectQueryString(
      'select item_value '+
        'from liberty100.vw_lib100_sn_payment_card@stagingdb '+
        'where servicepointno = :servicepointno and item_value is not null '+
        'order by item_value_date desc',
      ['servicepointno', otString, fSpan]);
  end;

  SetIDHTxt(ModeLookup.KeyValue);

  Self.Height := GetCalculatedHeight([Panel_Span,
                                      Group_Customer,
                                      Group_Meter,
                                      Group_Tariff,
                                      Group_Config,
                                      Group_Credit,
                                      Group_Card,
                                      Group_Sms,
                                      Panel_Bottom,
                                      // SJ  - 18/03/2021 - Added height of the new group box
                                      grbMeterChange]);


  Caption := 'Change to ' + MeterModeText[TMeterMode(ModeLookup.KeyValue)];

  if TMeterMode(LatestConfig.FieldByName('metermode').AsInteger) = mmCredit then
  begin
    ShowMessage('You are about to put the meter in ' + MeterModeText[TMeterMode(ModeLookup.KeyValue)] + #13+#13+
                'Please check configuration details on the next screen before continuing.'+#13+#13+
                'Please Ensure Initial Credit & Debt Amounts are Correct.');
  end
  else
  begin
    {$IFNDEF CRMTEST}
    ShowMessage('You are about to put the meter in ' + MeterModeText[TMeterMode(ModeLookup.KeyValue)] + #13+#13+'Please check configuration details on the next screen before continuing.');
    {$ENDIF}
  end;
end;

procedure TFrm_Smets_Change_Mode_Dcc.HandleEdtEffectiveDateVisibility();
begin
  if not fCanChangeEffectiveDate then
  begin
    if TCrmUtil.HasUserFeature(Userid, USER_FEATURE__BACKEDATE_CHANGE_CONFIG) then
    begin
      fCanChangeEffectiveDate  := true;
      edtEffectiveDate.Enabled := true;
    end
    else
    begin
      edtEffectiveDate.Enabled := false;
    end;
  end;
end;

{------------------------------------------------------------------------------}
Procedure TFrm_Smets_Change_Mode_Dcc.FormClose(Sender: TObject; var Action: TCloseAction);
Begin
  qrMeterChange.Close;
End;

{------------------------------------------------------------------------------}
function TFrm_Smets_Change_Mode_Dcc.GetReadCycle : string;
var
  meterMode : TMeterMode;
begin
  Result := '';

  case TMeterMode(LatestConfig.FieldByName('metermode').AsInteger) of
    mmCredit:     meterMode := mmPrepayment;
    mmPrepayment: meterMode := mmCredit;
  end;

  // Anna: This doesn't require transaction management as the stored function inserts into log with autonomous transactions
  try
    Result := gSqlUtil.SelectQueryString('select ods.pk_crmui_metering.get_readcycle(:p_customer_id_in, :p_metermode_in, :p_span_in) from dual',
      ['p_customer_id_in', otLong,    fCustomerId,
       'p_metermode_in',   otInteger, Ord(meterMode),
       'p_span_in',        otLong,    StrToInt64(fSpan)]);
  except
    on e:Exception do
    begin
      MessageDlg(e.Message, mtError, [mbOk], 0);
      exit;
    end;
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Change_Mode_Dcc.SetIDHTxt(aMeterMode: variant);
var
  smsContent : variant;
  ihdContent : variant;
begin
  if VarIsNull(aMeterMode) then
    exit;

  try
    gSqlUtil.ExecProc('ods.pk_crmui_metering.get_com_messages(:p_metermode_in, :p_cardno_in, :p_sms_content_out, :p_ihd_content_out)', TRANSACTION_NO,
      ['p_metermode_in',    otInteger, pdInput,  aMeterMode,
       'p_cardno_in',       otString,  pdInput,  Trim(CardNo.Text),
       'p_sms_content_out', otString,  pdOutput, @smsContent,
       'p_ihd_content_out', otString,  pdOutput, @ihdContent]);

    Sms.Text := VarToStr(smsContent);
    Ihd.Text := VarToStr(ihdContent);
  except
    on e:Exception do
    begin
      MessageDlg(e.Message, mtError, [mbOk], 0);
    end;
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Change_Mode_Dcc.GetFriendlyCredit;
var
  friendCreditDays : variant;
  friendCreditWeek : variant;
begin
  gSqlUtil.ExecProc('ods.pk_crmui_metering.get_s2_friendly_credit(:p_span_in, :p_deviceguid_in, :p_metermode_in, :p_friendly_credit_days_out, :p_friendly_credit_week_out)', TRANSACTION_NO,
    ['p_span_in',                  otString,  pdInput,  fSpan,
     'p_deviceguid_in',            otString,  pdInput,  LatestConfig.FieldByName('deviceguid').AsString,
     'p_metermode_in',             otInteger, pdInput,  IfThen(LatestConfig.FieldByName('metermode').AsInteger = Ord(mmCredit), Ord(mmPrepayment), Ord(mmCredit)),
     'p_friendly_credit_days_out', otString,  pdOutput, @friendCreditDays,
     'p_friendly_credit_week_out', otString,  pdOutput, @friendCreditWeek]);

  if VarToStr(friendCreditDays) <> '' then
    Fcd_Edit.Text := VarToStr(friendCreditDays);

  if VarToStr(friendCreditWeek) <> '' then
    Fcw_Edit.Text := VarToStr(friendCreditWeek);
end;

{------------------------------------------------------------------------------}
function TFrm_Smets_Change_Mode_Dcc.CheckNewPeakTimes(out oSendSms: boolean): boolean;
var
  flag       : variant;
  offPeakMsg : variant;
begin
  oSendSms := false;
  Result   := true;

  try
    gSqlUtil.ExecProc('ods.pk_crmui_metering.get_off_peak_times(:p_mpxn_in, :p_simplified_tariffid_in, :p_meter_mode_in, :p_is_bookingscreen_in, :p_popup_type_out, :p_off_peak_out)', TRANSACTION_NO,
      ['p_mpxn_in',                otLong,    pdInput,  StrToInt64(fSpan),
       'p_simplified_tariffid_in', otLong,    pdInput,  TariffLookupCom.KeyValue,
       'p_meter_mode_in',          otInteger, pdInput,  Modelookup.KeyValue,
       'p_is_bookingscreen_in',    otString,  pdInput,  'N',
       'p_popup_type_out',         otInteger, pdOutput, @flag,
       'p_off_peak_out',           otString,  pdOutput, @offPeakMsg]);

  except
    on e:Exception do
    begin
      MessageDlg('Unable to load new peak times. Please contact support.' + e.Message, mtError, [mbOk], 0);
      exit(false);
    end
  end;

  if VarIsNull(flag) then
    flag := 0;

  if (flag = 1) and (VarToStr(offPeakMsg) <> '') then
  begin
    if MessageDlg(Format('%s Continue?', [VarToStr(offPeakMsg)]), mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      oSendSms := true;
      Result   := true;
    end
    else
    begin
      Result := false;
    end;
  end;
end;

function TFrm_Smets_Change_Mode_Dcc.IsRightTariff: boolean;
  {----------------------------------------------------------------------------}
  procedure InsertQuestionnaireNote(aDoProceed: boolean);
  begin
    try
      gSqlUtil.ExecProc('ods.pk_crmui_tariffs.insert_questionnaire_note(:p_mpxn_in, :p_proceed_in)', TRANSACTION_NO,
        ['p_mpxn_in',    otLong,   pdInput, StrToInt64(fSpan),
         'p_proceed_in', otString, pdInput, IfThen(aDoProceed, 'Y', 'N')]);
    except
      on e:Exception do
      begin
        MessageDlg('Unable to add info to account stating.' + #13#13 + e.Message, mtError, [mbOk],0);
      end;
    end;
  end;
  {----------------------------------------------------------------------------}
var
  flag    : variant;
  msgText : variant;
begin
  Result := true;

  try
    gSqlUtil.ExecProc('ods.pk_crmui_tariffs.get_questionnaire_popup(:p_mpxn_in, :p_simplified_tariffid_in, :p_meter_mode_in, :p_popup_required_out, :p_msg_text_out)', TRANSACTION_NO,
      ['p_mpxn_in',                otLong,    pdInput,  StrToInt64(fSpan),
       'p_simplified_tariffid_in', otInteger, pdInput,  TariffLookupCom.KeyValue,
       'p_meter_mode_in',          otInteger, pdInput,  LatestConfig.FieldByName('metermode').AsInteger,
       'p_popup_required_out',     otString,  pdOutput, @flag,
       'p_msg_text_out',           otString,  pdOutput, @msgText]);

    if VarIsNull(flag) then
      flag := 'N';

    if (flag = 'Y') and (VarToStr(msgText) <> '') then
    begin
      Result := MessageDlg(VarToStr(msgText), mtConfirmation, [mbYes, mbNo], 0) = mrYes;
      InsertQuestionnaireNote(Result);
    end;
  except
    on e:Exception do
    begin
      MessageDlg('Unable to load tariff questionnaire. Please contact support.', mtError, [mbOk], 0);
      Result := false;
    end;
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Change_Mode_Dcc.SendBtnClick(Sender: TObject);
  {----------------------------------------------------------------------------}
  function IsValid(out oPaymentDebtRegister: double; out oErrorMessage: string): boolean;
  begin
    oPaymentDebtRegister := 0;
    oErrorMessage        := '';

    if TMeterMode(LatestConfig.FieldByName('metermode').AsInteger) = mmCredit then
    begin
      if Trim(Dr.Text) = '' then
      begin
        oErrorMessage := 'Debt Recovery Rate cannot be blank!';
        exit(false);
      end;

      if Trim(Ic.Text) = '' then
      begin
        oErrorMessage := 'Initial Credit Amount cannot be blank!';  // ISC-722 Anna: message text has been as it
                                                                    // wasn't aligned with the control caption
        exit(false);
      end;

      oPaymentDebtRegister := Id.Value;
    end
    else
    begin
      oPaymentDebtRegister := 0;
    end;

    if Trim(CustomerName.Text) = '' then
    begin
      oErrorMessage := 'Customer Name cannot be blank!';
      exit(false);
    end;

    if Trim(DeviceLookup.Text) = '' then
    begin
      oErrorMessage := 'Device Type cannot be blank!';
      exit(false);
    end;

    if Trim(ModeLookup.Text) = '' then
    begin
      oErrorMessage := 'Meter Mode cannot be blank!';
      exit(false);
    end;

    if Trim(TariffLookupCom.Text) = '' then
    begin
      oErrorMessage := 'Tariff Refno cannot be blank!';
      exit(false);
    end;

    if Trim(ReadSchedule.Text) = '' then
    begin
      oErrorMessage := 'Reading Schedule cannot be blank!';
      exit(false);
    end;

    if Trim(Fcd_Edit.Text) = '' then
    begin
      oErrorMessage := 'Friendly Credit Special Days cannot be blank!';
      exit(false);
    end;

    if Trim(Fcw_Edit.Text) = '' then
    begin
      oErrorMessage := 'Friendly Credit Week Config cannot be blank!';
      exit(false);
    end;

    if Trim(CardNo.Text) = '' then
    begin
      oErrorMessage := 'Please enter a valid Payment Card Number!';
      exit(false);
    end;

    if edtEffectiveDate.Visible and (edtEffectiveDate.Date > Date) then
    begin
      oErrorMessage := 'The selected effective date cannot be in the future. Please choose a valid date.';
      exit(false);
    end;

    if edtEffectiveDate.Visible and (edtEffectiveDate.Date < IncMonth(Date, -10)) then
    begin
      oErrorMessage := 'The selected effective date cannot be more than 10 months in the past. Please choose a valid date.';
      exit(false);
    end;

    if edtEffectiveDate.Date < fEfsdmsMtd then
    begin
      oErrorMessage := 'You cannot select a date prior to the last MTD date, please select another date and try again.';
      exit(false);
    end;

    Result := true;
  end;
  {----------------------------------------------------------------------------}
  procedure ExecuteChange(aPaymentDebtRegister: double);
  var
    tariffRefNo        : variant;
    simplifiedTariffId : variant;
  begin
    case fService of
      SERVICE_ELECTRICITY:
      begin
        tariffRefNo        := null;
        simplifiedTariffId := TariffLookupCom.KeyValue;
      end;
      SERVICE_GAS:
      begin
        tariffRefNo        := TariffLookupCom.KeyValue;
        simplifiedTariffId := null;
      end;
    end;

    try
      // Anna: this is called with TRANSACTION_NO as the stored procedure manages the transaction!!!
      gSqlUtil.ExecProc('ods.pk_crmui_metering.request_change_of_mode(:p_span_in, '+
                                                                     ':p_deviceguid_in, '+
                                                                     ':p_metermode_in, '+
                                                                     ':p_updatemeterbalance_in, '+
                                                                     ':p_debtrecoveryperpayment_in, '+
                                                                     ':p_paymentdebtregister_in, '+
                                                                     ':p_paymentcardid_in, '+
                                                                     ':p_tariffrefno_in, '+
                                                                     ':p_process_in, '+
                                                                     ':p_simplified_tariffid_in, '+
                                                                     ':p_mtd_effective_from_in)', TRANSACTION_NO,
        ['p_span_in',                   otLong,    pdInput, StrToInt64(fSpan),
         'p_deviceguid_in',             otString,  pdInput, Trim(DBGuid.Text),
         'p_metermode_in',              otInteger, pdInput, IfThen(TMeterMode(LatestConfig.FieldByName('metermode').AsInteger) = mmCredit, Ord(mmPrepayment), Ord(mmCredit)),
         'p_updatemeterbalance_in',     otFloat,   pdInput, Ic.Value,
         'p_debtrecoveryperpayment_in', otInteger, pdInput, Dr.Value,
         'p_paymentdebtregister_in',    otFloat,   pdInput, aPaymentDebtRegister,
         'p_paymentcardid_in',          otString,  pdInput, Trim(CardNo.Text),
         'p_tariffrefno_in',            otLong,    pdInput, tariffRefNo,
         'p_process_in',                otString,  pdInput, IfThen(chkWeekLimit.Checked, WEEKLY_LIMIT),
         'p_simplified_tariffid_in',    otLong,    pdInput, simplifiedTariffId,
         'p_mtd_effective_from_in',     otDate,    pdInput, edtEffectiveDate.Date]);

      {$IFNDEF CRMTEST}
      MessageDlg('Request for change of mode has been sent successfully.', mtInformation, [mbOk], 0);
      {$ENDIF}

      ModalResult := mrOk;
    except
      on e:Exception do
      begin
        MessageDlg('Error: Change of Mode request has failed: ' + #13#10 + e.Message, mtError, [mbOk], 0);
      end;

    end;
  end;
  {----------------------------------------------------------------------------}
  procedure SendOffPeakSms;
  const
    ERRORCODE_MISSING_CUSTOMER_MOBILE = 20005;
  begin
    try
      gSqlUtil.ExecProc('ods.pk_crmui_metering.send_offpeak_sms_text(:p_mpxn_in, :p_simplified_tariffid_in, :p_meter_mode_in)', TRANSACTION_YES,
        ['p_mpxn_in',                otLong,    pdInput, StrToInt64(fSpan),
         'p_simplified_tariffid_in', otLong,    pdInput, TariffLookupCom.KeyValue,
         'p_meter_mode_in',          otInteger, pdInput, LatestConfig.FieldByName('metermode').AsInteger]);
    except
      on e:EOracleError do
      begin
        if e.ErrorCode = ERRORCODE_MISSING_CUSTOMER_MOBILE then
        begin
          MessageDlg('Error sending SMS with peak and off peak times to Customer. Missing Customer Mobile Number.', mtError, [mbOk], 0);
        end
        else
        begin
          MessageDlg('Unable to send SMS to customer with peak and off peak times. ' + e.Message, mtError, [mbOk], 0);
        end;
      end;
    end;
  end;
  {----------------------------------------------------------------------------}
var
  paymentDebtRegister : double;
  errorMsg            : string;
  sendSms             : boolean;
begin
  sendSms := false;

  if not IsValid(paymentDebtRegister, errorMsg) then
  begin
    {$IFNDEF CRMTEST}
    MessageDlg(errorMsg, mtError, [mbOk], 0);
    {$ENDIF}
    exit;
  end;

  {$IFNDEF CRMTEST}
  if Messagedlg(Format('Do you want to %s?', [Self.Caption]), mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
    exit;
  {$ENDIF}

  case fService of
    SERVICE_ELECTRICITY:
    begin
      if IsChangeAllowed and CheckNewPeakTimes(sendSms) and IsRightTariff then
        ExecuteChange(paymentDebtRegister);

      if sendSms then
        SendOffPeakSms;
    end;
    SERVICE_GAS: ExecuteChange(paymentDebtRegister);
  end;
end;

{------------------------------------------------------------------------------}
// SJ & BSL - 19/03/2021 - Control the listed options in the Combo Box
procedure TFrm_Smets_Change_Mode_Dcc.MM_SRCEDataChange(Sender: TObject; Field: TField);
const
  MeterModeDesc : array[TMeterMode] of string = ('PRE-PAYMENT TO CREDIT', 'CREDIT TO PRE-PAYMENT');
begin
  gSqlUtil.SelectQuery(qrMeterChange,
    'select id, mode_change_reason, meter_mode, can_user_select, needs_validation '+
      'from crm.smets_mode_change_reason '+
      'where can_user_select = ''Y'' and upper(meter_mode) = :meter_mode_desc',
    ['meter_mode_desc', otString, MeterModeDesc[TMeterMode(Mm_Query.FieldByName('value').AsInteger)]]);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Change_Mode_Dcc.ModeLookupCloseUp(Sender: TObject);
begin
  DoModeChange;
end;

function TFrm_Smets_Change_Mode_Dcc.IsChangeAllowed: boolean;
begin
  result := IsChangeModeAllowed and IsTariffChangeAllowed;
end;

function TFrm_Smets_Change_Mode_Dcc.IsChangeModeAllowed: boolean;
var
  flag    : variant;
  msgText : variant;
begin

  try
    gSqlUtil.ExecProc('ods.pk_crmui_metering.pr_modechange_meter_limitation('+
      ':p_mpxn_in, :p_meter_mode_in, :p_userid_in, :p_modechange_yn_out, :p_text_out)', TRANSACTION_NO,
      ['p_mpxn_in',           otLong,    pdInput,  StrToInt64(fSpan),
       'p_meter_mode_in',     otInteger, pdInput,  ModeLookup.KeyValue,
       'p_userid_in',         otString,  pdInput, Userid,
       'p_modechange_yn_out', otString,  pdOutput, @flag,
       'p_text_out',          otString,  pdOutput, @msgText]);

    if VarIsNull(flag) then
      flag := 'N';

    if (flag = 'N') and (VarToStr(msgText) <> '') then
    begin
      MessageDlg(VarToStr(msgText), mtInformation, [mbOk], 0);
      Result := false;
    end
    else
    begin
      Result := true;
    end;
  except
    on e:Exception do
    begin
      MessageDlg('Unable to check change viability. Please contact support.', mtError, [mbOk], 0);
      Result := false;
    end;
  end;
end;

function TFrm_Smets_Change_Mode_Dcc.IsTariffChangeAllowed: boolean;
const
  e10RateTariffs    = [5];
  e7RateTariffs     = [3, 4];
  singleRateTariffs = [1, 2];
var
  msg: string;

begin
  result := false;
  if fCustomerCurrentTariff in e10RateTariffs then
    result := Integer(TariffLookupCoM.KeyValue) in e10RateTariffs
  else if fCustomerCurrentTariff in e7RateTariffs then
    result := Integer(TariffLookupCoM.KeyValue) in e7RateTariffs
  else if fCustomerCurrentTariff in singleRateTariffs then
    result := Integer(TariffLookupCoM.KeyValue) in singleRateTariffs;

  if (not result) then
  begin
    msg := 'We currently cannot change the tariff to E7/E10 as the customers heating/hot water will not work.';

    if USER_FEATURE__CHANGE_MULTI_SINGLE_RATE then
    begin
      Result := MessageDlg(Format('%s Continue?', [msg]),
        mtConfirmation, [mbYes, mbNo], 0) = mrYes
    end
    else
    begin
      MessageDlg(msg, mtError, [mbOk],0);
      Result := False;
    end;
  end;
end;
{------------------------------------------------------------------------------}
{$endregion TFrm_Smets_Change_Mode_Dcc}
{==============================================================================}

end.