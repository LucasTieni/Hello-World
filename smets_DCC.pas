unit smets_DCC;
interface
uses
  RXTooledit, RXCurrEdit, Windows, Messages, SysUtils, Variants, Classes,
  Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, RxLookup, ComCtrls, Mask, DB,
  OracleData, DBCtrls, Grids, Menus, ToolWin, shellapi, oracle,
  AdvObj, BaseGrid, AdvGrid, DBAdvGrid, AdvUtil, AdvSmoothCircularProgress,
  Vcl.ImgList, System.ImageList, System.Generics.Collections;

type
  TLoadStatus = (lsFailed, lsOk);

  TFrm_Smets_Dcc = class(TForm)
    LATESTCONFIG: TOracleDataSet;
    LATESTCONFIG_SRCE: TDataSource;
    MainMenu1: TMainMenu;
    Actions_Full: TMenuItem;
    Req_Vend: TMenuItem;
    Req_Read: TMenuItem;
    N2: TMenuItem;
    Smets_TXT: TMenuItem;
    N3: TMenuItem;
    SMETS_CONFIG: TMenuItem;
    View1: TMenuItem;
    DataflowHistory1: TMenuItem;
    VendHistory1: TMenuItem;
    CommsData1: TMenuItem;
    File1: TMenuItem;
    Exit1: TMenuItem;
    N4: TMenuItem;
    SMETS_COT: TMenuItem;
    SMETS_CANC: TMenuItem;
    Req_Debt: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    device_remove: TMenuItem;
    MeterReadings1: TMenuItem;
    EventAlarms1: TMenuItem;
    SMETS_IHD: TMenuItem;
    ihd_add: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    txtMessageHistory1: TMenuItem;
    ProfileData1: TMenuItem;
    Panel2: TPanel;
    Panel1: TPanel;
    Panel_Toolbar_Stuff: TPanel;
    Panel_toolbar_view: TPanel;
    ToolBar2: TToolBar;
    TB_8: TToolButton;
    TB_9: TToolButton;
    TB_10: TToolButton;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton5: TToolButton;
    ToolButton3: TToolButton;
    N9: TMenuItem;
    Panel_MPAN: TPanel;
    GatewayStatusTxt: TDBText;
    FreeVend: TMenuItem;
    smets_vend_add: TMenuItem;
    smets_vend_deduct: TMenuItem;
    ManageDEbt: TMenuItem;
    smets_debt_add: TMenuItem;
    smets_debt_deduct: TMenuItem;
    smets_debt_set: TMenuItem;
    Req_SNAP: TMenuItem;
    SMETS_MODE: TMenuItem;
    BitBtn1: TBitBtn;
    Group_Meter: TGroupBox;
    MeterNo: TDBText;
    Label5: TLabel;
    DBText5: TDBText;
    Label1: TLabel;
    MeterRemoved: TLabel;
    SupplyImage: TImage;
    Group_Customer: TGroupBox;
    Label7: TLabel;
    DBText2: TDBText;
    DBText3: TDBText;
    Label19: TLabel;
    o_cn: TDBText;
    DBText16: TDBText;
    BitBtn2: TBitBtn;
    BitBtn16: TBitBtn;
    Group_Block: TGroupBox;
    Label4: TLabel;
    Label6: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    DB_tariff: TDBText;
    DB_CT: TDBText;
    DB_MM: TDBText;
    DB_TT: TDBText;
    o_ct: TDBText;
    o_tt: TDBText;
    o_mm: TDBText;
    o_tr: TDBText;
    BTN_Tariff: TBitBtn;
    BTN_CT: TBitBtn;
    BTN_MODE: TBitBtn;
    BTN_TT: TBitBtn;
    Group_Prepay: TGroupBox;
    Label15: TLabel;
    Label17: TLabel;
    Label16: TLabel;
    Label18: TLabel;
    DB_PP: TDBText;
    DB_FCW: TDBText;
    DB_FCD: TDBText;
    DBVAT: TDBText;
    o_pp: TDBText;
    o_fcw: TDBText;
    o_fcd: TDBText;
    o_vat: TDBText;
    BTN_PP: TBitBtn;
    BTN_FCW: TBitBtn;
    BTN_FCD: TBitBtn;
    BTN_VAT: TBitBtn;
    Group_Config: TGroupBox;
    Label28: TLabel;
    Label29: TLabel;
    Label_gas: TLabel;
    Label_co2: TLabel;
    o_rs: TDBText;
    o_es: TDBText;
    o_gc: TDBText;
    o_co2: TDBText;
    DBText17: TDBText;
    DBText18: TDBText;
    efd_gas: TDBText;
    efd_co2: TDBText;
    BitBtn17: TBitBtn;
    BitBtn18: TBitBtn;
    btn_gas: TBitBtn;
    btn_co2: TBitBtn;
    Group_Balance: TGroupBox;
    DBText21: TDBText;
    DBText22: TDBText;
    DBText23: TDBText;
    al: TDBText;
    dl: TDBText;
    dr: TDBText;
    Label2: TLabel;
    Label33: TLabel;
    Label32: TLabel;
    Label10: TLabel;
    BitBtn23: TBitBtn;
    BitBtn22: TBitBtn;
    BitBtn21: TBitBtn;
    OD: TCurrencyEdit;
    AB: TCurrencyEdit;
    Smets_Loan: TMenuItem;
    device_enable_han: TMenuItem;
    SMETS_METER_TOOLS: TMenuItem;
    N11: TMenuItem;
    device_reset_token: TMenuItem;
    SiteInfo1: TMenuItem;
    TR: TBitBtn;
    N13: TMenuItem;
    N14: TMenuItem;
    Smets_debt_admin: TMenuItem;
    N15: TMenuItem;
    Smets_Vend_Admin: TMenuItem;
    N16: TMenuItem;
    N1: TMenuItem;
    Smets_Warehouse: TMenuItem;
    Req_Events: TMenuItem;
    smets_debt_suspend: TMenuItem;
    N10: TMenuItem;
    Suspended: TLabel;
    DBText1: TDBText;
    label_firmware: TLabel;
    BitBtn3: TBitBtn;
    ihd_pin: TMenuItem;
    N12: TMenuItem;
    REQ_Diag: TMenuItem;
    BOOST_SETTINGS: TMenuItem;
    N17: TMenuItem;
    COSStatus1: TMenuItem;
    Req_Mirror: TMenuItem;
    Gas_Wait_Timer: TTimer;
    Circle: TAdvSmoothCircularProgress;
    SignalList: TImageList;
    PANEL_COMMS: TPanel;
    SupplyState: TImage;
    PowerList: TImageList;
    CommsGrade: TOracleDataSet;
    btnViewFriendlyDayCreditConfig: TBitBtn;
    GenDS: TDataSource;
    OracleGenDS: TOracleDataSet;
    FCD_ID: TDBText;
    N18: TMenuItem;
    DisconnectSupply1: TMenuItem;
    ConnectSupplyArmed1: TMenuItem;
    M_SUPPLY_CONTROL: TMenuItem;
    DBGuid: TDBEdit;
    PPMIDMessages1: TMenuItem;
    BalanceDebtHistory1: TMenuItem;
    FC_Update: TMenuItem;
    ToolButton4: TToolButton;
    Group_DataFlow_History: TGroupBox;
    FlowGridDCC: TDBAdvGrid;
    Panel3: TPanel;
    FlowHistoryDCC: TOracleDataSet;
    Flow_Srce_DCC: TDataSource;
    ihd_rebind: TMenuItem;
    N19: TMenuItem;
    label_maketype: TLabel;
    o_mmt: TDBText;
    o_cms: TLabel;
    PanelCMS: TPanel;
    N20: TMenuItem;
    EnableEmergCredit: TMenuItem;
    MeterTypeLabel: TLabel;
    MeterTypeImage: TImage;
    LabelAmountRestore: TLabel;
    LabelLastDemError: TLabel;
    LabelDemandReq: TLabel;
    lbOffPeakTimes: TLabel;
    o_opt: TDBText;
    LabelWeekLimit: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure Req_VendClick(Sender: TObject);
    procedure Req_ReadClick(Sender: TObject);
    procedure SendTXTMessage1Click(Sender: TObject);
    procedure SMETS_CONFIGClick(Sender: TObject);
    procedure VendHistory1Click(Sender: TObject);
    procedure DataflowHistory1Click(Sender: TObject);
    procedure CommsData1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure TB_4Click(Sender: TObject);
    procedure TB_5Click(Sender: TObject);
    procedure TB_6Click(Sender: TObject);
    procedure TB_7Click(Sender: TObject);
    procedure TB_8Click(Sender: TObject);
    procedure TB_9Click(Sender: TObject);
    procedure TB_10Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn23Click(Sender: TObject);
    procedure BitBtn22Click(Sender: TObject);
    procedure BitBtn21Click(Sender: TObject);
    procedure BTN_TariffClick(Sender: TObject);
    procedure BTN_VATClick(Sender: TObject);
    procedure BTN_CTClick(Sender: TObject);
    procedure BTN_PPClick(Sender: TObject);
    procedure BTN_FCDClick(Sender: TObject);
    procedure BTN_FCWClick(Sender: TObject);
    procedure BTN_MODEClick(Sender: TObject);
    procedure BitBtn16Click(Sender: TObject);
    procedure BitBtn17Click(Sender: TObject);
    procedure BitBtn18Click(Sender: TObject);
    procedure btn_gasClick(Sender: TObject);
    procedure btn_co2Click(Sender: TObject);
    procedure SMETS_CANCClick(Sender: TObject);
    procedure tb_15Click(Sender: TObject);
    procedure BTN_TTClick(Sender: TObject);
    procedure SMETS_COTClick(Sender: TObject);
    procedure tb_13Click(Sender: TObject);
    procedure tb_14Click(Sender: TObject);
    procedure Req_DebtClick(Sender: TObject);
    procedure tb_11Click(Sender: TObject);
    procedure device_removeClick(Sender: TObject);
    procedure MeterReadings1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure EventAlarms1Click(Sender: TObject);
    procedure ihd_addClick(Sender: TObject);
    procedure txtMessageHistory1Click(Sender: TObject);
    procedure ProfileData1Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
    procedure smets_vend_addClick(Sender: TObject);
    procedure smets_vend_deductClick(Sender: TObject);
    procedure smets_debt_addClick(Sender: TObject);
    procedure smets_debt_deductClick(Sender: TObject);
    procedure smets_debt_setClick(Sender: TObject);
    procedure Req_SNAPClick(Sender: TObject);
    procedure smets_ihd_replaceClick(Sender: TObject);
    procedure smets_ihd_addClick(Sender: TObject);
    procedure smets_ihd_removeClick(Sender: TObject);
    procedure smets_txtClick(Sender: TObject);
    procedure SMETS_MODEClick(Sender: TObject);
    procedure Smets_LoanClick(Sender: TObject);
    procedure device_enable_hanClick(Sender: TObject);
    procedure device_reset_tokenClick(Sender: TObject);
    procedure SiteInfo1Click(Sender: TObject);
    procedure TRClick(Sender: TObject);
    procedure RemoveDevice;
    procedure MTR1Click(Sender: TObject);
    procedure RequestResetToken;
    procedure MTR2Click(Sender: TObject);
    procedure EnableThisHan;
    procedure MTR3Click(Sender: TObject);
    procedure Smets_debt_adminClick(Sender: TObject);
    procedure Smets_Vend_AdminClick(Sender: TObject);
    procedure Smets_WarehouseClick(Sender: TObject);
    procedure Req_EventsClick(Sender: TObject);
    procedure smets_debt_suspendClick(Sender: TObject);
    procedure ihd_pinClick(Sender: TObject);
    procedure REQ_DiagClick(Sender: TObject);
    procedure BOOST_SETTINGSClick(Sender: TObject);
    procedure COSStatus1Click(Sender: TObject);
    procedure RefreshTimerTimer(Sender: TObject);
    procedure AutoRefreshClick(Sender: TObject);
    procedure FlowGridDblClick(Sender: TObject);
    Procedure ShowDataflowHistory;
    procedure FlowGridGetCellColor(Sender: TObject; ARow, ACol: integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure FlowHistoryAfterOpen(DataSet: TDataSet);
    procedure Req_MirrorClick(Sender: TObject);
    procedure Gas_Wait_TimerTimer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnViewFriendlyDayCreditConfigClick(Sender: TObject);
    procedure DisconnectSupply1Click(Sender: TObject);
    procedure ConnectSupplyArmed1Click(Sender: TObject);
    procedure PPMIDMessages1Click(Sender: TObject);
    procedure BalanceDebtHistory1Click(Sender: TObject);
    procedure ToolButton4Click(Sender: TObject);
    procedure DelayButton;
    procedure Ihd_rebindClick(Sender: TObject);
    procedure EnableEmergCreditClick(Sender: TObject);
    procedure FC_UpdateClick(Sender: TObject);
  {$IFDEF CRMTEST}
  protected
  {$ELSE}
  strict private
  {$ENDIF}
    fService     : integer;
    fSpan        : string;
    fAgreementId : Int64;
    fLoadStatus  : TLoadStatus;
    fEfsdmsMtd   : TDate;

    procedure Refreshdata;
    procedure DoFlowHistory(aSpan: string);
    procedure SetLatestDemand;
    function IsManufacturerAllowed: boolean;
    function HandleNotAllowed(aComError: string): boolean;
    function HandleProcodeMeter(aComError: string): boolean;

    // Setting up user features
    procedure SetUserSystemFeatures;

    procedure DoSmets2Credit(aMode: integer);
    procedure DoSmets2Debt(aMode: integer);
    procedure ShowItem(aTable, aDescription, aTitleCaption: string);

    procedure PerformIhdRepair;
    function  GetMeterTypeIndex: integer;
    procedure RestrictStaffTariffPaymentsAndControls;
    function  ContainsFCData: boolean;

  public
    FormId : integer;

    constructor Create(aOwner: TComponent; aService: integer; aSpan: string; aAgreementId: Int64); reintroduce;
    class procedure ShowSmets(aOwner: TComponent; aService: integer; aSpan: string; aAgreementId: Int64 = 0);

    function CheckSpan(aService: integer; aSpan: string): boolean;

    property LoadStatus : TLoadStatus read fLoadStatus;
    { Public declarations }
  end;

var
  Frm_Smets_Dcc: TFrm_Smets_Dcc;

const
  XS_Code     = 'UTILITA';
  XS_test     = '1';
  XS_PRIORITY = '1';
  XS_RETRY    = '10';
  XS_APP      = 'CRM';

implementation
uses
  smets_manage_credit_dcc, DataModule, main, LoginUnit, smets_configuration_dcc, smets_data_item,
  smets_removedevice, smets_updates, SMETS_PROFILE, Processing, Common, DMImages, CrmCommon, UELSqlUtils,
  // added by maryam on 04/11/2015
  FriendlyDayCreditConfigViewer, System.StrUtils, UelMessage,

    // DCC forms
  smets_balance_debt_DCC,
  SMETS_READINGS_DCC,
  smets_flow_history_DCC,
  smets_alarms_DCC,
  SMETS_SITE_INFO_DCC,
  smets_data_item_DCC,
  smets_message_history_dcc,
  smets_debt_DCC,
  smets_on_demand_read_DCC,
  smets_change_mode_DCC,
  smets_device_dcc,
  Smets_Vend_Hist_Dcc,
  SmetsCommon,
  FriendlyCreditUpdate;

{$R *.dfm}

const
  ICON_LARGEIMAGES__NONE                 = -1;
  ICON_LARGEIMAGES__HH_METER             = 9;
  ICON_LARGEIMAGES__NHH_KEY_METER        = 21;
  ICON_LARGEIMAGES__NHH_CREDIT_METER     = 17;
  ICON_LARGEIMAGES__NHH_SMAR_CARD_METER  = 22;
  ICON_LARGEIMAGES__NHH_TOKEN_METER      = 23;
  ICON_LARGEIMAGES__SMETS1               = 239;
  ICON_LARGEIMAGES__SMETS2               = 205;
  ICON_LARGEIMAGES__SMETS1EA             = 313;
  ICON_LARGEIMAGES__RCAM                 = 314;
  ICON_LARGEIMAGES__STANDING_CHARGE      = 321;

{==============================================================================}
{$region 'Class: TFrm_Smets_Dcc'}
{------------------------------------------------------------------------------}
constructor TFrm_Smets_Dcc.Create(aOwner: TComponent; aService: integer; aSpan: string; aAgreementId: Int64);
begin
  inherited Create(aOwner);

  fService     := aService;
  fSpan        := aSpan;
  fAgreementId := aAgreementId;

  FormId       := 0;
end;

{------------------------------------------------------------------------------}
class procedure TFrm_Smets_Dcc.ShowSmets(aOwner: TComponent; aService: integer; aSpan: string; aAgreementId: Int64 = 0);
  {----------------------------------------------------------------------------}
  function LatestConfigExists(aSpan: string): boolean;
  begin
    Result := gSqlUtil.SelectQueryInteger(
      'select count(*) '+
        'from ods.vw_sn_current_values '+
        'where servicepointno = :servicepointno',
      ['servicepointno', otString, aSpan]) > 0;
  end;
  {----------------------------------------------------------------------------}
  // Returns with true if this window is already open with the same service and span, oFormId will
  // contain the form ID of the existing window
  function CheckInstanceOfSpan(aService: integer; aSpan: string; out oFormId: integer): boolean;
  var
    existingForm : TCrmForm;
  begin
    oFormId := 0;

    for existingForm in Frm_Main.CrmFormList do
    begin
      if (existingForm.Form is TFrm_Smets_Dcc) and TFrm_Smets_Dcc(existingForm.Form).CheckSpan(aService, aSpan) then
      begin
        oFormId := existingForm.FormId;
        exit(true);
      end;
    end;

    Result := false;
  end;
  {----------------------------------------------------------------------------}
var
  frm    : TFrm_Smets_Dcc;
  formId : integer;
begin
  if CheckInstanceOfSpan(aService, aSpan, formId) then
  begin
    frm := Frm_Main.CrmFormList.GetCrmFormByFormId(formId).Form as TFrm_Smets_Dcc;
    if Assigned(frm) then
    begin
      frm.BringToFront;
      frm.Activate;
    end;
  end
  else
  begin
    if LatestConfigExists(aSpan) then
    begin
      frm := TFrm_Smets_Dcc.Create(aOwner, aService, aSpan, aAgreementId);
      if frm.LoadStatus = lsFailed then
      begin
        FreeAndNil(frm);
        exit;
      end;

      Frm_Main.CrmFormList.Add(frm, frm.FormId);

      if not gNoCosPrompt then
        frm.Show;
    end
    else
      TFrm_Smets_Flow_History_Dcc.Start(Frm_Main, aService, aSpan, true);
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.FormCreate(Sender: TObject);
var
  iconIndex : integer;
begin
  try
    SetUserSystemFeatures;
    Refreshdata;

    MeterTypeImage.Picture.Assign(nil);

    iconIndex := GetMeterTypeIndex;
    if iconIndex > ICON_LARGEIMAGES__NONE then
      DM_Images.LargeImages.GetIcon(iconIndex, MeterTypeImage.Picture.Icon);

    fLoadStatus := lsOk;
  except
    on e:Exception do
    begin
      MessageDlg(e.Message, mtError, [mbOk], 0);
      fLoadStatus := lsFailed;
    end;
  end;

  RestrictStaffTariffPaymentsAndControls;
end;

{------------------------------------------------------------------------------}
function TFrm_Smets_Dcc.GetMeterTypeIndex: integer;
  {----------------------------------------------------------------------------}
  function GetIconIndexElec(var oMeterTypeText: string): integer;
  var
    q : TOracleDataSet;
  begin
    Result := ICON_LARGEIMAGES__NONE;

    q := gSqlUtil.SelectQuery(
      'select meter_type, man_make_type, retrieval_method, efsdmsmtd, has_sc_tariff '+
        'from (select meter_type, man_make_type, retrieval_method, efsdmsmtd, has_sc_tariff '+
                'from ods.vw_crm_mtds_elec '+
                'where mpancore = :span and meterid = :meterid '+
                'order by efsdmsmtd desc) '+
          'where rownum <= 1',
      ['span',    otString, fSpan,
       'meterid', otString, MeterNo.Caption]);
    try
      if q.FieldByName('has_sc_tariff').AsString = 'Y' then
      begin
        oMeterTypeText := 'Standing Charge Meter';
        Result         := ICON_LARGEIMAGES__STANDING_CHARGE;
      end
      else if q.FieldByName('meter_type').AsString = 'N' then
      begin
        oMeterTypeText := 'NHH Credit Meter';
        Result         := ICON_LARGEIMAGES__NHH_CREDIT_METER;
      end
      else if q.FieldByName('meter_type').AsString = 'S' then
      begin
        if q.FieldByName('retrieval_method').AsString = 'R' then
        begin
          oMeterTypeText := 'Remote Read Smart Meter';
          Result         := ICON_LARGEIMAGES__SMETS2;
        end
        else
        begin
          oMeterTypeText := 'NHH Smart Card Meter';
          Result := ICON_LARGEIMAGES__NHH_SMAR_CARD_METER;
        end;
      end
      else if (Copy(q.FieldByName('meter_type').AsString, 1, 4) = 'RCAM') or (q.FieldByName('man_make_type').AsString = 'PRI') or (Copy(q.FieldByName('meter_type').AsString, 1, 3) = 'NSS') then
      begin
        oMeterTypeText := 'Smart Meter';
        Result         := ICON_LARGEIMAGES__RCAM;
      end
      else if q.FieldByName('meter_type').AsString = 'S1EA' then
      begin
        oMeterTypeText := 'Smets 1 E&&A Smart Meter';
        Result := ICON_LARGEIMAGES__SMETS1EA
      end
      else if Copy(q.FieldByName('meter_type').AsString, 1, 2) = 'S1' then
      begin
        oMeterTypeText := 'Smart Meter';
        Result         := ICON_LARGEIMAGES__SMETS1;
      end
      else if Copy(q.FieldByName('meter_type').AsString, 1, 2) = 'S2' then
      begin
        oMeterTypeText := 'SMETS 2 Meter';
        Result         := ICON_LARGEIMAGES__SMETS2;
      end
      else if q.FieldByName('meter_type').AsString = 'T' then
      begin
        oMeterTypeText := 'NHH Token Meter';
        Result         := ICON_LARGEIMAGES__NHH_TOKEN_METER;
      end
      else if q.FieldByName('meter_type').AsString = 'K' then
      begin
        oMeterTypeText := 'NHH Key Meter';
        Result         := ICON_LARGEIMAGES__NHH_KEY_METER;
      end
      else if q.FieldByName('meter_type').AsString = 'H' then
      begin
        oMeterTypeText := 'HH Meter';
        Result         := ICON_LARGEIMAGES__HH_METER;
      end;

      fEfsdmsMtd := q.FieldByName('efsdmsmtd').AsDateTime;
    finally
      FreeAndNil(q);
    end;
  end;
  {----------------------------------------------------------------------------}
  function GetIconIndexGas(var oMeterTypeText: string): integer;
  var
    q               : TOracleDataSet;
    meterMechanism  : string;
    manufactureCode : string;
  begin
    Result         := ICON_LARGEIMAGES__NHH_CREDIT_METER;
    oMeterTypeText := 'Metric Meter';

    q := gSqlUtil.SelectQuery(
      'select metertype, metermechanism, manufacturecode, has_sc_tariff '+
        'from (select metertype, metermechanism, manufacturecode, has_sc_tariff '+
                'from ods.vw_crm_mtds_gas '+
                'where meter_point_reference = :span and serialnum = :meterid '+
                'order by startdate desc) '+
        'where rownum <= 1',
      ['span',    otString, fSpan,
       'meterid', otString, MeterNo.Caption]);
    try
      meterMechanism  := q.FieldByName('metermechanism').AsString;
      manufactureCode := q.FieldByName('manufacturecode').AsString;

      if (meterMechanism = 'CM') or (meterMechanism = 'ET') or (meterMechanism = 'MT') or
         (meterMechanism = 'PP') or (meterMechanism = 'TH') then
      begin
        Result := ICON_LARGEIMAGES__NHH_TOKEN_METER;
      end;

      if (manufactureCode = 'PRI') then
      begin
        if meterMechanism = 'S1' then
        begin
          Result := ICON_LARGEIMAGES__SMETS1;
        end
        else
        begin
          Result := ICON_LARGEIMAGES__RCAM;
        end;
      end;

      if (manufactureCode = 'SCM') then
      begin
        if meterMechanism = 'S1' then
        begin
          Result := ICON_LARGEIMAGES__SMETS1;
        end
        else if meterMechanism = 'NS' then
        begin
          Result := ICON_LARGEIMAGES__RCAM;
        end;
      end;

      if meterMechanism = 'S2' then
      begin
        oMeterTypeText := 'SMETS 2 Meter';
        Result := ICON_LARGEIMAGES__SMETS2;
      end;

      if meterMechanism = 'S1EA' then
      begin
        oMeterTypeText := 'Smets 1 E&&A Smart Meter';
        Result := ICON_LARGEIMAGES__SMETS1EA;
      end;

      if q.FieldByName('metertype').AsString = 'F' then
        oMeterTypeText := 'Imperial Meter';

      if q.FieldByName('has_sc_tariff').AsString = 'Y' then
      begin
        oMeterTypeText := 'Standing Charge Meter';
        Result         := ICON_LARGEIMAGES__STANDING_CHARGE;
      end

    finally
      FreeAndNil(q);
    end;
  end;
  {----------------------------------------------------------------------------}
var
  meterType : string;
begin
  case fService of
    SERVICE_ELECTRICITY: Result := GetIconIndexElec(meterType);
    SERVICE_GAS:         Result := GetIconIndexGas(meterType);
  end;

  MeterTypeLabel.Caption := meterType;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.SetUserSystemFeatures;
var
  featureIds : TList<integer>;
  i          : integer;
begin
  // Vend
  FreeVend.Visible           := false;
  Smets_Vend_Add.Visible     := false;
  Smets_Vend_Deduct.Visible  := false;
  Smets_Vend_Admin.Visible   := false;

  // Loan
  Smets_Loan.Visible         := false;

  // Debt
  ManageDebt.Visible         := false;
  Smets_Debt_Add.Visible     := false;
  Smets_Debt_Deduct.Visible  := false;
  Smets_Debt_Set.Visible     := false;
  Smets_Debt_Suspend.Visible := false;
  Smets_Debt_Admin.Visible   := false;

  // Req
  Req_Debt.Visible           := false;
  Req_Vend.Visible           := false;
  Req_Events.Visible         := false;
  Req_Mirror.Visible         := false;
  Req_Read.Visible           := false;
  Req_Snap.Visible           := false;
  Req_Diag.Visible           := false;

  Smets_Txt.Visible          := false;
  Smets_Config.Visible       := false;
  Smets_Mode.Visible         := false;
  Smets_Cot.Visible          := false;
  Smets_Canc.Visible         := false;
  Smets_Meter_Tools.Visible  := false;
  Device_Remove.Visible      := false;
  Device_Reset_Token.Visible := false;
  Device_Enable_Han.Visible  := false;
  M_Supply_Control.Visible   := false;
  FC_Update.Visible          := false;

  // IHD
  Smets_Ihd.Visible          := false;
  Ihd_Add.Visible            := false;
  Ihd_Rebind.Visible         := false;
  Ihd_Pin.Visible            := false;

  featureIds := TCrmUtil.GetUserFeatures(UserId);
  try
    for i := 0 to featureIds.Count-1 do
    begin
      case featureIds[i] of
        {1}
        USER_FEATURE__ADD_CREDIT_VEND:
        begin
          Actions_Full.Visible   := true;    // Anna: this initialised at design-time
          Smets_Vend_Add.Visible := true;
          FreeVend.Visible       := true;
        end;

        {2}
        USER_FEATURE__DEDUCT_CREDIT_VEND:
        begin
          Actions_Full.Visible      := true;    // Anna: this initialised at design-time
          Smets_Vend_Deduct.Visible := true;
          FreeVend.Visible          := true;
        end;

        {3}
        USER_FEATURE__SET_CREDIT_VEND:
        begin
          Actions_Full.Visible := true;    // Anna: this initialised at design-time
          FreeVend.Visible     := true;
        end;

        {4}
        USER_FEATURE__ADD_DEBT:
        begin
          Actions_Full.Visible   := true;    // Anna: this initialised at design-time
          Smets_Debt_Add.Visible := true;
          ManageDebt.Visible     := true;
        end;

        {5}
        USER_FEATURE__DEDUCT_DEBT:
        begin
          Actions_Full.Visible      := true;    // Anna: this initialised at design-time
          Smets_Debt_Deduct.Visible := true;
          ManageDebt.Visible        := true;
        end;

        {6}
        USER_FEATURE__SET_DEBT:
        begin
          Actions_Full.Visible   := true;    // Anna: this initialised at design-time
          Smets_Debt_Set.Visible := true;
          ManageDebt.Visible     := true;
        end;

        {7}
        USER_FEATURE__MANAGE_CREDIT_ADMIN:
        begin
          Actions_Full.Visible     := true;    // Anna: this initialised at design-time
          FreeVend.Visible         := true;
          Smets_Vend_Admin.Visible := true;
        end;

        {8}
        USER_FEATURE__MANAGE_DEBT_ADMIN:
        begin
          Actions_Full.Visible     := true;    // Anna: this initialised at design-time
          ManageDebt.Visible       := true;
          Smets_Debt_Admin.Visible := true;
        end;

        {9}
        USER_FEATURE__CHANGE_CONFIG:
        begin
          Actions_Full.Visible := true;    // Anna: this initialised at design-time
          Smets_Config.Visible := true;
        end;

        {13}
        USER_FEATURE__IHD_REPLACE:
        begin
          Actions_Full.Visible := true;    // Anna: this initialised at design-time
          Smets_Ihd.Visible    := true;
        end;

        {14}
        USER_FEATURE__IHD_ADD:
        begin
          Actions_Full.Visible := true;    // Anna: this initialised at design-time
          Ihd_Add.Visible      := true;
          Ihd_Rebind.Visible   := true;
          Smets_Ihd.Visible    := true;
        end;

        {15}
        USER_FEATURE__IHD_REMOVE:
        begin
          Actions_Full.Visible := true;    // Anna: this initialised at design-time
          Smets_Ihd.Visible    := true;
        end;

        {17}
        USER_FEATURE__LOAN_AMOUNT_VEND:
        begin
          Actions_Full.Visible := true;    // Anna: this initialised at design-time
          Smets_Loan.Visible   := true;
        end;

        {25}
        USER_FEATURE__ON_DEMAND_READING:
        begin
          Actions_Full.Visible := true;    // Anna: this initialised at design-time
          Req_Read.Visible     := true;
          Req_Mirror.Visible   := true;
        end;

        {26}
        USER_FEATURE__CHANGE_METER_MODE_PREPAY_CREDIT:
        begin
          Actions_Full.Visible := true;    // Anna: this initialised at design-time
          Smets_Mode.Visible   := true;
        end;

        {29}
        USER_FEATURE__CANCEL_DELAYED_CONFIG:
        begin
          Actions_Full.Visible := true;    // Anna: this initialised at design-time
          Smets_Canc.Visible   := true;
        end;

        {31}
        USER_FEATURE__SUSPEND_DEBT:
        begin
          Actions_Full.Visible       := true;    // Anna: this initialised at design-time
          Smets_Debt_Suspend.Visible := true;
        end;

        {148}
        USER_FEATURE__UPDATE_FRIENDLY_CREDIT_PERIOD:
        begin
          Actions_Full.Visible := true;
          FC_Update.Visible    := true;
        end;
      end;
    end;
  finally
    FreeAndNil(featureIds);
  end;

  Smets_Ihd.Visible := true;  // alwasys ON as PIN is default allowed
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BitBtn3Click(Sender: TObject);
begin
  Messagedlg('The Device Firmware option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {ShowItem('SP_METER_FIRMWARE', 'Device Firmware', 'Device Firmware');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.btnViewFriendlyDayCreditConfigClick(Sender: TObject);
begin
  if Fcd_Id.Caption = '' then
    exit;

  if ContainsFCData then
  begin
   TFrmFriendlyDayCreditConfigViewer.StartModal(Self, fSpan, LatestConfig.FieldByName('fcsdayconfigid').AsString, true);
  end
  else
   Messagedlg('Friendly Credit Days not defined.',mtinformation, [mbOK],0);

end;

{------------------------------------------------------------------------------}

function TFrm_Smets_Dcc.ContainsFCData : boolean;
var
  fctempDataSet : TOracleDataSet;
  sqlText : string;

begin
  Result := true;
  sqlText := 'ODS.PK_CRMUI_NDC.get_special_days(:P_MPXN, :PO_PC_SPECIAL_DAYS)';

  try
   fctempDataSet := gSqlUtil.CreateCursor(sqlText, TRANSACTION_NO,
    ['P_MPXN'             , otString, fSpan,
     'PO_PC_SPECIAL_DAYS' , otCursor, null]);

   if fctempDataSet.FieldByName('DATE_DESCRIPTION').AsString.IsEmpty     and
      fctempDataSet.FieldByName('FRIENDLY_CREDIT_DATE').AsString.IsEmpty and
      fctempDataSet.FieldByName('CONFIGURATION').AsString.IsEmpty        then
   begin
    Result := false;
   end;

  finally
   FreeAndNil(fctempDataSet);
  end;

end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BOOST_SETTINGSClick(Sender: TObject);
begin
  // TODO  -   THIS WAS DISABLED PRIOR TO DCC Changes
  // by maryam on 04/11/2015 according to the Martin's email on the same date with title  Update on Helpdesk Call:314312
  { Application.CreateForm(TFRM_SMETS_ManageBoostSettings, FRM_SMETS_ManageBoostSettings);
    try
    FRM_SMETS_ManageBoostSettings.doquery(FRM_SMETS_DCC.span.Text);
    FRM_SMETS_ManageBoostSettings.showmodal;
    finally
    FRM_SMETS_ManageBoostSettings.release;
    end;
    ShowDataflowHistory; }
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Refreshdata;
var
  option         : boolean;
  suspendedUntil : TDate;
  commStatus     : variant;
  formHeight     : integer;
  q              : TOracleDataSet;
  makeType       : string;
  meterModeDesc  : string;
begin
  Self.Caption         := 'DCC SMETS - SPAN ' + fSpan;
  MeterRemoved.Visible := false;

  SetSupplyTypeIcon(fService, SupplyImage);

  if fService = SERVICE_ELECTRICITY then
  begin
    Group_Config.Height := 50; //90;
    SiteInfo1.enabled   := true;
    SiteInfo1.visible   := true;
    Req_Mirror.enabled  := false;
    Req_Mirror.visible  := false;
  end
  else
  begin
    Group_Config.Height := 50; //126;
    SiteInfo1.enabled   := false;
    SiteInfo1.visible   := false;
    Req_Mirror.enabled  := true;
    Req_Mirror.visible  := true;
  end;

  // Gas Configuration
  //Label_gas.visible := fService = SERVICE_GAS;
  //o_gc.visible      := fService = SERVICE_GAS;
  //efd_gas.visible   := fService = SERVICE_GAS;
  //btn_gas.visible   := fService = SERVICE_GAS;

  Group_Balance.Caption := 'Account Balance Summary';

  gSqlUtil.SelectQuery(LatestConfig,
    'select suppliercode, jobrequestno, supplytype, supplytype_description, servicepointno, deviceno, devicetype, '+         //6
           'devicetype_description, deviceinstallationdate, devicemfgno, metertype, metertype_description, metermode, '+     //12
           'metermode_description, metermode_efd, meter_firmware, chargingtype, chargingtype_description, chargingtype_efd, '+//18
           'tariffrefno, tariffrefno_description, tariffrefno_efd, vatgroupid, vatgroupid_description, vatgroupid_efd, '+    //24
           'gasconfigid, gasconfigid_description, gas_config_efd, prepayconfigid, prepayconfigid_description, '+             //29
           'prepayconfig_efd, fcsdayconfigid, fcsdayconfigid_description, fcsdayconfig_efd, fcweekconfigid, '+               //34
           'fcweekconfigid_description, fcweekconfig_efd, rdgscheduleid, rdgscheduleid_description, rdgscheduleid_efd, '+    //39
           'evtnotifycfgid, evtnotifycfgid_description, evtnotifycfgid_efd, tarifftype, tarifftype_description, '+           //44
           'tarifftype_efd, co2configid, co2configid_description, co2configid_efd, paymentcardid, paymentcardid_efd, '+      //50
           'customername, customername_efd, gatewayno, accountbalance, accountbalance_as_of, outstandingdebt, '+             //56
           'outstandingdebt_as_of, debtrecoveryrate, debtrecoveryrate_efd, currencycode, tarifflabel, meter_removed, '+      //62
           'meter_removed_date, gateway_status, gateway_status_date, supply_status, supply_status_date, smiff_id, '+         //68
           'tariff_line_1, tariff_line_2, last_refreshed, deviceguid, customer_id, registration_status, registration_date, '+//75
           'manufact_make_type, commission_in_progress, off_peak_hours '+                                                    //78
      'from ods.vw_sn_current_values '+
      'where servicepointno = :servicepointno',
    ['servicepointno', otString, fSpan]);

  Suspended.Visible := LatestConfig.FieldByName('debtrecoveryrate').AsFloat <= 0;

  if Suspended.Visible then
  begin
    suspendedUntil    := Check_For_Suspended_Debt_Dcc(fSpan);
    Suspended.Caption := IfThen(suspendedUntil > 0, 'Debt SUSPENDED until ' + FormatDateTime('dd/mm/yyyy', suspendedUntil));
  end;

  // Account balance
  Ab.Value := LatestConfig.FieldByName('accountbalance').AsFloat;
  // Outstanding debt
  Od.Value := LatestConfig.FieldByName('outstandingdebt').AsFloat;

  // Tariff Code
  TR.Visible := LatestConfig.FieldByName('tariffrefno').AsString <> '';

  // IF NO WAN, METER TYPE UNKNWON THEN DISABLE SOME OPTIONS
  Req_Read.Enabled  := o_mm.Caption <> 'UNKNOWN'; // On DEmand;

  // if MeterisRemoved then disable buttons
  if LatestConfig.FieldByName('deviceno').AsString =
    LatestConfig.FieldByName('meter_removed').AsString then
  begin
    MeterRemoved.Caption := FormatDateTime('dd/mm/yyyy', LatestConfig.FieldByName('meter_removed_date').AsDateTime) + ' REMOVED';
    MeterRemoved.visible := true;
  end;

  // If Meter in Credit Mode, disable Creditdebt options
  // debt is managed in wse so no need to disable this option, but without payment card cant issue debt anyway.
  // if not prepay
  Group_Customer.Visible := false;
  Group_Prepay.Visible   := false;
  Group_Config.Visible   := false;

  if TMeterMode(LatestConfig.FieldByName('metermode').AsInteger) <> mmPrepayment then
  begin
    o_mm.Color            := clYellow;
    Group_Balance.Caption := 'CREDIT MODE - Last Known Account Balance Summary';
  end
  else
  begin
    o_mm.Color             := clBtnFace;
    Group_Balance.Visible  := true;
    Group_Config.Visible   := true;
    Group_Prepay.Visible   := true;
    Group_Customer.Visible := true;
  end;

  Circle.Width := 1;

  gSqlUtil.ExecProc('ods.pk_crmui_metering.get_comms_status(:p_mpxn_in, :p_commstatus_out)', TRANSACTION_NO,
    ['p_mpxn_in',        otString,  pdInput,  fSpan,
     'p_commstatus_out', otInteger, pdOutput, @commStatus]);

  if VarIsNull(commStatus) then
    commStatus := 0; // safety

  if Integer(commStatus) > 0 then
  begin
    Circle.Digits.Visible := true;
    Circle.Width          := 53;
  end;

  case fService of
    SERVICE_ELECTRICITY:
    begin
      Circle.Position        := 0;
      Circle.Width           := 1;
      Gas_Wait_Timer.Enabled := false;
    end;
    SERVICE_GAS:
    begin
      Gas_Wait_Timer.Enabled := true;
    end;
  end;

  if commStatus = 0 then
    Smets_TXT.Enabled := false;

  option                 := true; // Anna: unnecessary as it always true
  Group_Block.Height     := 126;
  Group_Prepay.Height    := 126;
  Self.Width             := 624;
  Label_Firmware.Visible := true;
  Group_Config.Visible   := option;

  Group_Block.Height     := 75;
  Group_Prepay.Height    := 75;
  Group_Config.Height    := 50;

  // FCW
  Label17.Visible := option;
  o_fcw.Visible   := option;
  DB_FCW.Visible  := option;
  BTN_FCW.Visible := option;

  BitBtn21.Visible := option;
  BitBtn22.Visible := option;
  BitBtn23.Visible := option;

  if (Abs(LatestConfig.FieldByName('accountbalance_as_of').AsFloat - LatestConfig.FieldByName('outstandingdebt_as_of').AsFloat) < 0.00001) and
     (Abs(LatestConfig.FieldByName('accountbalance_as_of').AsFloat - LatestConfig.FieldByName('debtrecoveryrate_efd').AsFloat) < 0.00001) then
  begin
    DBText22.Visible := false;
    DBText23.Visible := false;
  end
  else
  begin
    DBText22.Visible := true;
    DBText23.Visible := true;
  end;

  formHeight := 46 {menu height} + Panel_MPAN.Height + Group_Meter.Height;

  if Group_Customer.Visible then
    formHeight := formHeight + Group_Customer.Height;

  if Group_Block.Visible then
    formHeight := formHeight + Group_Block.Height;

  if Group_Prepay.Visible then
    formHeight := formHeight + Group_Prepay.Height;

  if Group_Config.Visible then
    formHeight := formHeight + Group_Config.Height;

  if Group_Balance.Visible then
    formHeight := formHeight + Group_Balance.Height;

  if Group_DataFlow_History.Visible then
    formHeight := formHeight + Group_DataFlow_History.Height;

  Self.Height := formHeight;

  {ISC-722 (Anna) This was in the original code, however it didn't do anything apart from opening a view
  q := gSqlUtil.SelectQuery(
    'select suppliercode, jobrequestno, supplytype, servicepointno, switchtime, priority, requestexpirydate, '+
           'requestretryinterval, requestingapp, requestingappuser, requestdate, testrequest, processstatus, supplierownref, '+
           'rq_responsedate, rq_process_status, rq_state, rq_errorcodes, cos_gatewayno, cos_deviceno, cos_old_supplier, '+
           'cos_suppliername, cos_wsetime, cos_createdon, cfg_jobrequestno, cfg_responsedate, cfg_process_status, cfg_state, '+
           'cfg_errorcodes, cfg_request_status, step_1_desc, step_2_desc, internal_status, transfer_done, g_or_e, '+
           'sn_suppliercode, sn_prepayconfigid, cfg_first_error, cos_status, cos_effective_date, payment_card_id, is_live_span '+
      'from liberty100.vw_lib100_cos_gain_status '+
      'where cfg_state is null and is_live_span is not null and servicepointno = :servicepointno',
    ['servicepointno', otString, fSpan]);
  try
    //
  finally
    FreeAndNil(q);
  end;}

  DoFlowHistory(fSpan);

  if LatestConfig.FieldByName('registration_status').AsString = 'LOST' then
  begin
    Self.Caption := Format('SPAN %s - LOST in %s on %s',
      [ fSpan, LatestConfig.FieldByName('metermode_description').AsString,
        FormatDateTime('dd/mm/yyyy',
        LatestConfig.FieldByName('registration_date').AsDateTime) ]);
  end;

  if LatestConfig.FieldByName('commission_in_progress').AsString = 'N' then
  begin
    o_cms.Caption  := 'COMMISSIONED';
    PanelCMS.Width := 90;
    PanelCMS.Left  := 512;
    PanelCMS.Color := clWebLightGreen;
  end
  else
  begin
    o_cms.Caption  := 'NOT COMMISSIONED';
    PanelCMS.Width := 115;
    PanelCMS.Left  := 487;
    PanelCMS.Color := clWebOrange;
  end;

  makeType                  := UpperCase(LatestConfig.FieldByName('manufact_make_type').AsString);
  meterModeDesc             := UpperCase(LatestConfig.FieldByName('metermode_description').AsString);
  EnableEmergCredit.Enabled := makeType.Contains('SECURE') and meterModeDesc.Contains('PREPAYMENT');
  EnableEmergCredit.Visible := EnableEmergCredit.Enabled;

  if (not String.IsNullOrEmpty(LatestConfig.FieldByName('off_peak_hours').AsString)) then
  begin
    Group_Block.Height:= Group_Block.Height + 25;
    o_opt.Visible := true;
    o_opt.Top := 75;
    lbOffPeakTimes.Visible := true;
    lbOffPeakTimes.Top := 75;
  end
  else
  begin
    o_opt.Visible := false;
    lbOffPeakTimes.Visible := false;
  end;

  SetLatestDemand;
end;

procedure TFrm_Smets_Dcc.SetLatestDemand;
var
  onDemandReq  : variant;
  amountReq    : variant;
  lastDemError : variant;
  weeklyMsg    : variant;
begin
  try
    gSqlUtil.ExecProc('ods.pk_crmui_metering.get_last_on_demand(:p_mpxn_in, :p_on_demand_out, :p_amount_restore_out, :p_error_message_out, :p_weekly_message_out)', TRANSACTION_NO,
      ['p_mpxn_in',            otLong,   pdInput,  StrToInt64(fSpan),
       'p_on_demand_out',      otString, pdOutput, @onDemandReq,
       'p_amount_restore_out', otFloat,  pdOutput, @amountReq,
       'p_error_message_out',  otString, pdOutput, @lastDemError,
       'p_weekly_message_out', otString, pdOutput, @weeklyMsg]);

    if VarToStr(lastDemError) = '' then
    begin
      LabelAmountRestore.Visible := true;
      LabelDemandReq.Visible     := true;
      LabelLastDemError.Visible  := false;

      LabelAmountRestore.Caption := 'Amount to Restore: ' + FloatToStr(amountReq);
      LabelDemandReq.Caption     := 'On Demand Request: ' + VarToStr(onDemandReq);
    end
    else
    begin
      LabelLastDemError.Caption  := lastDemError;
      LabelLastDemError.Width    := 200;
      LabelLastDemError.Height   := 22;

      LabelAmountRestore.Visible := false;
      LabelDemandReq.Visible     := false;
      LabelLastDemError.Visible  := true;
    end;

    LabelWeekLimit.Caption := VarToStr(weeklyMsg);
  except
    raise Exception.Create('Error: Unable to retrieve latest Demand Amount');
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Req_VendClick(Sender: TObject);
begin
  Messagedlg('The Vend History option is not currently developed for DCC meters ',mtinformation, [mbOK],0);
  // SMI TODO
 { if DO_SMETS_SUPPLIER_CHECK(Span.text) = false then
    exit;

  Application.CreateForm(TFRM_SMETS_VENDHISTORY_JRQ, FRM_SMETS_VENDHISTORY_JRQ);
  try
    FRM_SMETS_VENDHISTORY_JRQ.Refreshdata(Span.text,
      inttostr(Service.ItemIndex), MeterNo.caption);
    FRM_SMETS_VENDHISTORY_JRQ.T_Card.text := o_cn.caption;
    FRM_SMETS_VENDHISTORY_JRQ.ShowModal;
  finally
    FRM_SMETS_VENDHISTORY_JRQ.release;
  end;
  if FRM_SMETS_VENDHISTORY_JRQ.tag = 1 then
    ShowDataflowHistory;}
end;

procedure TFRM_SMETS_DCC.RestrictStaffTariffPaymentsAndControls;
var
  bRestrict: Boolean;
begin
  bRestrict := True;
  if Pos('STAFF', AnsiUpperCase(LATESTCONFIG.FieldByName('TARIFFREFNO_DESCRIPTION').AsString)) > 0 then
  begin
    if RESTRICT_STAFF_TARIFF_PAYMENTS_AND_CONTROLS = 'Y' then
    begin
      bRestrict := USER_FEATURE__UPDATE_STAFF_ACCOUNTS;
    end;
  end;

  FreeVend.Visible := bRestrict;
  Smets_Loan.Visible := bRestrict;
  ManageDEbt.Visible := bRestrict;
  EnableEmergCredit.Visible := bRestrict;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Req_ReadClick(Sender: TObject);
begin
  if TFrm_Smets_On_Demand_Read_Dcc.StartModal(Self, fService, fSpan, MeterNo.Caption) then
    ShowDataflowHistory;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.SendTXTMessage1Click(Sender: TObject);
begin
  // SMI TODO
    Messagedlg('The Send Text option is not currently developed for DCC meters ',mtinformation, [mbOK],0);
{  if DO_SMETS_SUPPLIER_CHECK(Span.text) = false then
    exit;

  if GetSmetsWanStatus(Span.text, '') <> 'ON' then
  Begin
    Messagedlg
      ('Unable to action reqeust. There does not appear to be any active Remote Comms to this Supply',
      mtwarning, [mbok], 0);
    exit;
  end;
  Application.CreateForm(TFRM_SMETS_TEXTMSG, FRM_SMETS_TEXTMSG);
  try
    FRM_SMETS_TEXTMSG.Refreshdata(Span.text, inttostr(Service.ItemIndex),
      MeterNo.caption);
    FRM_SMETS_TEXTMSG.sendnow.checked := true;
    FRM_SMETS_TEXTMSG.T_Date.enabled := false;
    FRM_SMETS_TEXTMSG.T_Time.enabled := false;
    FRM_SMETS_TEXTMSG.TYPELOOKUP.keyvalue := 0;
    FRM_SMETS_TEXTMSG.ShowModal;
  finally
    FRM_SMETS_TEXTMSG.release;
  end;
  if FRM_SMETS_TEXTMSG.tag = 1 then
    ShowDataflowHistory;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.SMETS_CONFIGClick(Sender: TObject);
begin
  if LatestConfig.FieldByName('metermode').IsNull then
  begin
    MessageDlg('Missing meter mode. The configuration cannot be changed. Contact support.', mtError, [mbOk], 0);
    exit;
  end;

  TFrm_Smets_Change_Configuration_Dcc.StartModal(
    Self,
    fService,
    fSpan,
    TMeterMode(LatestConfig.FieldByName('metermode').AsInteger),
    LatestConfig.FieldByName('customer_id').AsLargeInt,
    fAgreementId, fEfsdmsMtd);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.VendHistory1Click(Sender: TObject);
begin
  TFrm_Smets_Vends_Dcc.Start(Self, fService, fSpan);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.DataflowHistory1Click(Sender: TObject);
begin
  TFrm_Smets_Flow_History_Dcc.Start(Self, fService, fSpan, false);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.DelayButton;
begin
  BitBtn1.Enabled := false;
  Sleep(1000);
  BitBtn1.Enabled := true;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.AutoRefreshClick(Sender: TObject);
begin
  //RefreshTimer.enabled := AutoRefresh.checked;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.CommsData1Click(Sender: TObject);
begin
  // SMI TODO
  Messagedlg('The Comms Data option is not currently developed for DCC meters ',mtinformation, [mbOK],0);
{  ShowSmetsComms(inttostr(Service.ItemIndex), Span.text, '0');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.ConnectSupplyArmed1Click(Sender: TObject);
begin
  Messagedlg('The Re-Connect option is not currently developed for DCC meters ',mtinformation, [mbOK],0);
  // SMI TODO
  {if Messagedlg('Are you sure you wish to Re-connect Supply (Arm)?',
    mtconfirmation, [mbyes, mbno], 0) <> mryes then
    exit;
  if SupplyControl(inttostr(Service.ItemIndex), Span.text, MeterNo.caption, 'Y',
    '1') = true then
    ShowDataflowHistory;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.COSStatus1Click(Sender: TObject);
begin
  // FRM_SMETS_DCC_cos_gain.ShowModal;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Exit1Click(Sender: TObject);
begin
  Close;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.FlowGridDblClick(Sender: TObject);
begin
  DataflowHistory1Click(Sender);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.FlowGridGetCellColor(Sender: TObject; ARow, ACol: integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
var
  ebcolor, rowcolor: tcolor;
  statecode, reqstatuscode, processstatus, flow: string;
begin
  if (ARow = 0) or (ACol = 0) then
    exit;

  flow := FlowGridDCC.cells[1, ARow];
  statecode := FlowGridDCC.cells[5, ARow];
  reqstatuscode := FlowGridDCC.cells[4, ARow];
  processstatus := FlowGridDCC.cells[6, ARow];

  ebcolor := clwhite;
  rowcolor := clblue;

  if statecode = '999' then
  Begin
    ebcolor := clyellow;
    rowcolor := clblack;
  End;

  if processstatus = '-2' then
  Begin
    ebcolor := clyellow;
  end;

  if statecode = '200' then
  Begin
    if reqstatuscode = '0' then
    Begin
      rowcolor := clblack; // OK
      if flow = 'CoS Gain' then
        rowcolor := clgreen;
    end

    else if reqstatuscode = '-1' then
      rowcolor := clpurple
    else
      rowcolor := clred; // Failed
  end;

  if ebcolor = clyellow then
    AFont.Style := AFont.Style + [fsstrikeout]
  else
    AFont.Style := AFont.Style - [fsstrikeout];

  AFont.Color := rowcolor;
  ABrush.Color := ebcolor;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.FlowHistoryAfterOpen(DataSet: TDataSet);
begin
 //
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.ihd_pinClick(Sender: TObject);
begin
  Messagedlg('The Generate IDH Pin option is not currently developed for DCC meters ',mtinformation, [mbOK],0);
  // SMI TODO

{  FRM_IHD_PIN.MACADDRESS.text := FRM_IHD_PIN.GETIHDMAC(Span.text);
  FRM_IHD_PIN.ShowModal;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Ihd_rebindClick(Sender: TObject);
begin
  PerformIhdRepair;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.PerformIhdRepair;
begin
  TFrm_Smets_Device_Dcc.PerformRebind(Self, fService, fSpan, fAgreementId);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.TB_4Click(Sender: TObject);
begin
  Req_VendClick(Sender);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.TB_5Click(Sender: TObject);
begin
  Req_ReadClick(Sender);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.TB_6Click(Sender: TObject);
begin
  SendTXTMessage1Click(Sender);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.TB_7Click(Sender: TObject);
begin
  SMETS_CONFIGClick(Sender);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.TB_8Click(Sender: TObject);
begin
  DataflowHistory1Click(Sender);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.TB_9Click(Sender: TObject);
begin
  // SMI TODO
  VendHistory1Click(Sender);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Gas_Wait_TimerTimer(Sender: TObject);
var
  m: integer;
begin
  m := strtoint(formatdatetime('nn', now));
  if m < 30 then
    Circle.position := 30 - m
  else
    Circle.position := (60 - m);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.RefreshTimerTimer(Sender: TObject);
begin
  try
//
  except
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.TB_10Click(Sender: TObject);
begin
  Messagedlg('The Comms Data option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {CommsData1Click(Sender);}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BitBtn1Click(Sender: TObject);
begin
  ShowDataflowHistory;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BitBtn23Click(Sender: TObject);
begin
  ShowItem('', ' Account Balance', 'Total');
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.ShowItem(aTable, aDescription, aTitleCaption: string);
begin
  TFrm_Smets_Item_History.StartModal(
    Self,
    rtSmetDcc,
    Frm_Login.MainSession,
    fService,
    fSpan,
    MeterNo.Caption,
    '',
    aDescription,
    aTitleCaption);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BitBtn22Click(Sender: TObject);
begin
  TFrm_Smets_Balance_Debt_Dcc.StartModal(Self, fService, fSpan, MeterNo.Caption);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BitBtn21Click(Sender: TObject);
begin
  Messagedlg('The Debt Recovery Rate option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
{  ShowItem('SP_DEBT_RECOVRY_RATE', 'Debt Recovery Rate', 'Recovery Rate');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BTN_TariffClick(Sender: TObject);
begin
  TFrm_Smets_Item_History_Dcc.StartModal(Self, rtTariff, fSpan, MeterNo.Caption);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BTN_VATClick(Sender: TObject);
begin
  Messagedlg('The Vat option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {ShowItem('SP_VAT_GROUP_ID', 'Vat', 'Vat Code');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BTN_CTClick(Sender: TObject);
begin
  Messagedlg('The Meter Charging Type option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
{  ShowItem('SP_CHARGING_TYPE', 'Meter Charging Type', 'Code');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BTN_PPClick(Sender: TObject);
begin
  Messagedlg('The PrePayment Configuration option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
 { ShowItem('SP_PREPAY_CONFIG', 'Prepayment Configuration', 'Value');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BTN_FCDClick(Sender: TObject);
begin
  Messagedlg('The Friendly Credit Special Days option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
{  ShowItem('SP_FCS_DAY', 'Friendly Credit Special Days', 'Config Id');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BTN_FCWClick(Sender: TObject);
begin
  TFrm_Smets_Item_History_Dcc.StartModal(Self, rtFriendlyCreditWeek, fSpan, MeterNo.CAption);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BTN_MODEClick(Sender: TObject);
begin
  TFrm_Smets_Item_History_Dcc.StartModal(Self, rtMeterMode, fSpan, MeterNo.CAption);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BalanceDebtHistory1Click(Sender: TObject);
begin
  TFrm_Smets_Balance_Debt_Dcc.StartModal(Self, fService, fSpan, MeterNo.Caption);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BitBtn16Click(Sender: TObject);
begin
  Messagedlg('The Payment Card Number option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {ShowItem('SP_PAYMENT_CARD', 'Payment Card Number', 'Card Number');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BitBtn17Click(Sender: TObject);
begin
  Messagedlg('The Reading Schedule Configuration option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {ShowItem('SP_READ_SCHEDULE', 'Reading Schedule Configuration', 'Schedule Id');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BitBtn18Click(Sender: TObject);
begin
  Messagedlg('The Event Notifications Configuration option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {ShowItem('SP_EVENT_CONFIG', 'Event Notifications Configuration', 'Config Id');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.btn_gasClick(Sender: TObject);
begin
  Messagedlg('The Gas Configuration option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {ShowItem('SP_GAS_CONFIG', 'Gas Configuration', 'Config Id');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.btn_co2Click(Sender: TObject);
begin
  Messagedlg('The Co2 Configuration option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {ShowItem('SP_CO2_CONFIG', 'Co2 Configuration', 'Config Id');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.SMETS_CANCClick(Sender: TObject);
begin
  Messagedlg('The Cancel Delayed option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
{  if DO_SMETS_SUPPLIER_CHECK(Span.text) = false then
    exit;

  Application.CreateForm(TFRM_SMETS_CANCEL_DELAYED, FRM_SMETS_CANCEL_DELAYED);
  try
    FRM_SMETS_CANCEL_DELAYED.Refreshdata(Span.text, inttostr(Service.ItemIndex),
      MeterNo.caption);
    if FRM_SMETS_CANCEL_DELAYED.delayedquery.recordcount <> 0 then
      FRM_SMETS_CANCEL_DELAYED.ShowModal;
  finally
    FRM_SMETS_CANCEL_DELAYED.release;
  end;
  if FRM_SMETS_CANCEL_DELAYED.tag = 1 then
    ShowDataflowHistory;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.tb_15Click(Sender: TObject);
begin
  SMETS_CANCClick(Sender);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BTN_TTClick(Sender: TObject);
begin
  Messagedlg('The Tariff Type option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
 { ShowItem('SP_TARIFF_TYPE', 'Tariff Type', 'Tariff Type Id');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.SMETS_COTClick(Sender: TObject);
begin
  // SMI TODO
  Messagedlg('The COT option is not currently active for DCC meters ',mtinformation, [mbOK],0);

{  if DO_SMETS_SUPPLIER_CHECK(Span.text) = false then
    exit;

  Application.CreateForm(TFRM_SMETS_COT, FRM_SMETS_COT);
  try
    FRM_SMETS_COT.Refreshdata(Span.text, inttostr(Service.ItemIndex),
      MeterNo.caption);
    FRM_SMETS_COT.ShowModal;
  finally
    FRM_SMETS_COT.release;
  end;
  if FRM_SMETS_COT.tag = 1 then
    ShowDataflowHistory;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.tb_13Click(Sender: TObject);
begin
  SMETS_COTClick(Sender);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.tb_14Click(Sender: TObject);
begin
  Messagedlg('The IHD Replace option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {ihd_replaceClick(Sender);}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Req_DebtClick(Sender: TObject);
begin
   Messagedlg('The Debt option is not currently active for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO

{  if DO_SMETS_SUPPLIER_CHECK(Span.text) = false then
    exit;

  if GetSMetsMeterMode(Span.text) <> '1' then
  Begin
    Messagedlg
      ('This option is only allowed if Current SMETS Meter is in PRE-PAYMENT Mode',
      mtwarning, [mbok], 0);
    exit;
  End;

  if GetSmetsPaymentCardno(Span.text, '') = '' then
  Begin
    Messagedlg('This option is NOT allowed. No Topup Card is Registered',
      mtwarning, [mbok], 0);
    exit;
  End;

  If Messagedlg
    ('Do you wish to Request the current Debt Position for this Meter?',
    mtconfirmation, [mbyes, mbno], 0) <> mryes then
    exit;
  if RequestDebt(inttostr(Service.ItemIndex), Span.text, MeterNo.caption,
    o_cn.caption, '') = true then
    Messagedlg('Request for Debt position has been generated.', mtinformation,
      [mbok], 0);
  ShowDataflowHistory;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.REQ_DiagClick(Sender: TObject);
begin
  // SMI TODO
  Messagedlg('The Wan Status option is not currently active for DCC meters ',mtinformation, [mbOK],0);


{  if DO_SMETS_SUPPLIER_CHECK(Span.text) = false then
    exit;

  if GetSmetsWanStatus(Span.text, '') = 'OFF' then
  Begin
    Messagedlg
      ('Unable to Read Meter. There does not appear to be any active Remote Comms to this Supply',
      mtwarning, [mbok], 0);
    exit;
  end
  else if GetSmetsWanStatus(Span.text, '') <> 'ON' then
  Begin
    if Messagedlg
      ('Communications to this meter may be slow or unreliable. Continue anyway?',
      mtwarning, [mbyes, mbno], 0) <> mryes then
      exit;
  end;

  Application.CreateForm(TFRM_SMETS_METER_DIAG, FRM_SMETS_METER_DIAG);
  try
    if Service.ItemIndex = 0 then
    begin
      FRM_SMETS_METER_DIAG.GRP_aux.visible := true;
      FRM_SMETS_METER_DIAG.GRP_battery.visible := false;
    end
    else
    begin
      FRM_SMETS_METER_DIAG.GRP_aux.visible := false;
      FRM_SMETS_METER_DIAG.GRP_battery.visible := true;
    end;
    FRM_SMETS_METER_DIAG.device.caption := MeterNo.caption;
    FRM_SMETS_METER_DIAG.doquery;
    FRM_SMETS_METER_DIAG.Height := 240;
    FRM_SMETS_METER_DIAG.ShowModal;
  finally
    FRM_SMETS_METER_DIAG.release;
  end;
  ShowDataflowHistory;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.tb_11Click(Sender: TObject);
begin
  Req_DebtClick(Sender);
end;

procedure TFrm_Smets_Dcc.device_removeClick(Sender: TObject);
begin
  Messagedlg('The Remove Device option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
 { RemoveDevice;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.RemoveDevice;
var
  resOk : boolean;
begin
  // SMI TODO
  if not Do_Smets_Supplier_Check(fSpan) then
    exit;

  Application.CreateForm(TFRM_SMETS_REMOVE_DEVICE, FRM_SMETS_REMOVE_DEVICE);
  try
    FRM_SMETS_REMOVE_DEVICE.Refreshdata(fSpan, inttostr(fService), MeterNo.Caption, '');
    FRM_SMETS_REMOVE_DEVICE.IHD_GROUP.visible := false;
    FRM_SMETS_REMOVE_DEVICE.HAN_GROUP.visible := false;
    FRM_SMETS_REMOVE_DEVICE.checknow.checked := false;
    FRM_SMETS_REMOVE_DEVICE.caption := 'Remove Meter';
    FRM_SMETS_REMOVE_DEVICE.Height := 174;
    FRM_SMETS_REMOVE_DEVICE.ShowModal;
    resOk := FRM_SMETS_REMOVE_DEVICE.Tag = 1;
  finally
    FRM_SMETS_REMOVE_DEVICE.release;
  end;

  if resOk then
    ShowDataflowHistory;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.MeterReadings1Click(Sender: TObject);
begin
  TFrm_Smets_Readings_Dcc.Start(Self, fSpan);
end;

procedure TFrm_Smets_Dcc.FC_UpdateClick(Sender: TObject);
begin
  TFrm_Friendly_Credit_Update.StartModal(Self, fSpan);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.BitBtn2Click(Sender: TObject);
begin
  Messagedlg('The Customer Name option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {ShowItem('SP_CUSTOMER_NAME', 'Customer Name', 'Customer Name');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.EventAlarms1Click(Sender: TObject);
begin
  TFrm_Smets_Alarms_Dcc.Start(Self, fSpan);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.ihd_addClick(Sender: TObject);
begin
  TFrm_Smets_Device_Dcc.StartModal(Self, fService, fSpan, fAgreementId);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.txtMessageHistory1Click(Sender: TObject);
begin
  Messagedlg('The Show Smets SMS option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {ShowSmetsSMS(Span.text, '0');}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.PPMIDMessages1Click(Sender: TObject);
begin
  TFrm_Smets_Message_History_Dcc.StartModal(Self, fSpan);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.ProfileData1Click(Sender: TObject);
begin
  TFrm_Smets_Profile.Start(Self, fService, fSpan, MeterNo.Caption);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.ToolButton5Click(Sender: TObject);
begin
  ProfileData1Click(Sender);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.ToolButton3Click(Sender: TObject);
begin
  //txtMessageHistory1Click(Sender);
  PPMIDMessages1Click(Sender)
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.ToolButton4Click(Sender: TObject);
begin
  TFrm_Smets_Balance_Debt_Dcc.StartModal(Self, fService, fSpan, MeterNo.Caption);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.DoSmets2Credit(aMode: integer);
begin
  if GetMeterMode(fSpan) <> IntToStr(Ord(mmPrepayment)) then
  begin
    MessageDlg('This option is only allowed if Current SMETS2 Meter is in PRE-PAYMENT Mode.', mtError, [mbOk], 0);
    exit;
  end;

  if TSmetsUtil.HasRpFlags(LatestConfig.FieldByName('customer_id').AsLargeInt) and (not Frm_Common.SuperAuthorityCheck) then
  begin
    MessageDlg('Super Authority check has failed', mtError, [mbOk], 0);
    exit;
  end;

  TFrm_Smets_Manage_Credit_Dcc.StartModal(Self, fService, fSpan, MeterNo.Caption, aMode)
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.smets_vend_addClick(Sender: TObject);
begin
  DoSmets2Credit(0);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.smets_vend_deductClick(Sender: TObject);
begin
  DoSmets2Credit(1);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.DoSmets2Debt(aMode: integer);
  {----------------------------------------------------------------------------}
  function IsValid(aMode: integer; out oErrorMessage: string): boolean;
  var
    requestDate    : TDateTime;
    suspendedUntil : TDate;
    msg            : string;
  begin
    oErrorMessage := '';

    if GetMeterMode(fSpan) <> IntToStr(Ord(mmPrepayment)) then
    begin
      oErrorMessage := 'This option is only allowed if Current Meter is in PRE-PAYMENT Mode';
      exit(false);
    end;

    if GetPaymentCardNo(fSpan, '') = '' then
    begin
      oErrorMessage := 'This option is NOT allowed. No Topup Card is Registered';
      exit(false);
    end;

    if aMode = 5 then
    begin
      if Check_For_Suspended_Debt(fSpan, MeterNo.Caption, requestDate) <> 0 then
      begin
        oErrorMessage := 'There is already a suspended Debt on this account, '+
                         'that is to be reactivated on ' + FormatDateTime('dd/mm/yyyy', requestDate)+'.';
        exit(false);
      end;

      if Dr.Caption = '0' then
      begin
        oErrorMessage := 'You cannot suspend debt when current recovery rate is set to ZERO.';
        exit(false);
      end;
    end
    else
    begin
      suspendedUntil := Check_For_Suspended_Debt_Dcc(fSpan);
      if suspendedUntil > 0 then
      begin
        msg := 'There is a suspended Debt on this account. ' +
               'This must be Cancelled before any Debt Changes can be made.' + #13 + #13 +
               'Do you wish to Remove this Suspended Debt?';
        if MessageDlg(msg, mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
          exit(false); // No error message

        try
          RemoveSuspendedDebtDCC(fSpan, suspendedUntil);
        except
          on e:Exception do
          begin
            if Pos('API is waiting', oErrorMessage) > 0 then
              oErrorMessage := 'Suspend Debt API is waiting for response from TMA.'
            else
              oErrorMessage := Format('There was an error submitting this request. Please check data and try again. (%s)', [e.Message]);

            exit(false);
          end;
        end;
      end;
    end;

    Result := true;
  end;
  {----------------------------------------------------------------------------}
var
  errorMsg : string;
begin
  if not IsValid(aMode, errorMsg) then
  begin
    if errorMsg <> '' then
      MessageDlg(errorMsg, mtError, [mbOk], 0);

    exit;
  end;

  TFrm_Smets_Manage_Debt_Dcc.StartModal(
    Self,
    rqSmetsDcc,
    fService,
    fSpan,
    MeterNo.Caption,
    aMode,
    Od.Value,
    LatestConfig.FieldByName('debtrecoveryrate').AsFloat,
    true);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.smets_debt_addClick(Sender: TObject);
begin
  DoSmets2Debt(0);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.smets_debt_deductClick(Sender: TObject);
begin
  DoSmets2Debt(1);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.smets_debt_setClick(Sender: TObject);
begin
  Messagedlg('The Set Debt option is not currently developed for DCC meters ',mtinformation, [mbOK],0);
  // SMI TODO
  {  DoSmets2Debt(2);}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Req_SNAPClick(Sender: TObject);
var
  cardno: string;
begin
  Messagedlg('The Vend History option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO

  {if DO_SMETS_SUPPLIER_CHECK(Span.text) = false then
    exit;

  if GetSmetsWanStatus(Span.text, '') = 'OFF' then
  Begin
    Messagedlg
      ('Unable to Read Meter. There does not appear to be any active Remote Comms to this Supply',
      mtwarning, [mbok], 0);
    exit;
  end
  else if GetSmetsWanStatus(Span.text, '') <> 'ON' then
  Begin
    if Messagedlg
      ('Communications to this meter may be slow or unreliable. Continue anyway?',
      mtwarning, [mbyes, mbno], 0) <> mryes then
      exit;
  end;

  cardno := GetSmetsPaymentCardno(Span.text, '');
  if cardno = '' then
  Begin
    Messagedlg
      ('There Does not appear to be a Valid payment Card id on this supply',
      mtwarning, [mbok], 0);
    exit;
  end;
  If Messagedlg('Are you sure you wish to Read this Meter to obtain most recent'
    + #13 + 'Vend History, Debt, Account Balance && Meter Reading.' + #13 + #13
    + 'This may take a few minutes?', mtconfirmation, [mbyes, mbno], 0) <> mryes
  then
    exit;
  Request_VendHistory(inttostr(Service.ItemIndex), Span.text, MeterNo.caption,
    cardno, '', '', '');
  Insert_OnDemand_Read(inttostr(Service.ItemIndex), Span.text, MeterNo.caption,
    '', '', '0', datetostr(now), datetostr(now), '', '');

  Messagedlg('Your Request has been submitted.' + #13 +
    'Please Check the SMETS Dataflow History to check on the progress of this request.',
    mtinformation, [mbok], 0);
  ShowDataflowHistory;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.smets_ihd_replaceClick(Sender: TObject);
begin
  // SMI TODO
  Application.CreateForm(TFRM_SMETS_REMOVE_DEVICE, FRM_SMETS_REMOVE_DEVICE);
  try
    FRM_SMETS_REMOVE_DEVICE.Refreshdata(fSpan, inttostr(fService), MeterNo.caption, 'Y');
    FRM_SMETS_REMOVE_DEVICE.IHD_GROUP.visible := true;
    FRM_SMETS_REMOVE_DEVICE.HAN_GROUP.visible := true;
    FRM_SMETS_REMOVE_DEVICE.Height := 340;
    FRM_SMETS_REMOVE_DEVICE.checknow.checked := true;
    FRM_SMETS_REMOVE_DEVICE.caption := 'Replace IHD';
    FRM_SMETS_REMOVE_DEVICE.ShowModal;
  finally
    FRM_SMETS_REMOVE_DEVICE.release;
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.smets_ihd_addClick(Sender: TObject);
begin
  // SMI TODO
  Application.CreateForm(TFRM_SMETS_REMOVE_DEVICE, FRM_SMETS_REMOVE_DEVICE);
  try
    FRM_SMETS_REMOVE_DEVICE.Refreshdata(fSpan, inttostr(fService), MeterNo.caption, 'Y');
    FRM_SMETS_REMOVE_DEVICE.IHD_GROUP.visible := false;
    FRM_SMETS_REMOVE_DEVICE.HAN_GROUP.visible := true;
    FRM_SMETS_REMOVE_DEVICE.checknow.checked := true;
    FRM_SMETS_REMOVE_DEVICE.caption := 'Add IHD';
    FRM_SMETS_REMOVE_DEVICE.Height := 278;
    FRM_SMETS_REMOVE_DEVICE.ShowModal;
  finally
    FRM_SMETS_REMOVE_DEVICE.release;
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.smets_ihd_removeClick(Sender: TObject);
begin
  // SMI TODO
  Application.CreateForm(TFRM_SMETS_REMOVE_DEVICE, FRM_SMETS_REMOVE_DEVICE);
  try
    FRM_SMETS_REMOVE_DEVICE.Refreshdata(fSpan, inttostr(fService), MeterNo.caption, 'Y');
    FRM_SMETS_REMOVE_DEVICE.IHD_GROUP.visible := true;
    FRM_SMETS_REMOVE_DEVICE.HAN_GROUP.visible := false;
    FRM_SMETS_REMOVE_DEVICE.checknow.checked := true;
    FRM_SMETS_REMOVE_DEVICE.caption := 'Remove IHD';
    FRM_SMETS_REMOVE_DEVICE.Height := 238;
    FRM_SMETS_REMOVE_DEVICE.ShowModal;
  finally
    FRM_SMETS_REMOVE_DEVICE.release;
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.smets_txtClick(Sender: TObject);
begin
  // SMI TODO
  Messagedlg('The Text option is not currently developed for DCC meters ',mtinformation, [mbOK],0);


{  if GetSmetsWanStatus(Span.text, '') <> 'ON' then
  Begin
    Messagedlg
      ('Unable to action reqeust. There does not appear to be any active Remote Comms to this Supply',
      mtwarning, [mbok], 0);
    exit;
  end;
  Application.CreateForm(TFRM_SMETS_TEXTMSG, FRM_SMETS_TEXTMSG);
  try
    FRM_SMETS_TEXTMSG.Refreshdata(Span.text, inttostr(Service.ItemIndex),
      MeterNo.caption);
    FRM_SMETS_TEXTMSG.sendnow.checked := true;
    FRM_SMETS_TEXTMSG.T_Date.enabled := false;
    FRM_SMETS_TEXTMSG.T_Time.enabled := false;
    FRM_SMETS_TEXTMSG.TYPELOOKUP.keyvalue := 0;
    FRM_SMETS_TEXTMSG.ShowModal;
  finally
    FRM_SMETS_TEXTMSG.release;
  end;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.SMETS_MODEClick(Sender: TObject);
var
  msgForm    : TUELMessageForm;
  isEligible : boolean;
begin
  // Anna: this is quite slow (IsSmetsEAFullyMigrated), thus I added the warning message window
  // this function retrieves information from two views (electric/gas). They are the parts and both are queried.
  // This should be moved into a stored procedure/function and only that one should be used which
  // is required depending on fService

  msgForm := TUELMessageForm.ShowUELMessage(Self, mtWarning, 'Checking eligibility...' + #13#10#13#10 + 'This might take a few minutes!');
  try
    isEligible := not (IsSmetsEAFullyMigrated(fSpan, MeterNo.caption) and (GetMeterMode(fSpan) = '0') and
                      (not TCrmUtil.HasUserFeature(UserId, USER_FEATURE__E_A_CHANGE_METER_MODE_PREPAY_CREDIT)));
  finally
    FreeAndNil(msgForm);
  end;

  if not isEligible then
  begin
    MessageDlg('Unable to Mode Change a SMETS1 E&A meter. Meter must remain as Credit.', mtError, [mbOk], 0);
    exit;
  end;

  if not IsManufacturerAllowed then
    exit;

  if TFrm_Smets_Change_Mode_Dcc.StartModal(Self, fService, fSpan, MeterNo.Caption, LatestConfig.FieldByName('customer_id').AsLargeInt, fAgreementId, fEfsdmsMtd) then
    ShowDataflowHistory;
end;

{------------------------------------------------------------------------------}
function TFrm_Smets_Dcc.IsManufacturerAllowed: boolean;
var
  comAllowed : variant;
  msg        : variant;
begin
  Result := false;

  try
    gSqlUtil.ExecProc('ods.pk_crmui_metering.check_allow_change_of_mode(:p_deviceno_in, :p_isallow_out, :p_error_message_out)', TRANSACTION_YES,
      ['p_deviceno_in',       otString, pdInput,  MeterNo.Caption,
       'p_isallow_out',       otString, pdOutput, @comAllowed,
       'p_error_message_out', otString, pdOutput, @msg]);

    case IndexStr(VarToStr(comAllowed), ['Y', 'N', 'P']) of
      0: Result := true;
      1: Result := HandleNotAllowed(msg);
      2: Result := HandleProcodeMeter(msg);
    end;
  except
    on e: Exception do
      raise Exception.Create('Error: Unable to execute the Change of Mode. ' + e.Message);
  end;
end;

{------------------------------------------------------------------------------}
function TFrm_Smets_Dcc.HandleNotAllowed(aComError: string): boolean;
begin
  if USER_FEATURE__CHANGE_MODE_UNSUPPORTED then
  begin
    Result := MessageDlg('This meter has an unsupported manufacturer. Do you still want to execute Change of Mode?', mtConfirmation, [mbYes, mbNo], 0) = mrYes;
  end
  else
  begin
    MessageDlg(aComError, mtError, [mbOk], 0);
    Result := false;
  end;
end;

{------------------------------------------------------------------------------}
function TFrm_Smets_Dcc.HandleProcodeMeter(aComError: string): boolean;
  {----------------------------------------------------------------------------}
  procedure ShellOpen(const Url: string; const Params: string = '');
  begin
    ShellApi.ShellExecute(0, 'Open', PChar(Url), PChar(Params), nil, SW_SHOWNORMAL);
  end;
  {----------------------------------------------------------------------------}
var
  MsgDlgCustom: TForm;
  vMsg: String;
  vFormatedCoMError: String;
  vBtnYes: TButton;
  vbtnExit: TButton;
begin
  if USER_FEATURE__CHANGE_MODE_UNSUPPORTED then
    begin
      vMsg := 'This Meter is on procode,. ' +
        'Do you still want to execute Change of Mode?';
      Result := Messagedlg(vMsg, mtinformation, [mbyes, mbno], 0) = mryes
    end
  else
  begin
    vFormatedCoMError  := StringReplace(aCoMError, '. ', '. '#13#10'', [rfReplaceAll, rfIgnoreCase]);

    MsgDlgCustom := CreateMessageDialog(vFormatedCoMError, mtConfirmation, [mbyes, mbno]);
    try
      with MsgDlgCustom do
      begin
        Width     := 590;
        Font.Size := 10;
        Height    := Height + 20;
        Position  := poScreenCenter;

        vBtnYes := (FindComponent('Yes') as TButton);
        with vBtnYes do
        begin
          caption := 'Access Link';
          Width   := 135;
          Left    := 125;
          Top     := 75;
        end;

        vBtnExit := (FindComponent('No') as TButton);
        with vBtnExit do
        begin
          caption := 'Exit';
          Left    := 375;
          Top     := 75;
        end;
      end;

      try
        if MsgDlgCustom.ShowModal = mrYes then
        begin
          ShellOpen('https://utilita.helpjuice.com/dcc-managed-metering/procode?from_search=138913884');
        end;

      except
        on e: Exception do
          raise Exception.Create('Error: Unable to Open Browser');
      end;

    finally
      MsgDlgCustom.Free;
    end;

    Result := false;
  end;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Smets_LoanClick(Sender: TObject);
begin
  DoSmets2Credit(5);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.device_enable_hanClick(Sender: TObject);
begin
  Messagedlg('The Enable Han option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO

  {EnableThisHan;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.EnableEmergCreditClick(Sender: TObject);
var
  getEmergCreditFlag : variant;
  setMPXN, setCustId, setUserName, SqlText : string;

begin
  setMPXN := fSpan;
  setCustId := LATESTCONFIG.FieldByName('CUSTOMER_ID').asstring;
  setUserName := uppercase(FRM_Login.edtUsername.Text);

  if TSmetsUtil.HasRpFlags(LatestConfig.FieldByName('customer_id').AsLargeInt) and not Frm_Common.SuperAuthorityCheck then
  begin
    MessageDlg('Super Authority check has failed', mtError, [mbOk], 0);
    exit;
  end;

  try
    SqlText := 'ODS.PK_CRMUI_Metering.pr_enable_emergency_credit(';
    SqlText := SqlText + ':p_mpxn, ';
    SqlText := SqlText + ':p_customer_id, ';
    SqlText := SqlText + ':p_user_id, ';
    SqlText := SqlText + ':p_flag)';

    gSqlUtil.ExecProc(SQLText, TRANSACTION_YES,
    ['p_mpxn',        otString, pdInput, setMPXN,
     'p_customer_id', otString, pdInput, setCustId,
     'p_user_id',     otString, pdInput, setUserName,
     'p_flag',        otInteger,pdOutput,@getEmergCreditFlag]);

  except
    on e: Exception do
    begin
      MessageDlg('An error has ocurred - Enable Emergency Credit.', mtError, [mbOk], 0);
    end;
  end;

  if getEmergCreditFlag = 0 then
    Messagedlg('Emergency Credit requested successfully ',mtinformation, [mbOK],0);

  if getEmergCreditFlag = 1 then
    Messagedlg('Emergency Credit already requested ',mtError, [mbOK],0);

end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.EnableThisHan;
var
  mt, jbsid: string;
  clickedok: boolean;
begin
  // SMI TODO
  if not Do_Smets_Supplier_Check(fSpan) then
    exit;

  repeat
    jbsid := 'N/A';
    clickedok := InputQuery('JBS Job ID?',
      'Please Enter JBS JOB ID, if you wish the Codes to appear in JBS for Engineer',
      jbsid);
    if not clickedok then
      exit;
    If jbsid = '' then
      Messagedlg('Please Enter A JBS ID', mterror, [mbok], 0);
  until jbsid <> '';

  if fService = SERVICE_ELECTRICITY then
  Begin
    mt := '4';
    if Messagedlg
      ('Do you wish to enable HAN on Electric Meter, to enable other devices to be paired?'
      + #13 + 'JBS ID: ' + jbsid, mtinformation, [mbyes, mbno], 0) <> mryes then
      exit;
  end
  else if fService = SERVICE_GAS then
  begin
    mt := '5';
    if Messagedlg
      ('Do you wish to enable HAN on Gas Meter, to enable pairing with Electric Meter?'
      + #13 + 'JBS ID: ' + jbsid, mtinformation, [mbyes, mbno], 0) <> mryes then
      exit;
  end;

  enable_han(inttostr(fService), fSpan, MeterNo.caption, mt, jbsid);
  Messagedlg('Request has been submitted', mtinformation, [mbok], 0);

  ShowDataflowHistory;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.device_reset_tokenClick(Sender: TObject);
begin
  Messagedlg('The Reset Token option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {if DO_SMETS_SUPPLIER_CHECK(Span.text) = false then
    exit;
  RequestResetToken;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.RequestResetToken;
begin
  // SMI TODO
  iF MeterRemoved.caption = '' THEN
  Begin
    Messagedlg
      ('Meter must first be removed, in order to request a Meter Reset Token.',
      mtinformation, [mbok], 0);
    exit;
  End;

  if Messagedlg('Are you sure you wish to request a Meter Reset Token?',  mtconfirmation, [mbyes, mbno], 0) <> mryes then
    exit;

  RequestMEterResetToken(fSpan, inttostr(fService), MeterNo.caption);
  ShowDataflowHistory;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.SiteInfo1Click(Sender: TObject);
begin
  TFrm_Smets_SiteInfo_Dcc.Start(Self, fService, fSpan);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.TRClick(Sender: TObject);
var
  msg,bTitle,bDesc: string;
begin
  with main_data_module.tempquery1 do
  begin
    close;
    sql.clear;
    DeleteVariables;
    DeclareVariable(':p_mpxn', otString);
    DeclareVariable(':p_tariff_title', otString);
    DeclareVariable(':p_tariff_description', otString);
    Sql.add('BEGIN');
    Sql.add('ods.PK_CRMUI_TARIFFS.get_tariff_details(p_mpxn => :p_mpxn,');
    Sql.add('p_tariff_title => :p_tariff_title, p_tariff_description => :p_tariff_description);');
    Sql.add('END;');
    Setvariable('p_mpxn', fSpan);
    try
      Open;
      frm_login.mainsession.commit;
    except
      on e: Exception do
        raise Exception.Create('Unable to load data. Error: ' + e.Message);
    end;
  end;
  bTitle := main_data_module.tempquery1.GetVariable('p_tariff_title');
  bDesc := main_data_module.tempquery1.GetVariable('p_tariff_description');

  if bTitle.IsEmpty and bDesc.IsEmpty then
    msg := 'No results found.'
  else
    msg := bTitle + sLineBreak + sLineBreak +
           bDesc;

  Messagedlg(msg, mtinformation, [mbok], 0);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.MTR1Click(Sender: TObject);
begin
  {ISC-722 Anna: this is not used anywhere (menu item does not exist). However, if it would work
   then after showing the message which indicates this menu item is not available, but the next line
   exectues the device removal.

   I've left this here until further decision
  }

  Messagedlg('The Remove Device option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  { Anna (ISC-722) Removed as it has remained in the code unintentionally
   RemoveDevice;
   }
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.MTR2Click(Sender: TObject);
begin
  Messagedlg('The Request Reset Token option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {RequestResetToken;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.MTR3Click(Sender: TObject);
begin
  Messagedlg('The Enable this HAN option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
  {EnableThisHan;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Smets_debt_adminClick(Sender: TObject);
begin
  // SMI TODO
  Messagedlg('The Manage Debt option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  {if GetSMetsMeterMode(Span.text) <> '1' then
  Begin
    Messagedlg
      ('This option is only allowed if Current SMETS Meter is in PRE-PAYMENT Mode',
      mtwarning, [mbok], 0);
    exit;
  End;

  if GetSmetsPaymentCardno(Span.text, '') = '' then
  Begin
    Messagedlg('This option is not allowed. No Topup Card is Registered',
      mtwarning, [mbok], 0);
    exit;
  End;

  Application.CreateForm(TFRM_SMETS_MANAGE_DEBT, FRM_SMETS_MANAGE_DEBT);
  try
    FRM_SMETS_MANAGE_DEBT.Refreshdata(Span.text, inttostr(Service.ItemIndex),
      MeterNo.caption, '', 'F');
    FRM_SMETS_MANAGE_DEBT.ShowModal;
  finally
    FRM_SMETS_MANAGE_DEBT.release;
  end;
  if FRM_SMETS_MANAGE_DEBT.tag = 1 then
    ShowDataflowHistory;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Smets_Vend_AdminClick(Sender: TObject);
begin
  if TFrm_Smets_Manage_Credit_Dcc.StartModal(Self, fService, fSpan, MeterNo.Caption, -1) then
    ShowDataFlowHistory;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Smets_WarehouseClick(Sender: TObject);
{var
  ts, ps: string;}
begin
  Messagedlg('The Warehouse Tool option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO
{  ts := mastersource + 'WarehouseTool.exe';
  ps := '"' + frm_login.logon_db.caption + '" "' + frm_login.username.text +
    '" "' + frm_login.password.text + '"';
  shellexecute(Handle, 'open', pchar(ts), pchar(ps), nil, sw_shownormal);}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Req_EventsClick(Sender: TObject);
begin
  // SMI TODO
  Messagedlg('The Events option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  {if DO_SMETS_SUPPLIER_CHECK(Span.text) = false then
    exit;

  if GetSmetsWanStatus(Span.text, '') = 'OFF' then
  Begin
    Messagedlg
      ('Unable to Read Meter. There does not appear to be any active Remote Comms to this Supply',
      mtwarning, [mbok], 0);
    exit;
  end
  else if GetSmetsWanStatus(Span.text, '') <> 'ON' then
  Begin
    if Messagedlg
      ('Communications to this meter may be slow or unreliable. Continue anyway?',
      mtwarning, [mbok], 0) <> mryes then
      exit;
  end;

  Application.CreateForm(TFRM_SMETS_GETEVENTS, FRM_SMETS_GETEVENTS);
  try
    FRM_SMETS_GETEVENTS.sd.date := now - 7;
    FRM_SMETS_GETEVENTS.ed.date := now + 1;
    FRM_SMETS_GETEVENTS.c0.checked := true;
    FRM_SMETS_GETEVENTS.c2.checked := true;
    FRM_SMETS_GETEVENTS.c3.checked := true;
    FRM_SMETS_GETEVENTS.c4.checked := true;
    FRM_SMETS_GETEVENTS.c5.checked := true;
    FRM_SMETS_GETEVENTS.c6.checked := true;
    FRM_SMETS_GETEVENTS.c7.checked := true;
    FRM_SMETS_GETEVENTS.c8.checked := true;
    FRM_SMETS_GETEVENTS.c9.checked := true;
    FRM_SMETS_GETEVENTS.c10.checked := true;
    FRM_SMETS_GETEVENTS.tag := 0;
    FRM_SMETS_GETEVENTS.ShowModal;

    if FRM_SMETS_GETEVENTS.tag = 1 then
    Begin
      // SMI TODO
      RequestEvents(inttostr(Service.ItemIndex), Span.text, MeterNo.caption,
        datetostr(FRM_SMETS_GETEVENTS.sd.date),
        datetostr(FRM_SMETS_GETEVENTS.ed.date),
        FRM_SMETS_GETEVENTS.lst.caption);
      ShowDataflowHistory;
    end;
  finally
    FRM_SMETS_GETEVENTS.release;
  end;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.Req_MirrorClick(Sender: TObject);
begin
  Messagedlg('The Mirror Data option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO

  {if Service.ItemIndex = 0 then
  Begin
    Messagedlg('You can olny perform a Get Mirror Data Request on a GAS Meter.',
      mtwarning, [mbok], 0);
    exit;
  End;

  if DO_SMETS_SUPPLIER_CHECK(Span.text) = false then
    exit;

  if GetSMetsMeterMode(Span.text) <> '1' then
  Begin
    Messagedlg
      ('This option is only allowed if Current SMETS Meter is in PRE-PAYMENT Mode',
      mtwarning, [mbok], 0);
    exit;
  End;

  If Messagedlg('Do you wish to Request the Latest Mirror Data for this Meter?',
    mtconfirmation, [mbyes, mbno], 0) <> mryes then
    exit;
  Insert_OnDemand_MIRROR(inttostr(Service.ItemIndex), Span.text,
    MeterNo.caption);;
  ShowDataflowHistory;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.smets_debt_suspendClick(Sender: TObject);
begin
  DoSmets2Debt(5);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.ShowDataflowHistory;
begin
  BitBtn1.Enabled := false;
  Refreshdata;
  DelayButton;
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  // SMI TODO
  tag := 1;
  Gas_Wait_Timer.enabled := false;

  if FormId > 0 then
    Frm_Main.CrmFormList.ReleaseForm(FormId);
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.DisconnectSupply1Click(Sender: TObject);
begin
  Messagedlg('The Disconnect Supply option is not currently developed for DCC meters ',mtinformation, [mbOK],0);

  // SMI TODO

  {if Messagedlg('Are you sure you wish to Disconnect Supply?', mtconfirmation,
    [mbyes, mbno], 0) <> mryes then
    exit;

  if SupplyControl(inttostr(Service.ItemIndex), Span.text, MeterNo.caption, 'Y',
    '0') = true then
    ShowDataflowHistory;}
end;

{------------------------------------------------------------------------------}
procedure TFrm_Smets_Dcc.DoFlowHistory(aSpan: string);
begin
  gSqlUtil.CreateCursor(FlowHistoryDCC, 'ods.pk_crmui_metering.get_dataflow_history_latest(:p_mpxn_in, :p_res_out)', TRANSACTION_NO,
    ['p_mpxn_in', otString, aSpan,
     'p_res_out', otCursor, null]);
end;

{------------------------------------------------------------------------------}
function TFrm_Smets_Dcc.CheckSpan(aService: integer; aSpan: string): boolean;
begin
  Result := (fService = aService) and (fSpan = aSpan);
end;
{------------------------------------------------------------------------------}
{$endregion TFrm_Smets_Dcc}
{==============================================================================}

end.