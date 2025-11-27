unit smets_check_comms_DCC;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, AdvUtil,
  Vcl.Grids, AdvObj, BaseGrid, AdvGrid, DBAdvGrid, UelSqlUtils, CrmCommon,
  Oracle, OracleData,
  Data.DB, Vcl.ExtCtrls, AdvProgressBar, Vcl.Buttons;

type
  TFRM_CHECK_COMMS = class(TForm)
    LabelPremiseTxt: TLabel;
    GroupBoxPremise: TGroupBox;
    PageControlService: TPageControl;
    TabService: TTabSheet;
    GroupBoxServHist: TGroupBox;
    GridCheckComms: TDBAdvGrid;
    qryCheckComms: TOracleDataSet;
    srcCheckComms: TDataSource;
    Image4: TImage;
    qryCheckCommsTMA: TOracleQuery;
    ProgBarCheckComms: TAdvProgressBar;
    LabelProgComms: TLabel;
    Timer_Comms: TTimer;
    qryCheckCommsEnquiry: TOracleQuery;
    BtnReadLatest: TBitBtn;
    procedure BtnReadLatestClick(Sender: TObject);
    procedure GridCheckCommsGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure RefreshTimerComms(Sender: TObject);
    procedure FormCreate(Sender: TObject);

  strict private
    { Private declarations }
    fSetCustomerID, fSetAgreeID, fSetPremiseID: Int64;
    fSetPremiseInfo : string;
    procedure CheckComms(aGenEnquiry: Boolean);
    procedure CheckCommsTMA;
    procedure JobBooking;
    procedure reSetProgressItems;
    procedure GetRaiseEnquireDelay;
    function CheckCommsMessage(CheckCommsMsgId:integer):string;

  public
    { Public declarations }
    constructor Create(aOwner: TComponent; aCustomerId, aAgreementId, aPremiseId: Int64; aPremiseInfo: string); reintroduce;
    class procedure StartModal(aOwner: TComponent; aCustomerId, aAgreementId, aPremiseId: Int64; aPremiseInfo: string);

  end;

var
  FRM_CHECK_COMMS: TFRM_CHECK_COMMS;

  const rsLiveStatus        = 'LIVE';
  const rsCommsCheckStatus  = 'COMMS_CHECK_STATUS';
  const rsCommsLast12Hrs    = 'COMMS_LAST_12HOURS';

implementation

uses
  LoginUnit, DMImages, addtojbs, Common, DataModule;

{$R *.dfm}

constructor TFRM_CHECK_COMMS.Create(aOwner: TComponent; aCustomerId, aAgreementId, aPremiseId: Int64; aPremiseInfo: string);
begin
  inherited Create(aOwner);

  fSetCustomerID  := aCustomerId;
  fSetAgreeID     := aAgreementId;
  fSetPremiseID   := aPremiseId;
  fSetPremiseInfo := aPremiseInfo;
end;

class procedure TFRM_CHECK_COMMS.StartModal(aOwner: TComponent; aCustomerId, aAgreementId, aPremiseId: Int64; aPremiseInfo: string);
var
  frm : TFRM_CHECK_COMMS;
begin
  frm := TFRM_CHECK_COMMS.Create(aOwner, aCustomerId, aAgreementId, aPremiseId, aPremiseInfo);
  try
    frm.ShowModal;
  finally
    FreeAndNil(frm);
  end;
end;

procedure TFRM_CHECK_COMMS.FormCreate(Sender: TObject);
begin
  LabelPremiseTxt.Caption := fSetPremiseInfo;
  DM_Images.LargeImages.GetIcon(32, Image4.Picture.Icon);
  GetRaiseEnquireDelay;
  CheckComms(false);
end;

function TFRM_CHECK_COMMS.CheckCommsMessage(CheckCommsMsgId:integer):string;
begin
  result := '';

    case (CheckCommsMsgId) of
      0: result := 'Error: Check Comms was unable to load data';
      1: result := 'No Response from DCC, please book a Job from the Booking screen';
      2: result := 'All meters communicating. No need to book any Job for the customer';
      3: result := 'Site visit required. Book Check Comms visit now';
    end;

end;

procedure TFRM_CHECK_COMMS.CheckComms(aGenEnquiry: Boolean);
var
  sqlText: string;
  bBookJob, bGenSuccessMsg: Boolean;
  bCommStatusText: string;

begin
  bBookJob := false;
  bGenSuccessMsg := true;

  sqlText := 'ODS.pk_crmui_metering.pr_get_latest_check_comms_result(';
  sqlText := sqlText + ':p_customer_id,';
  sqlText := sqlText + ':po_results)';
  gSqlUtil.CreateCursor(qryCheckComms, sqlText, TRANSACTION_YES,
    ['p_customer_id', otString, fSetCustomerID,
     'po_results'   , otCursor, null]);

  with qryCheckComms do
  begin
    GridCheckComms.HideColumn(5);

    try

      if qryCheckComms.recordcount > 0 then
      begin

        while not qryCheckComms.eof do
        begin
          bCommStatusText := FieldByName(rsCommsCheckStatus).asString;

          if (bCommStatusText <> rsLiveStatus) then
          begin
            bGenSuccessMsg := false;

            if aGenEnquiry then
              bBookJob := true;

          end;

          qryCheckComms.Next;
        end;

      end
      else
      begin
        Messagedlg(CheckCommsMessage(1), MTWarning, [mbOk], 0);
        bGenSuccessMsg := false;
        aGenEnquiry  := false;
      end;

    except
      on e: Exception do
        raise Exception.Create(CheckCommsMessage(0));

    end;
    DeleteVariables;

  end;

  if aGenEnquiry then
  begin
    sqlText := 'ODS.PK_CRMUI_METERING.pr_raise_comms_check_enquiry(';
    sqlText := sqlText + ' :p_customer_id)';
    gSqlUtil.ExecProc(sqlText, TRANSACTION_YES,
    ['p_customer_id', otString, pdInput, fSetCustomerID]);

    ProgBarCheckComms.Position := 100;
    LabelProgComms.Caption := 'Operation is completed.';
    Screen.Cursor := crDefault;
  end;

  if bBookJob then
    JobBooking;

  if bGenSuccessMsg then
  begin
    Messagedlg(CheckCommsMessage(2), MTWarning, [mbOk], 0);
  end;

  reSetProgressItems;
  BtnReadLatest.Enabled := true;

end;

procedure TFRM_CHECK_COMMS.JobBooking;
begin

  if Messagedlg(CheckCommsMessage(3), mtconfirmation, [mbyes, mbno], 0) = mryes then
  begin

    if FRM_Common.isagreementlive(InttoStr(fSetAgreeID)) = false then
    begin
      Messagedlg('You must select a Premise on a LIVE agreement', MTWarning, [mbOk], 0);
      exit;
    end;

    Application.CreateForm(TFrm_add2Jbs, Frm_add2jbs);
    try
      Frm_add2jbs.clearfields;
      Frm_add2jbs.agreementId := InttoStr(fSetAgreeID);
      Frm_add2jbs.btnGenSuperAuthFlag := false;
      Frm_add2jbs.CheckCommsRestrictFlag := false;
      Frm_add2jbs.GetData(InttoStr(fSetPremiseID));
      Frm_add2jbs.Position := poDesktopCenter;
      Frm_add2jbs.showmodal;

    finally
      Frm_add2jbs.release;

    end;
  end;
  reSetProgressItems;

end;

procedure TFRM_CHECK_COMMS.GetRaiseEnquireDelay;
var
  raiseEnquiryDelay: Double;
begin
  raiseEnquiryDelay := gSqlUtil.SelectQueryDouble('select item_value from crm.standing_data where item_name = ''CRM_COMMS_CHECK_TIMEOUT''');

  Timer_Comms.Interval := Trunc((raiseEnquiryDelay / 50) * 1000);
end;

procedure TFRM_CHECK_COMMS.BtnReadLatestClick(Sender: TObject);
begin
  BtnReadLatest.Enabled := false;
  CheckCommsTMA;
end;

procedure TFRM_CHECK_COMMS.CheckCommsTMA;
var
  sqlText: string;

begin
  sqlText := 'ODS.PK_CRMUI_METERING.pr_perform_check_comms(';
  sqlText := sqlText + ' :p_customer_id)';
  gSqlUtil.ExecProc(sqlText, TRANSACTION_YES,
   ['p_customer_id', otString, pdInput, fSetCustomerID]);

  Timer_Comms.Enabled := true;
  LabelProgComms.Visible := true;
  LabelProgComms.Caption := 'Waiting for Check Comms... Operation in progress, please wait.';
  ProgBarCheckComms.Visible := true;
  ProgBarCheckComms.Position := 2;
  Screen.Cursor := crHourGlass;

end;

procedure TFRM_CHECK_COMMS.RefreshTimerComms(Sender: TObject);
begin
  ProgBarCheckComms.Position := ProgBarCheckComms.Position + 2;

  if ProgBarCheckComms.Position >= 98 then
  begin
    ProgBarCheckComms.Position := 99;
    Timer_Comms.Enabled := false;
    Screen.Cursor := crHourGlass;
    CheckComms(true);
  end;

end;

procedure TFRM_CHECK_COMMS.GridCheckCommsGetCellColor(Sender: TObject;
  ARow, ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
var
  boolCheckLast12Hours: Boolean;
  CheckCommStatus, strCheckLast12Hours: string;

begin
  if (ARow = 0) or (ACol = 0) then
    exit;

  strCheckLast12Hours := GridCheckComms.cells[5, 1];

  if not strCheckLast12Hours.IsEmpty then
  begin
    strCheckLast12Hours := GridCheckComms.cells[5, ARow];        // COMMS_CHECK_12HRS
    boolCheckLast12Hours := StrToBool(strCheckLast12Hours);

    CheckCommStatus := UpperCase(GridCheckComms.cells[3, ARow]); // COMMS_CHECK_STATUS

    if (CheckCommStatus = rsLiveStatus) and (boolCheckLast12Hours) then
    begin
      AFont.Color := clGreen;
    end;

    if (CheckCommStatus <> rsLiveStatus) and (not boolCheckLast12Hours) then
    begin
      AFont.Color := clRed;
    end;

  end;

  if strCheckLast12Hours.IsEmpty then
  begin
    abort;
  end;

end;

procedure TFRM_CHECK_COMMS.reSetProgressItems;
begin
  Screen.Cursor := crDefault;
  LabelProgComms.Visible := false;
  ProgBarCheckComms.Visible := false;
  ProgBarCheckComms.Position := 0;
end;

end.