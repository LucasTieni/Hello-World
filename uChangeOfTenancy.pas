unit uChangeOfTenancy;

interface
uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, AeroButtons,
  Vcl.StdCtrls, JvExStdCtrls, JvButton, JvControlPanelButton, JvExControls,
  JvSpin, JvColorBox, JvColorButton, Vcl.ComCtrls, Vcl.DockTabSet, Vcl.Tabs,
  Vcl.Buttons, System.DateUtils, AddAddress, System.Math,
  System.JSON, REST.Json, System.Generics.Collections, REST.Json.Types,
  LoginUnit, OracleData, Data.DB, Datasnap.DBClient, Vcl.DBCtrls, Vcl.Mask;

type
  TAgreement = class
  public
    [JSONName('agreement_id')]
    agreementId: Int64;
    [JSONName('agreement_type_id')]
    agreementTypeId: Integer;
    [JSONName('agreement_status_id')]
    agreementStatusId: Integer;
    [JSONName('agreement_start_date')]
    agreementStartDate: string;
    [JSONName('agreement_end_date')]
    agreementEndDate: string;
    [JSONName('billfrequency')]
    billFrequency: string;
    [JSONName('agreement_balance')]
    agreementBalance: string;
  end;

type
  TOutgoingSupply = class
  public
    [JSONName('agreement_id')]
    agreementId: Int64;
    [JSONName('service_id')]
    serviceId: Int64;
    [JSONName('service_type_id')]
    serviceTypeId: string;
    [JSONName('span')]
    span: Int64;
    [JSONName('span_type_id')]
    spanTypeId: string;
    [JSONName('span_start_date')]
    spanStartDate: string;
    [JSONName('span_end_date')]
    spanEndDate: string;
    [JSONName('order_status_id')]
    orderStatusId: Integer;
    isTma: Boolean;
  end;

  TIncomingSupply = class
  public
    ServiceId: Int64;
    ServiceTypeId: string;
    Span: Int64;
    AgreementId: Int64;
    MeterBalance: Double;
    MeterMode: string;
    DebtBalance: Double;
    RecoveryRate: Double;
    isTma: Boolean;
  end;

  TPrePayRead = class
  public
    [JSONName('meter_balance')]
    MeterBalance: Double;
    [JSONName('debt_balance')]
    DebtBalance: Double;
    [JSONName('emc_used')]
    EMCUsed: Double;
    [JSONName('last_updated')]
    LastUpdated: string;
  end;

  TConsumptionRead = class
  public
    [JSONName('consumption')]
    Consumption: Double;
    [JSONName('day')]
    Day: Double;
    [JSONName('night')]
    Night: Double;
    [JSONName('last_updated')]
    LastUpdated: string;
  end;

  TReadDetails = class
  public
    [JSONName('mpxn')]
    Mpxn: Int64;
    [JSONName('prepay_read')]
    PrePayReads: TArray<TPrePayRead>;
    [JSONName('consumption_read')]
    ConsumptionReads: TArray<TConsumptionRead>;
    function HasPrePay: Boolean;
    function GetPrePayRead: TPrePayRead;
    function HasConsumption: Boolean;
    function GetConsumptionRead: TConsumptionRead;
  end;

type
  TFrmChangeOfTenancy = class(TForm)
    pnlBack: TPanel;
    lblTopBar: TLabel;
    pgcCustomer: TPageControl;
    tsOutgoingCustomer: TTabSheet;
    tsIncomingCustomer: TTabSheet;
    DateTimeChangeTenant: TDateTimePicker;
    lblVacatedDate: TLabel;
    pgcAgreement: TPageControl;
    pnlButtons: TPanel;
    shpBtnSubmit: TShape;
    lblSubmit: TLabel;
    shpBtnCancel: TShape;
    lblCancel: TLabel;
    pbPerformingCot: TProgressBar;
    pnlOutgoingCustomerInfo: TPanel;
    lblForwardingAddress: TLabel;
    btnForwardingAddress: TButton;
    edtCustomerName: TEdit;
    shpCustomerName: TShape;
    lblCustomerName: TLabel;
    shpForwardingAddress: TShape;
    pnlVacateDate: TPanel;
    pgcAgreementIn: TPageControl;
    pnlCustomerDetails: TPanel;
    lblInCustomerName1: TLabel;
    shpInCustomerName1: TShape;
    lblInCustomerDoB1: TLabel;
    shpInCustomerDoB: TShape;
    lblInCustomerEmail: TLabel;
    shpInCustomerEmail: TShape;
    lblInCustomerMobile: TLabel;
    shpInCustomerMobileA: TShape;
    lblInCustomerName2: TLabel;
    shpInCustomerName2: TShape;
    shpInCustomerMobileB: TShape;
    shpInCustomerMobileC: TShape;
    lblInCustomerDoB2: TLabel;
    edtInCustomerMobileInfoA: TLabel;
    edtInCustomerMobileInfoB: TLabel;
    edtInCustomerMobileInfoC: TLabel;
    edtInCustomerMobileInfoD: TLabel;
    lblInCustomerContact: TLabel;
    edtInCustomerName1: TEdit;
    edtInCustomerEmail: TEdit;
    edtInCustomerMobileA: TEdit;
    DateInCustomerDoB: TDateTimePicker;
    cboInContactChoice: TComboBox;
    edtInCustomerName2: TEdit;
    edtInCustomerMobileB: TEdit;
    edtInCustomerMobileC: TEdit;
    edtOutCustWalletBalance: TEdit;
    lblOutCustWalletBalance: TLabel;
    shpOutCustWalletBalance: TShape;

    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure lblForwardingAddressClick(Sender: TObject);
    procedure btnSubmitClick(Sender: TObject);
    procedure edtInCustomerName1KeyPress(Sender: TObject; var Key: Char);
    procedure edtInCustomerMobileBKeyPress(Sender: TObject; var Key: Char);
    procedure ValidateInput<T>(aSender: TObject; var aKey: Char);
    procedure DecimalOnlyKeyPress(Sender: TObject; var Key: Char);
    procedure MeterBalanceChange(Sender: TObject);
    procedure DebtBalanceChange(Sender: TObject);
    procedure RecoveryRateChange(Sender: TObject);
{$IFDEF CRMTEST}
  protected
{$ELSE}
  private
{$ENDIF}
    fLabelSpanMap: TDictionary<TLabel, Int64>;
    fCustomerName: string;
    fOutgoingCustomerId: Int64;
    fIncomingCustomerId: Int64;
    fAgreements: TObjectList<TAgreement>;
    fOutgoingSupplies: TObjectList<TOutgoingSupply>;
    fIncomingSupplies: TObjectList<TIncomingSupply>;
    fReadDetailsMap: TDictionary<Int64, TReadDetails>;
    fCotSuccess: Boolean;
    fHasPerformedCot: Boolean;

    procedure RefreshData;
    procedure DoCotDateFilter(aSysDateTime: TDateTime);
    procedure GetWalletBalance;
    procedure GetAgreements;
    procedure GetSupplies;
    procedure GetDemandRequest;
    procedure GetReadsDetails;
    function GetReadDetails(const aSpan: Int64): TReadDetails;
    procedure GetAgreementBalance;
    procedure PerformCot(aSupply: TIncomingSupply);
    procedure CancelCot;
    procedure PopulateOutgoingAgreementsTabs;
    procedure PopulateOutgoingSupplyDetails;
    procedure PopulateIncomingAgreementsTabs;
    procedure PopulateIncomingSupplyDetails;
    procedure BuildIncomingSupplies;
    procedure CreateAgreementEditFields(aParentTab: TTabSheet; aAgreement: TAgreement);
    procedure CreateNoDataLabel(aParent: TWinControl; const aCaption: string);
    function CreatePanel(aParent: TWinControl; aAlign: TAlign; aAlignment: TAlignment; aWidth: integer): Tpanel;
    procedure CreateOutgoingSupplyFields(aParentTab: TTabSheet; aSupply: TOutgoingSupply);
    procedure CreateIncomingSupplyFields(aParentTab: TTabSheet; aSupply: TIncomingSupply);
    procedure CreateSupplyReadDetailsFields(aParentPnl: TPanel; aReadDetails: TReadDetails);
    procedure AddFieldSection(aParent: TWinControl; const aLabelCaption, aText: string;
      aIsReadOnly: Boolean = True; aKeyPressEvent: TKeyPressEvent = nil; aTag: NativeInt = 0;
      aChangeEvent: TNotifyEvent = nil);
    procedure AddTmaBottomLabel(aParent: TWinControl);
    procedure RefreshSupplyReadDetailsClick(Sender: TObject);

  public
    constructor Create(aOwner: TComponent; aOutgoingCustomerId: Int64;
      aCustomerName: string); reintroduce;
    class function StartModal(aOwner: TComponent;  aOutgoingCustomerId: Int64;
      aCustomerName: string): boolean;
end;

var
  FrmChangeOfTenancy: TFrmChangeOfTenancy;
  FrmCoTAddress: TFRM_Add_Address;

const
  GET_AGREEMENTS_SQL     = 'crm.pk_ui_customer.pr_get_agreements(:p_payload, :p_response)';
  GET_SUPPLIES_SQL       = 'crm.pk_ui_customer.pr_get_supplies(:p_agreement_payload, :p_response)';
  GET_READ_DETAILS_SQL   = 'ods.pk_crmui_changeoftenancy.get_latest_read_details(:p_payload, :p_response)';
  PERFORM_COT_SQL        = 'ods.pk_crmui_changeoftenancy.pr_proceed(:p_payload, :p_response)';
  GET_ONDEMAND_SQL       = 'ods.pk_crmui_changeoftenancy.pr_initialise(:p_payload, :p_response)';
  CANCEL_COT_SQL         = 'ods.pk_crmui_changeoftenancy.pr_cancel(:p_payload, :p_response)';
  GET_WALLET_BALANCE_SQL = 'crm.pk_ui_payment.pr_get_wallet_balance(:p_payload, :p_response)';

  PREPAY_MODE            = 'Prepayment';
  CREDIT_MODE            = 'Credit';

implementation
uses
  UelSqlUtils, CrmCommon, Oracle, DMImages;
{$R *.dfm}

constructor TFrmChangeOfTenancy.Create(aOwner: TComponent; aOutgoingCustomerId: Int64;
  aCustomerName: string);
begin
  inherited Create(aOwner);
  fCustomerName := aCustomerName;
  fOutgoingCustomerId := aOutgoingCustomerId;
  fIncomingCustomerId := fOutgoingCustomerId;
  edtCustomerName.Text := aCustomerName;
  fAgreements := TObjectList<TAgreement>.Create(true);
  fOutgoingSupplies := TObjectList<TOutgoingSupply>.Create(true);
  fIncomingSupplies := TObjectList<TIncomingSupply>.Create(true);
  fReadDetailsMap := TDictionary<Int64, TReadDetails>.Create();
  fLabelSpanMap := TDictionary<TLabel, Int64>.Create;
  fHasPerformedCot := False;
end;

procedure TFrmChangeOfTenancy.FormCreate(Sender: TObject);
begin
  RefreshData;
end;

procedure TFrmChangeOfTenancy.FormDestroy(Sender: TObject);
begin
  fAgreements.Free;
  fOutgoingSupplies.Free;
  fIncomingSupplies.Free;
  fReadDetailsMap.Free;
  fLabelSpanMap.Free;
end;

procedure TFrmChangeOfTenancy.btnSubmitClick(Sender: TObject);
var
  i: Integer;
  errorMessages: TStringList;
begin
  errorMessages := TStringList.Create;
  pbPerformingCot.Visible := True;
  Application.ProcessMessages;
  fCotSuccess := false;

  try
    for i := 0 to fIncomingSupplies.Count - 1 do
    begin
      if not fIncomingSupplies[i].isTma then
      begin
        try
          PerformCot(fIncomingSupplies[i]);

        except
          on E: Exception do
          begin
            errorMessages.Add(Format('Error processing supply %d: %s', [fIncomingSupplies[i].span, E.Message]));
          end;
        end;
      end;
    end;

    fHasPerformedCot := true;

    if errorMessages.Count > 0 then
      {$IFNDEF CRMTEST}
      MessageDlg('Errors occurred during processing:' + sLineBreak + errorMessages.Text, mtError, [mbOk], 0)
      {$ENDIF}
    else
    begin
      fCotSuccess := true;
      {$IFNDEF CRMTEST}
      MessageDlg('All supplies processed successfully!', mtInformation, [mbOk], 0);
      {$ENDIF}
    end;

  finally
    pbPerformingCot.Visible := False;
    errorMessages.Free;
  end;
end;

procedure TFrmChangeOfTenancy.lblForwardingAddressClick(Sender: TObject);
begin
  try
    FrmCoTAddress := TFRM_Add_Address.Create(nil);
    if MessageDlg('Do you have a forwarding address for the customer?',
        mtconfirmation, [mbyes, mbno], 0) = mryes then
    begin
      FrmCoTAddress := TFRM_Add_Address.Create(nil);
      FrmCoTAddress.Clearfields;
      FrmCoTAddress.Caption := 'Please Select Forwarding Address';
      FrmCoTAddress.ShowModal;
    end;
  finally
    FreeAndNil(FrmCoTAddress);
  end;
end;

procedure TFrmChangeOfTenancy.RefreshData;
begin
  DoCotDateFilter(Now);
  GetWalletBalance;
  GetAgreements;
  GetAgreementBalance;
  PopulateOutgoingAgreementsTabs;
  PopulateIncomingAgreementsTabs;
  GetSupplies;
  GetReadsDetails;
  GetDemandRequest;
  PopulateOutgoingSupplyDetails;
  BuildIncomingSupplies;
  PopulateIncomingSupplyDetails;
end;

procedure TFrmChangeOfTenancy.GetAgreements;
var
  payload: string;
  response: variant;
  jsonString: string;
  jsonResponse: TJSONObject;
  jsonArray: TJSONArray;
  i: Integer;
begin
  payload := Format('{"customer_id": "%d", "agreement_status": "All"}', [fOutgoingCustomerId]);
  try
    gSqlUtil.ExecProc(GET_AGREEMENTS_SQL, TRANSACTION_NO,
      ['p_payload', otString, pdInput, payload,
       'p_response', otString, pdOutput, @response]);

    jsonString := VarToStr(response);
    jsonResponse := TJSONObject.ParseJSONValue(jsonString) as TJSONObject;

    if not Assigned(jsonResponse) then
    begin
      raise Exception.Create('Unable to load current agreements. Please contact support.');
    end;

    jsonArray := jsonResponse.GetValue<TJSONArray>('agreements');
    fAgreements.Clear;

    for i := 0 to jsonArray.Count - 1 do
    begin
      fAgreements.Add(TJson.JsonToObject<TAgreement>(jsonArray.Items[i] as TJSONObject));
    end;
  finally
    jsonResponse.Free;
  end;
end;

procedure TFrmChangeOfTenancy.GetSupplies;
var
  payload: string;
  response: variant;
  jsonString: string;
  jsonResponse: TJSONObject;
  jsonArray: TJSONArray;
  i: integer;
  agreementIndex: integer;
  supply: TOutgoingSupply;
  vSupplyServID: Int64;

begin
  if (fAgreements = nil) or (fAgreements.Count = 0) then
  begin
    Exit;
  end;

  fOutgoingSupplies.Clear;
  for agreementIndex := 0 to fAgreements.Count - 1 do
  begin
    vSupplyServID := 0;
    payload := Format('{"agreement_id": "%d", "supply_status": "Live"}', [fAgreements[agreementIndex].agreementId]);

    try
      gSqlUtil.ExecProc(GET_SUPPLIES_SQL, TRANSACTION_NO,
        ['p_agreement_payload', otString, pdInput, payload,
         'p_response', otString, pdOutput, @response]);

      jsonString := VarToStr(response);
      jsonResponse := TJSONObject.ParseJSONValue(jsonString) as TJSONObject;

      if not Assigned(jsonResponse) then
      begin
        raise Exception.Create('Unable to load current supplies. Please contact support.');
      end;

     if jsonResponse.TryGetValue('supplies', jsonArray) and (jsonArray.Count > 0) then
        jsonArray.Items[0].TryGetValue('service_id', vSupplyServID);

      if (vSupplyServID > 0) then
      begin
        for i := 0 to jsonArray.Count - 1 do
        begin
          supply := TJson.JsonToObject<TOutgoingSupply>(jsonArray.Items[i] as TJSONObject);
          supply.agreementId := fAgreements[agreementIndex].agreementId;
          fOutgoingSupplies.Add(supply);
        end;
      end;

    finally
      jsonResponse.Free;
    end;

  end;
end;

function TFrmChangeOfTenancy.GetReadDetails(const aSpan: Int64): TReadDetails;
var
  payload: string;
  response: variant;
  jsonString: string;
  jsonResponse: TJSONObject;
begin
  Result := nil;
  if fReadDetailsMap.TryGetValue(aSpan, Result) then
    Exit;

  payload := Format('{"mpxn": "%d"}', [aSpan]);
  try
    gSqlUtil.ExecProc(GET_READ_DETAILS_SQL, TRANSACTION_NO,
      ['p_payload', otString, pdInput, payload,
       'p_response', otString, pdOutput, @response]);

    jsonString := VarToStr(response);
    jsonResponse := TJSONObject.ParseJSONValue(jsonString) as TJSONObject;

    if not Assigned(jsonResponse) then
      raise Exception.Create('Unable to load read details. Please contact support.');

    try
      Result := TJson.JsonToObject<TReadDetails>(jsonResponse);
      if not Assigned(Result) then
        raise Exception.Create('Failed to parse read details');

      fReadDetailsMap.AddOrSetValue(aSpan, Result);

    except
      on E: Exception do
      begin
        FreeAndNil(Result);
        raise;
      end;
    end;

  finally
    if Assigned(jsonResponse) then
      jsonResponse.Free;
  end;
end;

procedure TFrmChangeOfTenancy.GetReadsDetails;
var
  i: Integer;
  readDetail: TReadDetails;
begin
  if (fOutgoingSupplies = nil) or (fOutgoingSupplies.Count = 0) then
    Exit;

  fReadDetailsMap.Clear;

  for i := 0 to fOutgoingSupplies.Count - 1 do
  begin
    try
      readDetail := GetReadDetails(fOutgoingSupplies[i].span);
    except
      on E: Exception do
      begin
        {$IFNDEF CRMTEST}
        MessageDlg('Error getting read details for span: ' + IntToStr(fOutgoingSupplies[i].span), mtError, [mbOk], 0)
        {$ENDIF}
      end;
    end;
  end;
end;
procedure TFrmChangeOfTenancy.GetWalletBalance;
var
  payload, warning: string;
  response: variant;
  jsonObject: TJSONObject;
  jsonString: string;
  jsonResponse: TJSONObject;
  balance: Double;
begin
  try
    jsonObject := TJSONObject.Create;
    try
      jsonObject.AddPair('customer_id', TJSONNumber.Create(fOutgoingCustomerId));
      payload := jsonObject.ToString;

      gSqlUtil.ExecProc(GET_WALLET_BALANCE_SQL, TRANSACTION_NO,
        ['p_payload', otString, pdInput, payload,
         'p_response', otString, pdOutput, @response]);

      jsonString := VarToStr(response);
      jsonResponse := TJSONObject.ParseJSONValue(jsonString) as TJSONObject;

      if Assigned(jsonResponse) then
      begin
        if jsonResponse.TryGetValue<Double>('wallet_balance', balance) then
        begin
          edtOutCustWalletBalance.Text := ASCII_POUND + FormatFloat('0.00', balance);
        end
        else
        begin
          if jsonResponse.TryGetValue<string>('warning', warning) then
          begin
            edtOutCustWalletBalance.Text := warning;
          end
          else
          begin
            edtOutCustWalletBalance.Text := 'Customer does not have wallet';
          end;
        end;
      end;

    finally
      FreeAndNil(jsonObject);
      FreeAndNil(jsonResponse);
    end;
  except
    on E: Exception do
    begin
      {$IFNDEF CRMTEST}
      MessageDlg('Error retrieving wallet balance: ' + E.Message, mtError, [mbOk], 0);
      {$ENDIF}
      edtOutCustWalletBalance.Text := 'Error retrieving wallet balance';
    end;
  end;
end;
procedure TFrmChangeOfTenancy.GetAgreementBalance;
var
  jsonString, vAgreeBalance, vAgreeID, sqlText, vBillFreq: string;
  errorMessages: TStringList;
  jsonResponse: TJSONObject;
  i: integer;
  vAgreeBalFormat : Double;
  qrySalesLedger : TOracleDataSet;
begin
  if (fAgreements = nil) or (fAgreements.Count = 0) then
  begin
    exit;
  end;

  errorMessages := TStringList.Create;
  for i := 0 to fAgreements.Count - 1 do
  begin

    try
      qrySalesLedger := TOracleDataSet.Create(nil);

      try
        vAgreeID  := InttoStr(fAgreements[i].agreementId);
        vBillFreq := fAgreements[i].billFrequency;
        sqlText  := 'SELECT SALESLEDGER.fn_agreement_balance (';
        sqlText  := sqlText + QuotedStr(vAgreeID) + ',' + QuotedStr(vBillFreq) + ') ';
        sqlText  := sqlText + 'FROM dual';

        qrySalesLedger := gSqlUtil.SelectQuery(sqlText);
        jsonString     := VarToStr(qrySalesLedger.fields[0].text);
        jsonResponse   := TJSONObject.ParseJSONValue(jsonString) as TJSONObject;

        if not Assigned(jsonResponse) then
        begin
          errorMessages.Add(Format('Error processing supply %d: %s',
          [fAgreements[i].agreementId, 'Unable to retrieve agreement balance data.']));
        end
        else
        begin
          vAgreeBalance := LowerCase(jsonResponse.GetValue<string>('pence'));
          vAgreeBalFormat := StrtoFloat(vAgreeBalance);

          if vAgreeBalFormat <> 0 then
          begin
            vAgreeBalance := ASCII_POUND + FormatCurr('0.####', (StrtoCurr(vAgreeBalance) / 100));
          end
          else
            vAgreeBalance := ASCII_POUND + '0.00';

          fAgreements[i].agreementBalance := vAgreeBalance;
        end;

      except
        on E: Exception do
        begin
          errorMessages.Add('Error: Agreement Balance. Please contact support.');
        end;
      end;

    finally
      FreeAndNil(jsonResponse);
      FreeAndNil(qrySalesLedger);
    end;
  end;

  if errorMessages.Count > 0 then
  begin
    {$IFNDEF CRMTEST}
    MessageDlg('Errors occurred retrieving the agreement balance:' + sLineBreak + errorMessages.Text, mtError, [mbOk], 0)
    {$ENDIF}
  end;

  FreeAndNil(errorMessages);
end;

procedure TFrmChangeOfTenancy.GetDemandRequest;
var
  payload, jsonString, status, message, meterInTma: string;
  errorMessages, tmaSupplies : TStringList;
  response: variant;
  jsonResponse, jsonObj: TJSONObject;
  i: integer;
begin
  if (fAgreements = nil) or (fAgreements.Count = 0) then
  begin
    Exit;
  end;

  errorMessages := TStringList.Create;
  tmaSupplies := TStringList.Create;

  for i := 0 to fOutgoingSupplies.Count - 1 do
  begin
    try
      jsonObj := TJSONObject.Create;
      jsonObj.AddPair('mpxn', TJSONNumber.Create(fOutgoingSupplies[i].span));
      jsonObj.AddPair('username', FRM_LOGIN.edtUsername.Text);
      payload := JsonObj.ToString;

      try
        gSqlUtil.ExecProc(GET_ONDEMAND_SQL, TRANSACTION_NO,
          ['p_payload', otString, pdInput, payload,
           'p_response', otString, pdOutput, @response]);

        jsonString := VarToStr(response);
        jsonResponse := TJSONObject.ParseJSONValue(jsonString) as TJSONObject;
        status := LowerCase(jsonResponse.GetValue<string>('status'));
        message := jsonResponse.GetValue<string>('message');

        if not Assigned(jsonResponse) then
        begin
          errorMessages.Add(Format('Error processing supply %d: %s', [fOutgoingSupplies[i].span,
           'Unable to trigger on demand request.']));
        end;

        if status = 'ok' then
        begin
          if jsonResponse.TryGetValue<string>('meter_in_tma', meterInTma) then
          begin
            fOutgoingSupplies[i].isTma := (UpperCase(meterInTma) = 'Y');
            if fOutgoingSupplies[i].isTma then
            begin
              tmaSupplies.Add(IntToStr(fOutgoingSupplies[i].span));
            end;
          end;
        end
        else
        begin
          if jsonResponse.TryGetValue<string>('meter_in_tma', meterInTma) then
          begin
            errorMessages.Add(Format('Error processing supply %d: %s - %s', [fOutgoingSupplies[i].span, message, meterInTma]));
          end
          else
          begin
            errorMessages.Add(Format('Error processing supply %d: %s', [fOutgoingSupplies[i].span, message]));
          end;
        end;

      except
        on E: Exception do
        begin
          errorMessages.Add('Error: Trigger on demand request.');
        end;
      end;

    finally
      FreeAndNil(jsonObj);
      FreeAndNil(jsonResponse);
    end;
  end;

  if tmaSupplies.Count > 0 then
  begin
    {$IFNDEF CRMTEST}
    MessageDlg('This customer has some meters (supplies) that cannot be submitted as the Meter is on TMA Adaptor.' + sLineBreak +
               'Please submit the request through TMA Portal.' + sLineBreak + sLineBreak +
               'TMA Supplies: ' + tmaSupplies.CommaText, mtWarning, [mbOk], 0);
    {$ENDIF}
  end;

  if errorMessages.Count > 0 then
  begin
    {$IFNDEF CRMTEST}
    MessageDlg('Errors occurred during on demand request:' + sLineBreak + errorMessages.Text, mtError, [mbOk], 0)
    {$ENDIF}
  end;

  FreeAndNil(errorMessages);
  FreeAndNil(tmaSupplies);
end;

procedure TFrmChangeOfTenancy.CancelCot;
var
  payload, jsonString, status, message: string;
  errorMessages: TStringList;
  response: variant;
  jsonResponse, jsonObj: TJSONObject;
  i: integer;
begin
  if (fOutgoingSupplies = nil) or (fOutgoingSupplies.Count = 0) then
  begin
    Exit;
  end;

  errorMessages := TStringList.Create;

  for i := 0 to fOutgoingSupplies.Count - 1 do
  begin
    try
      jsonObj := TJSONObject.Create;
      jsonObj.AddPair('mpxn', TJSONNumber.Create(fOutgoingSupplies[i].span));
      jsonObj.AddPair('username', FRM_LOGIN.edtUsername.Text);
      payload := JsonObj.ToString;

      try
        gSqlUtil.ExecProc(CANCEL_COT_SQL, TRANSACTION_NO,
        ['p_payload', otString, pdInput, payload,
         'p_response', otString, pdOutput, @response]);

        jsonString := VarToStr(response);
        jsonResponse := TJSONObject.ParseJSONValue(jsonString) as TJSONObject;
        status := LowerCase(jsonResponse.GetValue<string>('status'));
        message := jsonResponse.GetValue<string>('message');

        if not Assigned(jsonResponse) then
        begin
          errorMessages.Add(Format('Error cancelling CoT for supply %d: %s', [fOutgoingSupplies[i].span,
           'Unable to cancel change of tenancy.']));
        end;

        if status <> 'ok' then
          errorMessages.Add(Format('Error cancelling CoT for supply %d: %s', [fOutgoingSupplies[i].span, message]));

      except
        on E: Exception do
        begin
          errorMessages.Add('Error: Failed to cancel change of tenancy.' + e.Message);
        end;
      end;

    finally
      FreeAndNil(jsonObj);
      FreeAndNil(jsonResponse);
    end;
  end;

  if errorMessages.Count > 0 then
  begin
    {$IFNDEF CRMTEST}
    MessageDlg('Errors occurred during CoT cancellation:' + sLineBreak + errorMessages.Text, mtError, [mbOk], 0)
    {$ENDIF}
  end;

  FreeAndNil(errorMessages);
end;

procedure TFrmChangeOfTenancy.PerformCot(aSupply: TIncomingSupply);
var
  payload: string;
  jsonObj: TJSONObject;
  response: variant;
  jsonResponse: TJSONObject;
  status, message: string;
begin
  jsonObj := TJSONObject.Create;

  try
    jsonObj.AddPair('new_customer_id', TJSONNull.Create); // TODO -> TJSONNumber.Create(fIncomingCustomerId)
    jsonObj.AddPair('mpxn', TJSONNumber.Create(aSupply.span));
    jsonObj.AddPair('cot_date', FormatDateTime('DD-MM-YYYY', DateTimeChangeTenant.Date));
    jsonObj.AddPair('username', UserID);
    jsonObj.AddPair('debt_balance', FloattoStr(aSupply.debtBalance));
    jsonObj.AddPair('debt_recovery_rate', FloattoStr(aSupply.recoveryRate));

    if aSupply.meterMode = PREPAY_MODE then
    begin
      jsonObj.AddPair('required_meter_balance', FloatToStr(aSupply.meterBalance));
    end;

    payload := jsonObj.ToString;

    try
      gSqlUtil.ExecProc(PERFORM_COT_SQL, TRANSACTION_NO,
        ['p_payload', otString, pdInput, payload,
         'p_response', otString, pdOutput, @response]);

      jsonResponse := TJSONObject.ParseJSONValue(VarToStr(response)) as TJSONObject;
      status := LowerCase(jsonResponse.GetValue<string>('status'));
      message := jsonResponse.GetValue<string>('message');

    except
      on E: Exception do
      begin
        raise Exception.Create('Unknown error. Contact support.');
      end;
    end;

    if status <> 'ok' then
      raise Exception.Create(message);

  finally
    if Assigned(jsonResponse) then
      jsonResponse.Free;

    jsonObj.Free;
  end;
end;
procedure TFrmChangeOfTenancy.PopulateOutgoingAgreementsTabs;
var
  i: integer;
  newTab: TTabSheet;
  agreement: TAgreement;
begin
  pgcAgreement.ActivePage := nil;
  if (fAgreements = nil) or (fAgreements.Count = 0) then
  begin
    CreateNoDataLabel(pgcAgreement, 'No agreements for this customer');
    Exit;
  end;

  for i := 0 to fAgreements.Count - 1 do
  begin
    agreement := fAgreements[i];
    newTab := TTabSheet.Create(pgcAgreement);
    newTab.PageControl := pgcAgreement;
    newTab.Caption := IntToStr(agreement.agreementId);
    newTab.Tag := NativeInt(agreement.agreementId);
    newTab.ImageIndex := 71;
    CreateAgreementEditFields(newTab, agreement);
  end;
end;

procedure TFrmChangeOfTenancy.PopulateOutgoingSupplyDetails;
var
  i, j: Integer;
  newSupplyTab: TTabSheet;
  supply: TOutgoingSupply;
  agreementTab: TTabSheet;
  pgcMeter: TPageControl;
  matchingSupplies: TList<TOutgoingSupply>;
begin
  for i := 0 to pgcAgreement.PageCount - 1 do
  begin
    agreementTab := pgcAgreement.Pages[i];
    matchingSupplies := TList<TOutgoingSupply>.Create;

    try
      for j := 0 to fOutgoingSupplies.Count - 1 do
      begin
        if NativeInt(fOutgoingSupplies[j].agreementId) = agreementTab.Tag then
          matchingSupplies.Add(fOutgoingSupplies[j]);
      end;

      pgcMeter := TPageControl.Create(agreementTab);
      pgcMeter.Parent := agreementTab;
      pgcMeter.Align := alBottom;
      pgcMeter.TabOrder := 1;
      pgcMeter.Tag := agreementTab.Tag;
      pgcMeter.Images := DM_Images.LargeImages;
      pgcMeter.Height := 220;

      if matchingSupplies.Count = 0 then
      begin
        CreateNoDataLabel(pgcMeter, 'No supplies for this agreement');
        Continue;
      end;
      for supply in matchingSupplies do
      begin
        newSupplyTab := TTabSheet.Create(pgcMeter);
        newSupplyTab.PageControl := pgcMeter;
        newSupplyTab.Caption := IntToStr(supply.span);
        if supply.serviceTypeId = 'E' then
          newSupplyTab.ImageIndex := 5
        else if supply.serviceTypeId = 'G' then
          newSupplyTab.ImageIndex := 66;
        CreateOutgoingSupplyFields(newSupplyTab, supply);
      end;
    finally
      matchingSupplies.Free;
    end;
  end;
end;

procedure TFrmChangeOfTenancy.PopulateIncomingAgreementsTabs;
var
  i: integer;
  newTabIn: TTabSheet;
  agreementIn: TAgreement;
begin
  pgcAgreementIn.ActivePage := nil;

  if (fAgreements = nil) or (fAgreements.Count = 0) then
  begin
    CreateNoDataLabel(pgcAgreementIn, 'No agreements for this customer');
    Exit;
  end;

  pgcAgreementIn.Images := DM_Images.LargeImages;

  for i := 0 to fAgreements.Count - 1 do
  begin
    agreementIn := fAgreements[i];
    newTabIn := TTabSheet.Create(pgcAgreementIn);
    newTabIn.PageControl := pgcAgreementIn;
    newTabIn.Caption := 'Pending Agreement ' + IntToStr(i+1);
    newTabIn.Tag := NativeInt(agreementIn.agreementId);
    newTabIn.ImageIndex := 71;
  end;
end;

procedure TFrmChangeOfTenancy.PopulateIncomingSupplyDetails;
var
  i, j: Integer;
  newSupplyTab: TTabSheet;
  supply: TIncomingSupply;
  agreementTabIn: TTabSheet;
  pgcMeter: TPageControl;
  matchingSupplies: TList<TIncomingSupply>;
begin
  for i := 0 to pgcAgreementIn.PageCount - 1 do
  begin

    agreementTabIn := pgcAgreementIn.Pages[i];
    matchingSupplies := TList<TIncomingSupply>.Create;

    try
      for j := 0 to fIncomingSupplies.Count - 1 do
      begin
        if NativeInt(fIncomingSupplies[j].agreementId) = agreementTabIn.Tag then
          matchingSupplies.Add(fIncomingSupplies[j]);
      end;

      pgcMeter := TPageControl.Create(agreementTabIn);
      pgcMeter.Parent := agreementTabIn;
      pgcMeter.Align := alBottom;
      pgcMeter.TabOrder := 1;
      pgcMeter.Tag := agreementTabIn.Tag;
      pgcMeter.Images := DM_Images.LargeImages;

      if matchingSupplies.Count = 0 then
      begin
        CreateNoDataLabel(pgcMeter, 'No supplies for this agreement');
        Continue;
      end;

      for supply in matchingSupplies do
      begin
        newSupplyTab := TTabSheet.Create(pgcMeter);
        newSupplyTab.PageControl := pgcMeter;
        newSupplyTab.Caption := IntToStr(supply.span);

        if supply.serviceTypeId = 'E' then
          newSupplyTab.ImageIndex := 5
        else if supply.serviceTypeId = 'G' then
          newSupplyTab.ImageIndex := 66;

        CreateIncomingSupplyFields(newSupplyTab, supply);
      end;
    finally
      matchingSupplies.Free;
    end;
  end;
end;

procedure TFrmChangeOfTenancy.DoCotDateFilter(aSysDateTime: TDateTime);
begin
  DateTimeChangeTenant.Date := DateOf(aSysDateTime);
  DateTimeChangeTenant.MinDate := DateOf(aSysDateTime) - 30;
  DateTimeChangeTenant.MaxDate := DateOf(Now);
end;

procedure TFrmChangeOfTenancy.edtInCustomerMobileBKeyPress(Sender: TObject;
  var Key: Char);
begin
  ValidateInput<integer>(Sender, Key);
end;

procedure TFrmChangeOfTenancy.edtInCustomerName1KeyPress(Sender: TObject;
  var Key: Char);
begin
  ValidateInput<string>(Sender, Key);
end;

procedure TFrmChangeOfTenancy.CreateAgreementEditFields(aParentTab: TTabSheet; aAgreement: TAgreement);
var
  pnlLeft, pnlRight, pnlTop: TPanel;
begin
  pnlLeft   := CreatePanel(aParentTab, alLeft, taLeftJustify, aParentTab.Width div 2);
  pnlRight  := CreatePanel(aParentTab, alClient, taLeftJustify, aParentTab.Width div 2);
  pnlTop    := CreatePanel(aParentTab, alTop, taLeftJustify, aParentTab.Width);

  AddFieldSection(pnlTop  , 'Agreement Balance', aAgreement.agreementBalance);
  AddFieldSection(pnlLeft , 'Agreement ID', Format('%d', [aAgreement.agreementId]));
  AddFieldSection(pnlLeft , 'Agreement Status ID', Format('%d', [aAgreement.agreementStatusId]));
  AddFieldSection(pnlRight, 'Start Date', aAgreement.agreementStartDate);
  AddFieldSection(pnlRight, 'End Date', aAgreement.agreementEndDate);
end;

procedure TFrmChangeOfTenancy.AddFieldSection(aParent: TWinControl; const aLabelCaption, aText: string;
  aIsReadOnly: Boolean = True; aKeyPressEvent: TKeyPressEvent = nil; aTag: NativeInt = 0;
  aChangeEvent: TNotifyEvent = nil);
var
  pnl: TPanel;
  le: TLabeledEdit;
  topPosition: Integer;
  i: Integer;
begin
  topPosition := 0;

  for i := 0 to aParent.ControlCount - 1 do
  begin
    if (aParent.Controls[i] is TPanel) then
    begin
      topPosition := Max(topPosition, TPanel(aParent.Controls[i]).Top + TPanel(aParent.Controls[i]).Height);
    end;
  end;

  pnl := TPanel.Create(aParent);
  pnl.Parent := aParent;
  pnl.Align := alNone;
  pnl.BevelOuter := bvNone;
  pnl.Height := 45;
  pnl.Width := aParent.Width - 10;
  pnl.Top := topPosition;
  pnl.Left := 5;

  le := TLabeledEdit.Create(pnl);
  le.Parent := pnl;
  le.EditLabel.Caption := aLabelCaption;
  le.Text := aText;
  le.SetBounds(10, 18, 180, 24);
  le.Margins.Left := 5;
  le.Margins.Top := 3;
  le.Margins.Right := 5;
  le.Margins.Bottom := 3;
  le.Width := 180;
  le.ReadOnly := aIsReadOnly;
  le.Ctl3D := False;
  le.Font.Size := 9;
  le.Tag := aTag;

  if Assigned(aKeyPressEvent) then
    le.OnKeyPress := aKeyPressEvent;

  if Assigned(aChangeEvent) then
    le.OnChange := aChangeEvent;
end;


procedure TFrmChangeOfTenancy.AddTmaBottomLabel(aParent: TWinControl);
var
  pnl: TPanel;
  lblTma: TLabel;
begin
  pnl := TPanel.Create(aParent);
  pnl.Parent := aParent;
  pnl.Align := alBottom;
  pnl.BevelOuter := bvNone;
  pnl.Height := 35;

  lblTma := TLabel.Create(pnl);
  lblTma.Parent := pnl;
  lblTma.Align := alClient;
  lblTma.Alignment := taLeftJustify;
  lblTma.Layout := tlCenter;
  lblTma.Caption := 'Meter is on TMA Adaptor, submit the request through TMA Portal.';
  lblTma.Font.Color := clRed;
  lblTma.Font.Style := [fsBold];
  lblTma.Font.Size := 9;
  lblTma.Margins.Left := 10;
  lblTma.AlignWithMargins := True;
end;

function TFrmChangeOfTenancy.CreatePanel(aParent: TWinControl; aAlign: TAlign; aAlignment: TAlignment; aWidth: integer): TPanel;
var
  panel: TPanel;
begin
  panel := TPanel.Create(aParent);
  panel.Parent := aParent;
  panel.Align := aAlign;
  panel.Alignment := aAlignment;
  panel.AlignWithMargins := true;
  panel.Width := aWidth;
  panel.BevelKind := bkNone;
  panel.BevelOuter := bvNone;
  result := panel;
end;

procedure TFrmChangeOfTenancy.CreateNoDataLabel(aParent: TWinControl; const aCaption: string);
var
  lblNoData: TLabel;
begin
  lblNoData := TLabel.Create(aParent);
  lblNoData.Parent := aParent;
  lblNoData.Align := alClient;
  lblNoData.Alignment := taCenter;
  lblNoData.Layout := tlCenter;
  lblNoData.Font.Style := [fsBold];
  lblNoData.Font.Color := clBlack;
  lblNoData.Color := clWhite;
  lblNoData.Caption := aCaption;
  lblNoData.BringToFront;
end;

procedure TFrmChangeOfTenancy.CreateOutgoingSupplyFields(aParentTab: TTabSheet; aSupply: TOutgoingSupply);
var
  pnlSupplyDetails, pnlSupplyReadDetails: TPanel;
  readDetails: TReadDetails;
begin
  pnlSupplyDetails := CreatePanel(aParentTab, alLeft, taLeftJustify, 200);
  pnlSupplyReadDetails := CreatePanel(aParentTab, alClient, taLeftJustify, 400);

  AddFieldSection(pnlSupplyDetails, 'Span', IntToStr(aSupply.span));
  AddFieldSection(pnlSupplyDetails, 'Span Start Date', aSupply.spanStartDate);
  AddFieldSection(pnlSupplyDetails, 'Span End Date', aSupply.spanStartDate);
  if aSupply.isTma then
  begin
    AddTmaBottomLabel(aParentTab);
  end;

  if fReadDetailsMap.TryGetValue(aSupply.span, readDetails) then
  begin
    CreateSupplyReadDetailsFields(pnlSupplyReadDetails, readDetails);
  end
  else
  begin
    AddFieldSection(pnlSupplyReadDetails, 'Status', 'No meter read details available');
  end;
end;

procedure TFrmChangeOfTenancy.CreateSupplyReadDetailsFields(aParentPnl: TPanel; aReadDetails: TReadDetails);
var
  shpBtnRefresh: TShape;
  lblRefresh: TLabel;
  pnl: TPanel;
  pnlLeftColumn, pnlRightColumn: TPanel;
  lastPanelBottom: Integer;
  i: Integer;
begin
  aParentPnl.DisableAlign;
  try
    pnlLeftColumn := CreatePanel(aParentPnl, alLeft, taLeftJustify, aParentPnl.Width div 2);
    pnlRightColumn := CreatePanel(aParentPnl, alClient, taLeftJustify, aParentPnl.Width div 2);

    if not (aReadDetails.HasPrePay or aReadDetails.HasConsumption) then
    begin
      AddFieldSection(pnlLeftColumn, 'Status', 'No meter read details available');
    end
    else
    begin
      if aReadDetails.HasPrePay then
      begin
        AddFieldSection(pnlLeftColumn, 'Meter Balance', FormatFloat('0.00', aReadDetails.GetPrePayRead.MeterBalance));
        AddFieldSection(pnlLeftColumn, 'Debt Balance', FormatFloat('0.00', aReadDetails.GetPrePayRead.DebtBalance));

        AddFieldSection(pnlRightColumn, 'Last Updated', aReadDetails.GetPrePayRead.LastUpdated);
      end;

      if aReadDetails.HasConsumption then
      begin
        AddFieldSection(pnlLeftColumn, 'Consumption', FormatFloat('0.00', aReadDetails.GetConsumptionRead.Consumption));
        AddFieldSection(pnlLeftColumn, 'Day / Peak', FormatFloat('0.00', aReadDetails.GetConsumptionRead.Day));

        AddFieldSection(pnlRightColumn, 'Night / Off Peak', FormatFloat('0.00', aReadDetails.GetConsumptionRead.Night));
        AddFieldSection(pnlRightColumn, 'Last Updated', aReadDetails.GetConsumptionRead.LastUpdated);
      end;

      lastPanelBottom := 0;
      for i := 0 to pnlRightColumn.ControlCount - 1 do
      begin
        if (pnlRightColumn.Controls[i] is TPanel) then
        begin
          lastPanelBottom := Max(lastPanelBottom, TPanel(pnlRightColumn.Controls[i]).Top + TPanel(pnlRightColumn.Controls[i]).Height);
        end;
      end;

      pnl := TPanel.Create(pnlRightColumn);
      pnl.Parent := pnlRightColumn;
      pnl.Align := alNone;
      pnl.BevelOuter := bvNone;
      pnl.Height := 45;
      pnl.Width := pnlRightColumn.Width - 10;
      pnl.Top := lastPanelBottom;
      pnl.Left := 5;

      shpBtnRefresh := TShape.Create(pnl);
      shpBtnRefresh.Parent := pnl;
      shpBtnRefresh.Align := alNone;
      shpBtnRefresh.Height := 25;
      shpBtnRefresh.Width := 150;
      shpBtnRefresh.Top := 10;
      shpBtnRefresh.Left := pnl.Left + 20;
      shpBtnRefresh.Brush.Color := clHotLight;
      shpBtnRefresh.Pen.Color := clHotLight;
      shpBtnRefresh.Shape := stRoundRect;

      lblRefresh := TLabel.Create(pnl);
      lblRefresh.Parent := pnl;
      lblRefresh.Top := shpBtnRefresh.Top + 5;
      lblRefresh.Left := shpBtnRefresh.Left + 55;
      lblRefresh.Height := shpBtnRefresh.Height;
      lblRefresh.Width := shpBtnRefresh.Width;
      lblRefresh.Alignment := taCenter;
      lblRefresh.AlignWithMargins := false;
      lblRefresh.Caption := 'Refresh';
      lblRefresh.Font.Charset := ANSI_CHARSET;
      lblRefresh.Font.Color := 16119285;
      lblRefresh.Font.Height := -12;
      lblRefresh.Font.Name := 'Urbanist';
      lblRefresh.Font.Style := [fsBold];
      lblRefresh.Layout := tlCenter;
      lblRefresh.Cursor := crHandPoint;
      lblRefresh.OnClick := RefreshSupplyReadDetailsClick;
      lblRefresh.BringToFront;
    end;

  finally
    aParentPnl.EnableAlign;
  end;
end;

procedure TFrmChangeOfTenancy.ValidateInput<T>(aSender: TObject; var aKey: Char);
var
  edit: TCustomEdit;
  inputText, newText: string;
  intValue: Integer;
  floatValue: Double;
  valid: Boolean;
begin
  if (aKey = #8) or (aKey = #127) or (aKey = #0) then
    exit;

  if aSender is TCustomEdit then
    edit := TCustomEdit(aSender)
  else
    exit;

  inputText := edit.Text;
  newText := inputText + aKey;
  valid := False;

  if TypeInfo(T) = TypeInfo(Integer) then
  begin
    valid := TryStrToInt(newText, intValue);
  end
  else if TypeInfo(T) = TypeInfo(Double) then
  begin
    valid := TryStrToFloat(newText, floatValue);
  end
  else if TypeInfo(T) = TypeInfo(string) then
  begin
    valid := aKey in ['A'..'Z', 'a'..'z'];
  end;

  if not valid then
    aKey := #0;
end;

procedure TFrmChangeOfTenancy.DecimalOnlyKeyPress(Sender: TObject; var Key: Char);
begin
  ValidateInput<Double>(Sender, Key);
end;

procedure TFrmChangeOfTenancy.MeterBalanceChange(Sender: TObject);
var
  editField: TLabeledEdit;
  incomingSupply: TIncomingSupply;
  meterBalance: Double;
  validText: string;
  i, decimalPoints: Integer;
  cursorPos: Integer;
begin
  editField := Sender as TLabeledEdit;
  incomingSupply := TIncomingSupply(editField.Tag);

  if (incomingSupply <> nil) and TryStrToFloat(editField.Text, meterBalance) then
  begin
    incomingSupply.meterBalance := meterBalance;
  end;
end;

procedure TFrmChangeOfTenancy.DebtBalanceChange(Sender: TObject);
var
  editField: TLabeledEdit;
  incomingSupply: TIncomingSupply;
  debtBalance: Double;

begin
  editField := Sender as TLabeledEdit;
  incomingSupply := TIncomingSupply(editField.Tag);

  if (incomingSupply <> nil) and TryStrToFloat(editField.Text, debtBalance) then
  begin
    incomingSupply.debtBalance := debtBalance;
  end;

end;

procedure TFrmChangeOfTenancy.RecoveryRateChange(Sender: TObject);
var
  editField: TLabeledEdit;
  incomingSupply: TIncomingSupply;
  vRecoveryRate: double;

begin
  editField := Sender as TLabeledEdit;
  incomingSupply := TIncomingSupply(editField.Tag);

  if (incomingSupply <> nil) and TryStrToFloat(editField.Text, vRecoveryRate) then
  begin

    if vRecoveryRate > 100 then
    begin
      editField.Text := '100';
      vRecoveryRate  := 100;
    end;

    incomingSupply.recoveryRate := vRecoveryRate;
  end;

end;

procedure TFrmChangeOfTenancy.CreateIncomingSupplyFields(aParentTab: TTabSheet; aSupply: TIncomingSupply);
var
  pnlSupplyDetailsA, pnlSupplyDetailsB, pnlSupplyDetailsC: TPanel;
begin
  pnlSupplyDetailsA := CreatePanel(aParentTab, alLeft, taLeftJustify, 200);
  pnlSupplyDetailsB := CreatePanel(aParentTab, alClient, taLeftJustify, 200);
  pnlSupplyDetailsC := CreatePanel(aParentTab, alRight, taLeftJustify, 200);

  AddFieldSection(pnlSupplyDetailsA, 'Meter Mode', aSupply.meterMode);
  if aSupply.isTma then
  begin
    AddTmaBottomLabel(aParentTab);
  end;

  if aSupply.meterMode = PREPAY_MODE then
  begin
    AddFieldSection(
      pnlSupplyDetailsA,
      'Meter Balance',
      Format('%.2f', [aSupply.meterBalance]),
      False,
      DecimalOnlyKeyPress,
      NativeInt(aSupply),
      MeterBalanceChange
    );

    AddFieldSection(
      pnlSupplyDetailsB,
      'Debt Balance',
      '0.00',
      false,
      DecimalOnlyKeyPress,
      NativeInt(aSupply),
      DebtBalanceChange
    );

    AddFieldSection(
      pnlSupplyDetailsC,
      'Recovery Rate (%)',
      '20',
      false,
      DecimalOnlyKeyPress,
      NativeInt(aSupply),
      RecoveryRateChange
    );
  end;
end;

procedure TFrmChangeOfTenancy.BuildIncomingSupplies;
var
  i: Integer;
  incomingSupply: TIncomingSupply;
  hasPrePaySupply: Boolean;
begin
  if (fOutgoingSupplies = nil) or (fOutgoingSupplies.Count = 0) then
  begin
    Exit;
  end;

  fIncomingSupplies.Clear;
  hasPrePaySupply := False;

  for i := 0 to fOutgoingSupplies.Count - 1 do
  begin
    incomingSupply := TIncomingSupply.Create;

    incomingSupply.serviceId := fOutgoingSupplies[i].serviceId;
    incomingSupply.serviceTypeId := fOutgoingSupplies[i].serviceTypeId;
    incomingSupply.span := fOutgoingSupplies[i].span;
    incomingSupply.agreementId := fOutgoingSupplies[i].agreementId;
    incomingSupply.debtBalance := 0;
    incomingSupply.recoveryRate := 20;
    incomingSupply.isTma := fOutgoingSupplies[i].isTma;

    if fReadDetailsMap.ContainsKey(fOutgoingSupplies[i].span) and
       fReadDetailsMap[fOutgoingSupplies[i].span].HasPrePay then
    begin
      incomingSupply.meterMode := PREPAY_MODE;
      hasPrePaySupply := True;
    end
    else
    begin
      incomingSupply.meterMode := CREDIT_MODE;
    end;

    fIncomingSupplies.Add(incomingSupply);
  end;

end;

procedure TFrmChangeOfTenancy.RefreshSupplyReadDetailsClick(Sender: TObject);
var
  span: Int64;
  refreshedReadDetails: TReadDetails;
  lblRefresh: TLabel;
  pnlRightColumn, pnlMainParent, pnlSupplyReadDetails: TPanel;
  tabSheet: TTabSheet;
  i: Integer;
begin
  lblRefresh := TLabel(Sender);
  pnlRightColumn := lblRefresh.Parent.Parent as TPanel;
  pnlSupplyReadDetails := pnlRightColumn.Parent as TPanel;
  tabSheet := pnlSupplyReadDetails.Parent as TTabSheet;
  span := StrToInt64(tabSheet.Caption);

  while pnlSupplyReadDetails.ControlCount > 0 do
    pnlSupplyReadDetails.Controls[0].Free;

  if fReadDetailsMap.ContainsKey(span) then
  begin
    fReadDetailsMap.Remove(span);
  end;

  refreshedReadDetails := GetReadDetails(span);
  CreateSupplyReadDetailsFields(pnlSupplyReadDetails, refreshedReadDetails);
  Application.ProcessMessages;
end;

function TReadDetails.HasPrePay: Boolean;
begin
  Result := (Length(PrePayReads) > 0);
end;

function TReadDetails.GetPrePayRead: TPrePayRead;
begin
  if HasPrePay then
    Result := PrePayReads[0]
  else
    Result := nil;
end;

function TReadDetails.HasConsumption: Boolean;
begin
  Result := (Length(ConsumptionReads) > 0);
end;

function TReadDetails.GetConsumptionRead: TConsumptionRead;
begin
  if HasConsumption then
    Result := ConsumptionReads[0]
  else
    Result := nil;
end;

class function TFrmChangeOfTenancy.StartModal(aOwner: TComponent;
  aOutgoingCustomerId: Int64; aCustomerName: string): boolean;
var
  frm: TFrmChangeOfTenancy;
begin
  try
    frm := TFrmChangeOfTenancy.Create(aOwner, aOutgoingCustomerId, aCustomerName);
    frm.Position := poScreenCenter;
    frm.BringToFront;
    Result := frm.showmodal = mrOk;

  finally
    if not frm.fHasPerformedCot then
    begin
      frm.CancelCot;
    end;

    FreeAndNil(frm);
  end;
end;

end.