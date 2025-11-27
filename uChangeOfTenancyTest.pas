unit uChangeTenancyTests;

interface
uses
  DUnitX.TestFramework, LoginUnit, Main, CrmCommon, DMImages, vcl.Forms,
  UELSqlUtils, System.SysUtils, SmetsCommon, uChangeOfTenancy,
  System.DateUtils, AddAddress, Common, AddNewCustomer, vcl.Controls, System.StrUtils,
  Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.StdCtrls;


type
  TFrmChangeOfTenancySub = class(TFrmChangeOfTenancy);
  [TestFixture]
  TChangeTenancyTests = class
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure DoValidDateChangeTenancySuccess;

    [Test]
    [TestCase('Test B - Check for Invalid Date - Prior 30 days', '{today}-35')]
    [TestCase('Test C - Check for Invalid Date - After 1 day', '{today}+1')]
    procedure DoValidDateChangeTenancyFail(aSysDateTimeTest : TDateTime);

    [Test]
    procedure DoViewEditNameSuccess;

    [Test]
    procedure DoViewEditForwardAddressSuccess;

    [Test]
    procedure OpenForwardAddressScreenSuccess;

    [Test]
    procedure GetAgreementsSuccess;

    [Test]
    procedure GetSuppliesSuccess;

    [Test]
    procedure PerformCotExecutionSuccess;

    [Test]
    procedure PopulateOutgoingAgreementsTabsSuccess;

    [Test]
    procedure PopulateOutgoingSupplyDetailsSuccess;

    [Test]
    procedure PerformCoTForAllSuppliesSuccessfully;

    [Test]
    procedure GetReadDetailsSuccess;

    [Test]
    procedure CreateSupplyReadDetailsFieldsSuccess;

    [Test]
    procedure RefreshSupplyReadDetailsClickSuccess;

    [Test]
    procedure GetDemandRequestSuccess;

    [Test]
    procedure GetAgreementBalanceSuccess;

    [Test]
    procedure TestCancelCotWhenNotPerformed;

    [Test]
    procedure TestNoCancelCotWhenPerformed;

    [Test]
    procedure ExecuteCancelCotWithoutError;

    [Test]
    procedure PopulateIncomingAgreementsTabsSuccess;

    [Test]
    procedure PopulateIncomingSupplyDetailsSuccess;

    [Test]
    procedure TestIntegerValidation;

    [Test]
    procedure TestDoubleValidation;

    [Test]
    procedure TestStringValidation;

    [Test]
    procedure TestDecimalOnlyKeyPress;

    [Test]
    procedure TestReadDetailsHasPrePay;

    [Test]
    procedure TestReadDetailsGetPrePayRead;

    [Test]
    procedure TestReadDetailsHasConsumption;

    [Test]
    procedure TestReadDetailsGetConsumptionRead;

    [Test]
    procedure TestBuildIncomingSupplies;

    [Test]
    procedure TestMeterBalanceChange;

    [Test]
    procedure TestDebtBalanceChange;

    [Test]
    procedure TestRecoveryRateChange;

    [Test]
    procedure TestOutgoingCustomerWalletBalance;

    [Test]
    procedure TestOutgoingCustomerWalletBalanceWarning;

    [Test]
    procedure TestTmaSupplyIdentification;
  end;
implementation

var
  FrmChangeOfTenancySub : TFrmChangeOfTenancySub;

procedure TChangeTenancyTests.Setup;
begin
  Application.CreateForm(TFRM_MAIN, FRM_MAIN);
  Application.CreateForm(TDM_Images, DM_Images);
  Application.CreateForm(TFRM_Login, FRM_Login);
  FRM_LOGIN.MainSession.LogonUsername := 'TEST01';
  FRM_LOGIN.MainSession.LogonPassword := 'TEST01';
  FRM_LOGIN.MainSession.LogonDatabase := 'DEVCRM';
  FRM_LOGIN.MainSession.Connected     := true;
  UserID := FRM_LOGIN.MainSession.LogonUsername;
  gSqlUtil := TSqlUtil.create(FRM_Login.mainsession);
end;
procedure TChangeTenancyTests.TearDown;
begin
  FreeAndNil(gSqlUtil);
  FreeAndNil(FRM_MAIN);
  FreeAndNil(DM_Images);
  FreeAndNil(FRM_Login);
end;

procedure TChangeTenancyTests.DoValidDateChangeTenancySuccess;
var
  Expected, Actual : boolean;
  FrmChangeOfTenancySub : TFrmChangeOfTenancySub;
  vSysDateTimeTest : TDateTime;

begin
  Expected := true;
  vSysDateTimeTest := Date;

  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');
  FrmChangeOfTenancySub.DoCotDateFilter(vSysDateTimeTest);

  Actual := FrmChangeOfTenancySub.DateTimeChangeTenant.Date = vSysDateTimeTest;
  Assert.AreEqual(Expected, Actual, 'Invalid Date Selected');
  Actual := FrmChangeOfTenancySub.DateTimeChangeTenant.MinDate = (vSysDateTimeTest - 30);
  Assert.AreEqual(Expected, Actual, 'Invalid MinDate Selected');
  Actual := DatetoStr(FrmChangeOfTenancySub.DateTimeChangeTenant.MaxDate) = DatetoStr(vSysDateTimeTest);
  Assert.AreEqual(Expected, Actual, 'Invalid MaxDate Selected');

  FreeAndNil(FrmChangeOfTenancySub);
end;

procedure TChangeTenancyTests.DoValidDateChangeTenancyFail(aSysDateTimeTest : TDateTime);
var
  Expected, Actual : boolean;
  FrmChangeOfTenancySub : TFrmChangeOfTenancySub;
  CurrentDate: TDateTime;
begin
  Actual := false;
  CurrentDate := DateOf(Now);

  if Pos('{today}-35', DateToStr(aSysDateTimeTest)) > 0 then
    aSysDateTimeTest := CurrentDate - 35
  else if Pos('{today}+1', DateToStr(aSysDateTimeTest)) > 0 then
    aSysDateTimeTest := CurrentDate + 1;

  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');
  try
    FrmChangeOfTenancySub.DoCotDateFilter(aSysDateTimeTest);
  except
    Actual := true;
  end;
  Assert.IsTrue(Actual, 'Incorrect date range');
  FreeAndNil(FrmChangeOfTenancySub);
end;

procedure TChangeTenancyTests.DoViewEditNameSuccess;
var
  Expected, Actual : boolean;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');
  // TEST TO VIEW CUSTOMER NAME
  try
    Actual :=  FrmChangeOfTenancySub.edtCustomerName.Text <> '';
  except
    Actual := false;
  end;
  Assert.IsTrue(Actual, 'Customer name empty');
  // TEST TO EDIT CUSTOMER NAME
  try
    FrmChangeOfTenancySub.edtCustomerName.Text := LeftStr(FrmChangeOfTenancySub.edtCustomerName.Text, 7);
    Actual := FrmChangeOfTenancySub.edtCustomerName.Text = 'Larry P';
  except
    Actual := false;
  end;
  Assert.IsTrue(Actual, 'Customer name edit fail');
  FreeAndNil(FrmChangeOfTenancySub);
end;

procedure TChangeTenancyTests.DoViewEditForwardAddressSuccess;
var
  Expected, Actual : boolean;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');
  FrmCoTAddress := TFRM_Add_Address.Create(nil);
  // TEST TO VIEW FORWARDING ADDRESS TEXT
  try
    FrmCoTAddress.Clearfields;
    FrmCoTAddress.Ad1.Text      := 'Procode Test';
    FrmCoTAddress.PostCode.Text := 'ABC 99DD';
    Actual := (FrmCoTAddress.Ad1.Text <> '') and
                (FrmCoTAddress.PostCode.Text <> '');
  except
    Actual := false;
  end;
  Assert.IsTrue(Actual, 'Customer Forward Address view and edit fail');
  // TEST TO EDIT FORWARDING ADDRESS TEXT
  try
    FrmCoTAddress.Clearfields;
    FrmCoTAddress.Ad1.Text      := 'Procode Test';
    FrmCoTAddress.PostCode.Text := 'ABC 99DD';
    FrmCoTAddress.Ad1.Text      := LeftStr(FrmCoTAddress.Ad1.Text, 9);
    FrmCoTAddress.PostCode.Text := LeftStr(FrmCoTAddress.PostCode.Text, 5);
    Actual := (FrmCoTAddress.Ad1.Text = 'Procode T') and
                (FrmCoTAddress.PostCode.Text = 'ABC 9') ;
  except
    Actual := false;
  end;
  Assert.IsTrue(Actual, 'Customer Forward Address view and edit fail');
  FreeAndNil(FrmChangeOfTenancySub);
  FreeAndNil(FrmCoTAddress);
end;

procedure TChangeTenancyTests.OpenForwardAddressScreenSuccess;
var
  Expected, Actual : boolean;
begin
  Actual := false;
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');
  // TEST TO CHECK FORWARD ADDRESS UNIT OPENED
  try
    FrmCoTAddress := TFRM_Add_Address.Create(nil);
    Actual := FrmCoTAddress <> nil;
  finally
    FreeAndNil(FrmCoTAddress);
  end;
  Assert.IsTrue(Actual, 'Forwarding Address screen open fail');
  FreeAndNil(FrmChangeOfTenancySub);
end;

procedure TChangeTenancyTests.GetAgreementsSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    FrmChangeOfTenancySub.GetAgreements;

    Assert.IsNotNull(FrmChangeOfTenancySub.fAgreements, 'Agreements list should not be nil');

  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.GetSuppliesSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    FrmChangeOfTenancySub.GetAgreements;
    FrmChangeOfTenancySub.GetSupplies;

    Assert.IsNotNull(FrmChangeOfTenancySub.fOutgoingSupplies, 'Supplies list should not be nil');

  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.GetDemandRequestSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    FrmChangeOfTenancySub.GetAgreements;
    FrmChangeOfTenancySub.GetSupplies;
    FrmChangeOfTenancySub.GetDemandRequest;

    Assert.Pass('On Demand Request executed without exceptions');

  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.PerformCotExecutionSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
  testSupply: TIncomingSupply;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    testSupply := TIncomingSupply.Create;
    try
      testSupply.span := 2000022710812;
      testSupply.meterMode := CREDIT_MODE;

      try
        FrmChangeOfTenancySub.PerformCot(testSupply);
      except
        on E: Exception do
          Assert.Fail('PerformCot threw an exception: ' + E.Message);
      end;
    finally
      testSupply.Free;
    end;
  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.PopulateOutgoingAgreementsTabsSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    FrmChangeOfTenancySub.GetAgreements;

    FrmChangeOfTenancySub.PopulateOutgoingAgreementsTabs;

    if FrmChangeOfTenancySub.fAgreements.Count > 0 then
      Assert.IsTrue(FrmChangeOfTenancySub.pgcAgreement.PageCount > 0,
        'Agreement tabs should have been created')
    else
      Assert.IsNotNull(FrmChangeOfTenancySub.pgcAgreement,
        'Page control should exist even if empty');
  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.PopulateOutgoingSupplyDetailsSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    FrmChangeOfTenancySub.GetAgreements;
    FrmChangeOfTenancySub.GetSupplies;

    FrmChangeOfTenancySub.PopulateOutgoingAgreementsTabs;

    FrmChangeOfTenancySub.PopulateOutgoingSupplyDetails;

    Assert.Pass('PopulateOutgoingSupplyDetails executed without exceptions');
  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;


procedure TChangeTenancyTests.PopulateIncomingAgreementsTabsSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    FrmChangeOfTenancySub.GetAgreements;

    FrmChangeOfTenancySub.PopulateIncomingAgreementsTabs;

    if FrmChangeOfTenancySub.fAgreements.Count > 0 then
      Assert.IsTrue(FrmChangeOfTenancySub.pgcAgreementIn.PageCount > 0,
        'Agreement tabs should have been created')
    else
      Assert.IsNotNull(FrmChangeOfTenancySub.pgcAgreementIn,
        'Page control should exist even if empty');
  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.PopulateIncomingSupplyDetailsSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    FrmChangeOfTenancySub.GetAgreements;
    FrmChangeOfTenancySub.GetSupplies;

    FrmChangeOfTenancySub.PopulateIncomingAgreementsTabs;

    FrmChangeOfTenancySub.PopulateIncomingSupplyDetails;

    Assert.Pass('PopulateIncomingSupplyDetails executed without exceptions');
  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.PerformCoTForAllSuppliesSuccessfully;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');
  try
    FrmChangeOfTenancySub.btnSubmitClick(nil);

    Assert.Pass('btnSubmitClick executed without exceptions');

  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.GetReadDetailsSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
  testSpan: Int64;
  readDetails: TReadDetails;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    testSpan := 2000022710812;

    readDetails := FrmChangeOfTenancySub.GetReadDetails(testSpan);

    Assert.IsNotNull(readDetails, 'ReadDetails should not be nil');
    Assert.IsTrue(FrmChangeOfTenancySub.fReadDetailsMap.ContainsKey(testSpan),
      'Read details should be stored in the map');

  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.CreateSupplyReadDetailsFieldsSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
  testSpan: Int64;
  readDetails: TReadDetails;
  testPanel: TPanel;
  prePayRead: TPrePayRead;
  consumptionRead: TConsumptionRead;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    testSpan := 2000022710812;
    testPanel := TPanel.Create(nil);
    testPanel.Parent := FrmChangeOfTenancySub;

    readDetails := TReadDetails.Create;
    try
      readDetails.Mpxn := testSpan;

      SetLength(readDetails.PrePayReads, 1);
      prePayRead := TPrePayRead.Create;
      prePayRead.MeterBalance := 100.50;
      prePayRead.DebtBalance := 25;
      prePayRead.EMCUsed := 10;
      prePayRead.LastUpdated := FormatDateTime('dd/mm/yyyy hh:nn:ss', Now);
      readDetails.PrePayReads[0] := prePayRead;

      SetLength(readDetails.ConsumptionReads, 1);
      consumptionRead := TConsumptionRead.Create;
      consumptionRead.Consumption := 450;
      consumptionRead.Day := 200;
      consumptionRead.Night := 100;
      consumptionRead.LastUpdated := FormatDateTime('dd/mm/yyyy hh:nn:ss', Now);
      readDetails.ConsumptionReads[0] := consumptionRead;

      FrmChangeOfTenancySub.fReadDetailsMap.AddOrSetValue(testSpan, readDetails);

      FrmChangeOfTenancySub.CreateSupplyReadDetailsFields(testPanel, readDetails);

      Assert.IsTrue(testPanel.ControlCount > 0,
        'CreateSupplyReadDetailsFields should create components');

    finally
      testPanel.Free;
    end;
  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.RefreshSupplyReadDetailsClickSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
  testSpan: Int64;
  readDetails: TReadDetails;
  testPanel, parentPanel: TPanel;
  testTabSheet: TTabSheet;
  lblRefresh: TLabel;
  shpRefresh: TShape;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    testSpan := 2000022710812;

    testTabSheet := TTabSheet.Create(nil);
    testTabSheet.Parent := FrmChangeOfTenancySub;
    testTabSheet.Caption := IntToStr(testSpan);

    parentPanel := TPanel.Create(nil);
    parentPanel.Parent := testTabSheet;

    testPanel := TPanel.Create(nil);
    testPanel.Parent := parentPanel;
    testPanel.Name := 'pnlSupplyReadDetails';

    shpRefresh := TShape.Create(nil);
    shpRefresh.Parent := testPanel;

    lblRefresh := TLabel.Create(nil);
    lblRefresh.Parent := testPanel;

    try
      readDetails := FrmChangeOfTenancySub.GetReadDetails(testSpan);
      Assert.IsNotNull(readDetails, 'ReadDetails should not be nil');
      Assert.IsTrue(FrmChangeOfTenancySub.fReadDetailsMap.ContainsKey(testSpan),
        'Read details should be stored in the map');

      FrmChangeOfTenancySub.fReadDetailsMap.Remove(testSpan);
      Assert.IsFalse(FrmChangeOfTenancySub.fReadDetailsMap.ContainsKey(testSpan),
        'Read details should be removed from map during refresh');

      readDetails := FrmChangeOfTenancySub.GetReadDetails(testSpan);
      Assert.IsNotNull(readDetails, 'Read details should be refreshed successfully');
      Assert.IsTrue(FrmChangeOfTenancySub.fReadDetailsMap.ContainsKey(testSpan),
        'Read details should be added back to map after refresh');

      Assert.Pass('Refresh read details test completed successfully');

    finally
      testTabSheet.Free;
    end;

  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;


procedure TChangeTenancyTests.GetAgreementBalanceSuccess;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
  i: integer;
  Actual, Expected : boolean;
  vAgreeBalance : Double;

begin
  Expected := true;
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    try
      FrmChangeOfTenancySub.GetAgreements;
      FrmChangeOfTenancySub.GetAgreementBalance;

      for i := 0 to FrmChangeOfTenancySub.fAgreements.Count - 1 do
      begin
        vAgreeBalance := StrtoFloat(FrmChangeOfTenancySub.fAgreements[i].agreementBalance.Replace(ASCII_POUND,''));
        Actual := (not vAgreeBalance.IsNan) and (Abs(vAgreeBalance) >= 0);
      end;

    except
      Actual := false;
    end;

    Assert.AreEqual(Expected, Actual, 'Agreement Balance error.');

  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.TestCancelCotWhenNotPerformed;
var
  frm: TFrmChangeOfTenancySub;
begin
  frm := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    Assert.IsFalse(frm.fHasPerformedCot, 'CancelCot should be called when CoT not performed');
  finally
    FreeAndNil(frm);
  end;
end;

procedure TChangeTenancyTests.TestNoCancelCotWhenPerformed;
var
  frm: TFrmChangeOfTenancySub;
begin
  frm := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');
  try
    frm.btnSubmitClick(nil);
    Assert.IsTrue(frm.fHasPerformedCot, 'CancelCot should not be called when CoT was performed');
  finally
    FreeAndNil(frm);
  end;
end;

procedure TChangeTenancyTests.ExecuteCancelCotWithoutError;
var
  FrmChangeOfTenancySub: TFrmChangeOfTenancySub;
begin
  FrmChangeOfTenancySub := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Larry Page');

  try
    FrmChangeOfTenancySub.CancelCot;

    Assert.Pass('Cancel Change of Tenancy executed without exceptions');
  finally
    FreeAndNil(FrmChangeOfTenancySub);
  end;
end;

procedure TChangeTenancyTests.TestIntegerValidation;
var
  frm: TFrmChangeOfTenancySub;
  edit: TEdit;
  key: Char;
begin
  frm := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Test User');
  try
    edit := TEdit.Create(frm);
    edit.Parent := frm;

    key := '5';
    frm.ValidateInput<Integer>(edit, key);
    Assert.AreEqual('5', key, 'Digits should be allowed for Integer validation');

    key := 'A';
    frm.ValidateInput<Integer>(edit, key);
    Assert.AreEqual(#0, key, 'Letters should be rejected for Integer validation');

    key := '#';
    frm.ValidateInput<Integer>(edit, key);
    Assert.AreEqual(#0, key, 'Symbols should be rejected for Integer validation');

    key := #8;
    frm.ValidateInput<Integer>(edit, key);
    Assert.AreEqual(#8, key, 'Backspace should always be allowed');

    edit.Text := '2147483647';
    key := '9';
    frm.ValidateInput<Integer>(edit, key);
    Assert.AreEqual(#0, key, 'Input exceeding Integer bounds should be rejected');
  finally
    edit.Free;
    frm.Free;
  end;
end;

procedure TChangeTenancyTests.TestDoubleValidation;
var
  frm: TFrmChangeOfTenancySub;
  edit: TEdit;
  key: Char;
begin
  frm := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Test User');
  try
    edit := TEdit.Create(frm);
    edit.Parent := frm;

    key := '5';
    frm.ValidateInput<Double>(edit, key);
    Assert.AreEqual('5', key, 'Digits should be allowed for Double validation');

    key := '.';
    frm.ValidateInput<Double>(edit, key);
    Assert.AreEqual('.', key, 'Decimal point should be allowed for Double validation');

    key := 'A';
    frm.ValidateInput<Double>(edit, key);
    Assert.AreEqual(#0, key, 'Letters should be rejected for Double validation');

    key := '#';
    frm.ValidateInput<Double>(edit, key);
    Assert.AreEqual(#0, key, 'Symbols except decimal point should be rejected for Double validation');

    key := #8;
    frm.ValidateInput<Double>(edit, key);
    Assert.AreEqual(#8, key, 'Backspace should always be allowed');

    edit.Text := '3.14';
    key := '.';
    frm.ValidateInput<Double>(edit, key);
    Assert.AreEqual(#0, key, 'Only one decimal point should be allowed');

    edit.Text := '9999999999999999';
    key := '9';
    frm.ValidateInput<Double>(edit, key);
    Assert.AreEqual('9', key, 'Very large numbers should be allowed for Double');
  finally
    edit.Free;
    frm.Free;
  end;
end;

procedure TChangeTenancyTests.TestStringValidation;
var
  frm: TFrmChangeOfTenancySub;
  edit: TEdit;
  key: Char;
begin
  frm := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Test User');
  try
    edit := TEdit.Create(frm);
    edit.Parent := frm;

    key := 'A';
    frm.ValidateInput<string>(edit, key);
    Assert.AreEqual('A', key, 'Letters should be allowed for String validation');

    key := 'a';
    frm.ValidateInput<string>(edit, key);
    Assert.AreEqual('a', key, 'Lowercase letters should be allowed for String validation');

    key := '5';
    frm.ValidateInput<string>(edit, key);
    Assert.AreEqual(#0, key, 'Digits should be rejected for String validation');

    key := '#';
    frm.ValidateInput<string>(edit, key);
    Assert.AreEqual(#0, key, 'Symbols should be rejected for String validation');

    key := #8;
    frm.ValidateInput<string>(edit, key);
    Assert.AreEqual(#8, key, 'Backspace should always be allowed');
  finally
    edit.Free;
    frm.Free;
  end;
end;

procedure TChangeTenancyTests.TestDecimalOnlyKeyPress;
var
  frm: TFrmChangeOfTenancySub;
  sender: TObject;
  key: Char;
begin
  frm := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Test User');
  try
    sender := TEdit.Create(frm);
    TEdit(sender).Parent := frm;

    key := '5';
    frm.DecimalOnlyKeyPress(sender, key);
    Assert.AreEqual('5', key, 'Digits should be allowed for Decimal validation');

    key := '.';
    frm.DecimalOnlyKeyPress(sender, key);
    Assert.AreEqual('.', key, 'Decimal point should be allowed for Decimal validation');

    key := 'A';
    frm.DecimalOnlyKeyPress(sender, key);
    Assert.AreEqual(#0, key, 'Letters should be rejected for Decimal validation');

    key := #8;
    frm.DecimalOnlyKeyPress(sender, key);
    Assert.AreEqual(#8, key, 'Backspace should always be allowed');
  finally
    TEdit(sender).Free;
    frm.Free;
  end;
end;

procedure TChangeTenancyTests.TestReadDetailsHasPrePay;
var
  readDetails: TReadDetails;
begin
  readDetails := TReadDetails.Create;
  try
    SetLength(readDetails.PrePayReads, 0);
    Assert.IsFalse(readDetails.HasPrePay, 'HasPrePay should return false with no PrePayReads');

    SetLength(readDetails.PrePayReads, 1);
    readDetails.PrePayReads[0] := TPrePayRead.Create;
    readDetails.PrePayReads[0].MeterBalance := 100.50;
    Assert.IsTrue(readDetails.HasPrePay, 'HasPrePay should return true with PrePayReads');

    SetLength(readDetails.PrePayReads, 0);
    Assert.IsFalse(readDetails.HasPrePay, 'HasPrePay should return false with empty PrePayReads array');
  finally
    readDetails.Free;
  end;
end;

procedure TChangeTenancyTests.TestReadDetailsGetPrePayRead;
var
  readDetails: TReadDetails;
  prePayRead: TPrePayRead;
begin
  readDetails := TReadDetails.Create;
  try
    SetLength(readDetails.PrePayReads, 1);
    prePayRead := TPrePayRead.Create;
    prePayRead.MeterBalance := 100.50;
    prePayRead.DebtBalance := 25.75;
    prePayRead.LastUpdated := '01/05/2025';
    readDetails.PrePayReads[0] := prePayRead;

    Assert.IsNotNull(readDetails.GetPrePayRead, 'GetPrePayRead should not return nil');
    Assert.AreEqual<Double>(100.50, readDetails.GetPrePayRead.MeterBalance, 'GetPrePayRead should return the correct PrePayRead');
    Assert.AreEqual<Double>(25.75, readDetails.GetPrePayRead.DebtBalance, 'GetPrePayRead should return the correct PrePayRead');
    Assert.AreEqual<string>('01/05/2025', readDetails.GetPrePayRead.LastUpdated, 'GetPrePayRead should return the correct PrePayRead');

    SetLength(readDetails.PrePayReads, 0);
    Assert.IsNull(readDetails.GetPrePayRead, 'GetPrePayRead should return nil with empty PrePayReads array');
  finally
    readDetails.Free;
  end;
end;

procedure TChangeTenancyTests.TestReadDetailsHasConsumption;
var
  readDetails: TReadDetails;
begin
  readDetails := TReadDetails.Create;
  try
    SetLength(readDetails.ConsumptionReads, 0);
    Assert.IsFalse(readDetails.HasConsumption, 'HasConsumption should return false with no ConsumptionReads');

    SetLength(readDetails.ConsumptionReads, 1);
    readDetails.ConsumptionReads[0] := TConsumptionRead.Create;
    readDetails.ConsumptionReads[0].Consumption := 450;
    Assert.IsTrue(readDetails.HasConsumption, 'HasConsumption should return true with ConsumptionReads');

    SetLength(readDetails.ConsumptionReads, 0);
    Assert.IsFalse(readDetails.HasConsumption, 'HasConsumption should return false with empty ConsumptionReads array');
  finally
    readDetails.Free;
  end;
end;

procedure TChangeTenancyTests.TestReadDetailsGetConsumptionRead;
var
  readDetails: TReadDetails;
  consumptionRead: TConsumptionRead;
begin
  readDetails := TReadDetails.Create;
  try
    SetLength(readDetails.ConsumptionReads, 1);
    consumptionRead := TConsumptionRead.Create;
    consumptionRead.Consumption := 450;
    consumptionRead.Day := 200;
    consumptionRead.Night := 250;
    consumptionRead.LastUpdated := '01/05/2025';
    readDetails.ConsumptionReads[0] := consumptionRead;

    Assert.IsNotNull(readDetails.GetConsumptionRead, 'GetConsumptionRead should not return nil');
    Assert.AreEqual<Double>(450, readDetails.GetConsumptionRead.Consumption, 'GetConsumptionRead should return the correct ConsumptionRead');
    Assert.AreEqual<Double>(200, readDetails.GetConsumptionRead.Day, 'GetConsumptionRead should return the correct ConsumptionRead');
    Assert.AreEqual<Double>(250, readDetails.GetConsumptionRead.Night, 'GetConsumptionRead should return the correct ConsumptionRead');
    Assert.AreEqual<string>('01/05/2025', readDetails.GetConsumptionRead.LastUpdated, 'GetConsumptionRead should return the correct ConsumptionRead');

    SetLength(readDetails.ConsumptionReads, 0);
    Assert.IsNull(readDetails.GetConsumptionRead, 'GetConsumptionRead should return nil with empty ConsumptionReads array');
  finally
    readDetails.Free;
  end;
end;

procedure TChangeTenancyTests.TestBuildIncomingSupplies;
var
  frm: TFrmChangeOfTenancySub;
  outgoingSupply: TOutgoingSupply;
  readDetails: TReadDetails;
  i: Integer;
  foundPrepay, foundCredit: Boolean;
  outgoingId, incomingId: Int64;
begin
  frm := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Test User');
  try
    if frm.fOutgoingSupplies.Count = 0 then
    begin
      outgoingSupply := TOutgoingSupply.Create;
      outgoingSupply.serviceId := 12345;
      outgoingSupply.serviceTypeId := 'E';
      outgoingSupply.span := 2000022710812;
      outgoingSupply.agreementId := 98765;

      frm.fOutgoingSupplies.Add(outgoingSupply);
    end;

    outgoingSupply := frm.fOutgoingSupplies[0];
    outgoingId := outgoingSupply.serviceId;

    readDetails := TReadDetails.Create;
    readDetails.Mpxn := outgoingSupply.span;
    SetLength(readDetails.PrePayReads, 1);
    readDetails.PrePayReads[0] := TPrePayRead.Create;
    readDetails.PrePayReads[0].MeterBalance := 100.50;

    frm.fReadDetailsMap.Clear;
    frm.fReadDetailsMap.AddOrSetValue(outgoingSupply.span, readDetails);

    frm.fIncomingSupplies.Clear;

    frm.BuildIncomingSupplies;

    Assert.AreEqual(frm.fOutgoingSupplies.Count, frm.fIncomingSupplies.Count,
      'Should create the same number of incoming supplies as outgoing supplies');

    Assert.IsTrue(frm.fIncomingSupplies.Count > 0, 'Should create at least one incoming supply');

    foundPrepay := False;
    for i := 0 to frm.fIncomingSupplies.Count - 1 do
    begin
      if frm.fIncomingSupplies[i].span = outgoingSupply.span then
      begin
        Assert.AreEqual(outgoingSupply.serviceId, frm.fIncomingSupplies[i].serviceId, 'serviceId should match');
        Assert.AreEqual(outgoingSupply.serviceTypeId, frm.fIncomingSupplies[i].serviceTypeId, 'serviceTypeId should match');
        Assert.AreEqual(outgoingSupply.agreementId, frm.fIncomingSupplies[i].agreementId, 'agreementId should match');
        Assert.AreEqual(PREPAY_MODE, frm.fIncomingSupplies[i].meterMode, 'meterMode should be PREPAY_MODE');
        foundPrepay := True;
        Break;
      end;
    end;

    Assert.IsTrue(foundPrepay, 'Should find matching incoming supply with PREPAY_MODE');

    frm.fIncomingSupplies.Clear();
    frm.fReadDetailsMap.Clear();

    readDetails := TReadDetails.Create;
    readDetails.Mpxn := outgoingSupply.span;
    SetLength(readDetails.PrePayReads, 0);

    frm.fReadDetailsMap.AddOrSetValue(outgoingSupply.span, readDetails);

    frm.BuildIncomingSupplies;

    Assert.AreEqual(frm.fOutgoingSupplies.Count, frm.fIncomingSupplies.Count,
      'Should create the same number of incoming supplies as outgoing supplies');

    Assert.IsTrue(frm.fIncomingSupplies.Count > 0, 'Should create at least one incoming supply');

    foundCredit := False;
    for i := 0 to frm.fIncomingSupplies.Count - 1 do
    begin
      if frm.fIncomingSupplies[i].span = outgoingSupply.span then
      begin
        Assert.AreEqual(outgoingSupply.serviceId, frm.fIncomingSupplies[i].serviceId, 'serviceId should match');
        Assert.AreEqual(outgoingSupply.serviceTypeId, frm.fIncomingSupplies[i].serviceTypeId, 'serviceTypeId should match');
        Assert.AreEqual(outgoingSupply.agreementId, frm.fIncomingSupplies[i].agreementId, 'agreementId should match');
        Assert.AreEqual(CREDIT_MODE, frm.fIncomingSupplies[i].meterMode, 'meterMode should be CREDIT_MODE');
        foundCredit := True;
        Break;
      end;
    end;

    Assert.IsTrue(foundCredit, 'Should find matching incoming supply with CREDIT_MODE');

  finally
    FreeAndNil(frm);
  end;
end;

procedure TChangeTenancyTests.TestMeterBalanceChange;
var
  frm: TFrmChangeOfTenancySub;
  incomingSupply: TIncomingSupply;
  edit: TLabeledEdit;
begin
  frm := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Test User');
  try
    incomingSupply := TIncomingSupply.Create;
    incomingSupply.meterBalance := 0;

    edit := TLabeledEdit.Create(frm);
    edit.Parent := frm;
    edit.Text := '123.45';
    edit.Tag := NativeInt(incomingSupply);

    frm.MeterBalanceChange(edit);

    Assert.AreEqual<Double>(123.45, incomingSupply.meterBalance, 'Meter balance should be updated correctly');

    edit.Text := 'abc';
    frm.MeterBalanceChange(edit);

    Assert.AreEqual<Double>(123.45, incomingSupply.meterBalance, 'Meter balance should not change with invalid input');

    edit.Tag := 0;
    edit.Text := '999.99';
    frm.MeterBalanceChange(edit);

    Assert.Pass('MeterBalanceChange handled nil incomingSupply gracefully');

  finally
    incomingSupply.Free;
    FreeAndNil(frm);
  end;
end;

procedure TChangeTenancyTests.TestDebtBalanceChange;
var
  frm: TFrmChangeOfTenancySub;
  incomingSupply: TIncomingSupply;
  edit: TLabeledEdit;

begin
  frm := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Test User');
  incomingSupply := TIncomingSupply.Create;

  try
    incomingSupply.debtBalance := 0;
    edit := TLabeledEdit.Create(frm);
    edit.Parent := frm;
    edit.Text := '73.84';
    edit.Tag := NativeInt(incomingSupply);

    frm.DebtBalanceChange(edit);

    Assert.AreEqual<Double>(73.84, incomingSupply.DebtBalance, 'Debt balance must be updated correctly.');

    edit.Text := 'W@d9!';
    frm.DebtBalanceChange(edit);

    Assert.AreEqual<Double>(73.84, incomingSupply.DebtBalance, 'Debt balance must not change with invalid input.');

    edit.Tag := 0;
    edit.Text := '999.99';
    frm.DebtBalanceChange(edit);

    Assert.Pass('Debt Balance Change executed successfully.');

  finally
    FreeAndNil(incomingSupply);
    FreeAndNil(frm);
  end;

end;

procedure TChangeTenancyTests.TestRecoveryRateChange;
var
  frm: TFrmChangeOfTenancySub;
  incomingSupply: TIncomingSupply;
  edit: TLabeledEdit;

begin
  frm := TFrmChangeOfTenancySub.Create(nil, 1313306223, 'Test User');
  incomingSupply := TIncomingSupply.Create;

  try
    incomingSupply.recoveryRate := 20;

    edit := TLabeledEdit.Create(frm);
    edit.Parent := frm;
    edit.Text := '18.8';
    edit.Tag := NativeInt(incomingSupply);

    frm.RecoveryRateChange(edit);

    Assert.AreEqual<Double>(18.8, incomingSupply.RecoveryRate, 'Recovery Rate must be updated correctly.');

    edit.Text := 'W@d9!';
    frm.RecoveryRateChange(edit);

    Assert.AreEqual<Double>(18.8, incomingSupply.RecoveryRate, 'Recovery Rate must not change with invalid input.');

    edit.Tag := 0;
    edit.Text := '99.99';
    frm.RecoveryRateChange(edit);

    Assert.Pass('Recovery Rate executed successfully.');

  finally
    FreeAndNil(incomingSupply);
    FreeAndNil(frm);
  end;
end;

procedure TChangeTenancyTests.TestOutgoingCustomerWalletBalance;
var
  frm: TFrmChangeOfTenancySub;
  Actual: Boolean;
begin
  frm := TFrmChangeOfTenancySub.Create(nil, 1000818150, 'Test User');

  try
    frm.GetWalletBalance;

    Actual := (frm.edtOutCustWalletBalance.Text <> '');
    Assert.IsTrue(Actual, '�4.50');

  finally
    FreeAndNil(frm);
  end;
end;

[Test]
procedure TChangeTenancyTests.TestOutgoingCustomerWalletBalanceWarning;
var
  frm: TFrmChangeOfTenancySub;
  Actual: Boolean;
begin
  frm := TFrmChangeOfTenancySub.Create(nil, 3146310133, 'Test User');

  try
    frm.GetWalletBalance;

    Actual := (frm.edtOutCustWalletBalance.Text = 'Customer does not have wallet.');
    Assert.IsTrue(Actual, 'Wallet balance should display "Customer does not have wallet" for warning case');

  finally
    FreeAndNil(frm);
  end;
end;

procedure TChangeTenancyTests.TestTmaSupplyIdentification;
var
  frm: TFrmChangeOfTenancySub;
  i: Integer;
  testCustomerId: Int64;
  outgoingTmaFound, incomingTmaFound: Boolean;
begin
  testCustomerId := 1170872439;
  outgoingTmaFound := False;
  incomingTmaFound := False;

  frm := TFrmChangeOfTenancySub.Create(nil, testCustomerId, 'Test User TMA');

  try
    Assert.IsTrue(frm.fOutgoingSupplies.Count > 0, 'Should have outgoing supplies');

    for i := 0 to frm.fOutgoingSupplies.Count - 1 do
    begin
      outgoingTmaFound := True;
      Assert.IsTrue(frm.fOutgoingSupplies[i].isTma,
        Format('Outgoing supply for meter %d should be TMA', [frm.fOutgoingSupplies[i].span]));
    end;

    Assert.IsTrue(outgoingTmaFound, 'Outgoing supplies should be found');

    Assert.IsNotNull(frm.fIncomingSupplies, 'Incoming supplies should not be nil');
    Assert.IsTrue(frm.fIncomingSupplies.Count > 0, 'Should have incoming supplies');

    for i := 0 to frm.fIncomingSupplies.Count - 1 do
    begin
      incomingTmaFound := True;
      Assert.IsTrue(frm.fIncomingSupplies[i].isTma,
        Format('Incoming supply for meter %d should be TMA', [frm.fIncomingSupplies[i].span]));
    end;

    Assert.IsTrue(incomingTmaFound, 'Incoming supplies should be found');

    Assert.Pass('TMA test completed successfully. All meters found in both outgoing and incoming as TMA');

  finally
    FreeAndNil(frm);
  end;
end;

initialization
 {$IFDEF CUSTOMER}
   TDUnitX.RegisterTestFixture(TChangeTenancyTests);
 {$ENDIF}
end.