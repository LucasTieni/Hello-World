unit uSmetsCheckCommsDccTest;

interface

uses
  DUnitX.TestFramework, LoginUnit, Main, CrmCommon, DMImages, vcl.Forms,
  UELSqlUtils, System.SysUtils, smets_check_comms_DCC, Oracle, OracleData,
  Vcl.ExtCtrls, Vcl.StdCtrls, Common;

type
  TFRM_CHECK_COMMSSub = class(TFRM_CHECK_COMMS);
  [TestFixture]
  TSmetsCheckCommsDccTests = class
  private
    FrmCheckCommsSub: TFRM_CHECK_COMMSSub;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure CreateFormSuccessfully;

    [Test]
    procedure TestGetRaiseEnquireDelayWithValidData;

    [Test]
    procedure TestGetRaiseEnquireDelayCalculation;

    [Test]
    procedure TestTimerIntervalSetCorrectly;

  end;

implementation

procedure TSmetsCheckCommsDccTests.Setup;
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

procedure TSmetsCheckCommsDccTests.TearDown;
begin
  if Assigned(FrmCheckCommsSub) then
    FreeAndNil(FrmCheckCommsSub);
  FreeAndNil(gSqlUtil);
  FreeAndNil(FRM_MAIN);
  FreeAndNil(DM_Images);
  FreeAndNil(FRM_Login);
end;

procedure TSmetsCheckCommsDccTests.CreateFormSuccessfully;
var
  TestCustomerId, TestAgreementId, TestPremiseId: Int64;
  TestPremiseInfo: string;
begin
  TestCustomerId := 1313306223;
  TestAgreementId := 1234567890;
  TestPremiseId := 9876543210;
  TestPremiseInfo := 'Test Premise Info';

  FrmCheckCommsSub := TFRM_CHECK_COMMSSub.Create(nil, TestCustomerId, TestAgreementId, TestPremiseId, TestPremiseInfo);
  try
    Assert.IsNotNull(FrmCheckCommsSub, 'Form should be created successfully');
  finally
    FreeAndNil(FrmCheckCommsSub);
  end;
end;

procedure TSmetsCheckCommsDccTests.TestGetRaiseEnquireDelayWithValidData;
var
  TestCustomerId, TestAgreementId, TestPremiseId: Int64;
  TestPremiseInfo: string;
  InitialInterval: Integer;
begin
  TestCustomerId := 1313306223;
  TestAgreementId := 1234567890;
  TestPremiseId := 9876543210;
  TestPremiseInfo := 'Test Premise Info';

  FrmCheckCommsSub := TFRM_CHECK_COMMSSub.Create(nil, TestCustomerId, TestAgreementId, TestPremiseId, TestPremiseInfo);
  try
    // Store initial timer interval (should be 0 before GetRaiseEnquireDelay is called)
    InitialInterval := FrmCheckCommsSub.Timer_Comms.Interval;

    // The method is called during FormCreate, but let's verify the timer interval was set
    Assert.IsTrue(FrmCheckCommsSub.Timer_Comms.Interval > 0,
      'Timer interval should be set to a positive value after GetRaiseEnquireDelay is called');

  finally
    FreeAndNil(FrmCheckCommsSub);
  end;
end;

procedure TSmetsCheckCommsDccTests.TestGetRaiseEnquireDelayCalculation;
var
  TestCustomerId, TestAgreementId, TestPremiseId: Int64;
  TestPremiseInfo: string;
  ExpectedInterval: Integer;
  RaiseEnquiryDelay: Double;
begin
  TestCustomerId := 1313306223;
  TestAgreementId := 1234567890;
  TestPremiseId := 9876543210;
  TestPremiseInfo := 'Test Premise Info';

  // Get the actual value from the database
  RaiseEnquiryDelay := gSqlUtil.SelectQueryDouble('select item_value from crm.standing_data where item_name = ''CRM_COMMS_CHECK_TIMEOUT''');

  // Calculate expected interval based on the formula: Trunc((raiseEnquiryDelay / 50) * 1000)
  ExpectedInterval := Trunc((RaiseEnquiryDelay / 50) * 1000);

  FrmCheckCommsSub := TFRM_CHECK_COMMSSub.Create(nil, TestCustomerId, TestAgreementId, TestPremiseId, TestPremiseInfo);
  try
    // FormCreate calls GetRaiseEnquireDelay automatically
    Assert.AreEqual<Integer>(ExpectedInterval, FrmCheckCommsSub.Timer_Comms.Interval,
      'Timer interval should match the calculated value from the database');

  finally
    FreeAndNil(FrmCheckCommsSub);
  end;
end;

procedure TSmetsCheckCommsDccTests.TestTimerIntervalSetCorrectly;
var
  TestCustomerId, TestAgreementId, TestPremiseId: Int64;
  TestPremiseInfo: string;
  ActualInterval: Integer;
begin
  TestCustomerId := 1313306223;
  TestAgreementId := 1234567890;
  TestPremiseId := 9876543210;
  TestPremiseInfo := 'Test Premise Info';

  FrmCheckCommsSub := TFRM_CHECK_COMMSSub.Create(nil, TestCustomerId, TestAgreementId, TestPremiseId, TestPremiseInfo);
  try
    ActualInterval := FrmCheckCommsSub.Timer_Comms.Interval;

    // Verify the timer interval is within a reasonable range (should be milliseconds)
    Assert.IsTrue(ActualInterval > 0, 'Timer interval should be greater than 0');
    Assert.IsTrue(ActualInterval < 1000000, 'Timer interval should be less than 1000000 milliseconds');

  finally
    FreeAndNil(FrmCheckCommsSub);
  end;
end;

initialization
  {$IFDEF METERING}
  TDUnitX.RegisterTestFixture(TSmetsCheckCommsDccTests);
  {$ENDIF}

end.
