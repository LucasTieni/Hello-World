unit uSmetChangeModeDcc;

interface

uses
  DUnitX.TestFramework, LoginUnit, Main, CrmCommon, DataModule, DMImages, vcl.Forms, Vcl.ExtCtrls,
  UELSqlUtils, System.SysUtils, System.UITypes, SmetsCommon, Smets_Change_Mode_DCC;

type
  TFrm_Smets_Change_Mode_DccSub = class(TFrm_Smets_Change_Mode_Dcc);

  [TestFixture]
  TTestCaseSmetsDCC = class
  private
    Frm_Smets_Change_Mode_DccSub : TFrm_Smets_Change_Mode_DccSub;

  public
    [SetupFixture]
    procedure SetupFixture;

    [TearDownFixture]
    procedure TearDownFixture;

    [Test]
    [TestCase('Submit change of mode dcc OK', '0,3080000036655,3032313674,303231367401,03/02/2025')]
    procedure TestSubmitChangeOk(aService: integer; aSpan: string; aCustomerId, aAgreementId: Int64; aEfsdmsMtd: TDate);

        [Test]
    [TestCase('Submit change of mode fails when efective date before last Mtd date', '0,1200039581670,3032313674,303231367401,03/02/2025')]
    procedure TestEffectiveDateBeforeLastMtdDate(aService: integer; aSpan: string; aCustomerId, aAgreementId: Int64; aEfsdmsMtd: TDate);

    [Test]
    [TestCase('IsChangeModeAllowed returns true when mode change is allowed', '0,1200039581670,3032313674,303231367401,03/02/2025')]
    procedure TestIsChangeModeAllowedTrue(aService: integer; aSpan: string; aCustomerId, aAgreementId: Int64; aEfsdmsMtd: TDate);

    [Test]
    [TestCase('IsChangeModeAllowed returns false when mode change is not allowed', '0,1100015218988,3032313674,303231367401,03/02/2025')]
    procedure TestIsChangeModeAllowedFalse(aService: integer; aSpan: string; aCustomerId, aAgreementId: Int64; aEfsdmsMtd: TDate);

  end;

implementation

procedure TTestCaseSmetsDCC.SetupFixture;
begin

  Application.CreateForm(TFRM_MAIN, FRM_MAIN);
  Application.CreateForm(TDM_Images, DM_Images);
  Application.CreateForm(TFRM_Login, FRM_Login);
  Application.CreateForm(TMain_Data_Module, Main_Data_Module);

  FRM_LOGIN.MainSession.LogonUsername := 'TEST01';
  FRM_LOGIN.MainSession.LogonPassword := 'TEST01';
  FRM_LOGIN.MainSession.LogonDatabase := 'DEVCRM';
  FRM_LOGIN.MainSession.Connected     := true;
  UserID := FRM_LOGIN.MainSession.LogonUsername;

  gSqlUtil := TSqlUtil.create(FRM_Login.mainsession);

end;

procedure TTestCaseSmetsDCC.TearDownFixture;
begin
  FreeAndNil(gSqlUtil);
  FreeAndNil(FRM_MAIN);
  FreeAndNil(DM_Images);
  FreeAndNil(FRM_Login);
  FreeAndNil(Main_Data_Module);
end;

procedure TTestCaseSmetsDCC.TestEffectiveDateBeforeLastMtdDate(aService: integer; aSpan: string; aCustomerId, aAgreementId: Int64; aEfsdmsMtd: TDate);
var
  expected, actual: TModalResult;
begin
  try
    expected := mrNone;
    Frm_Smets_Change_Mode_DccSub := TFrm_Smets_Change_Mode_DccSub.Create(nil, aService, aSpan, '1', aCustomerId, aAgreementID, aEfsdmsMtd);
    Frm_Smets_Change_Mode_DccSub.edtEffectiveDate.Date := StrToDate('03/01/2025');

    Frm_Smets_Change_Mode_DccSub.SendBtn.Click;

    actual := Frm_Smets_Change_Mode_DccSub.ModalResult;
    Assert.AreEqual(expected, actual, 'Change fails when should');
  finally
    FreeAndNil(Frm_Smets_Change_Mode_DccSub);
  end;
end;

procedure TTestCaseSmetsDCC.TestSubmitChangeOk(aService: integer; aSpan: string; aCustomerId, aAgreementId: Int64; aEfsdmsMtd: TDate);
var
  expected, actual: TModalResult;
begin
  try
    expected := mrOk;
    Frm_Smets_Change_Mode_DccSub := TFrm_Smets_Change_Mode_DccSub.Create(nil, aService, aSpan, '1', aCustomerId, aAgreementID, aEfsdmsMtd);

    Frm_Smets_Change_Mode_DccSub.SendBtn.Click;

    actual := Frm_Smets_Change_Mode_DccSub.ModalResult;
    Assert.AreEqual(expected, actual, 'Change was not sucess when should');
  finally
    FreeAndNil(Frm_Smets_Change_Mode_DccSub);
  end;
end;

procedure TTestCaseSmetsDCC.TestIsChangeModeAllowedTrue(aService: integer; aSpan: string; aCustomerId, aAgreementId: Int64; aEfsdmsMtd: TDate);
var
  expected, actual: boolean;
begin
  try
    expected := true;
    Frm_Smets_Change_Mode_DccSub := TFrm_Smets_Change_Mode_DccSub.Create(nil, aService, aSpan, '1', aCustomerId, aAgreementID, aEfsdmsMtd);

    actual := Frm_Smets_Change_Mode_DccSub.IsChangeModeAllowed;

    Assert.AreEqual(expected, actual, 'IsChangeModeAllowed should return true when mode change is allowed');
  finally
    FreeAndNil(Frm_Smets_Change_Mode_DccSub);
  end;
end;

procedure TTestCaseSmetsDCC.TestIsChangeModeAllowedFalse(aService: integer; aSpan: string; aCustomerId, aAgreementId: Int64; aEfsdmsMtd: TDate);
var
  expected, actual: boolean;
begin
  try
    expected := false;
    Frm_Smets_Change_Mode_DccSub := TFrm_Smets_Change_Mode_DccSub.Create(nil, aService, aSpan, '1', aCustomerId, aAgreementID, aEfsdmsMtd);

    actual := Frm_Smets_Change_Mode_DccSub.IsChangeModeAllowed;

    Assert.AreEqual(expected, actual, 'IsChangeModeAllowed should return false when mode change is not allowed');
  finally
    FreeAndNil(Frm_Smets_Change_Mode_DccSub);
  end;
end;




initialization
{$IFDEF METERING}
  TDUnitX.RegisterTestFixture(TTestCaseSmetsDCC);
{$ENDIF}

end.