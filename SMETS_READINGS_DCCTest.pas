unit uSmetsReadingsDccTest;

interface
{$DEFINE CRMTEST}
uses
  DUnitX.TestFramework, LoginUnit, Main, CrmCommon, DMImages, vcl.Forms,
  UELSqlUtils, System.SysUtils, SmetsCommon, SMETS_READINGS_DCC,
  System.DateUtils, Common, vcl.Controls, System.StrUtils,
  Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.StdCtrls, DB, Vcl.Dialogs;

type
  TFrm_Smets_Readings_DccSub = class(TFrm_Smets_Readings_Dcc);
  [TestFixture]
  TSMETSReadingsDCCTests = class

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure CreateFormSuccessfully;

    [Test]
    procedure GetMetersSuccess;

    [Test]
    procedure GetReadingsSuccess;

    [Test]
    procedure RefreshDataSuccess;
    
    [Test]
    procedure TabsOrderedCorrectly;
  end;
implementation
var
  Frm_Smets_Readings_DccSub: TFrm_Smets_Readings_DccSub;

  procedure TSMETSReadingsDCCTests.Setup;
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

procedure TSMETSReadingsDCCTests.TearDown;
begin
  FreeAndNil(gSqlUtil);
  FreeAndNil(FRM_MAIN);
  FreeAndNil(DM_Images);
  FreeAndNil(FRM_Login);
end;

procedure TSMETSReadingsDCCTests.CreateFormSuccessfully;
var
  TestSpan: string;
  Form: TFrm_Smets_Readings_DccSub;
begin
  TestSpan := '2000022710812';
  Form := TFrm_Smets_Readings_DccSub.Create(nil, TestSpan);
  try
    try
      Assert.IsNotNull(Form, 'Form should be created successfully');
      Assert.AreEqual(TestSpan, Form.fSpan, 'Form should have the correct span value');
    except
      on E: exception do
        showmessage('Unable to retrieve Meter Readings: '+E.Message);
    end;
  finally
    FreeAndNil(Form);
  end;
end;

procedure TSMETSReadingsDCCTests.GetMetersSuccess;
var
  TestSpan: string;
  Form: TFrm_Smets_Readings_DccSub;
begin
  TestSpan := '2000022710812';
  Form := TFrm_Smets_Readings_DccSub.Create(nil, TestSpan);
  try
    try
      Form.GetMeters;
    except
      on E: Exception do
        Assert.Fail('GetMeters completed with exceptions: ' + E.Message);
    end;
  finally
    FreeAndNil(Form);
  end;
end;

procedure TSMETSReadingsDCCTests.GetReadingsSuccess;
var
  TestSpan: string;
  Form: TFrm_Smets_Readings_DccSub;
begin
  TestSpan := '2000022710812';
  Form := TFrm_Smets_Readings_DccSub.Create(nil, TestSpan);
  try
    try
      Form.GetReadings;
    except
      on E: Exception do
        Assert.Fail('GetReadings completed with exceptions: ' + E.Message);
    end;
  finally
    FreeAndNil(Form);
  end;
end;

procedure TSMETSReadingsDCCTests.RefreshDataSuccess;
var
  TestSpan: string;
  Form: TFrm_Smets_Readings_DccSub;
begin
  TestSpan := '2000022710812';
  Form := TFrm_Smets_Readings_DccSub.Create(nil, TestSpan);
  try
    try
      Form.RefreshData;
      Assert.IsTrue(Form.tabRegisters.Tabs.Count = 0,
        'After RefreshData with no valid data, there should be no tabs');
    except
      on E: Exception do
        Assert.Fail('RefreshData threw an exception: ' + E.Message);
    end;
  finally
    FreeAndNil(Form);
  end;
end;

procedure TSMETSReadingsDCCTests.TabsOrderedCorrectly;
var
  TestSpan: string;
  Form: TFrm_Smets_Readings_DccSub;
  MockData: TOracleDataSet;
begin
  TestSpan := '2000022710812';
  Form := TFrm_Smets_Readings_DccSub.Create(nil, TestSpan);
  try
    try
      // Setup mock data with Day, Night and Total readings
      Form.RegistersQuery.Active := True;
      
      // Force tab creation with specified TPR names
      Form.fHasE7AndE10Reads := False;
      
      // Create a string list to simulate the data that BuildTabs would process
      with Form.RegistersQuery do
      begin
        // Clear any existing data
        Close;
        Fields.Clear;
        
        // Add the TPR_NAME field (needed by BuildTabs)
        FieldDefs.Clear;
        FieldDefs.Add('TPR_NAME', ftString, 20);
        FieldDefs.Add('FLOW', ftString, 20);
        FieldDefs.Add('READDATE', ftDateTime);
        FieldDefs.Add('KWH', ftFloat);
        CreateDataSet;
        
        // Add records with Day, Night, and Total
        Append;
        FieldByName('TPR_NAME').AsString := 'Day';
        Post;
        
        Append;
        FieldByName('TPR_NAME').AsString := 'Night';
        Post;
        
        Append;
        FieldByName('TPR_NAME').AsString := 'Total';
        Post;
        
        First; // Position at first record
      end;
      
      // Call BuildTabs which should order the tabs
      Form.BuildTabs;
      
      // Assert that tabs are created in the correct order: Day / Night / Total
      Assert.AreEqual(3, Form.tabRegisters.Tabs.Count, 'Should have 3 tabs');
      Assert.AreEqual('Day', Form.tabRegisters.Tabs[0], 'First tab should be Day');
      Assert.AreEqual('Night', Form.tabRegisters.Tabs[1], 'Second tab should be Night');
      Assert.AreEqual('Total', Form.tabRegisters.Tabs[2], 'Third tab should be Total');
    except
      on E: Exception do
        Assert.Fail('Tab ordering test failed with exception: ' + E.Message);
    end;
  finally
    FreeAndNil(Form);
  end;
end;

initialization
  {$IFDEF METERING}
  TDUnitX.RegisterTestFixture(TSMETSReadingsDCCTests);
  {$ENDIF}
end.
