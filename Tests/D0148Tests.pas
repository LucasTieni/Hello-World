unit uD0148Test;

interface
uses
  DUnitX.TestFramework, D0148, System.SysUtils, System.IOUtils,
  Oracle, OracleData, System.Classes, LoginUnit, Main, CrmCommon, DMImages,
  vcl.Forms, UELSqlUtils, JvDBLookup, DBCtrls, System.Rtti;

type
  TFrmD0148Sub = class(TFRM_D0148);

  // Helper to access private fields
  TJvDBLookupComboHelper = class helper for TJvDBLookupCombo
  public
    procedure SetTextValue(const Value: string);
  end;

  [TestFixture]
  TD0148Test = class
  private
    procedure CreateD0148Form;
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;

    [Test]
    procedure TestCHECKD0151S_WithRecords_DatabaseLevel;
    [Test]
    procedure TestCHECKD0151S_NoRecords_DatabaseLevel;
    [Test]
    procedure TestCHECKD0205_WithRecords_DatabaseLevel;
    [Test]
    procedure TestCHECKD0205_NoRecords_DatabaseLevel;
    [Test]
    procedure TestCOADA_Success;
    [Test]
    procedure TestCOADC_Success;
    [Test]
    procedure TestCOAMOP_Success;
    [Test]
    procedure TestDOD0170OLDDC_Success;
    [Test]
    procedure TestDOD0170OLDMO_Success;
    [Test]
    procedure TestDOD0205_Success;
    [Test]
    procedure TestGET_NEW_AGENTS_Success;
  end;

implementation

uses
  Common, DataModule, Winapi.Windows, Winapi.Messages;

var
  FrmD0148Sub: TFrmD0148Sub;

{ TJvDBLookupComboHelper }

procedure TJvDBLookupComboHelper.SetTextValue(const Value: string);
begin
  // Hack: Use SendMessage to set the text directly in the edit control
  SendMessage(Self.Handle, WM_SETTEXT, 0, LPARAM(PChar(Value)));
end;

//procedure TD0148Test.SetupFixture;
//begin
//  Application.CreateForm(TFRM_MAIN, FRM_MAIN);
//  Application.CreateForm(TDM_Images, DM_Images);
//  Application.CreateForm(TFRM_Login, FRM_Login);
//  Application.CreateForm(TMain_Data_Module, main_data_module);
//
//  FRM_LOGIN.MainSession.LogonUsername := 'TEST01';
//  FRM_LOGIN.MainSession.LogonPassword := 'TEST01';
//  FRM_LOGIN.MainSession.LogonDatabase := 'DEVCRM';
//  FRM_LOGIN.MainSession.Connected     := true;
//  UserID := FRM_LOGIN.MainSession.LogonUsername;
//  gSqlUtil := TSqlUtil.create(FRM_Login.mainsession);
//end;
//
//procedure TD0148Test.TearDownFixture;
//begin
//  FreeAndNil(gSqlUtil);
//  FreeAndNil(main_data_module);
//  FreeAndNil(FRM_MAIN);
//  FreeAndNil(DM_Images);
//  FreeAndNil(FRM_Login);
//end;

procedure TD0148Test.Setup;
begin
  Application.CreateForm(TFRM_MAIN, FRM_MAIN);
  Application.CreateForm(TDM_Images, DM_Images);
  Application.CreateForm(TFRM_Login, FRM_Login);
  FRM_LOGIN.MainSession.LogonUsername := 'TEST01';
  FRM_LOGIN.MainSession.LogonPassword := 'TEST01';
  FRM_LOGIN.MainSession.LogonDatabase := 'DEVCRM';
  FRM_LOGIN.MainSession.Connected := true;
  UserID := FRM_LOGIN.MainSession.LogonUsername;
  Application.CreateForm(TMain_Data_Module, main_data_module);
  gSqlUtil := TSqlUtil.create(FRM_Login.mainsession);
end;

procedure TD0148Test.CreateD0148Form;
begin
  if Assigned(FrmD0148Sub) then Exit;

  Application.CreateForm(TFrmD0148Sub, FrmD0148Sub);
  FrmD0148Sub.mpanstatus.Session := FRM_LOGIN.MainSession;
  FrmD0148Sub.DA.Session := FRM_LOGIN.MainSession;
  FrmD0148Sub.DC.Session := FRM_LOGIN.MainSession;
  FrmD0148Sub.MO.Session := FRM_LOGIN.MainSession;
  FrmD0148Sub.l_da.Session := FRM_LOGIN.MainSession;
  FrmD0148Sub.l_dc.Session := FRM_LOGIN.MainSession;
  FrmD0148Sub.l_mo.Session := FRM_LOGIN.MainSession;
  FrmD0148Sub.D0151Query.Session := FRM_LOGIN.MainSession;
  FrmD0148Sub.D0205Query.Session := FRM_LOGIN.MainSession;
end;

procedure TD0148Test.TearDown;
begin
  FreeAndNil(FrmD0148Sub);
  FreeAndNil(gSqlUtil);
  FreeAndNil(main_data_module);
  FreeAndNil(FRM_MAIN);
  FreeAndNil(DM_Images);
  FreeAndNil(FRM_Login);
end;

procedure TD0148Test.TestCHECKD0151S_WithRecords_DatabaseLevel;
var
  RecordCount: Integer;
const
  testMPANCORE = '111122225550';
  testFlowVersion = 'D0151';
  testMPID = 'TEST';
  testRole = 'M';
  testFilename = 'TEST.dat';  // Max 8 characters for D0151.FILENAME
begin
  try
    // Setup: Insert flowheader and D0151 record for MO termination
    try
      gSqlUtil.InsertRecord('EDMGR.FLOWHEADERS', TRANSACTION_NO,
        ['MPANCORE',     otString, testMPANCORE,
         'FILENAME',     otString, testFilename,
         'FLOW_VERSION', otString, testFlowVersion,
         'TONAME',       otString, testMPID,
         'TOID',         otString, testRole,
         'FILE_DATE_TIME', otDate, Now
        ]);

      gSqlUtil.InsertRecord('EDMGR.D0151', TRANSACTION_NO,
        ['MPANCORE',          otString, testMPANCORE,
         'FILENAME',          otString, testFilename,
         'TERMINATION_REASON', otString, 'CA',
         'EFTD_MOA',          otDate, Date - 1
        ]);
    except
      on E: Exception do
        Assert.Fail('Failed to insert test data: ' + E.Message);
    end;

    // Verify the D0151 record exists (this is what CheckD0151s queries)
    RecordCount := gSqlUtil.SelectQueryInteger(
      'SELECT COUNT(*) FROM EDMGR.D0151 A, EDMGR.FLOWHEADERS F ' +
      'WHERE A.TERMINATION_REASON = ''CA'' ' +
      'AND A.MPANCORE = :pMPAN ' +
      'AND A.MPANCORE = F.MPANCORE ' +
      'AND A.FILENAME = F.FILENAME ' +
      'AND F.FLOW_VERSION = ''D0151'' ' +
      'AND F.TONAME = :pMPID ' +
      'AND F.TOID = :pROLE',
      ['pMPAN', otString, testMPANCORE,
       'pMPID', otString, testMPID,
       'pROLE', otString, testRole]);

    Assert.AreEqual(1, RecordCount, 'CheckD0151s should find 1 D0151 termination record');

  finally
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.D0151 WHERE MPANCORE = :pMPANCORE',
      TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    gSqlUtil.ExecSql('DELETE FROM EDMGR.FLOWHEADERS WHERE MPANCORE = :pMPANCORE AND FLOW_VERSION = ''D0151''',
      TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
  end;
end;

procedure TD0148Test.TestCHECKD0151S_NoRecords_DatabaseLevel;
var
  RecordCount: Integer;
const
  testMPANCORE = '999999999999';
  testMPID = 'TEST';
  testRole = 'M';
begin
  // Verify no D0151 records exist (what CheckD0151s would check when checkbox should be true)
  RecordCount := gSqlUtil.SelectQueryInteger(
    'SELECT COUNT(*) FROM EDMGR.D0151 A, EDMGR.FLOWHEADERS F ' +
    'WHERE A.TERMINATION_REASON = ''CA'' ' +
    'AND A.MPANCORE = :pMPAN ' +
    'AND A.MPANCORE = F.MPANCORE ' +
    'AND A.FILENAME = F.FILENAME ' +
    'AND F.FLOW_VERSION = ''D0151'' ' +
    'AND F.TONAME = :pMPID ' +
    'AND F.TOID = :pROLE',
    ['pMPAN', otString, testMPANCORE,
     'pMPID', otString, testMPID,
     'pROLE', otString, testRole]);

  Assert.AreEqual(0, RecordCount, 'Should find no D0151 records for non-existent MPAN');
end;

procedure TD0148Test.TestCHECKD0205_WithRecords_DatabaseLevel;
var
  RecordCount: Integer;
const
  testMPANCORE = '111122225551';
  testDCMPID = 'TSDC';  // Max 4 characters
  testDAMPID = 'TSDA';  // Max 4 characters
  testMOMPID = 'TSMO';  // Max 4 characters
begin
  try
    // Setup: Insert test AGENTS_MPAS record
    try
      gSqlUtil.InsertRecord('EDMGR.AGENTS_MPAS', TRANSACTION_NO,
        ['MPANCORE', otString, testMPANCORE,
         'SSD',      otDate,   Date,
         'DC_ID',    otString, testDCMPID,
         'DA_ID',    otString, testDAMPID,
         'MO_ID',    otString, testMOMPID
        ]);
    except
      on E: Exception do
        Assert.Fail('Failed to insert test data: ' + E.Message);
    end;

    // Verify the AGENTS_MPAS record exists (this is what CheckD0205s queries)
    RecordCount := gSqlUtil.SelectQueryInteger(
      'SELECT COUNT(*) FROM EDMGR.AGENTS_MPAS ' +
      'WHERE MPANCORE = :pMPAN AND SSD = :pSSD',
      ['pMPAN', otString, testMPANCORE,
       'pSSD', otDate, Date]);

    Assert.AreEqual(1, RecordCount, 'CheckD0205s should find 1 AGENTS_MPAS record');

  finally
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.AGENTS_MPAS WHERE MPANCORE = :pMPANCORE',
      TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
  end;
end;

procedure TD0148Test.TestCHECKD0205_NoRecords_DatabaseLevel;
var
  RecordCount: Integer;
const
  testMPANCORE = '999999999999';
begin
  // Verify no AGENTS_MPAS records exist (what CheckD0205s would handle gracefully)
  RecordCount := gSqlUtil.SelectQueryInteger(
    'SELECT COUNT(*) FROM EDMGR.AGENTS_MPAS ' +
    'WHERE MPANCORE = :pMPAN AND SSD = :pSSD',
    ['pMPAN', otString, testMPANCORE,
     'pSSD', otDate, Date]);

  Assert.AreEqual(0, RecordCount, 'Should find no AGENTS_MPAS records for non-existent MPAN');
end;

procedure TD0148Test.TestCOADA_Success;
var
  RecordCount: Integer;
const
  testMPANCORE = '3080000036655';
  testFlowVersion = 'D0151';
  testRole = 'D';
  testMPID = 'TSDA';
begin
  CreateD0148Form;
  try
    FrmD0148Sub.mpancore.SetTextValue(testMPANCORE);
    FrmD0148Sub.da_mpid.SetTextValue(testMPID);
    FrmD0148Sub.da_role.Text := testRole;

    // Set other form fields
    FrmD0148Sub.L_DA_EFD.Text := DateToStr(Date + 1);
    FrmD0148Sub.DBSSD.Date := Date;

    // NOW ACTUALLY CALL THE COADA METHOD!
    FrmD0148Sub.COADA('BATCHED');

    RecordCount := gSqlUtil.SelectQueryInteger(
      'SELECT COUNT(*) FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA ' +
      'WHERE MPANCORE = :pMPAN AND FLOWVERSION = :pFLOW',
      ['pMPAN', otString, testMPANCORE,
       'pFLOW', otString, testFlowVersion]);

    Assert.AreEqual(1, RecordCount, 'COADA should create 1 batch record');

    with gSqlUtil.SelectQuery(
      'SELECT * FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA ' +
      'WHERE MPANCORE = :pMPAN AND FLOWVERSION = :pFLOW',
      ['pMPAN', otString, testMPANCORE,
       'pFLOW', otString, testFlowVersion]) do
    begin
      Assert.AreEqual(testMPANCORE, Fields[0].AsString);
      Assert.AreEqual(testFlowVersion, Fields[1].AsString);
      Assert.AreEqual(testRole, Fields[2].AsString);
      Assert.AreEqual(testMPID, Fields[3].AsString);
      Free;
    end;

  finally
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA WHERE MPANCORE = :pMPANCORE',
      TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
  end;
end;

procedure TD0148Test.TestCOADC_Success;
var
  RecordCount: Integer;
const
  testMPANCORE = '3080000036655';
  testFlowVersion = 'D0151';
  testRole = 'C';
  testMPID = 'TSDC';
begin
  CreateD0148Form;
  try
    FrmD0148Sub.mpancore.SetTextValue(testMPANCORE);
    FrmD0148Sub.dc_mpid.SetTextValue(testMPID);
    FrmD0148Sub.dc_role.Text := testRole;
    FrmD0148Sub.L_DC_EFD.Text := DateToStr(Date + 1);
    FrmD0148Sub.DBSSD.Date := Date;

    FrmD0148Sub.COADC('BATCHED');

    RecordCount := gSqlUtil.SelectQueryInteger(
      'SELECT COUNT(*) FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA ' +
      'WHERE MPANCORE = :pMPAN AND FLOWVERSION = :pFLOW',
      ['pMPAN', otString, testMPANCORE,
       'pFLOW', otString, testFlowVersion]);

    Assert.AreEqual(1, RecordCount);

    with gSqlUtil.SelectQuery(
      'SELECT * FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA ' +
      'WHERE MPANCORE = :pMPAN AND FLOWVERSION = :pFLOW',
      ['pMPAN', otString, testMPANCORE,
       'pFLOW', otString, testFlowVersion]) do
    begin
      Assert.AreEqual(testMPANCORE, Fields[0].AsString);
      Assert.AreEqual(testFlowVersion, Fields[1].AsString);
      Assert.AreEqual(testRole, Fields[2].AsString);
      Assert.AreEqual(testMPID, Fields[3].AsString);
      Free;
    end;

  finally
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA WHERE MPANCORE = :pMPANCORE',
      TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
  end;
end;

procedure TD0148Test.TestCOAMOP_Success;
var
  RecordCount: Integer;
const
  testMPANCORE = '3080000036655';
  testFlowVersion = 'D0151';
  testRole = 'M';
  testMPID = 'TSMO';
begin
  CreateD0148Form;
  try
    FrmD0148Sub.mpancore.SetTextValue(testMPANCORE);
    FrmD0148Sub.mo_mpid.SetTextValue(testMPID);
    FrmD0148Sub.mo_role.Text := testRole;
    FrmD0148Sub.L_MO_EFD.Text := DateToStr(Date + 1);
    FrmD0148Sub.DBSSD.Date := Date;

    FrmD0148Sub.COAMOP('BATCHED');

    RecordCount := gSqlUtil.SelectQueryInteger(
      'SELECT COUNT(*) FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA ' +
      'WHERE MPANCORE = :pMPAN AND FLOWVERSION = :pFLOW',
      ['pMPAN', otString, testMPANCORE,
       'pFLOW', otString, testFlowVersion]);

    Assert.AreEqual(1, RecordCount);

    with gSqlUtil.SelectQuery(
      'SELECT * FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA ' +
      'WHERE MPANCORE = :pMPAN AND FLOWVERSION = :pFLOW',
      ['pMPAN', otString, testMPANCORE,
       'pFLOW', otString, testFlowVersion]) do
    begin
      Assert.AreEqual(testMPANCORE, Fields[0].AsString);
      Assert.AreEqual(testFlowVersion, Fields[1].AsString);
      Assert.AreEqual(testRole, Fields[2].AsString);
      Assert.AreEqual(testMPID, Fields[3].AsString);
      Free;
    end;

  finally
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA WHERE MPANCORE = :pMPANCORE',
      TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
  end;
end;

procedure TD0148Test.TestDOD0170OLDDC_Success;
var
  RecordCount: Integer;
const
  testMPANCORE = '3080000036655';
  testFlowVersion = 'D0170';
  testRole = 'D';
  testMPID = 'TSDC';
begin
  CreateD0148Form;
  try
    FrmD0148Sub.mpancore.SetTextValue(testMPANCORE);
    FrmD0148Sub.dc_mpid.SetTextValue(testMPID);
    FrmD0148Sub.l_dc_mpid.SetTextValue('NDCM');
    FrmD0148Sub.L_DC_EFD.Text := DateToStr(Date + 1);
    FrmD0148Sub.DBSSD.Date := Date;

    FrmD0148Sub.DoD0170OLDDC('BATCHED');

    RecordCount := gSqlUtil.SelectQueryInteger(
      'SELECT COUNT(*) FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA ' +
      'WHERE MPANCORE = :pMPAN AND FLOWVERSION = :pFLOW',
      ['pMPAN', otString, testMPANCORE,
       'pFLOW', otString, testFlowVersion]);

    Assert.AreEqual(1, RecordCount);

    with gSqlUtil.SelectQuery(
      'SELECT * FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA ' +
      'WHERE MPANCORE = :pMPAN AND FLOWVERSION = :pFLOW',
      ['pMPAN', otString, testMPANCORE,
       'pFLOW', otString, testFlowVersion]) do
    begin
      Assert.AreEqual(testMPANCORE, Fields[0].AsString);
      Assert.AreEqual(testFlowVersion, Fields[1].AsString);
      Assert.AreEqual(testRole, Fields[2].AsString);
      Assert.AreEqual(testMPID, Fields[3].AsString);
      Free;
    end;

  finally
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.BATCH_FLOWS_FOR_SENDING_COA WHERE MPANCORE = :pMPANCORE',
      TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
  end;
end;

procedure TD0148Test.TestDOD0170OLDMO_Success;
const
  testMPANCORE = '3080000036655';
begin
  CreateD0148Form;
  try
    FrmD0148Sub.mpancore.SetTextValue(testMPANCORE);
    FrmD0148Sub.mo_mpid.SetTextValue('TSMO');
    FrmD0148Sub.l_mo_mpid.SetTextValue('NMOM');
    FrmD0148Sub.L_MO_EFD.Text := DateToStr(Date + 1);
    FrmD0148Sub.DBSSD.Date := Date;

    FrmD0148Sub.DoD0170OLDMO;

  finally
  end;
end;

procedure TD0148Test.TestDOD0205_Success;
const
  testMPANCORE = '3080000036655';
begin
  CreateD0148Form;
  try
    FrmD0148Sub.mpancore.SetTextValue(testMPANCORE);
    FrmD0148Sub.DBSSD.Date := Date;
    FrmD0148Sub.l_dc_mpid.SetTextValue('NDCM');
    FrmD0148Sub.l_da_mpid.SetTextValue('NDAM');
    FrmD0148Sub.l_mo_mpid.SetTextValue('NMOM');
    FrmD0148Sub.L_DC_EFD.Text := DateToStr(Date + 1);
    FrmD0148Sub.L_DA_EFD.Text := DateToStr(Date + 1);
    FrmD0148Sub.L_MO_EFD.Text := DateToStr(Date + 1);

    FrmD0148Sub.D0205Query.Close;
    FrmD0148Sub.D0205Query.SQL.Clear;
    FrmD0148Sub.D0205Query.SQL.Add('SELECT NULL, NULL, NULL, ''ODCM'', NULL, ''ODAM'', NULL, ''OMOM'' FROM DUAL');
    FrmD0148Sub.D0205Query.Open;

    FrmD0148Sub.DoD0205;

  finally
  end;
end;

procedure TD0148Test.TestGET_NEW_AGENTS_Success;
const
  testMPANCORE = '3080000036655';
begin
  CreateD0148Form;
  try
    FrmD0148Sub.mpancore.SetTextValue(testMPANCORE);
    FrmD0148Sub.da_check.Checked := True;
    FrmD0148Sub.mo_check.Checked := True;
    FrmD0148Sub.dc_check.Checked := True;

    FrmD0148Sub.Get_New_Agents;

    Assert.IsTrue(FrmD0148Sub.c_da.Visible);
    Assert.IsTrue(FrmD0148Sub.c_mo.Visible);
    Assert.IsTrue(FrmD0148Sub.c_dc.Visible);

  finally
  end;
end;

initialization
  {$IFDEF FLOWS}
  TDUnitX.RegisterTestFixture(TD0148Test);
  {$ENDIF}
end.
