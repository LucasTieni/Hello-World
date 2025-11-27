unit ExportTests;

interface

uses
  DUnitX.TestFramework, Export, System.SysUtils, System.IOUtils, UELSession, Oracle;

type
  [TestFixture]
  TExportTest = class
  private

  public
    [SetupFixture]
    procedure SetupFixture;
    [TearDownFixture]
    procedure TearDownFixture;
    [Test]
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;
    [Test]
    procedure D0131HHDCTest;
    [Test]
    procedure D0131HHDC_AddressChangeTest;
    [Test]
    procedure D0131MOTest;
    [Test]
    procedure D0131MO_AddressChangeTest;
    [Test]
    procedure D0131NHHDCTest;
    [Test]
    procedure D0131NHHDC_AddressChangeTest;
    [Test]
    procedure Do_E_D0205sTest;
  end;

implementation
uses
  DFMCommon, MockHelpers, Main, CopyProgress, busy, Common, DataModule, Processing;

procedure TExportTest.SetupFixture;
begin
  CreateSQLUtil('TEST01','TEST01','DEVCRM');
  FRM_Main := TFRM_Main.Create(Nil);
  Main_Data_Module := TMain_Data_Module.Create(nil);
  FRM_Common := TFRM_Common.Create(nil);
  FRM_PROCESSING := TFRM_PROCESSING.Create(nil);
  Main.H_OUTGOING := TPath.GetTempPath;
  Main.H_Mode := 'TEST' //JIC
end;

procedure TExportTest.TearDownFixture;
begin
  FreeAndNil(Main_Data_Module);
  FreeAndNil(FRM_Common);
  FreeAndNil(FRM_PROCESSING);
  FreeAndNil(FRM_Main);
  FreeSQLUtil;
end;

procedure TExportTest.Setup;
begin
  FRM_Export := TFRM_Export.Create(nil);
  FRM_File_Progress := TFRM_File_Progress.Create(nil);
  FRM_File_Progress.Hide;
  FRM_File_Progress.Height := 0;
  FRM_File_Progress.Width := 0;
  FRM_BUSY := TFRM_BUSY.Create(nil);
  FRM_BUSY.Height := 0;
  FRM_BUSY.Width := 0;
end;

procedure TExportTest.TearDown;
begin
  FreeAndNil(FRM_Export);
  FreeAndNil(FRM_File_Progress);
  FreeAndNil(FRM_BUSY);
end;

procedure TExportTest.D0131HHDCTest;
var
  Expected, Actual: Boolean;
  lOutputFileName: string;
const
  testMPANCORE = '111122223333';
begin
  Expected := True;

  try
    try
       gSqlUtil.InsertRecord('EDMGR.MPAN_STATUS',TRANSACTION_NO,
      ['MPANCORE',          otString, testMPANCORE,
       'CONFIRMED_DC_ID',   otString, '1111',
       'CONFIRMED_DC_ROLE', otString, 'C',
       'REGSTATUS',         otString, 'REGISTERED'
      ]); //D0131_dc is null
    except
    end;

    try
      gSqlUtil.InsertRecord('EDMGR.MPAS_CURRENT_ADDR',TRANSACTION_NO,
        ['MPANCORE',          otString, testMPANCORE,
         'METERING_POINT_ADDRESS1', otString, 'TEST METERING_POINT_ADDRESS1'
        ]
      );
    except
    end;

    FRM_Export.AgentsQuery.Filter := 'MPANCORE = '+QuotedStr(testMPANCORE);
    FRM_Export.AgentsQuery.Filtered := True;
    FRM_Export.D0131HHDC;
    lOutputFileName := Export.OutPutFilename;
    Actual := FileExists(lOutputFileName);
    Assert.AreEqual(Expected,Actual,', file not created');
    Actual := FindStringInFile(lOutputFileName,'D0131001');
    Assert.AreEqual(Expected,Actual,', flow version not found');
    Actual := FindStringInFile(lOutputFileName,'TEST METERING_POINT_ADDRESS1');
    Assert.AreEqual(Expected,Actual,', address not found');
  finally
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAS_CURRENT_ADDR WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES,
      ['pMPANCORE', otString, pdInput, testMPANCORE]
    );
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAN_STATUS WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES,
      ['pMPANCORE', otString, pdInput, testMPANCORE]
    );
    if (lOutputFileName <> '') and FileExists(lOutputFileName) then
      DeleteFile(lOutputFileName);
  end;
end;

procedure TExportTest.D0131HHDC_AddressChangeTest;
var
  Expected, Actual: Boolean;
  lOutputFileName: string;
const
  testMPANCORE = '111122223334';
  testAddress = 'TEST ADDR CHANGE ADDR1';
begin
  Expected := True;
  try
    try
      gSqlUtil.InsertRecord('EDMGR.MPAN_STATUS', TRANSACTION_NO,
        ['MPANCORE',          otString,   testMPANCORE,
         'CONFIRMED_DC_ID',   otString,   '1111',
         'CONFIRMED_DC_ROLE', otString,   'C',
         'REGSTATUS',         otString,   'REGISTERED',
         'D0131_dc',          otDate, StrToDate('01/01/1950')
        ]);
    except
    end;

    try
      gSqlUtil.InsertRecord('EDMGR.MPAS_CURRENT_ADDR', TRANSACTION_NO,
        ['MPANCORE',                otString, testMPANCORE,
         'METERING_POINT_ADDRESS1', otString, testAddress
        ]);
    except
    end;
    FRM_Export.AgentsQuery.Filter := 'MPANCORE = '+QuotedStr(testMPANCORE);
    FRM_Export.AgentsQuery.Filtered := True;
    FRM_Export.D0131HHDC_Addresschange;

    lOutputFileName := Export.OutPutFilename;
    Actual := FileExists(lOutputFileName);
    Assert.AreEqual(Expected, Actual, 'File for D0131 address change not created');

    Actual := FindStringInFile(lOutputFileName, 'D0131001');
    Assert.AreEqual(Expected, Actual, 'Flow version not found in D0131 address change file');

    Actual := FindStringInFile(lOutputFileName, testAddress);
    Assert.AreEqual(Expected, Actual, 'Address not found in D0131 address change file');
  finally
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAS_CURRENT_ADDR WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAN_STATUS WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    if (lOutputFileName <> '') and FileExists(lOutputFileName) then
      DeleteFile(lOutputFileName);
  end;
end;

procedure TExportTest.D0131MOTest;
var
  Expected, Actual: Boolean;
  lOutputFileName: string;
const
  testMPANCORE = '111122223335';
  testAddress = 'TEST D0131MO ADDRESS';
  testAgentID = '2222';
begin
  Expected := True;
  lOutputFileName := '';

  try
    try
      // Setup: Insert test data for a standard D0131 MOP export.
      gSqlUtil.InsertRecord('EDMGR.MPAN_STATUS', TRANSACTION_NO,
        ['MPANCORE',          otString,   testMPANCORE,
         'CONFIRMED_MO_ID',   otString,   testAgentID,
         'CONFIRMED_MO_ROLE', otString,   'M',
         'REGSTATUS',         otString,   'REGISTERED'
         // D0131_MO is null by default
        ]);
    except
    end;

    try
      gSqlUtil.InsertRecord('EDMGR.MPAS_CURRENT_ADDR', TRANSACTION_NO,
        ['MPANCORE',                otString, testMPANCORE,
         'METERING_POINT_ADDRESS1', otString, testAddress
        ]);
    except
    end;

    FRM_Export.AgentsQuery.Filter := 'MPANCORE = '+QuotedStr(testMPANCORE);
    FRM_Export.AgentsQuery.Filtered := True;
    // Execute: Call the procedure under test.
    FRM_Export.D0131MO;

    // Assert: Verify the output file and its contents.
    lOutputFileName := Export.OutPutFilename;
    Actual := FileExists(lOutputFileName);
    Assert.AreEqual(Expected, Actual, 'File for D0131 MOP not created');

    Actual := FindStringInFile(lOutputFileName, 'D0131001');
    Assert.AreEqual(Expected, Actual, 'Flow version not found in D0131 MOP file');

    Actual := FindStringInFile(lOutputFileName, testAddress);
    Assert.AreEqual(Expected, Actual, 'Address not found in D0131 MOP file');

  finally
    // Teardown: Clean up database and files.
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAS_CURRENT_ADDR WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAN_STATUS WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    if (lOutputFileName <> '') and FileExists(lOutputFileName) then
      DeleteFile(lOutputFileName);
  end;
end;

procedure TExportTest.D0131MO_AddressChangeTest;
var
  Expected, Actual: Boolean;
  lOutputFileName: string;
const
  testMPANCORE = '111122223336';
  testAddress = 'TEST D0131MO ADDR CHANGE';
  testAgentID = '2222';
begin
  Expected := True;
  lOutputFileName := '';

  try
    // Setup: Insert test data for a D0131 MOP address change export.
    try
      gSqlUtil.InsertRecord('EDMGR.MPAN_STATUS', TRANSACTION_NO,
        ['MPANCORE',          otString,   testMPANCORE,
         'CONFIRMED_MO_ID',   otString,   testAgentID,
         'CONFIRMED_MO_ROLE', otString,   'M',
         'REGSTATUS',         otString,   'REGISTERED',
         'D0131_MO',          otDate, StrToDate('01/01/1950')
        ]);
    except
    end;

    try
      gSqlUtil.InsertRecord('EDMGR.MPAS_CURRENT_ADDR', TRANSACTION_NO,
        ['MPANCORE',                otString, testMPANCORE,
         'METERING_POINT_ADDRESS1', otString, testAddress
        ]);
    except
    end;

    // Execute: Call the procedure under test.
    FRM_Export.D0131MO_addresschange;

    // Assert: Verify the output file and its contents.
    lOutputFileName := Export.OutPutFilename;
    Actual := FileExists(lOutputFileName);
    Assert.AreEqual(Expected, Actual, 'File for D0131 MOP address change not created');

    Actual := FindStringInFile(lOutputFileName, 'D0131001');
    Assert.AreEqual(Expected, Actual, 'Flow version not found in D0131 MOP address change file');

    Actual := FindStringInFile(lOutputFileName, testAddress);
    Assert.AreEqual(Expected, Actual, 'Address not found in D0131 MOP address change file');
  finally
    // Teardown: Clean up database and files.
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAS_CURRENT_ADDR WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAN_STATUS WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    if (lOutputFileName <> '') and FileExists(lOutputFileName) then
      DeleteFile(lOutputFileName);
  end;
end;

procedure TExportTest.D0131NHHDCTest;
var
  Expected, Actual: Boolean;
  lOutputFileName: string;
const
  testMPANCORE = '111122223337';
  testAddress = 'TEST D0131NHHDC ADDRESS';
  testAgentID = '3333';
begin
  Expected := True;
  lOutputFileName := '';

  try
    // Setup: Insert test data for a standard D0131 NHHDC export.
    try
      gSqlUtil.InsertRecord('EDMGR.MPAN_STATUS', TRANSACTION_NO,
        ['MPANCORE',          otString,   testMPANCORE,
         'CONFIRMED_DC_ID',   otString,   testAgentID,
         'CONFIRMED_DC_ROLE', otString,   'D', // NHH DC
         'REGSTATUS',         otString,   'REGISTERED'
         // D0131_dc is null by default
        ]);
    except
    end;
    try
      gSqlUtil.InsertRecord('EDMGR.MPAS_CURRENT_ADDR', TRANSACTION_NO,
        ['MPANCORE',                otString, testMPANCORE,
         'METERING_POINT_ADDRESS1', otString, testAddress
        ]);
    except
    end;

    // Execute: Call the procedure under test.
    FRM_Export.D0131NHHDC;

    // Assert: Verify the output file and its contents.
    lOutputFileName := Export.OutPutFilename;
    Actual := FileExists(lOutputFileName);
    Assert.AreEqual(Expected, Actual, 'File for D0131 NHHDC not created');

    Actual := FindStringInFile(lOutputFileName, 'D0131001');
    Assert.AreEqual(Expected, Actual, 'Flow version not found in D0131 NHHDC file');

    Actual := FindStringInFile(lOutputFileName, testAddress);
    Assert.AreEqual(Expected, Actual, 'Address not found in D0131 NHHDC file');

  finally
    // Teardown: Clean up database and files.
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAS_CURRENT_ADDR WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAN_STATUS WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    if (lOutputFileName <> '') and FileExists(lOutputFileName) then
      DeleteFile(lOutputFileName);
  end;
end;

procedure TExportTest.D0131NHHDC_AddressChangeTest;
var
  Expected, Actual: Boolean;
  lOutputFileName: string;
const
  testMPANCORE = '111122223338';
  testAddress = 'TEST D0131NHHDC ADDR CHANGE';
  testAgentID = '3333';
begin
  Expected := True;
  lOutputFileName := '';

  try
    // Setup: Insert test data for a D0131 NHHDC address change export.
    try
      gSqlUtil.InsertRecord('EDMGR.MPAN_STATUS', TRANSACTION_NO,
        ['MPANCORE',          otString,   testMPANCORE,
         'CONFIRMED_DC_ID',   otString,   testAgentID,
         'CONFIRMED_DC_ROLE', otString,   'D', // NHH DC
         'REGSTATUS',         otString,   'REGISTERED',
         'D0131_dc',          otDate, StrToDate('01/01/1950')
        ]);
    except
    end;

    try
      gSqlUtil.InsertRecord('EDMGR.MPAS_CURRENT_ADDR', TRANSACTION_NO,
        ['MPANCORE',                otString, testMPANCORE,
         'METERING_POINT_ADDRESS1', otString, testAddress
        ]);
    except
    end;

    // Execute: Call the procedure under test.
    FRM_Export.D0131NHHDC_Addresschange;

    // Assert: Verify the output file and its contents.
    lOutputFileName := Export.OutPutFilename;
    Actual := FileExists(lOutputFileName);
    Assert.AreEqual(Expected, Actual, 'File for D0131 NHHDC address change not created');

    Actual := FindStringInFile(lOutputFileName, 'D0131001');
    Assert.AreEqual(Expected, Actual, 'Flow version not found in D0131 NHHDC address change file');

    Actual := FindStringInFile(lOutputFileName, testAddress);
    Assert.AreEqual(Expected, Actual, 'Address not found in D0131 NHHDC address change file');

  finally
    // Teardown: Clean up database and files.
    gSqlUtil.Rollback;
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAS_CURRENT_ADDR WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    gSqlUtil.ExecSql('DELETE FROM EDMGR.MPAN_STATUS WHERE MPANCORE = :pMPANCORE',TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    if (lOutputFileName <> '') and FileExists(lOutputFileName) then
      DeleteFile(lOutputFileName);
  end;
end;

procedure TExportTest.Do_E_D0205sTest;
var
  Expected, Actual: Boolean;
  lOutputFileName: string;
  lRecordCount: Integer;
  lStatus: string;
  lDateGenerated: TDateTime;
const
  testMPANCORE = '111122223339';
  testAgentID = 'MPAS';
  testAgentRole = 'P';
  testFlowLine = '432|INSTNO|SP04|111122223339|20230101||A||845|||';
begin
  Expected := True;
  lOutputFileName := '';

  try
    // Setup: Insert a record into the batch table for D0205 export.
    // Using try..except to align with existing test patterns, though it can hide setup issues.
    try
      gSqlUtil.InsertRecord('EDMGR.BATCH_FLOWS_FOR_SENDING_ALL', TRANSACTION_NO,
        ['MPANCORE',      otString, testMPANCORE,
         'FLOWVERSION',   otString, 'D0205',
         'TO_ROLE',       otString, testAgentRole,
         'TO_MPID',       otString, testAgentID,
         'LINE_2',        otString, testFlowLine,
         'STATUS',        otString, 'R'
        ]);
    except
    end;

    // Execute: Call the procedure under test.
    FRM_Export.Do_E_D0205s;

    // Assert: Verify the output file and its contents.
    lOutputFileName := Export.OutPutFilename;
    Actual := FileExists(lOutputFileName);
    Assert.AreEqual(Expected, Actual, 'File for D0205 not created');

    Actual := FindStringInFile(lOutputFileName, 'D0205001');
    Assert.AreEqual(Expected, Actual, 'Flow version not found in D0205 file');

    Actual := FindStringInFile(lOutputFileName, '752|');
    Assert.AreEqual(Expected, Actual, 'Group 752 not found in D0205 file');

    Actual := FindStringInFile(lOutputFileName, '|SP04|' + testMPANCORE);
    Assert.AreEqual(Expected, Actual, 'Flow line content not found in D0205 file');

    // Verify the record was updated in the database using gSqlUtil
    lRecordCount := gSqlUtil.SelectQueryInteger(
      'SELECT COUNT(*) FROM EDMGR.BATCH_FLOWS_FOR_SENDING_ALL WHERE MPANCORE = :pMPAN AND FLOWVERSION = ''D0205''',
      ['pMPAN', otString, testMPANCORE]);
    Assert.AreEqual(1, lRecordCount, 'Test record not found after execution.');

    lStatus := gSqlUtil.SelectQueryString(
      'SELECT STATUS FROM EDMGR.BATCH_FLOWS_FOR_SENDING_ALL WHERE MPANCORE = :pMPAN AND FLOWVERSION = ''D0205''',
      ['pMPAN', otString, testMPANCORE]);
    Assert.AreEqual('S', lStatus, 'Status should be updated to S.');

    lDateGenerated := gSqlUtil.SelectQueryDateTime(
      'SELECT DATE_GENERATED FROM EDMGR.BATCH_FLOWS_FOR_SENDING_ALL WHERE MPANCORE = :pMPAN AND FLOWVERSION = ''D0205''',
      ['pMPAN', otString, testMPANCORE]);
    Assert.IsTrue(lDateGenerated > 0, 'DATE_GENERATED should be populated.');
  finally
    // Teardown: Clean up database and files.
    gSqlUtil.ExecSql('DELETE FROM EDMGR.BATCH_FLOWS_FOR_SENDING_ALL WHERE MPANCORE = :pMPANCORE AND FLOWVERSION = ''D0205''',
                     TRANSACTION_YES, ['pMPANCORE', otString, pdInput, testMPANCORE]);
    if (lOutputFileName <> '') and FileExists(lOutputFileName) then
      DeleteFile(lOutputFileName);
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TExportTest);

end.