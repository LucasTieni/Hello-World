unit uFriendlyCreditConfigTests;

interface

uses
  DUnitX.TestFramework, uFriendlyCreditConfig, Main, DataModule, LoginUnit,
  Vcl.ExtCtrls, UelSqlUtils, CrmCommon, System.SysUtils, DMImages, vcl.Forms;

type
  TFRM_Friendly_Credit_ConfigSub = class(TFRM_Friendly_Credit_Config);
  [TestFixture]
  TFCPeriodConfig = class
  private
    FrmFriendlyCreditConfigsSub : TFRM_Friendly_Credit_ConfigSub;
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;
    [Test]
    procedure DoOpenFriendlyCreditConfigsSuccess;
    [Test]
    procedure TestCellSelectionValidation;
    [Test]
    procedure TestNumericInputValidation;
    [Test]
    procedure TestRequiredDistributionEditing;
    [Test]
    procedure TestSaveChangesWithModifiedData;
  end;

implementation

procedure TFCPeriodConfig.Setup;
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

procedure TFCPeriodConfig.TearDown;
begin
  FreeAndNil(gSqlUtil);
  FreeAndNil(FRM_MAIN);
  FreeAndNil(DM_Images);
  FreeAndNil(FRM_Login);
end;

procedure TFCPeriodConfig.DoOpenFriendlyCreditConfigsSuccess;
var
  Expected, Actual : boolean;

begin
  Expected := true;

  FrmFriendlyCreditConfigsSub := TFRM_Friendly_Credit_ConfigSub.Create(nil);

  Actual := FrmFriendlyCreditConfigsSub.Refreshdata;
  Assert.AreEqual(Expected, Actual, 'Friendly Credit Data not properly loaded.');

  FreeAndNil(FrmFriendlyCreditConfigsSub);
end;

procedure TFCPeriodConfig.TestCellSelectionValidation;
var
  CanSelect: Boolean;
begin
  FrmFriendlyCreditConfigsSub := TFRM_Friendly_Credit_ConfigSub.Create(nil);
  try
    FrmFriendlyCreditConfigsSub.Refreshdata;
    
    // Test that column 0 (End Time) is not selectable for data rows
    CanSelect := True;
    FrmFriendlyCreditConfigsSub.GridFCPeriodSelectCell(FrmFriendlyCreditConfigsSub.GridFCPeriod, 0, 1, CanSelect);
    Assert.IsFalse(CanSelect, 'Column 0 should not be selectable for data rows');
    
    // Test that column 1 (Required Distribution %) is selectable for data rows
    CanSelect := False;
    FrmFriendlyCreditConfigsSub.GridFCPeriodSelectCell(FrmFriendlyCreditConfigsSub.GridFCPeriod, 1, 1, CanSelect);
    Assert.IsTrue(CanSelect, 'Column 1 should be selectable for data rows');
    
    // Test that header row is not selectable
    CanSelect := True;
    FrmFriendlyCreditConfigsSub.GridFCPeriodSelectCell(FrmFriendlyCreditConfigsSub.GridFCPeriod, 1, 0, CanSelect);
    Assert.IsFalse(CanSelect, 'Header row should not be selectable');
    
  finally
    FreeAndNil(FrmFriendlyCreditConfigsSub);
  end;
end;

procedure TFCPeriodConfig.TestNumericInputValidation;
var
  Key: Char;
begin
  FrmFriendlyCreditConfigsSub := TFRM_Friendly_Credit_ConfigSub.Create(nil);
  try
    FrmFriendlyCreditConfigsSub.Refreshdata;
    
    // Simulate being in the editable cell (column 1, row 1)
    FrmFriendlyCreditConfigsSub.GridFCPeriod.Col := 1;
    FrmFriendlyCreditConfigsSub.GridFCPeriod.Row := 1;
    
    // Test valid numeric input
    Key := '5';
    FrmFriendlyCreditConfigsSub.GridFCPeriodKeyPress(FrmFriendlyCreditConfigsSub.GridFCPeriod, Key);
    Assert.AreEqual('5', Key, 'Valid numeric character should be allowed');
    
    // Test decimal point
    Key := '.';
    FrmFriendlyCreditConfigsSub.GridFCPeriodKeyPress(FrmFriendlyCreditConfigsSub.GridFCPeriod, Key);
    Assert.AreEqual('.', Key, 'Decimal point should be allowed');
    
    // Test backspace
    Key := #8;
    FrmFriendlyCreditConfigsSub.GridFCPeriodKeyPress(FrmFriendlyCreditConfigsSub.GridFCPeriod, Key);
    Assert.AreEqual(#8, Key, 'Backspace should be allowed');
    
    // Test invalid character (letter)
    Key := 'a';
    FrmFriendlyCreditConfigsSub.GridFCPeriodKeyPress(FrmFriendlyCreditConfigsSub.GridFCPeriod, Key);
    Assert.AreEqual(#0, Key, 'Invalid character should be blocked');
    
  finally
    FreeAndNil(FrmFriendlyCreditConfigsSub);
  end;
end;

procedure TFCPeriodConfig.TestRequiredDistributionEditing;
var
  InitialValue, NewValue: string;
  InitialModifiedCount: Integer;
begin
  FrmFriendlyCreditConfigsSub := TFRM_Friendly_Credit_ConfigSub.Create(nil);
  try
    FrmFriendlyCreditConfigsSub.Refreshdata;
    
    // Store initial state
    if FrmFriendlyCreditConfigsSub.GridFCPeriod.RowCount > 1 then
    begin
      InitialValue := FrmFriendlyCreditConfigsSub.GridFCPeriod.Cells[1, 1];
      InitialModifiedCount := FrmFriendlyCreditConfigsSub.ModifiedRows.Count;
      
      // Simulate editing a cell
      NewValue := '25.5';
      FrmFriendlyCreditConfigsSub.GridFCPeriodSetEditText(
        FrmFriendlyCreditConfigsSub.GridFCPeriod, 1, 1, NewValue);
      
      // Check that the row was marked as modified
      Assert.AreEqual(InitialModifiedCount + 1, FrmFriendlyCreditConfigsSub.ModifiedRows.Count,
        'Modified rows count should increase after editing');
      
      // Check that the specific row is in the modified list
      Assert.IsTrue(FrmFriendlyCreditConfigsSub.ModifiedRows.IndexOf('1') >= 0,
        'Row 1 should be marked as modified');
      
      // Check that save button is enabled
      Assert.IsTrue(FrmFriendlyCreditConfigsSub.btnSaveChanges.Enabled,
        'Save button should be enabled after modifications');
    end;
    
  finally
    FreeAndNil(FrmFriendlyCreditConfigsSub);
  end;
end;

procedure TFCPeriodConfig.TestSaveChangesWithModifiedData;
begin
  FrmFriendlyCreditConfigsSub := TFRM_Friendly_Credit_ConfigSub.Create(nil);
  try
    FrmFriendlyCreditConfigsSub.Refreshdata;
    
    // Initially save button should be disabled
    Assert.IsFalse(FrmFriendlyCreditConfigsSub.btnSaveChanges.Enabled,
      'Save button should initially be disabled');
    
    // Simulate editing a cell to enable save button
    if FrmFriendlyCreditConfigsSub.GridFCPeriod.RowCount > 1 then
    begin
      FrmFriendlyCreditConfigsSub.GridFCPeriodSetEditText(
        FrmFriendlyCreditConfigsSub.GridFCPeriod, 1, 1, '30.0');
      
      // Save button should now be enabled
      Assert.IsTrue(FrmFriendlyCreditConfigsSub.btnSaveChanges.Enabled,
        'Save button should be enabled after modifications');
      
      // Test that UpdateSaveButtonState works correctly
      FrmFriendlyCreditConfigsSub.ModifiedRows.Clear;
      FrmFriendlyCreditConfigsSub.UpdateSaveButtonState;
      Assert.IsFalse(FrmFriendlyCreditConfigsSub.btnSaveChanges.Enabled,
        'Save button should be disabled when no modifications exist');
    end;
    
  finally
    FreeAndNil(FrmFriendlyCreditConfigsSub);
  end;
end;

initialization
 {$IFDEF METERING}
   TDUnitX.RegisterTestFixture(TFCPeriodConfig);
 {$ENDIF}

end.