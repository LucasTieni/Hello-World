object FRM_D0148: TFRM_D0148
  Left = 313
  Top = 102
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Change Of Agent(s) D0148/D0205/D0170/D0151 Generator'
  ClientHeight = 390
  ClientWidth = 539
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object LABEL_EXAMPLE: TLabel
    Left = 6
    Top = 278
    Width = 34
    Height = 11
    Caption = 'Ready.'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -9
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object EXAMPLES: TLabel
    Left = 472
    Top = 298
    Width = 50
    Height = 13
    Caption = 'EXAMPLES'
    Visible = False
  end
  object GroupBox8: TGroupBox
    Left = 0
    Top = 157
    Width = 539
    Height = 88
    Align = alTop
    Caption = 'OLD / EXISTING Agent Details'
    TabOrder = 0
    object GroupBox9: TGroupBox
      Left = 352
      Top = 15
      Width = 175
      Height = 71
      Caption = 'Existing DA Details'
      TabOrder = 0
      object Label8: TLabel
        Left = 8
        Top = 44
        Width = 46
        Height = 13
        Caption = 'EFD_DAA'
      end
      object Label9: TLabel
        Left = 8
        Top = 20
        Width = 56
        Height = 13
        Caption = 'Exisiting DA'
      end
      object da_mpid: TJvDBLookupCombo
        Left = 108
        Top = 16
        Width = 59
        Height = 21
        LookupField = 'FROM_NAME'
        LookupDisplay = 'FROM_NAME'
        LookupSource = da_srce
        TabOrder = 0
      end
      object DA_EFD: TDBDateEdit
        Left = 62
        Top = 40
        Width = 106
        Height = 21
        DataField = 'EFF_FROM_DATE'
        DataSource = da_srce
        NumGlyphs = 2
        TabOrder = 1
      end
      object da_role: TDBEdit
        Left = 82
        Top = 16
        Width = 23
        Height = 21
        DataField = 'FROMID'
        DataSource = da_srce
        TabOrder = 2
      end
    end
    object GroupBox10: TGroupBox
      Left = 177
      Top = 15
      Width = 175
      Height = 71
      Caption = 'Existing MO Details'
      TabOrder = 1
      object Label10: TLabel
        Left = 8
        Top = 44
        Width = 48
        Height = 13
        Caption = 'EFD_MOA'
      end
      object Label7: TLabel
        Left = 8
        Top = 20
        Width = 58
        Height = 13
        Caption = 'Exisiting MO'
      end
      object mo_mpid: TJvDBLookupCombo
        Left = 108
        Top = 16
        Width = 59
        Height = 21
        LookupField = 'FROM_NAME'
        LookupDisplay = 'FROM_NAME'
        LookupSource = mo_srce
        TabOrder = 0
      end
      object MO_EFD: TDBDateEdit
        Left = 62
        Top = 40
        Width = 106
        Height = 21
        DataField = 'EFF_FROM_DATE'
        DataSource = mo_srce
        NumGlyphs = 2
        TabOrder = 1
      end
      object mo_role: TDBEdit
        Left = 82
        Top = 16
        Width = 23
        Height = 21
        DataField = 'FROMID'
        DataSource = mo_srce
        TabOrder = 2
      end
    end
    object GroupBox11: TGroupBox
      Left = 2
      Top = 15
      Width = 175
      Height = 71
      Caption = 'Existing DC Details'
      TabOrder = 2
      object Label11: TLabel
        Left = 8
        Top = 20
        Width = 56
        Height = 13
        Caption = 'Exisiting DC'
      end
      object Label12: TLabel
        Left = 8
        Top = 44
        Width = 46
        Height = 13
        Caption = 'EFD_DCA'
      end
      object dc_mpid: TJvDBLookupCombo
        Left = 108
        Top = 16
        Width = 59
        Height = 21
        LookupField = 'FROM_NAME'
        LookupDisplay = 'FROM_NAME'
        LookupSource = dc_srce
        TabOrder = 0
      end
      object DC_EFD: TDBDateEdit
        Left = 62
        Top = 40
        Width = 106
        Height = 21
        DataField = 'EFF_FROM_DATE'
        DataSource = dc_srce
        NumGlyphs = 2
        TabOrder = 1
      end
      object dc_role: TDBEdit
        Left = 82
        Top = 16
        Width = 23
        Height = 21
        DataField = 'FROMID'
        DataSource = dc_srce
        TabOrder = 2
      end
    end
  end
  object GroupBox3: TGroupBox
    Left = 0
    Top = 69
    Width = 539
    Height = 88
    Align = alTop
    Caption = 'NEW / LATEST Agent Details'
    TabOrder = 1
    object C_DA: TGroupBox
      Left = 352
      Top = 15
      Width = 175
      Height = 71
      Caption = 'NEW DA Details'
      TabOrder = 0
      object Label3: TLabel
        Left = 8
        Top = 20
        Width = 68
        Height = 13
        Caption = 'NEW DA MPID'
      end
      object Label4: TLabel
        Left = 8
        Top = 44
        Width = 48
        Height = 13
        Caption = 'EFD_MOA'
      end
      object l_da_mpid: TJvDBLookupCombo
        Left = 108
        Top = 16
        Width = 59
        Height = 21
        LookupField = 'FROM_NAME'
        LookupDisplay = 'FROM_NAME'
        LookupSource = l_da_srce
        TabOrder = 0
      end
      object L_DA_EFD: TDBDateEdit
        Left = 62
        Top = 40
        Width = 105
        Height = 21
        DataField = 'EFF_FROM_DATE'
        DataSource = l_da_srce
        NumGlyphs = 2
        TabOrder = 1
      end
      object l_da_role: TDBEdit
        Left = 82
        Top = 16
        Width = 23
        Height = 21
        DataField = 'FROMID'
        DataSource = l_da_srce
        TabOrder = 2
      end
    end
    object C_MO: TGroupBox
      Left = 177
      Top = 15
      Width = 175
      Height = 71
      Caption = 'NEW MO Details'
      TabOrder = 1
      object Label1: TLabel
        Left = 8
        Top = 20
        Width = 70
        Height = 13
        Caption = 'NEW MO MPID'
      end
      object Label2: TLabel
        Left = 8
        Top = 44
        Width = 48
        Height = 13
        Caption = 'EFD_MOA'
      end
      object l_mo_mpid: TJvDBLookupCombo
        Left = 108
        Top = 16
        Width = 59
        Height = 21
        LookupField = 'FROM_NAME'
        LookupDisplay = 'FROM_NAME'
        LookupSource = l_mo_srce
        TabOrder = 0
      end
      object L_MO_EFD: TDBDateEdit
        Left = 62
        Top = 40
        Width = 105
        Height = 21
        DataField = 'EFF_FROM_DATE'
        DataSource = l_mo_srce
        NumGlyphs = 2
        TabOrder = 1
      end
      object l_mo_role: TDBEdit
        Left = 82
        Top = 16
        Width = 23
        Height = 21
        DataField = 'FROMID'
        DataSource = l_mo_srce
        TabOrder = 2
      end
    end
    object C_DC: TGroupBox
      Left = 2
      Top = 15
      Width = 175
      Height = 71
      Caption = 'NEW DC Details'
      TabOrder = 2
      object Label5: TLabel
        Left = 8
        Top = 20
        Width = 68
        Height = 13
        Caption = 'NEW DC MPID'
      end
      object Label6: TLabel
        Left = 8
        Top = 44
        Width = 46
        Height = 13
        Caption = 'EFD_DCA'
      end
      object l_dc_mpid: TJvDBLookupCombo
        Left = 108
        Top = 16
        Width = 59
        Height = 21
        LookupField = 'FROM_NAME'
        LookupDisplay = 'FROM_NAME'
        LookupSource = l_dc_srce
        TabOrder = 0
      end
      object L_DC_EFD: TDBDateEdit
        Left = 62
        Top = 40
        Width = 105
        Height = 21
        DataField = 'EFF_FROM_DATE'
        DataSource = l_dc_srce
        NumGlyphs = 2
        TabOrder = 1
      end
      object l_dc_role: TDBEdit
        Left = 82
        Top = 16
        Width = 23
        Height = 21
        DataField = 'FROMID'
        DataSource = l_dc_srce
        TabOrder = 2
      end
    end
  end
  object GroupBox1: TGroupBox
    Left = 0
    Top = 0
    Width = 539
    Height = 69
    Align = alTop
    Caption = 'MPAN Details'
    TabOrder = 2
    object DBText1: TDBText
      Left = 230
      Top = 22
      Width = 177
      Height = 17
      DataField = 'REGSTATUS'
      DataSource = DataSource1
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label13: TLabel
      Left = 6
      Top = 20
      Width = 32
      Height = 13
      Caption = 'MPAN:'
    end
    object Label14: TLabel
      Left = 6
      Top = 44
      Width = 23
      Height = 13
      Caption = 'SSD:'
    end
    object Label15: TLabel
      Left = 162
      Top = 22
      Width = 57
      Height = 13
      Caption = 'Reg Status:'
    end
    object DBSSD: TDBDateEdit
      Left = 42
      Top = 40
      Width = 105
      Height = 21
      DataField = 'SSD'
      DataSource = DataSource1
      NumGlyphs = 2
      TabOrder = 0
    end
    object dc_check: TCheckBox
      Left = 428
      Top = 10
      Width = 97
      Height = 17
      Caption = 'Change DC'
      Checked = True
      State = cbChecked
      TabOrder = 1
      OnClick = dc_checkClick
    end
    object mo_check: TCheckBox
      Left = 428
      Top = 28
      Width = 97
      Height = 17
      Caption = 'Change MO'
      Checked = True
      State = cbChecked
      TabOrder = 2
      OnClick = mo_checkClick
    end
    object da_check: TCheckBox
      Left = 428
      Top = 46
      Width = 97
      Height = 17
      Caption = 'Change DA'
      Checked = True
      State = cbChecked
      TabOrder = 3
      OnClick = da_checkClick
    end
    object MPANCORE: TJvDBLookupCombo
      Left = 42
      Top = 16
      Width = 105
      Height = 21
      LookupField = 'MPANCORE'
      LookupDisplay = 'MPANCORE'
      LookupSource = DataSource1
      TabOrder = 4
      OnChange = MPANCOREChange
    end
  end
  object CreateBTN: TBitBtn
    Left = 8
    Top = 294
    Width = 120
    Height = 25
    Caption = 'Process this MPAN'
    Enabled = False
    TabOrder = 3
    OnClick = CreateBTNClick
  end
  object D0151_OLDMO: TCheckBox
    Left = 216
    Top = 250
    Width = 111
    Height = 17
    Caption = 'D0151 to Old MO'
    TabOrder = 4
  end
  object D0148Check: TCheckBox
    Left = 280
    Top = 298
    Width = 71
    Height = 17
    Caption = 'D0148'#39's'
    TabOrder = 5
  end
  object D0151_OLDDC: TCheckBox
    Left = 34
    Top = 250
    Width = 111
    Height = 17
    Caption = 'D0151 to Old DC'
    TabOrder = 6
  end
  object D0151_OLDDA: TCheckBox
    Left = 394
    Top = 250
    Width = 111
    Height = 17
    Caption = 'D0151 to Old DA'
    TabOrder = 7
  end
  object BatchGroup: TGroupBox
    Left = 0
    Top = 324
    Width = 539
    Height = 66
    Align = alBottom
    Caption = 'Batch Progress'
    TabOrder = 8
    object RC: TLabel
      Left = 142
      Top = 22
      Width = 14
      Height = 13
      Caption = 'RC'
    end
    object RunBTN: TBitBtn
      Left = 10
      Top = 17
      Width = 120
      Height = 25
      Caption = 'Process ENTIRE list'
      TabOrder = 0
      OnClick = RunBTNClick
    end
    object ProgressBar1: TProgressBar
      Left = 2
      Top = 47
      Width = 535
      Height = 17
      Align = alBottom
      TabOrder = 1
    end
  end
  object D0205_Update: TCheckBox
    Left = 142
    Top = 298
    Width = 133
    Height = 17
    Caption = 'D205 Update MPAS'
    TabOrder = 9
  end
  object MO: TOracleDataSet
    SQL.Strings = (
      'SELECT'
      ' DISTINCT '
      ' D.MPANCORE, '
      ' F.FROM_NAME,'
      ' F.FROMID,'
      ' D.EFF_FROM_DATE,'
      ' f.file_date_time  '
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F,'
      '       (select D.MPANCORE,max(file_date_time) fdt'
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '     F.FROMID ='#39'M'#39' AND'
      '    d.mpancore=:MPAN and (from_name<>:fromname)'
      'group by D.MPANCORE) g'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '    F.FROMID ='#39'M'#39
      '    and d.mpancore=:MPAN and (from_name<>:fromname)'
      '    and d.mpancore=g.mpancore'
      '    and f.file_date_time=g.fdt'
      'ORDER BY'
      '    D.MPANCORE ASC'
      '    '
      '    '
      ''
      ''
      '    '
      '    '
      '    '
      ''
      ''
      ''
      '')
    Optimize = False
    Variables.Data = {
      04000000020000000A0000003A004D00500041004E00050000000E0000003131
      30303033393530343539320000000000120000003A00460052004F004D004E00
      41004D0045000500000005000000554D4F4C0000000000}
    QBEDefinition.QBEFieldDefs = {
      0500000004000000100000004D00500041004E0043004F005200450001000000
      00001A0000004500460046005F00460052004F004D005F004400410054004500
      0100000000000C000000460052004F004D004900440001000000000012000000
      460052004F004D005F004E0041004D004500010000000000}
    Cursor = crSQLWait
    Session = FRM_Login.MainSession
    Left = 250
    Top = 198
  end
  object mo_srce: TDataSource
    DataSet = MO
    Left = 280
    Top = 196
  end
  object DC: TOracleDataSet
    SQL.Strings = (
      'SELECT'
      ' DISTINCT '
      ' D.MPANCORE, '
      ' F.FROM_NAME,'
      ' F.FROMID,'
      ' D.EFF_FROM_DATE,'
      ' f.file_date_time  '
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F,'
      '       (select D.MPANCORE,max(file_date_time) fdt'
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '     F.FROMID in ('#39'C'#39','#39'D'#39') AND'
      '    d.mpancore=:MPAN and (from_name<>:fromname)'
      'group by D.MPANCORE) g'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '    F.FROMID in ('#39'C'#39','#39'D'#39')'
      '    and d.mpancore=:MPAN and (from_name<>:fromname)'
      '    and d.mpancore=g.mpancore'
      '    and f.file_date_time=g.fdt'
      'ORDER BY'
      '    D.MPANCORE ASC'
      '    '
      '    '
      ''
      ''
      '    '
      '    '
      '    '
      ''
      ''
      ''
      '')
    Optimize = False
    Variables.Data = {
      04000000020000000A0000003A004D00500041004E00050000000E0000003131
      30303033393530343539320000000000120000003A00460052004F004D004E00
      41004D0045000500000005000000554D4F4C0000000000}
    QBEDefinition.QBEFieldDefs = {
      0500000004000000100000004D00500041004E0043004F005200450001000000
      00001A0000004500460046005F00460052004F004D005F004400410054004500
      0100000000000C000000460052004F004D004900440001000000000012000000
      460052004F004D005F004E0041004D004500010000000000}
    Cursor = crSQLWait
    Session = FRM_Login.MainSession
    Left = 192
    Top = 194
  end
  object dc_srce: TDataSource
    DataSet = DC
    Left = 216
    Top = 196
  end
  object DA: TOracleDataSet
    SQL.Strings = (
      'SELECT'
      ' DISTINCT '
      ' D.MPANCORE, '
      ' F.FROM_NAME,'
      ' F.FROMID,'
      ' D.EFF_FROM_DATE,'
      ' f.file_date_time  '
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F,'
      '       (select D.MPANCORE,max(file_date_time) fdt'
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '     F.FROMID in ('#39'A'#39','#39'B'#39') AND'
      '    d.mpancore=:MPAN and (from_name<>:fromname)'
      'group by D.MPANCORE) g'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '    F.FROMID in ('#39'A'#39','#39'B'#39')'
      '    and d.mpancore=:MPAN and (from_name<>:fromname)'
      '    and d.mpancore=g.mpancore'
      '    and f.file_date_time=g.fdt'
      'ORDER BY'
      '    D.MPANCORE ASC'
      '    '
      '    '
      ''
      ''
      '    '
      '    '
      '    '
      ''
      ''
      ''
      '')
    Optimize = False
    Variables.Data = {
      04000000020000000A0000003A004D00500041004E00050000000E0000003131
      30303033393530343539320000000000120000003A00460052004F004D004E00
      41004D0045000500000005000000554D4F4C0000000000}
    QBEDefinition.QBEFieldDefs = {
      0500000004000000100000004D00500041004E0043004F005200450001000000
      00001A0000004500460046005F00460052004F004D005F004400410054004500
      0100000000000C000000460052004F004D004900440001000000000012000000
      460052004F004D005F004E0041004D004500010000000000}
    Cursor = crSQLWait
    Session = FRM_Login.MainSession
    Left = 310
    Top = 194
  end
  object da_srce: TDataSource
    DataSet = DA
    Left = 340
    Top = 194
  end
  object l_dc: TOracleDataSet
    SQL.Strings = (
      'SELECT'
      ' DISTINCT '
      ' D.MPANCORE, '
      ' F.FROM_NAME,'
      ' F.FROMID,'
      ' D.EFF_FROM_DATE  '
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '    F.FROMID in ('#39'C'#39','#39'D'#39')'
      'AND '
      ' (D.MPANCORE,D.EFF_FROM_DATE,F.FILE_DATE_TIME)'
      'in'
      '(select D.MPANCORE,max(D.EFF_FROM_DATE),max(file_date_time)'
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '     F.FROMID in ('#39'C'#39','#39'D'#39') AND'
      '    d.mpancore=:MPAN'
      'group by D.MPANCORE'
      ')'
      'ORDER BY'
      '    D.MPANCORE ASC'
      ''
      '')
    Optimize = False
    Variables.Data = {
      04000000010000000A0000003A004D00500041004E00050000000E0000003131
      30303033393530343539320000000000}
    QBEDefinition.QBEFieldDefs = {
      0500000004000000100000004D00500041004E0043004F005200450001000000
      00001A0000004500460046005F00460052004F004D005F004400410054004500
      0100000000000C000000460052004F004D004900440001000000000012000000
      460052004F004D005F004E0041004D004500010000000000}
    Cursor = crSQLWait
    Session = FRM_Login.MainSession
    Left = 238
    Top = 114
  end
  object l_dc_srce: TDataSource
    DataSet = l_dc
    Left = 262
    Top = 114
  end
  object l_mo: TOracleDataSet
    SQL.Strings = (
      'SELECT'
      ' DISTINCT '
      ' D.MPANCORE, '
      ' F.FROM_NAME,'
      ' F.FROMID,'
      ' D.EFF_FROM_DATE  '
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '    F.FROMID in ('#39'M'#39')'
      'AND '
      ' (D.MPANCORE,D.EFF_FROM_DATE,F.FILE_DATE_TIME)'
      'in'
      '(select D.MPANCORE,max(D.EFF_FROM_DATE),max(file_date_time)'
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '     F.FROMID in ('#39'M'#39') AND'
      '    d.mpancore=:MPAN'
      'group by D.MPANCORE'
      ')'
      'ORDER BY'
      '    D.MPANCORE ASC'
      ''
      '')
    Optimize = False
    Variables.Data = {
      04000000010000000A0000003A004D00500041004E00050000000E0000003131
      30303033393530343539320000000000}
    QBEDefinition.QBEFieldDefs = {
      0500000004000000100000004D00500041004E0043004F005200450001000000
      00001A0000004500460046005F00460052004F004D005F004400410054004500
      0100000000000C000000460052004F004D004900440001000000000012000000
      460052004F004D005F004E0041004D004500010000000000}
    Cursor = crSQLWait
    Session = FRM_Login.MainSession
    Left = 298
    Top = 114
  end
  object l_mo_srce: TDataSource
    DataSet = l_mo
    Left = 326
    Top = 114
  end
  object l_da: TOracleDataSet
    SQL.Strings = (
      'SELECT'
      ' DISTINCT '
      ' D.MPANCORE, '
      ' F.FROM_NAME,'
      ' F.FROMID,'
      ' D.EFF_FROM_DATE  '
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '    F.FROMID in ('#39'A'#39','#39'B'#39')'
      'AND '
      ' (D.MPANCORE,D.EFF_FROM_DATE,F.FILE_DATE_TIME)'
      'in'
      '(select D.MPANCORE,max(D.EFF_FROM_DATE),max(file_date_time)'
      'FROM'
      '    EDMGR.D0011 D,'
      '    EDMGR.FLOWHEADERS F'
      'WHERE'
      '    D.FILENAME = F.FILENAME AND'
      '    D.MPANCORE = F.MPANCORE AND'
      '     F.FROMID in ('#39'A'#39','#39'B'#39') AND'
      '    d.mpancore=:MPAN'
      'group by D.MPANCORE'
      ')'
      'ORDER BY'
      '    D.MPANCORE ASC'
      '')
    Optimize = False
    Variables.Data = {
      04000000010000000A0000003A004D00500041004E00050000000E0000003131
      30303033393530343539320000000000}
    QBEDefinition.QBEFieldDefs = {
      0500000004000000100000004D00500041004E0043004F005200450001000000
      00001A0000004500460046005F00460052004F004D005F004400410054004500
      0100000000000C000000460052004F004D004900440001000000000012000000
      460052004F004D005F004E0041004D004500010000000000}
    Cursor = crSQLWait
    Session = FRM_Login.MainSession
    Left = 354
    Top = 114
  end
  object l_da_srce: TDataSource
    DataSet = l_da
    Left = 386
    Top = 112
  end
  object MPANSTATUS: TOracleDataSet
    SQL.Strings = (
      'select mpancore,ssd,regstatus from edmgr.mpan_status')
    Optimize = False
    QBEDefinition.QBEFieldDefs = {
      0500000003000000100000004D00500041004E0043004F005200450001000000
      0000060000005300530044000100000000001200000052004500470053005400
      4100540055005300010000000000}
    AfterQuery = MPANSTATUSAfterQuery
    Session = FRM_Login.MainSession
    Left = 288
    Top = 12
  end
  object DataSource1: TDataSource
    DataSet = MPANSTATUS
    Left = 316
    Top = 12
  end
  object MainMenu1: TMainMenu
    Left = 422
    Top = 256
    object ools1: TMenuItem
      Caption = 'Tools'
      object CheckMPANSwhereD0155CoAwithinlast30days2: TMenuItem
        Caption = 'Check MPANS where D0155 CoA within last X days'
        OnClick = CheckMPANSwhereD0155CoAwithinlast30days2Click
      end
      object N1: TMenuItem
        Caption = '-'
      end
      object LoadQuery1: TMenuItem
        Caption = 'Load Query'
        OnClick = LoadQuery1Click
      end
      object N2: TMenuItem
        Caption = '-'
      end
      object SMRSUpdates1: TMenuItem
        Caption = 'SMRS Updates'
        object UpdateMPASDC1: TMenuItem
          Caption = 'Update MPAS DC'
          OnClick = UpdateMPASDC1Click
        end
      end
    end
    object Examples1: TMenuItem
      Caption = 'Examples'
      Enabled = False
      Visible = False
      object Show1: TMenuItem
        Caption = 'Show'
        OnClick = Show1Click
      end
    end
  end
  object D0151Query: TOracleDataSet
    Optimize = False
    Cursor = crSQLWait
    Session = FRM_Login.MainSession
    Left = 336
    Top = 262
  end
  object OpenDialog1: TOpenDialog
    InitialDir = 'c:\'
    Title = 'Select SQL Query'
    Left = 212
    Top = 42
  end
  object D0205Query: TOracleDataSet
    Optimize = False
    Cursor = crSQLWait
    Session = FRM_Login.MainSession
    Left = 376
    Top = 286
  end
end