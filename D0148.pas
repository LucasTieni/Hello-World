unit D0148;

interface

uses

  RXTooledit,
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, DBCtrls, Buttons, DB, OracleData,
  ComCtrls, Menus, RXDBCtrl, JvExControls, JvDBLookup;


type
  TFRM_D0148 = class(TForm)
    MO: TOracleDataSet;
    mo_srce: TDataSource;
    DC: TOracleDataSet;
    dc_srce: TDataSource;
    DA: TOracleDataSet;
    da_srce: TDataSource;
    l_dc: TOracleDataSet;
    l_dc_srce: TDataSource;
    l_mo: TOracleDataSet;
    l_mo_srce: TDataSource;
    l_da: TOracleDataSet;
    l_da_srce: TDataSource;
    LABEL_EXAMPLE: TLabel;
    GroupBox8: TGroupBox;
    GroupBox9: TGroupBox;
    Label8: TLabel;
    Label9: TLabel;
    da_mpid: TJvDBLookupCombo;
    DA_EFD: TDBDateEdit;
    da_role: TDBEdit;
    GroupBox10: TGroupBox;
    Label10: TLabel;
    Label7: TLabel;
    mo_mpid: TJvDBLookupCombo;
    MO_EFD: TDBDateEdit;
    mo_role: TDBEdit;
    GroupBox11: TGroupBox;
    Label11: TLabel;
    Label12: TLabel;
    dc_mpid: TJvDBLookupCombo;
    DC_EFD: TDBDateEdit;
    dc_role: TDBEdit;
    GroupBox3: TGroupBox;
    C_DA: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    l_da_mpid: TJvDBLookupCombo;
    L_DA_EFD: TDBDateEdit;
    l_da_role: TDBEdit;
    C_MO: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    l_mo_mpid: TJvDBLookupCombo;
    L_MO_EFD: TDBDateEdit;
    l_mo_role: TDBEdit;
    C_DC: TGroupBox;
    Label5: TLabel;
    Label6: TLabel;
    l_dc_mpid: TJvDBLookupCombo;
    L_DC_EFD: TDBDateEdit;
    l_dc_role: TDBEdit;
    GroupBox1: TGroupBox;
    DBSSD: TDBDateEdit;
    dc_check: TCheckBox;
    mo_check: TCheckBox;
    da_check: TCheckBox;
    DBText1: TDBText;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    CreateBTN: TBitBtn;
    EXAMPLES: TLabel;
    MPANSTATUS: TOracleDataSet;
    DataSource1: TDataSource;
    MainMenu1: TMainMenu;
    Examples1: TMenuItem;
    Show1: TMenuItem;
    D0151_OLDMO: TCheckBox;
    D0148Check: TCheckBox;
    D0151_OLDDC: TCheckBox;
    D0151_OLDDA: TCheckBox;
    D0151Query: TOracleDataSet;
    MPANCORE: TJvDBLookupCombo;
    ools1: TMenuItem;
    LoadQuery1: TMenuItem;
    OpenDialog1: TOpenDialog;
    BatchGroup: TGroupBox;
    RC: TLabel;
    RunBTN: TBitBtn;
    ProgressBar1: TProgressBar;
    D0205_Update: TCheckBox;
    D0205Query: TOracleDataSet;
    N1: TMenuItem;
    CheckMPANSwhereD0155CoAwithinlast30days2: TMenuItem;
    SMRSUpdates1: TMenuItem;
    N2: TMenuItem;
    UpdateMPASDC1: TMenuItem;
    procedure dc_checkClick(Sender: TObject);
    procedure mo_checkClick(Sender: TObject);
    procedure da_checkClick(Sender: TObject);
    procedure Get_New_Agents;
    procedure CreateBTNClick(Sender: TObject);
    procedure ShowThisMpan;
    procedure OutputExamples;
    procedure Show1Click(Sender: TObject);
    procedure COAMOP(batched:string);
    procedure COADC(batched:string);
    procedure COADA(batched:string);
    procedure CheckD0151s;
    procedure CheckD0205s;
    procedure MPANCOREChange(Sender: TObject);
    procedure LoadQuery1Click(Sender: TObject);
    procedure MPANSTATUSAfterQuery(Sender: TOracleDataSet);
    procedure RunBTNClick(Sender: TObject);
    procedure DoD0205;
    procedure DoD0170OLDMO;
    procedure DoD0170OLDDC(Batched:string);
    procedure CheckMPANSwhereD0155CoAwithinlast30days1Click(
      Sender: TObject);
    procedure ShowRecent(Days:integer);
    procedure CheckMPANSwhereD0155CoAwithinlast30days2Click(
      Sender: TObject);
    procedure ChangeMPASDC;
    procedure UpdateMPASDC1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FRM_D0148: TFRM_D0148;
  alreadychecked,candofiles:boolean;

implementation
 uses main,loginunit, D0148_Template, Common, D0205, DataModule;
{$R *.dfm}



procedure Tfrm_d0148.Get_New_Agents;
begin
 
 if da_check.Checked then c_da.Visible:=true
 else c_da.visible:=false;

 if mo_check.Checked then c_mo.Visible:=true
 else c_mo.visible:=false;

 if dc_check.Checked then c_dc.Visible:=true
 else c_dc.visible:=false;

 with l_dc do
 Begin
  close;
  setvariable('MPAN',mpancore.text);
  if dc_check.checked then
  Begin
   open;
   l_dc_mpid.KeyValue:=l_dc.fields[1].text;
  end;
 End;

 with l_Mo do
 Begin
  close;
  setvariable('MPAN',mpancore.text);
  if mo_check.checked then
  Begin
   open;
   l_mo_mpid.KeyValue:=l_mo.fields[1].text;
  end;
 End;

 with l_da do
 Begin
  close;
  setvariable('MPAN',mpancore.text);
  if da_check.checked then
  Begin
   open;
   l_da_mpid.KeyValue:=l_da.fields[1].text;
  end;
 End;

 // Current Details
 with dc do
 Begin
  close;
  setvariable('MPAN',mpancore.text);
  if c_dc.visible=true then setvariable('FROMNAME',l_dc_mpid.text)
  else setvariable('FROMNAME','XXXX');
  open;
  dc_mpid.KeyValue:=dc.fields[1].text;
 end;

 with mo do
 Begin
  close;
  setvariable('MPAN',mpancore.text);
  if c_mo.visible=true then setvariable('FROMNAME',l_mo_mpid.text)
  else setvariable('FROMNAME','XXXX');
  open;
  mo_mpid.KeyValue:=mo.fields[1].text;
 end;

 with da do
 Begin
  close;
  setvariable('MPAN',mpancore.text);
  if c_da.visible=true then setvariable('FROMNAME',l_da_mpid.text)
  else setvariable('FROMNAME','XXXX');
  open;
  da_mpid.KeyValue:=da.fields[1].text;
 end;

 if dc_mpid.text='' then
 begin
//  If Messagedlg('There does not appear to be a change in DC. Cancel Change of DC Request?',mtconfirmation,[MBYES,MBNO],0)=mryes then
//  Begin
   dc_check.checked:=false;
//   exit;
//  end;
 end;

 if MO_mpid.text='' then
 begin
//  If Messagedlg('There does not appear to be a change in MO. Cancel Change of MO Request?',mtconfirmation,[MBYES,MBNO],0)=mryes then
//  Begin
   mo_check.checked:=false;
//   exit;
//  end;
 end;

 if da_mpid.text='' then
 begin
 // If Messagedlg('There does not appear to be a change in DA. Cancel Change of DA Request?',mtconfirmation,[MBYES,MBNO],0)=mryes then
 // Begin                      }
   da_check.checked:=false;
//   exit;
//  end;
 end;

 examples.caption:=''; 
 ////////////////////////////////////////////////////////////////////
// Change of ONE agent ONLY SINGLE INSTANCE
////////////////////////////////////////////////////////////////////
label_example.caption:='';

if (c_dc.Visible=false) and (c_mo.visible=true) and (c_da.Visible=false) then
Begin
 // Change Of MOP ONLY =
 // A to new mop
 // B to incumbent DC
 label_example.caption:='Change of MO ONLY. Send D0148 to NEW MOP example A, and D0148 to Incumbent DC example B';
 examples.caption:='AB';
end;

if (c_dc.Visible=true) and (c_mo.visible=false) and (c_da.Visible=false) then
Begin
// Change of DC ONLY =
   // C to incumbent MO
   // D to New DC or E to new dc if change of DA
 label_example.caption:='Change of DC ONLY. Send D0148 to Incumbent MO example C and D0148 to New DC example D';
 examples.caption:='CD';
end;

if (c_dc.Visible=false) and (c_mo.visible=false) and (c_da.Visible=true) then
Begin
// Change of DA only =
   // F to incumbent DC
 label_example.caption:='Change of DA ONLY. Send D0148 to Incumbent DC Example F';
  examples.caption:='F';
end;

// Change of DA only =
   // F to incumbent DC

////////////////////////////////////////////////////////////////////
// Change of MO and DC  SINGLE INSTANCE
////////////////////////////////////////////////////////////////////
   // G to new mop
   // H to New DC or I to new dc if change of DA

////////////////////////////////////////////////////////////////////
// Change of MO and DA  SINGLE INSTANCE
////////////////////////////////////////////////////////////////////
   // A to new mop
   // J to Incumbent DC

////////////////////////////////////////////////////////////////////
// Change of DC and DA  SINGLE INSTANCE
////////////////////////////////////////////////////////////////////
   // C to incumbent MO
   // K to New DC or L to new dc if change of DA

if (c_dc.Visible=true) and (c_mo.visible=false) and (c_da.Visible=true) then
Begin
 label_example.caption:='Change of DC,&& DA. Send D0148 to Existing MO Example C, and D0148 to NEW DC Example K.';
 examples.caption:='CK';
end;


////////////////////////////////////////////////////////////////////
// Change of DC, MO, DA Single Instance
////////////////////////////////////////////////////////////////////
if (c_dc.Visible=true) and (c_mo.visible=true) and (c_da.Visible=true) then
Begin
 label_example.caption:='Change of DC,MO && DA. Send D0148 to NEW MO Example G, and D0148 to NEW DC Example M.';
 examples.caption:='GM';
end;

 if examples.caption='' then createbtn.Enabled:=false
 else createbtn.enabled:=true;
 if createbtn.enabled=true then d0148check.Checked:=true;

// Now Check if a D0151 sent to Old Agent
CheckD0151s;
// Now Check If MPAS Agents are Correct
CheckD0205s;
end;

procedure TFRM_D0148.dc_checkClick(Sender: TObject);
begin
 get_new_agents;
end;

procedure TFRM_D0148.mo_checkClick(Sender: TObject);
begin
 get_new_agents;
end;

procedure TFRM_D0148.da_checkClick(Sender: TObject);
begin
 get_new_agents;
end;

procedure TFRM_D0148.CreateBTNClick(Sender: TObject);
begin
 If Messagedlg('Process this MPAN?',mtconfirmation,[mbyes,mbno],0)<> mryes then exit;
 if frm_common.authoritycheck=false then exit;
 createbtn.enabled:=false;
 OutputExamples;
 createbtn.enabled:=true;
 Messagedlg('MPAN Processed.',mtinformation,[mbok],0);
end;

procedure TFRM_D0148.OutputExamples;
var
z:integer;
begin
 CANDOFILES:=TRUE;
 // If processing of Flow Failed, then probably already done so dont action again.
 if CANDOFILES=FALSE then EXIT;


 if (D0151_OLDDC.checked=true) and (C_DC.visible=true) then
 Begin
  COADC('BATCHED');
 end;
  // If processing of Flow Failed, then probably already done so dont action again.
 if CANDOFILES=FALSE then EXIT;

 if (D0151_OLDMO.checked=true) and (C_MO.visible=true) then
 Begin
  COAMOP('BATCHED');
 end;
  // If processing of Flow Failed, then probably already done so dont action again.
 if CANDOFILES=FALSE then EXIT;

 if (D0151_OLDDA.checked=true) and (C_DA.visible=true) then COADA('BATCHED');
  // If processing of Flow Failed, then probably already done so dont action again.
 if CANDOFILES=FALSE then EXIT;

 if (D0205_update.checked=true) then DoD0205;
  // If processing of Flow Failed, then probably already done so dont action again.
 if CANDOFILES=FALSE then EXIT;

 if d0148check.checked=false then exit;

 Application.CreateForm(TFrm_D0148_Template, Frm_D0148_Template);
 try
 for z:=1 to length(examples.caption) do
 Begin
  Frm_D0148_Template.mpancore.text:=mpancore.text;
  Frm_D0148_Template.SSD.text:=dbssd.text;
  Frm_D0148_Template.examplelookup.Text:=examples.caption[z];
  Frm_D0148_Template.showgroups(examples.caption[z]);
  frm_d0148_template.tag:=0;
  Frm_D0148_Template.CreateFlow('BATCHED');
  if frm_d0148_template.tag=1 then candofiles:=false;
 end;
 finally
  Frm_D0148_Template.release;
 end;

 if (dc_Check.checked=true) and (candofiles=true) then DoD0170OLDDC('BATCHED');
 if (mo_check.checked=true) and (candofiles=true) then DoD0170OLDMO;
end;

procedure TFRM_D0148.ShowThisMpan;
Begin
 with main_data_module.tempquery do
 Begin
  close;
  sql.clear;
  sql.add('select flow_version,file_date_time from edmgr.flowheaders where FLOW_VERSION like ''C%''');
  sql.add('and mpancore='''+mpancore.text+'''');
  sql.add('order by 2 desc');
  open;
 end;
 if main_data_module.tempquery.recordcount<>0 then
 Begin
  if main_data_module.tempquery.fields[0].text='COAMO' then
  begin
   da_check.checked:=false;
   dc_check.checked:=false;
   mo_check.checked:=true;
  end;
  if (main_data_module.tempquery.fields[0].text='CANDC') or (main_data_module.tempquery.fields[0].text='CAHDC') then
  begin
   da_check.checked:=false;
   dc_check.checked:=true;
   mo_check.checked:=false;
  end;
  if (main_data_module.tempquery.fields[0].text='CANDA') or (main_data_module.tempquery.fields[0].text='CAHDA') then
  begin
   da_check.checked:=true;
   dc_check.checked:=false;
   mo_check.checked:=false;
  end;
 end
 else
 Begin
  dc_check.checked:=true;
  mo_check.Checked:=true;
  da_check.checked:=true;
 end;
 get_new_agents;
end;

procedure TFRM_D0148.Show1Click(Sender: TObject);
begin
 Application.CreateForm(TFrm_D0148_Template, Frm_D0148_Template);
 try
  Frm_D0148_Template.mpancore.text:=mpancore.text;
  Frm_D0148_Template.SSD.text:=dbssd.text;
  Frm_D0148_Template.examplelookup.text:=' ';
  Frm_D0148_Template.showmodal;
 finally
  Frm_D0148_template.release;
 end;
end;

procedure TFRM_D0148.COAMOP(batched:string);
Var
MPAN,AgentID,AgentROLE,REASON,flowdate,fileid,filename,msg:String;
SSD,EFTSSD:Tdatetime;
f:textfile;
s:string;
linecount,mpancount:integer;
Begin
 screen.Cursor:=crhourglass;
 MPAN:=mpancore.text;
 EFTSSD:=strtodate(L_MO_EFD.text)-1;
 AGENTID:=MO_MPID.text;
 AGENTROLE:=MO_ROLE.text;
 SSD:=DBSSD.date;
 REASON:='CA';
 // Header Record
 Flowdate:=formatdatetime('YYYYMMDDHHNNSS',now);
 if batched<>'BATCHED' then
 Begin
  fileid:=FRM_Common.GetNextFileID;
  Filename:=frm_common.GETVALUE('FILE_ELEC_OUT')+'D0151_'+AGENTID+'_'+Fileid+'.usr';
   AssignFile(F, Filename);
  Rewrite(F);
  // Write header record
  S:='ZHV|';
  s:=s+FileID+'|';                              // File ID
  s:=s+'D0151001|';
  s:=s+X_MPIDROLE+'|';                          // FromRole
  s:=s+X_MPID+'|';                              // From ID
  s:=s+AGENTROLE+'|';                                    // Recipient Role
  if H_Mode='TEST' then s:=s+H_REC + '|'
  else s:=s+AGENTID+ '|';
  s:=s+flowdate+'|';                            // Date Now
  s:=s+H_APP+'|';                               // Application generation Flow
  s:=s+'|';
  s:=s+'|';
  s:=s+H_TESTFLAG+'|';                          // Live/Test Flag
  writeln(f,s);
 end;
 linecount:=0;
 mpancount:=1;
 msg:='';
 s:='297|'+MPAN+'|'+FormatDateTime('YYYYMMDD',SSD)+'|'+Reason+'||'; // Termination will always be LC for Loss of Contract to Supply (COS)
 if batched<>'BATCHED' then writeln(f,s)
 else msg:=msg+s+#13+#10;
 s:='298|'+FormatDateTime('YYYYMMDD',EFTSSD)+'|';// If Mop Termination
 if batched<>'BATCHED' then writeln(f,s)
 else msg:=msg+s;

 if batched<>'BATCHED' then
 Begin
  // Write Flow Footer
  s:='ZPT|';
  s:=s+FileID+'|';
  s:=s+inttostr(2)+'|';                 // Number Of Lines In File
  s:=s+'|';
  s:=s+inttostr(MPANCOUNT)+'|';                // Number of MPANS
  Flowdate:=formatdatetime('YYYYMMDDHHNNSS',now);
  s:=s+flowdate+'|';
  writeln(f,s);
  Closefile(f);;
 end
 else
 Begin
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('insert into edmgr.batch_flows_for_sending_coa values('''+mpan+''',''D0151'','''+AGENTROLE+''','''+AGENTID+''',null,'''+MSG+''',null,''R'')');
   try
    execute;
   except
    candofiles:=false;
   end;
  End;
  frm_login.mainsession.commit;
 end;

 screen.cursor:=crdefault;
end;

procedure TFRM_D0148.COADC(batched:string);
Var
MPAN,AgentID,AgentROLE,REASON,flowdate,fileid,filename,msg:String;
SSD,EFTSSD:Tdatetime;
f:textfile;
s:string;
linecount,mpancount:integer;
Begin
 screen.Cursor:=crhourglass;
 MPAN:=mpancore.text;
 EFTSSD:=strtodate(L_DC_EFD.text)-1;
 AGENTID:=DC_MPID.text;
 AGENTROLE:=DC_ROLE.text;
 SSD:=DBSSD.date;
 REASON:='CA';
 if batched<>'BATCHED' then
 Begin
  // Header Record
  Flowdate:=formatdatetime('YYYYMMDDHHNNSS',now);
  fileid:=FRM_Common.GetNextFileID;
  Filename:=frm_common.GETVALUE('FILE_ELEC_OUT')+'D0151_'+AGENTID+'_'+Fileid+'.usr';
  AssignFile(F, Filename);
  Rewrite(F);
  // Write header record
  S:='ZHV|';
  s:=s+FileID+'|';                              // File ID
  s:=s+'D0151001|';
  s:=s+X_MPIDROLE+'|';                          // FromRole
  s:=s+X_MPID+'|';                              // From ID
  s:=s+AGENTROLE+'|';                                    // Recipient Role
  if H_Mode='TEST' then s:=s+H_REC + '|'
  else s:=s+AGENTID+ '|';
  s:=s+flowdate+'|';                            // Date Now
  s:=s+H_APP+'|';                               // Application generation Flow
  s:=s+'|';
  s:=s+'|';
  s:=s+H_TESTFLAG+'|';                          // Live/Test Flag
  writeln(f,s);
 end;
 linecount:=0;
 mpancount:=1;
 msg:='';
 s:='297|'+MPAN+'|'+FormatDateTime('YYYYMMDD',SSD)+'|'+Reason+'||'; // Termination will always be LC for Loss of Contract to Supply (COS)
  if batched<>'BATCHED' then writeln(f,s)
 else msg:=msg+s+#13+#10;
 s:='299|'+FormatDateTime('YYYYMMDD',EFTSSD)+'|';// If Mop Termination
  if batched<>'BATCHED' then writeln(f,s)
 else msg:=msg+s;
 if batched<>'BATCHED' then
 Begin
  // Write Flow Footer
  s:='ZPT|';
  s:=s+FileID+'|';
  s:=s+inttostr(2)+'|';                 // Number Of Lines In File
  s:=s+'|';
  s:=s+inttostr(MPANCOUNT)+'|';                // Number of MPANS
  Flowdate:=formatdatetime('YYYYMMDDHHNNSS',now);
  s:=s+flowdate+'|';
  writeln(f,s);
  Closefile(f);
 end
 else
  Begin
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('insert into edmgr.batch_flows_for_sending_coa values('''+mpan+''',''D0151'','''+AGENTROLE+''','''+AGENTID+''',null,'''+MSG+''',null,''R'')');
   try
    execute;
   except
    candofiles:=false;
   end;  
  End;
  frm_login.mainsession.commit;
 end;
 screen.cursor:=crdefault;
end;

procedure TFRM_D0148.COADA(batched:string);
Var
MPAN,AgentID,AgentROLE,REASON,flowdate,fileid,filename,msg:String;
SSD,EFTSSD:Tdatetime;
f:textfile;
s:string;
linecount,mpancount:integer;
Begin
 screen.Cursor:=crhourglass;
 MPAN:=mpancore.text;
 EFTSSD:=strtodate(L_DA_EFD.text)-1;
 AGENTID:=DA_MPID.text;
 AGENTROLE:=DA_ROLE.text;
 SSD:=DBSSD.date;
 REASON:='CA';
 if batched<>'BATCHED' then
 Begin
  // Header Record
  Flowdate:=formatdatetime('YYYYMMDDHHNNSS',now);
  fileid:=FRM_Common.GetNextFileID;
  Filename:=frm_common.GETVALUE('FILE_ELEC_OUT')+'D0151_'+AGENTID+'_'+Fileid+'.usr';
  AssignFile(F, Filename);
  Rewrite(F);
  // Write header record
  S:='ZHV|';
  s:=s+FileID+'|';                              // File ID
  s:=s+'D0151001|';
  s:=s+X_MPIDROLE+'|';                          // FromRole
  s:=s+X_MPID+'|';                              // From ID
  s:=s+AGENTROLE+'|';                                    // Recipient Role
  if H_Mode='TEST' then s:=s+H_REC + '|'
  else s:=s+AGENTID+ '|';
  s:=s+flowdate+'|';                            // Date Now
  s:=s+H_APP+'|';                               // Application generation Flow
  s:=s+'|';
  s:=s+'|';
  s:=s+H_TESTFLAG+'|';                          // Live/Test Flag
  writeln(f,s);
 end;
 linecount:=0;
 mpancount:=1;
 msg:='';
 s:='297|'+MPAN+'|'+FormatDateTime('YYYYMMDD',SSD)+'|'+Reason+'||'; // Termination will always be LC for Loss of Contract to Supply (COS)
  if batched<>'BATCHED' then writeln(f,s)
 else msg:=msg+s+#13+#10;
 s:='300|'+FormatDateTime('YYYYMMDD',EFTSSD)+'|';// If Mop Termination
 if batched<>'BATCHED' then writeln(f,s)
 else msg:=msg+s;

 if batched<>'BATCHED' then
 Begin
  // Write Flow Footer
  s:='ZPT|';
  s:=s+FileID+'|';
  s:=s+inttostr(2)+'|';                 // Number Of Lines In File
  s:=s+'|';
  s:=s+inttostr(MPANCOUNT)+'|';                // Number of MPANS
  Flowdate:=formatdatetime('YYYYMMDDHHNNSS',now);
  s:=s+flowdate+'|';
  writeln(f,s);
  Closefile(f);
 end
 else
  Begin
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('insert into edmgr.batch_flows_for_sending_coa values('''+mpan+''',''D0151'','''+AGENTROLE+''','''+AGENTID+''',null,'''+MSG+''',null,''R'')');
   try
    execute;
   except
    candofiles:=false;
   end;  
  End;
  frm_login.mainsession.commit;
 end;
 screen.cursor:=crdefault;
end;

procedure TFRM_D0148.CheckD0151s;
begin
 d0151_olddc.Checked:=false;
 d0151_oldmo.Checked:=false;
 d0151_oldda.Checked:=false;
 // Check if DC Terminated
 if c_dc.visible=true then
 Begin
  with D0151query do
  Begin
   close;
   sql.clear;
   sql.add('Select a.* from edmgr.d0151 A,edmgr.flowheaders F where A.termination_reason=''CA''');
   sql.add('and A.mpancore='''+mpancore.text+'''');
   sql.add('and A.eftd_dca=to_date('''+datetostr(strtodate(L_DC_EFD.text)-1)+''',''DD/MM/YYYY'')');
   sql.add('and a.mpancore=f.mpancore');
   sql.add('and a.filename=f.filename');
   sql.add('and f.flow_version=''D0151''');
   sql.add('and f.toname='''+dc_mpid.text+'''');
   sql.add('and f.toid='''+dc_role.text+'''');
   open;
   if d0151query.recordcount=0 then D0151_olddc.Checked:=true
   else d0151_olddc.Checked:=false;
  end;
 End;

  // Check if MO Terminated
 if c_mo.visible=true then
 Begin
  with D0151query do
  Begin
   close;
   sql.clear;
   sql.add('Select a.* from edmgr.d0151 A,edmgr.flowheaders F where A.termination_reason=''CA''');
   sql.add('and A.mpancore='''+mpancore.text+'''');
   sql.add('and A.eftd_moa=to_date('''+datetostr(strtodate(L_MO_EFD.text)-1)+''',''DD/MM/YYYY'')');
   sql.add('and a.mpancore=f.mpancore');
   sql.add('and a.filename=f.filename');
   sql.add('and f.flow_version=''D0151''');
   sql.add('and f.toname='''+mo_mpid.text+'''');
   sql.add('and f.toid='''+mo_role.text+'''');
   open;
   if d0151query.recordcount=0 then D0151_oldmo.Checked:=true
   else d0151_oldmo.Checked:=false;
  end;
 End;

  // Check if da Terminated
 if c_da.visible=true then
 Begin
  with D0151query do
  Begin
   close;
   sql.clear;
   sql.add('Select a.* from edmgr.d0151 A,edmgr.flowheaders F where A.termination_reason=''CA''');
   sql.add('and A.mpancore='''+mpancore.text+'''');
   sql.add('and A.eftd_daa=to_date('''+datetostr(strtodate(L_DA_EFD.text)-1)+''',''DD/MM/YYYY'')');
   sql.add('and a.mpancore=f.mpancore');
   sql.add('and a.filename=f.filename');
   sql.add('and f.flow_version=''D0151''');
   sql.add('and f.toname='''+da_mpid.text+'''');
   sql.add('and f.toid='''+da_role.text+'''');
   open;
   if d0151query.recordcount=0 then D0151_oldda.Checked:=true
   else d0151_oldda.Checked:=false;
  end;
 End;
end;

procedure TFRM_D0148.CheckD0205s;
begin
 d0205_update.Checked:=false;
 // Check All MPAS Agents
 // Check if DC Terminated
 with D0205query do
 Begin
  close;
  sql.clear;
  sql.add('Select * from edmgr.agents_mpas');
  sql.add('where mpancore='''+mpancore.text+'''');
  sql.add('and ssd=to_date('''+DBSSD.text+''',''DD/MM/YYYY'')');
  open;
 end;
 if d0205query.recordcount=0 then exit;
 if (l_dc_mpid.text<>'') and (l_dc_mpid.text<>d0205query.fields[3].text) then d0205_update.checked:=true; // check DC
 if (l_da_mpid.text<>'') and (l_da_mpid.text<>d0205query.fields[5].text) then d0205_update.checked:=true; // check DA
 if (l_mo_mpid.text<>'') and (l_mo_mpid.text<>d0205query.fields[7].text) then d0205_update.checked:=true; // check MO
end;



procedure TFRM_D0148.MPANCOREChange(Sender: TObject);
begin
 alreadychecked:=false;
 ShowThisMPAN;
end;

procedure TFRM_D0148.LoadQuery1Click(Sender: TObject);
begin
 if opendialog1.execute=false then exit;
 mpanstatus.close;
 mpanstatus.sql.LoadFromFile(opendialog1.filename);
 mpanstatus.open;
 mpanstatus.first;
 mpancore.keyvalue:=mpanstatus.fields[0].text;
end;

procedure TFRM_D0148.MPANSTATUSAfterQuery(Sender: TOracleDataSet);
begin
 rc.caption:=inttostr(mpanstatus.recordcount);
 if mpanstatus.recordcount>1 then
 Begin
  frm_d0148.Height:=448;
  runbtn.visible:=true;
  batchgroup.Visible:=true;
 end
 else
 Begin
  runbtn.visible:=false;
  batchgroup.Visible:=false;
  frm_d0148.Height:=376;
 end;
end;

procedure TFRM_D0148.RunBTNClick(Sender: TObject);
begin
 If Messagedlg('Process Entire List?',mtconfirmation,[mbyes,mbno],0)<> mryes then exit;
 if frm_common.authoritycheck=false then exit;
 createbtn.Enabled:=false;
 runbtn.enabled:=false;
 screen.cursor:=crhourglass;
 progressbar1.position:=0;
 progressbar1.max:=mpanstatus.recordcount;
 while not mpanstatus.eof do
 Begin
  mpancore.KeyValue:=mpanstatus.fields[0].text;
 // d0148check.checked:=false;
  outputexamples;
  progressbar1.Position:=progressbar1.position+1;
  application.processmessages;
  mpanstatus.Next;
 End;
 progressbar1.position:=0;
 screen.cursor:=crdefault;
 createbtn.enabled:=true;
 runbtn.enabled:=true;
 Messagedlg('All Flows Processed.',mtinformation,[mbok],0);
end;

procedure TFRM_D0148.DoD0205;
Begin
 if not Assigned(FRM_D0205) then Application.CreateForm(TFRM_D0205, FRM_D0205);
 FRM_D0205.clearfields;
 FRM_D0205.MMPAN.text:=MPANCore.text;
 FRM_D0205.MSSD.text:=DBSSD.text;
 if (l_dc_mpid.text<>'') and (l_dc_mpid.text<>d0205query.fields[3].text) then
 Begin
  FRM_D0205.MDC.text:=l_dc_mpid.text; // DC
  FRM_D0205.MDC_T.text:='N'; // base on measurement class
  FRM_D0205.DC_DATE.text:=l_dc_efd.text; // DC
 end;
 if (l_da_mpid.text<>'') and (l_da_mpid.text<>d0205query.fields[5].text) then
 Begin
  FRM_D0205.MDA.text:=l_da_mpid.text; // DA
  FRM_D0205.MDA_T.text:='N'; //base on measurement class
  FRM_D0205.DA_DATE.text:=l_da_efd.text; // DA
 end;
 if (l_mo_mpid.text<>'') and (l_mo_mpid.text<>d0205query.fields[7].text) then
 Begin
  FRM_D0205.MMO.text:=l_mo_mpid.text; // MO
  FRM_D0205.MMO_T.text:='N'; //base on measurement class
  FRM_D0205.MO_DATE.text:=l_mo_efd.text; // MO
 end;
 FRM_D0205.createsingleD0205('BATCHED');
End;

procedure TFRM_D0148.DoD0170OLDMO;
Var
MPAN,AgentID,AgentROLE,REASON,flowdate,fileid,filename:String;
SSD,EFTSSD:Tdatetime;
f:textfile;
s:string;
linecount,mpancount:integer;
Begin
 screen.Cursor:=crhourglass;
 MPAN:=mpancore.text;
 EFTSSD:=strtodate(L_MO_EFD.text); // New MO ID
 M_MPID:=l_mo_mpid.text;            // Existing MO
 AGENTID:=mo_mpid.text;            // Existing MO
 AGENTROLE:='M';                   // MO Role
 SSD:=DBSSD.date;

 // Header Record
 Flowdate:=formatdatetime('YYYYMMDDHHNNSS',now);
 fileid:=FRM_Common.GetNextFileID;
 Filename:=frm_common.GETVALUE('FILE_ELEC_OUT')+'D0170_'+AGENTID+'_'+Fileid+'.usr';
 AssignFile(F, Filename);
 Rewrite(F);
 // Write header record
 S:='ZHV|';
 s:=s+FileID+'|';                              // File ID
 s:=s+'D0170001|';
 s:=s+X_MPIDROLE+'|';                          // FromRole
 s:=s+X_MPID+'|';                              // From ID
 s:=s+AGENTROLE+'|';                                    // Recipient Role
 if H_Mode='TEST' then s:=s+H_REC + '|'
 else s:=s+AGENTID+ '|';
 s:=s+flowdate+'|';                            // Date Now
 s:=s+H_APP+'|';                               // Application generation Flow
 s:=s+'|';
 s:=s+'|';
 s:=s+H_TESTFLAG+'|';                          // Live/Test Flag
 writeln(f,s);
 linecount:=0;
 mpancount:=1;
 s:='350|'+FormatDateTime('YYYYMMDD',EFTSSD)+'|06|PLEASE FORWARD MTDS TO NEW MO|'+M_mpid+'||';
 inc(OutputFlowFlowcount);
 inc(OutputFlowLineCount);
 writeln(f,s);
 s:='351|'+MPAN+'|';
 inc(OutputFlowLineCount);
 writeln(f,s);
  // Write Flow Footer
 s:='ZPT|';
 s:=s+FileID+'|';
 s:=s+inttostr(2)+'|';                 // Number Of Lines In File
 s:=s+'|';
 s:=s+inttostr(MPANCOUNT)+'|';                // Number of MPANS
 Flowdate:=formatdatetime('YYYYMMDDHHNNSS',now);
 s:=s+flowdate+'|';
 writeln(f,s);
 Closefile(f);;
 screen.cursor:=crdefault;
end;

procedure TFRM_D0148.DoD0170OLDDC(Batched:string);
Var
MPAN,AgentID,AgentROLE,REASON,flowdate,fileid,filename,line_1,line_2:String;
SSD,EFTSSD:Tdatetime;
f:textfile;
s:string;
linecount,mpancount:integer;
Begin
 screen.Cursor:=crhourglass;
 MPAN:=mpancore.text;
 EFTSSD:=strtodate(L_DC_EFD.text); // New DC ID
 M_MPID:=l_dc_mpid.text;            // Existing DC
 AGENTID:=dc_mpid.text;            // Existing DC
 AGENTROLE:='D';                   // DC Role
 SSD:=DBSSD.date;

 // Header Record
 if batched<>'BATCHED' then
 Begin
  Flowdate:=formatdatetime('YYYYMMDDHHNNSS',now);
  fileid:=FRM_Common.GetNextFileID;
  Filename:=frm_common.GETVALUE('FILE_ELEC_OUT')+'D0170_'+AGENTID+'_'+Fileid+'.usr';
  AssignFile(F, Filename);
  Rewrite(F);
  // Write header record
  S:='ZHV|';
  s:=s+FileID+'|';                              // File ID
  s:=s+'D0170001|';
  s:=s+X_MPIDROLE+'|';                          // FromRole
  s:=s+X_MPID+'|';                              // From ID
  s:=s+AGENTROLE+'|';                                    // Recipient Role
  if H_Mode='TEST' then s:=s+H_REC + '|'
  else s:=s+AGENTID+ '|';
  s:=s+flowdate+'|';                            // Date Now
  s:=s+H_APP+'|';                               // Application generation Flow
  s:=s+'|';
  s:=s+'|';
  s:=s+H_TESTFLAG+'|';                          // Live/Test Flag
  writeln(f,s);
 end;
 line_1:='';
 line_2:='';
 linecount:=0;
 mpancount:=1;
 s:='350|'+FormatDateTime('YYYYMMDD',EFTSSD)+'|07|PLEASE FORWARD READ HISTORY TO NEW DC||'+M_mpid+'|';
 inc(OutputFlowFlowcount);
 inc(OutputFlowLineCount);

 if batched<>'BATCHED' then writeln(f,s)
 else line_1:=s;

 s:='351|'+MPAN+'|';
 inc(OutputFlowLineCount);

 if batched<>'BATCHED' then writeln(f,s)
 else line_2:=s;

  // Write Flow Footer
 if batched<>'BATCHED' then
 Begin
  s:='ZPT|';
  s:=s+FileID+'|';
  s:=s+inttostr(2)+'|';                 // Number Of Lines In File
  s:=s+'|';
  s:=s+inttostr(MPANCOUNT)+'|';                // Number of MPANS
  Flowdate:=formatdatetime('YYYYMMDDHHNNSS',now);
  s:=s+flowdate+'|';
  writeln(f,s);
  Closefile(f);;
 end
 else
 begin
 with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('insert into edmgr.batch_flows_for_sending_coa values('''+mpan+''',''D0170'','''+AGENTROLE+''','''+AGENTID+''','''+LINE_1+''','''+LINE_2+''',null,''R'')');
   try
    execute;
   except
    candofiles:=false;
   end;
  End;
 End;
 frm_login.mainsession.commit;
 screen.cursor:=crdefault;
end;
procedure TFRM_D0148.CheckMPANSwhereD0155CoAwithinlast30days1Click(
  Sender: TObject);
begin
 Showrecent(30);
end;

procedure TFRM_D0148.showrecent(days:integer);
Begin
 with mpanstatus do
 Begin
  close;
  sql.clear;
  sql.add('select mpancore,ssd,regstatus from edmgr.mpan_status');
  sql.add('where (mpancore) in (select distinct mpancore from edmgr.flowheaders where flow_version in(''CAHDA'',''CAHDC'',''CANDA'',''CANDC'',''COAMO'')');
  sql.add('and file_Date_time>sysdate-'+inttostr(days));
  //sql.add('and (mpancore) not in (select distinct mpancore from edmgr.batch_flows_for_sending_coa');
  //sql.add('where (date_generated is null or date_generated>sysdate-30))');
  sql.add(')');
  sql.add('order by mpancore');
  open;
  first;
 end;
 if mpanstatus.recordcount=0 then exit;
 mpancore.keyvalue:=mpanstatus.fields[0].text;
end;

procedure TFRM_D0148.CheckMPANSwhereD0155CoAwithinlast30days2Click(
  Sender: TObject);
var
defaultdays:integer;
newdays:string;
clickedok:boolean;
begin
 Repeat
 defaultdays:=30;
 ClickedOK := InputQuery('How far back in days?', 'Please Enter Number', NewDays);
 if not ClickedOK then exit;
 If NewDays='' then MessageDlg('Please Enter A Number',mtERROR,[MBOK],0);
 until newDays<>'';
 defaultdays:=strtoint(newDays);
 Showrecent(defaultdays);
end;

procedure TFRM_D0148.ChangeMPASDC;
begin
 if not Assigned(FRM_D0205) then Application.CreateForm(TFRM_D0205, FRM_D0205);
 with main_data_module.generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select mpas_mpancore,mpas_ssd,latest_dc_id,latest_dc_efd');
  sql.add('from edmgr.snapshot_agent_view');
  sql.add('where MPAS_REG_STATUS=''REGISTERED'' AND MPAS_gsp=''_G'' AND MPAS_DC_ID <> LATEST_DC_ID');
  sql.add('order by 1');
  open;
 End;
 while not main_data_module.generalquery.Eof do
 Begin
  FRM_D0205.clearfields;
  FRM_D0205.MMPAN.text:=main_data_module.generalquery.fields[0].text;
  FRM_D0205.MSSD.text:=main_data_module.generalquery.fields[1].text;
  FRM_D0205.MDC.text:=main_data_module.generalquery.fields[2].text; // DC
  FRM_D0205.MDC_T.text:='N';
  FRM_D0205.DC_DATE.text:=main_data_module.generalquery.fields[3].text; // DC
  FRM_D0205.createsingleD0205('BATCHED');
  main_data_module.generalquery.next;
 end;
end;

procedure TFRM_D0148.UpdateMPASDC1Click(Sender: TObject);
begin
if messagedlg('Update Incorrect DC on MPAS?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;
ChangeMPASDC;
end;

end.
