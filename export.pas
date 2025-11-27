unit Export;

interface
uses
  Windows, Messages, sysutils,Classes, Graphics, Controls, Forms, Dialogs,
  FmxUtils, StdCtrls, Mask, DBCtrls, Db, ComCtrls, Grids, DBGrids,
  RXDBCtrl, rxlookup,registry, FileCtrl, ExtCtrls, Oracle,
  OracleData, Spin, Buttons, CheckLst,idSMTPBase,
  IDText,IDAttachmentFile,math,ioutils, UELSession;


type
  // FSP-2218 - Adding BRN
  // FSP-905  - Move the generation of gas NOM files into DFM
  // TGasShipper is a Record where gsAQI = 0, gsCNC = 1, etc.
  TGasShipper = (gsAQI, gsBRN, gsCNC, gsCNF, gsEMC, gsGEA, gsMAI, gsMAM, gsMID, gsMSI,
    gsNOM, gsONJ, gsONU, gsORJ, gsRFA, gsRRP, gsSFN, gsSPC, gsSPI, gsUBR, gsUDR, gsUMR, gsWAO);

  TFRM_Export = class(TForm)
    MPANSQuery: TOracleDataSet;
    QueryRefCode: TOracleDataSet;
    D0148Memo: TMemo;
    GeneralQuery: TOracleDataSet;
    MOsQuery: TOracleDataSet;
    AgentsQuery: TOracleDataSet;
    SequenceQuery: TOracleDataSet;
    MPANLIST: TOracleDataSet;
    D0297List: TOracleDataSet;
    D0297ListDEFAULT: TOracleDataSet;
    MiscQuery: TOracleDataSet;
    CustQuery: TOracleDataSet;
    DC1: TOracleDataSet;
    DC2: TOracleDataSet;
    DC3: TOracleDataSet;
    MOTerm: TOracleDataSet;
    MOQuery: TOracleDataSet;
    RequestedMO: TOracleDataSet;
    da1: TOracleDataSet;
    da2: TOracleDataSet;
    da3: TOracleDataSet;
    AddressQuery: TOracleDataSet;
    D0302Query: TOracleDataSet;
    DetailsQuery: TOracleDataSet;
    GroupBox3: TGroupBox;
    PageControl1: TPageControl;
    Tabsheet_Export_E: TTabSheet;
    PageControl2: TPageControl;
    TabSheet_EXPORT_EX: TTabSheet;
    XCheckList: TCheckListBox;
    TabSheet_EXPORT_EM: TTabSheet;
    MChecklist: TCheckListBox;
    Panel2: TPanel;
    GroupBox1: TGroupBox;
    LoadAllX: TCheckBox;
    ExportBtn: TBitBtn;
    CancelBtn: TBitBtn;
    StatusBar1: TStatusBar;
    LoadAllM: TCheckBox;
    MOPSITEADDRESS: TOracleDataSet;
    GetRemovedMeters: TOracleDataSet;
    D0150Info: TOracleDataSet;
    MetersQuery: TOracleDataSet;
    EACDetails: TOracleDataSet;
    TabSheet_Export_G: TTabSheet;
    PageControl3: TPageControl;
    TabSheet_EXPORT_GS: TTabSheet;
    GAS_SUPPLIER_LIST: TCheckListBox;
    LoadAllG: TCheckBox;
    PARMSFILE: TMemo;
    OUTPUTMEMO: TMemo;
    D0205Query: TOracleDataSet;
    COA: TOracleDataSet;
    TempQuery: TOracleDataSet;
    ParmsQuery: TOracleDataSet;
    MopParms: TOracleDataSet;
    OutputFile: TMemo;
    TAB_EXP_SHIPPER: TTabSheet;
    Gas_Shipper_list: TCheckListBox;
    CheckReading: TOracleDataSet;
//    procedure CreateD0055s;
    procedure COSD0148;
    procedure POP148MO(Example,MPANCORE,SSD:String);
    procedure POP148DC(Example,MPANCORE,SSD:String);
    procedure newD0148btnClick(Sender: TObject);
  //  Procedure GenerateAgentAppointments;
    Procedure CreateAppointments(Agent,AgentRole,mtc,nhh,COAEFD,GSP,SMART_COS:String);
   // procedure D0153155BtnClick(Sender: TObject);
    procedure TerminateAppointments;
    procedure CreateD0151Flow(Reason:string);
    procedure IssueD0005readRequests;
    function  ReadCommaField():string;
    procedure CreateD0297;
    procedure CreateD0297toDefault;
    procedure newD0131s;
    procedure D0131NHHDC;
    procedure D0131nhhdc_addresschange;
    procedure D0131HHDC;
    procedure D0131hhdc_addresschange;
    procedure D0131MO;
    procedure D0131MO_addresschange;
    procedure NHHDAD0153;
    procedure NHHDAD0153COA;
    procedure NHHDCD0155;
    procedure NHHDCD0155COA;
    procedure NHHMOD0155;
    procedure MOD0155COA;
    procedure HHDAD0153;
    procedure HHDCD0155;
    procedure HHMOD0155;
    procedure TerminateMOPLoss;
    procedure TerminateDCLoss;
    Procedure TerminateDALoss;
    procedure TerminateMOlossOBJ;
    procedure TerminateMOPSMI;
    procedure TerminateMOPWrongSSD;
    procedure TerminateDClossOBJ;
    procedure Terminate_ACCU_DC_WRONG_APP_DATE;
    procedure TerminateDAlossOBJ;
    procedure TerminateDClossCOA;
    procedure TerminateDAlossCOA;
    procedure CreateFlowHeader(FlowVersion,Rec_MPID,Rec_Role:string; bResetProgressBar: boolean = True);
    procedure CreateFlowHeaderMOP(FlowVersion,Rec_MPID,Rec_Role:string);
    procedure CreateFlowHeaderMOPPARMS(FlowVersion,role,mpid,MON:string);
    procedure CreateFlowFooter(bResetProgressBar: boolean = True);
    procedure CreateFlowFooterPARMS;
    procedure CreateD0010(filter:string);
    procedure CreateD0010fromD0188;
    procedure CreateD0010fromDials;
    procedure CreateD0071;
    Procedure CreateD0005toMOP(AGENTID,MPAN,Reason,ReasonCode:string; Adate:TdateTime);
    Procedure IdentifyD0302(AgentRole,MPANS,GROUPS:String);
    procedure TerminateMOPdisc;
    procedure TerminateDCdisc;
    Procedure TerminateDAdisc;
    procedure CreateSingleD0183(rcode:string);
    procedure DoDailyExports;
//    procedure CreateSingleD0055(AREGILocation,FileOrDB: string);
    procedure QueryAgents(MPAN:String);
    procedure CreateOutstandingD0302s;
    procedure CreateOutstandingD0225s;
    procedure Deletebatches(Dflow:string);
    procedure DeleteBlankSpecialNeeds;
    procedure CreateMOPD0011s;
    procedure CreateMOPD0261s;
    procedure CreateMOPD0170s;
    procedure CreateDISTD0170s;
    procedure SelectAllXFlows;
    Procedure SelectAllMFlows;
    Procedure SelectAllGFlows;
    procedure ExportBtnClick(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
    procedure ExportD0149_X;
    procedure ExportD0149_D;
    procedure ExportD0149_R;
    procedure ExportD0149_M;
    Procedure ExportThisD0149Query;
    procedure ExportD0150_X;
    procedure ExportD0150_D;
    procedure ExportD0150_R;
    procedure ExportD0150_M;
    Procedure ExportThisD0150Query;
    procedure ExportD0313_X;
    procedure ExportD0313_D;
    procedure ExportD0313_R;
    procedure ExportD0313_M;
    Procedure ExportThisD0313Query;
    Procedure ExportD0312_P;
    procedure ExportD0303_Appoint;
    procedure ExportD0303_DeAppoint;
    Procedure ExportThisD0303;
    procedure CreateMOPD0135;
    procedure CreateMOPD0139;
    Procedure ExportThisD0135;
    Procedure ExportThisD0139;
    procedure CreateMOPD0010;
    Procedure ExportThisD0010;
    procedure CreateMOPD0002;
    Procedure ExportThisD0002;
    procedure ExportEmails;
    procedure CreateMOPD0224;
    procedure CreateD0311s;
    procedure Do_E_SupplierFiles;
    procedure Do_E_MOPFiles;
    procedure Do_G_Supplier_Files;
    procedure Do_G_Shipper_Files;
    procedure DoMopParms(Year,Month:string);
    procedure DoMopParms_SP05(Year,Month:string);
    procedure DoMopParms_SP06(Year,Month:string);
    procedure DoMopParms_NM03(Year,Month:string);
    procedure DoMopParms_NM04(Year,Month:string);
    procedure DoMopParms_NM12(Year,Month:string);
    procedure DoMopParms_SP11(Year,Month:string);
    procedure DoMopParms_SP14(Year,Month:string);
    procedure DoMopParms_SP15(Year,Month:string);
    Procedure DoSupplierParms_P0135(YEAR,MONTH:string);
    Procedure DoSupplierParms_P0142(YEAR,MONTH:string);
    procedure DOPARMSCHECKSUM;
    procedure ExportD0304_DIST;
    procedure ExportD0304_SUPP;
    procedure ExportD0304_MPAS;
    procedure DoD0142s;
    Procedure ExportThisD0304;
    procedure ExportD0386Unrelated;
    Procedure Do_E_D0205s;
//    procedure CreatebATCHD0055;
    procedure CreateD0151COA;
    procedure CreateD0148COA;
    procedure CreateD0205COA;
    procedure CreateD0170COA;
    procedure CreateD0190;
    procedure UpdateStatusBar(MSg,Banner:string);
    function GetAppDate(COADATE,SSD:string):string;
    function EstimateFinalRead08D(MPRN, RegisterID: string;  EndDate: TDateTime): string;
    function GetRegisteredReading(MPRN, RegisterID: string;  EndDate: TDateTime): String;
    procedure FormShow(Sender: TObject);
   private
     procedure ExportD0386;
    { Private declarations }
    function IndexOf(s: string): Integer;
  public
    { Public declarations }
  end;

var
 FRM_Export: TFRM_Export;
 F,ErrorF:Textfile;
 flowdate,MPANCORE,EFSD,Agent,EFT:string;
 Fileid :string;
 Lines,mpancount:Integer;
 prevdate:Tdate;
 Filename:String;
 Finifile:TRegIniFile;
 agenttype,oldagentid,oldagentrole:string;
 CHeader,Cfooter:boolean;
 D0131Mpancount,D0131LineCount,
 D0225Mpancount,D0225LineCount:integer;
 loop:integer;
 Outputflow:Textfile;
 OutPutFlowIdentifier:String;
 OutPutFilename:String;
 OutPutFlowLineCount, OutputFlowFlowCount:Integer;
 toparty:string;
 FileDirectory,FileIdDesc: string;
implementation

uses
  Loginunit,main, TransferUnit, CopyProgress, Import, ExportGas,
  DataModule, Common, email_component, busy, StrUtils, Utility, DfmCommon, Logger,
  UELUtils, UELSqlUtils, System.Variants, DFMSession;
{$R *.DFM}

type
  // FSP-3079-->)
  // Electricity checklist item identifiers
  TXCheckListItems = (
    { 0}xchliCOAFChangeOfAgentBatchedFile,
    { 1}xchliD0005ReadingRequests,
    { 2}xchliD0010_D0071ReadingsToDCAndSupplier,
    { 3}xchliD0052AffirmationOfMeteringSystemDetails,
    { 4}xchliD0064ObjectionRequests,
    { 5}xchliD0131AddressChangeHHDC,
    { 6}xchliD0131AddressChangeMOP,
    { 7}xchliD0131AddressChangeNHHDC,
    { 8}xchliD0132SupplyDisconnectionDeatails,
    { 9}xchliD0142RequestToChangeInstallMetering,
    {10}xchliD0148AgentDetails,
    {11}xchliD0151DisconnectionDA,
    {12}xchliD0151DisconnectionDC,
    {13}xchliD0151DisconnectionMO,
    {14}xchliD0151LossDA,
    {15}xchliD0151LossDC,
    {16}xchliD0151LossMO,
    {17}xchliD0151ObjectionDA,
    {18}xchliD0151ObjectionDC,
    {19}xchliD0151ObjectionMO,
    {20}xchliD0153AppointmentDA,
    {21}xchliD0155AppointmentDC,
    {22}xchliD0155AppointmentMO,
    {23}xchliD0190PrepayKeyRequests,
    {24}xchliD0205MPASUpdates,
    {25}xchliD0225SpecialNeeds,
    {26}xchliD0301ErroneousTransfers,
    {27}xchliD0302CustomerDetailsDC,
    {28}xchliD0302CustomerDetailsDist,
    {29}xchliD0302CustomerDetailsMO,
    {30}xchliD0306RequestForDebtInformation,
    {31}xchliD0307DebtInformation,
    {32}xchliD0308ConfirmationOfCustomerDebtTransfer,
    {33}xchliD0309ConfirmationOfDebtAssigned,
    {34}xchliD0311NosieFlow,
    {35}xchliD0358RegistrationWithdrawlRequest,
    {36}xchliD0381MeteringPointAddressUpdates,
    {37}xchliD0386ManageMeteringPointRelationships,
    {38}xchliD2026DUOSRemittanceAdvice,
    {39}xchliPARMSSupplierPARMSReports);
  // -->FSP-3079

//procedure TFRM_Export.CreateD0055s;
//var AREGILocation, AREGIFile, x,lno : string;
//    SearchRec : TSearchRec;
//    ch:char;
//begin
//
//
// cursor:=crhourglass;                               // Show busy cursor
// Finifile:=TReginiFile.Create(apptitle);
// AregiLocation:=FIniFile.ReadString('File Locations','AREGI SRCE','C:\');
// Aregilocation:=aregilocation+'Validated\';
// FiniFile.free;
// x:=AREGILocation+'*.txt';
// if FindFirst(x, faAnyFile, SearchRec) = 0 then     // Go through all the files in the directory and process
// begin
//  repeat
//   AREGIFile:=AREGILocation+SearchRec.Name;         // Full name of the file found
//   assignfile(F,AREGIFile);                         // Open the Aregi file
//   reset(F);
//   s:='';
//   while not eof(f) do
//   Begin
//    read(f,ch);
//    s:=s+ch;
//   End;
//                                     // Read contents of Regi File into string S
//   closefile(F);
//
//   AREGIRECORD:=S;                                 // Close File
//   FRM_Transfer_Aregi.readregivalues;              // Breakdown string s into fields
//   //CreateSingleD0055(AREGILOCATION,'FILE');      // 22/03/2018 - Disable Single File Generation
//   CreateSingleD0055(AREGILOCATION,'DB');          // 22/04/2018  - Write Files to Database instead
//   until (FindNext(SearchRec) <> 0);
// end;
// frm_login.MainSession.Commit;
//
// FRM_EXPORT.CREATEBATCHD0055;                      // Output DB files into one Batched File
//
// cursor := crdefault;
//end;

//procedure TFRM_Export.CreateSingleD0055(AREGILocation,FileOrDB : string);
//////////////////////////////////////////////////////////////////////////////////
//// Create A D0055 based on the field values from Regi File                    //
//////////////////////////////////////////////////////////////////////////////////
//var
// S,ToParty,PC,SSC,SMRS,ssd:String;
//begin
//  TOPARTY:=FRM_common.getsmrs(f14);                  // Identify Correct MPAS based on MPANCore
//  FRM_File_Progress.progressbar.position:=0;
//
//  frm_main.FileMEMO.Lines.Clear;
//  if fileorDB='FILE' then
//  begin
//   SMRS:=frm_common.smrssequenceno(TOPARTY);
//   CreateFlowHeader('D0055001',ToParty,'P');
//   s:='733|'+SMRS+'|';
//   frm_main.WriteLinetoFile(s);      // Now write Group 733 File Sequence Number
//  end;
//
//
//  S:='';
//  S:=S+'126'+'|';                                    // Group= 126
//  S:=S+frm_common.nextinstructionnumber+'|';                   // Write Instruction number
//
//  if F19='T' then S :=S+'SP01'+'|'                   // SP01 if new connection
//  else S:=S+'SP04'+'|';                              // SP04 for Change Of Supplier
//  S:=S+F14+'|';                                      // MPAN CORE
//
//  // If SSD is less than 2 days time adjust to earliest SSD
//  SSD:=copy(f15,7,2)+'/'+copy(f15,5,2)+'/'+copy(f15,1,4);
//  if strtodate(ssd)<now+2 then f15:=Formatdatetime('YYYYMMDD',now+2);
//
//  S:=S+F15+'|';                                      // Effective from date (SSD)
//
//  // Change implemented by lee kitchen 17/11/2003
//  // Dont include energisation status, always inherit from MPAS.
//
//  // Definately Don't Include in a New Connection
//  // ** Note a D0205 will need sending to update Energisation status, this can only be done on
//  // ** receipt of D0150 from MO.
//  // ** Change implemented as a result of TA2000 Testing 17/01/02
//
//  ///////////////////////////////////////////////////////////////////////////////////////////////////
//  // 17/03/2004 change made by Lee, as a result of MRASCo testing and St.Clements Validation rules //
//  // For New Connection, energisation status MUST be populated if a DA appointment exists.         //
//  // In this case Populate with a 'D' and update to 'E' later, upon receipt of D0150               //
//  ///////////////////////////////////////////////////////////////////////////////////////////////////
//
//  if F19='T' then S:=S+'D'+'|'                       // De-Energise NEW Connections
//  else S:=S+'|';                                     // Don't Include Energisation Status if COS
//  if (F49 ='A') or (F49='B') then AgentType:='N'     // If Measurment Class='A' or 'B' then NHH
//  else AgentType:='H';                               // Else Half Hourly
//
//  //***************************************************************************//
//  // Completion of MTC,PC and SSC combinations
//  // If populating any of the above then ALL 3 must be populated
//  // HH should be *845* null and null
//  // Instances when these would be populated are on New connection or
//  // CMC coincident with COS.
//  // When CMS/COS the items do not need to be populated, however a 205 would be required later.
//  // By default items will be populated on CMC.
//
//  // CMC will be identified by the Install/remove meter flag set to Y (See field F56)
//
//  if F19 ='T' then                                   // IF NEW CONNECTION
//  Begin
//   S:=S+F49+'|';                                     // Measurement Class ID
//   if length(f22)=2 then pc:=f22[2]                  // Profile Class
//   else pc:=f22;
//   f22:=pc;
//   if length(f48)=3 then ssc:='0'+f48                // SSC
//   else ssc:=f48;
//   f48:=ssc;
//   If agenttype ='N' then
//   S:=S+F21+'|'+F22+'|'+F48+'|'                      // NHH  New Connection PC,MTC,SSC to be populated
//   else S:=S+F21+'|||';                              // HH New Connection MTC populated, PC+SSC are NULL
//  End
//  else
//
//  if F19 <>'T' then                                  // IF CHANGE OF SUPPLIER
//  Begin
//    if F56<> 'Y' then S:=S+'||||'                    // If not CMC  pc null, MTC=null, SSC null
//    else                                             // CMC coincident with COS
//    begin
//     If agenttype ='N' then
//     S:=S+F49+'|'+F21+'|'+F22+'|'+F48+'|'            // if NHH   PC,MTC,SSC to be populated
//     else S:=S+F49+'|'+F21+'|||';                    // else HH  PC null, MTC=845, SSC null
//    end;
//  End;
//
//  S:=S+F45+'|'+AgentType+'|';                        // Data Aggregator ID +Type
//  S:=S+F46+'|'+AgentType+'|';                        // Data Collector ID  +Type
//  S:=S+F47+'|'+AgentType+'|';                        // Meter Operator ID  +Type
//  S:=S+F17+'|';                                      // Changes of Tenancy Indicator
//  frm_main.WriteLinetoFile( S);                            // Write this record to File
//  outputflowlinecount:=2;
//  outputflowflowcount:=1;
//  if fileorDB='FILE' then
//  begin
//   CreateFlowFooter;
//  end
//  else
//  begin
//    // Write To Database
//    with main_data_module.updatequery do
//    begin
//      close;
//      sql.clear;
//      sql.Add('insert into EDMGR.BATCH_FLOWS_FOR_SENDING_D0055 values (');
//      sql.Add(''''+ToParty+''','''+frm_main.filememo.Text+''',sysdate,null)');
//      execute;
//    end;
//  end;
//end;

procedure TFRM_Export.COSD0148;
var
  SSD, formattedssd: string;
  Recipient_MO, Recipient_DC: string;
  AGENTID, aMPAN: string;
  memoln: Integer;
  qryAgents, qryMOs: TOracleDataSet;
begin
  // ***************************************************************************//
  // This is where we are going to write the new procedure for the D0148        //
  // dataflows as a result of DTC 6.2                                           //
  // Primarily this change affects how the D0148s are populated & when they     //
  // are to be sent to the MO or DC                                             //
  // NOTE: THE FOLLOWING CODE HERE ONLY APPLIES TO 'CHANGE OF SUPPLIER' PROCESS //
  // ***************************************************************************//

  // ***************************************************************************//
  // METER OPERATOR FLOWS                                                       //
  // First identify all MPANS that require a flow sending to the Meter Operator //
  // See Section 8 of DTC 6.2 Page C-13                                         //
  // ***************************************************************************//

  // First get a list of all Meter Operators
  try
    qryAgents := gSqlUtil.CreateCursor
        ('EDMGR.PK_NHH_APPOINT_MO_UMOL.PR_GET_MOP_AGENTS_D0148(:p_dataset)', TRANSACTION_NO,
        [':p_dataset', otCursor, null]);
    try
      // for each MO agent create a file
      while not qryAgents.Eof do
      begin
        //Recipient_MO := '';
        OutputFlowFlowCount                    := 0;
        OutputFlowLineCount                    := 0;
        AGENTID                                := qryAgents.FieldByName('confirmed_mo_id').AsString;
        FRM_File_Progress.ProgressBar.Position := 0;
        // create output file Header
        CreateFlowHeader('D0148001', AGENTID, qryAgents.FieldByName('confirmed_mo_role').AsString);

        // Create Details
        //****************************************************************************//
        //                      8.1.1 to MO Single Instance **                        //
        //  No flow yet sent to MO, but both MO&DC have confirmed their appointments  //
        //****************************************************************************//
        try
          qryMOs := gSqlUtil.CreateCursor
            ('EDMGR.PK_NHH_APPOINT_MO_UMOL.PR_MO_DC_APPOINT_CONF(:p_conf_mo_id, :p_dataset)', TRANSACTION_NO,
            [':p_conf_mo_id', otString, AGENTID,
             ':p_dataset', otCursor, null]);

          try
            // Write records to this file
            FRM_File_Progress.ProgressBar.Max      := qryMOs.RecordCount;
            FRM_File_Progress.ProgressBar.Position := 0;

            // Repeat for Each MPAN
            while not qryMOs.Eof do
            begin
              FRM_File_Progress.D_File.Caption           := 'File:';
              FRM_File_Progress.LabelCount.Caption       := 'MOP + DC';
              FRM_File_Progress.ProgressBar.Position     := FRM_File_Progress.ProgressBar.Position + 1;
              aMPAN                                      := qryMOs.FieldByName('mpancore').AsString;
              FRM_File_Progress.Statusbar.Panels[0].Text := 'MPAN: ' + aMPAN;
              FRM_File_Progress.Statusbar.Update;
              QueryAgents(aMPAN);
              POP148MO('G', aMPAN, qryMOs.FieldByName('ssd').AsString); // RUN EXAMPLE G
              qryMOs.Next;
            end;

          finally
            FreeAndNil(qryMos);
          end;

        except
          on E: EOracleError do
          if E.ErrorCode <> 1 then
            FRM_common.DisplayOracleError(E.Message, FlowVersion, 'Q',
              Filename);
        end;

        ///////////////////////////////////////////////////////////////////////
        //                                                                   //
        // Added for new connection                                          //
        //                                                                   //
        ///////////////////////////////////////////////////////////////////////
        try
          qryMOs := gSqlUtil.CreateCursor
            ('EDMGR.PK_NHH_APPOINT_MO_UMOL.PR_MO_DC_APPOINT_CONF_NC(:p_conf_mo_id, :p_dataset)', TRANSACTION_NO,
            [':p_conf_mo_id', otString, AGENTID,
             ':p_dataset', otCursor, null]);

          try
            // Write records to this file
            FRM_File_Progress.ProgressBar.Max      := qryMOs.RecordCount;
            FRM_File_Progress.ProgressBar.Position := 0;

            // Repeat for Each MPAN
            while not qryMOs.Eof do
            begin
              FRM_File_Progress.D_File.Caption           := 'File:';
              FRM_File_Progress.LabelCount.Caption       := 'New Connections';
              FRM_File_Progress.ProgressBar.Position     := FRM_File_Progress.ProgressBar.Position + 1;
              aMPAN                                      := qryMOs.FieldByName('mpancore').AsString;
              FRM_File_Progress.Statusbar.Panels[0].Text := 'MPAN: ' + aMPAN;
              FRM_File_Progress.Statusbar.Update;
              QueryAgents(aMPAN);
              POP148MO('C', aMPAN, qryMOs.FieldByName('ssd').AsString); //  RUN EXAMPLE C
              qryMOs.Next;
            end;

          finally
            FreeAndNil(qryMOs);
          end;

        except
          on E: EOracleError do
          if E.ErrorCode <> 1 then
            FRM_common.DisplayOracleError(E.Message, FlowVersion, 'Q',
              Filename);
        end;

        //****************************************************************************//
        //                            8.1.2 to MO Flow 1 **                           //
        //  No flow yet sent to MO, but MO HAS confirmed but DC has not yet confirmed //
        //****************************************************************************//
        try
          qryMOs := gSqlUtil.CreateCursor
            ('EDMGR.PK_NHH_APPOINT_MO_UMOL.PR_DC_APPOINT_NOT_CONF(:p_conf_mo_id, :p_dataset)', TRANSACTION_NO,
            [':p_conf_mo_id', otString, AGENTID,
             ':p_dataset', otCursor, null]);

          try
            // Write records to this file
            FRM_File_Progress.ProgressBar.Max      := qryMOs.RecordCount;
            FRM_File_Progress.ProgressBar.Position := 0;

            // Repeat for Each MPAN
            while not qryMOs.Eof do
            begin
              FRM_File_Progress.D_File.Caption           := 'File:';
              FRM_File_Progress.LabelCount.Caption       := 'MOP Only';
              FRM_File_Progress.ProgressBar.Position     := FRM_File_Progress.ProgressBar.Position + 1;
              aMPAN                                      := qryMOs.FieldByName('mpancore').AsString;
              FRM_File_Progress.Statusbar.Panels[0].Text := 'MPAN: ' + aMPAN;
              FRM_File_Progress.Statusbar.Update;
              QueryAgents(aMPAN);
              POP148MO('O', aMPAN, qryMOs.FieldByName('ssd').AsString); // RUN EXAMPLE O
              qryMOs.Next;
            end;

          finally
            FreeAndNil(qryMOs);
          end;

        except
          on E: EOracleError do
          if E.ErrorCode <> 1 then
            FRM_common.DisplayOracleError(E.Message, FlowVersion, 'Q',
              Filename);
        end;

        // Cannot test this process until after after first run
        //****************************************************************************//
        //                            8.1.2 to MO Flow 2 **                           //
        //  MO already notified of OLD MO ONLY, now they need to know who the DC is   //
        //****************************************************************************//
        try
          qryMOs := gSqlUtil.CreateCursor
            ('EDMGR.PK_NHH_APPOINT_MO_UMOL.PR_DC_APPOINT(:p_conf_mo_id, :p_dataset)', TRANSACTION_NO,
            [':p_conf_mo_id', otString, AGENTID,
             ':p_dataset', otCursor, null]);

          try
            // Write records to this file
            FRM_File_Progress.ProgressBar.Max := qryMOs.RecordCount;
            FRM_File_Progress.ProgressBar.Position := 0;

            // Repeat for Each MPAN
            while not qryMOs.Eof do
            begin
              FRM_File_Progress.D_File.Caption := 'File:';
              FRM_File_Progress.LabelCount.Caption := 'DC only';
              FRM_File_Progress.ProgressBar.Position :=
                FRM_File_Progress.ProgressBar.Position + 1;
              aMPAN := qryMOs.FieldByName('mpancore').AsString;
              FRM_File_Progress.Statusbar.Panels[0].Text := 'MPAN: ' + aMPAN;
              FRM_File_Progress.Statusbar.Update;
              QueryAgents(aMPAN);
              POP148MO('C', aMPAN, qryMOs.FieldByName('ssd').AsString); // RUN EXAMPLE C
              qryMOs.Next;
            end;


          finally
            FreeAndNil(qryMOs);
          end;

        except
          on E: EOracleError do
          if E.ErrorCode <> 1 then
            FRM_common.DisplayOracleError(E.Message, FlowVersion, 'Q',
              Filename);
        end;

        // Create Footer Record
        CreateFlowfooter;
        //if mpancount=0 then deletefile(filename);
        qryAgents.Next;
      end;

    finally
      FreeAndNil(qryAgents);
    end;

  except
    on E: EOracleError do
      if E.ErrorCode <> 1 then
        FRM_common.DisplayOracleError(E.Message, FlowVersion, 'Q', Filename);
  end;

 //****************************************************************************//
 //                           DATA COLLECTOR FLOWS                             //
 //  First identify all MPANS that require a flow sending to the Data Collector//
 //                    See Section 8.2.1 of DTC 6.2 Page C-13                  //
 //****************************************************************************//
 // First get a list of all data collectors
 with agentsquery do
 begin
  close;
  sql.clear;
  sql.add('Select Distinct(Confirmed_dc_id), confirmed_dc_role from edmgr.mpan_status');
  sql.add('Where confirmed_dc_id is not null');
   sql.add('and Regstatus in (''REGISTERED'',''FUTURE LOSS'',''LOSS_PENDING'',''LOST'')');
    sql.add('and measurement_class in (''A'',''C'',''E'',''F'',''G'')');
  sql.add('and dc_check=''Y''');
  sql.add('and (DC_CHECK=''Y'' and MO_CHECK=''Y'' and DA_CHECK=''Y''');
  sql.add('and D0148_DC_OLDDC is null and D0148_DC_NEWMO is null and D0148_DC_NEWDA is null)');
  sql.add('or (DC_CHECK=''Y'' and MO_CHECK is null and DA_CHECK is null');
  sql.add('and D0148_DC_OLDDC is null)');
  sql.add('or (DC_CHECK=''Y'' and MO_CHECK =''Y'' and DA_CHECK is null');
  sql.add('and D0148_DC_OLDDC is not null and d0148_dc_newmo is null)');
  sql.add('or (DC_CHECK=''Y'' and MO_CHECK =''Y'' and DA_CHECK =''Y''');
  sql.add('and D0148_DC_OLDDC is not null and d0148_dc_newmo is not null and d0148_dc_newda is null)');
  sql.add('or (DC_CHECK=''Y'' and DA_CHECK =''Y'' and MO_CHECK is null');
  sql.add('and D0148_DC_OLDDC is not null and d0148_dc_newda is null)');
  sql.add('or (DC_CHECK=''Y'' and DA_CHECK =''Y'' and MO_CHECK =''Y''');
  sql.add('and D0148_DC_OLDDC is not null and d0148_dc_newda is not null and D0148_DC_newmo is null)');
  sql.add('or (DC_CHECK=''Y'' and MO_CHECK =''Y'' and DA_CHECK is null');
  sql.add('and D0148_DC_OLDDC is null and d0148_dc_newmo is null)');
  sql.add('or (DC_CHECK=''Y'' and MO_CHECK =''Y'' and DA_CHECK =''Y''');
  sql.add('and D0148_DC_OLDDC is not null and d0148_dc_newmo is not null and d0148_dc_newda is null)');
  sql.add('or (DC_CHECK=''Y'' and DA_CHECK =''Y'' and MO_CHECK is null');
  sql.add('and D0148_DC_OLDDC is null and d0148_dc_newda is null)');
  sql.add('or (DC_CHECK=''Y'' and MO_CHECK =''Y'' and DA_CHECK =''Y''');
  sql.add('and D0148_DC_OLDDC is not null and d0148_dc_newda is not null and D0148_DC_newmo is null)');
  open;
 end;
 // for each DC agent and role create a file
 while not agentsquery.eof do
 begin
  Recipient_dc:='';
  OutputflowFlowCount:=0;
  OutputFlowLineCount:=0;
  Agentid:=Agentsquery.fields[0].text;
  // create output file Header
  FRM_File_Progress.progressbar.position:=0;
  CreateFlowHeader('D0148001',AgentID,AgentsQuery.fields[1].text);
  // Create Details

  ////////////////////////////////////////////////////////////////////////////////////////////////////
  // 28/01/02 D0148 DC flows CHANGED BY LEE SIMPLIFIED
  ////////////////////////////////////////////////////////////////////////////////////////////////////

  // get a list of MPANS for DC requiring D0148s to be sent
  with MOsquery do
  begin
   close;
   sql.clear;
   sql.add('Select MPANCORE,SSD,Pes_Code,Confirmed_DC_ID,Confirmed_DC_Role,');
   sql.add('D0148_dc_olddc,D0148_dc_newmo,D0148_dc_newDA,');
   sql.add('DC_Check,mo_check,da_check from edmgr.mpan_status');
   sql.add('Where ');
   sql.add('measurement_class in (''A'',''C'',''E'',''F'',''G'')');

   sql.add('and (D0148_dc_olddc is null');

   // This if either MO or DA confirmed (Partial)
   {sql.add('or (mo_check=''Y'' and d0148_dc_newmo is null)');
   sql.add('or (da_check=''Y'' and d0148_dc_newda is null))');}
   // Commented Out by Lee & Chris 17/05/2012

    // This if Both MO or DA confirmed (FULL)
   sql.add('and (mo_check=''Y'' and d0148_dc_newmo is null)');
   sql.add('and (da_check=''Y'' and d0148_dc_newda is null))');


   sql.add('and (DC_check=''Y''');
   sql.add('and Regstatus in (''REGISTERED'',''FUTURE LOSS'',''LOSS PENDING'',''LOST'')');
   sql.add('and new_connection=''F''');        // and Not New Connections
   SQL.ADD('and confirmed_DC_id='''+Agentid+'''');// and going to the Confirmed MO
   sql.add('and confirmed_DC_role='''+AgentsQuery.fields[1].text+''')');
   sql.add('order by MPANCORE');
   open;
  end;
  OutputFlowFlowCount:=0;
  FRM_File_Progress.progressbar.max:=mosquery.recordcount;
  FRM_File_Progress.progressbar.position:=0;
  while not mosquery.eof do
  begin
   // write group D0270
   // format the startdate into a dataflow format
   FRM_File_Progress.d_file.caption:='File:';
   FRM_File_Progress.labelcount.caption:='COS';
   FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
   FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MOsquery.fields[0].text;
   Application.ProcessMessages;
   queryagents(MOsquery.fields[0].text);
   mpancore:=mosquery.fields[0].text;
   ssd:=mosquery.fields[1].text;
   FormattedSSD:=Formatdatetime('YYYYMMDD',strtodate(SSD));
   D0148Memo.clear;
   begin
    D0148Memo.lines.add('270|'+MPANCORE+'|'+FormattedSSD+'|');
    inc(OutputFlowFlowCount);
   end;
   //****************************************************************************//
   //                         Check for DC confirmation                          //
   //            IF DC_D0148_DC_old_DC is null                                   //
   //****************************************************************************//
   if (mosquery.fields[5].text='') then POP148DC('Q',MOsquery.fields[0].text,MOsquery.fields[1].text);  //  RUN EXAMPLE Q

   //****************************************************************************//
   //                         Check for MO confirmation                          //
   //              If D0148_dc_newmo is null and MO Check='Y'                    //
   //****************************************************************************//
   if (mosquery.fields[6].text='') and (mosquery.fields[9].text='Y') then POP148DC('B',MOsquery.fields[0].text,MOsquery.fields[1].text);  //  RUN EXAMPLE B
   //****************************************************************************//
   //                         Check for DA confirmation                          //
   //             If D0148_dc_newda is null and DA_check='Y'                     //
   //****************************************************************************//
   if (mosquery.fields[7].text='') and (mosquery.fields[10].text='Y') then POP148DC('F',MOsquery.fields[0].text,MOsquery.fields[1].text);  //  RUN EXAMPLE F

   for Memoln:=1 to D0148memo.lines.count do
   begin
    S:=D0148Memo.lines[memoln-1];
    frm_main.WriteLinetoFile(s);
    inc(OutputFlowLineCount);
   end;
   mosquery.next; // Do Next MPAN
  end;


  ///////////////////////////////////////////////////////////////////////////////
  //
  // new connection for DC //////////////////////////////////////////////////////
  ///////////////////////////////////////////////////////////////////////////////
  with MOsquery do // select all MPANS
  begin            // DCCHECK='Y' and MOCHECK='Y' and DACHECK='Y'
    close;          // and DC-OLDDC is null and DC-NEWMO is NULL and DC-NEWDA is NULL
   sql.clear;
   sql.add('Select MPANCORE,SSD,Pes_Code,Confirmed_DC_ID,Confirmed_DC_Role from edmgr.mpan_status');
   sql.add('Where DC_CHECK=''Y'' and MO_CHECK=''Y'' and DA_CHECK=''Y''');
     sql.add('and measurement_class in (''A'',''C'',''E'',''F'',''G'')');
   sql.add('and D0148_DC_OLDDC is null and D0148_DC_NEWMO is null and D0148_DC_NEWDA is null');
    sql.add('and Regstatus in (''REGISTERED'',''FUTURE LOSS'',''LOSS PENDING'',''LOST'')');
   sql.add('and new_connection=''T''');        // and New Connections
   SQL.ADD('and confirmed_DC_id='''+Agentid+'''');// and going to the Confirmed DC
   sql.add('and confirmed_DC_role='''+AgentsQuery.fields[1].text+'''');
   open;
  end;

  FRM_File_Progress.progressbar.max:=mosquery.recordcount;
  FRM_File_Progress.progressbar.position:=0;
  while not mosquery.eof do
  begin
   // write group D0270
   // format the startdate into a dataflow format
   FRM_File_Progress.d_file.caption:='File:';
   FRM_File_Progress.labelcount.caption:='New Connection';
   FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
   FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MOsquery.fields[0].text;
   Application.ProcessMessages;
   queryagents(MOsquery.fields[0].text);
   mpancore:=mosquery.fields[0].text;
   ssd:=mosquery.fields[1].text;
   FormattedSSD:=Formatdatetime('YYYYMMDD',strtodate(SSD));
   D0148Memo.clear;
   begin
    D0148Memo.lines.add('270|'+MPANCORE+'|'+FormattedSSD+'|');
    inc(OutputFlowFlowCount);
   end;
   POP148DC('J',MOsquery.fields[0].text,MOsquery.fields[1].text);  //  RUN EXAMPLE Q
   for Memoln:=1 to D0148memo.lines.count do
   begin
    S:=D0148Memo.lines[memoln-1];
    frm_main.WriteLinetoFile(s);
    inc(outputflowlinecount);
   end;
   mosquery.next; // Do Next MPAN
  end;
  // Create Footer Record
  CreateFlowfooter;
  //if mpancount=0 then deletefile(filename);
  agentsquery.next;
 end;
end;


//****************************************************************************//
//                     New D0148 Populator for MO flows                       //
//****************************************************************************//
Procedure TFRM_Export.POP148MO(Example,MPANCORE,SSD:String);
var
FormattedSSD:string;
OLDMO:STRING;
memoln:integer;
Begin

 // format the startdate into a dataflow format
 FormattedSSD:=Formatdatetime('YYYYMMDD',strtodate(SSD));
 D0148Memo.clear;
 // Group 270 G O C   // MPAN+SSD
 if (example='G') or (example='O') or (Example='C') then
 begin
  D0148Memo.lines.add('270|'+MPANCORE+'|'+FormattedSSD+'|');
 end;

 // Group 271 G   C   // New DC and Status
 if (example='G') or (Example='C') then
 begin
  with generalquery do
  begin
   close;
   sql.clear;
   sql.add('Select Confirmed_DC_ID from edmgr.mpan_status where mpancore='''+MPANCORE+'''');
   open;
  end;
  D0148Memo.lines.add('271|'+Generalquery.fields[0].text+'|N|');
 end;

 // Group 272 G   C   // New DC EFfective from Date
 if (example='G') or (Example='C') then
 begin
  // Date Needs to be appointment effective from Date.
  if dc2.fields[3].text <>'' then
   D0148Memo.lines.add('272|'+formatdatetime('YYYYMMDD',strtodatetime(dc2.fields[3].text))+'|')
   else D0148Memo.lines.add('272|'+formattedssd+'|');// e.g. If D0012 but no D0011
 end;

 // Group 274 G O    // Old MOP details
 if (example='G') or (example='O') then
 begin
  // Try and establish who the OLD MO is from the D0260
  with generalquery do
  begin
   close;
   sql.clear;
   sql.add('Select OLD_MO from edmgr.mpan_status where mpancore='''+MPANCORE+'''');
   open;
  end;
  oldmo:=generalquery.fields[0].text;
  if oldmo='' then
  begin
   // we dont know who old MO was so do we guess?
   // Some assume same MO we are appointing
   with generalquery do
   begin
    close;
    sql.clear;
    sql.add('Select confirmed_MO_ID from edmgr.mpan_status where mpancore='''+MPANCORE+'''');
    open;
   end;
   oldmo:=generalquery.fields[0].text
  end;
  D0148Memo.lines.add('274|'+OLDMO+'|O|');
 end;

 // Group 276 G O    // old MOP date
 if (example='G') or (example='O') then
 begin
  /////////////////////////////////////////////////////
  ///////effective from date -1 ///////////////////////
  /////////////////////////////////////////////////////
  D0148Memo.lines.add('276|'+formatdatetime('YYYYMMDD',strtodatetime(moquery.fields[3].text)-1)+'|');
 end;
 inc(OutputflowFlowCount);
 for Memoln:=1 to D0148memo.lines.count do
 begin
 S:=D0148Memo.lines[memoln-1];
 frm_main.WriteLinetoFile(s);
 inc(outputflowlinecount);
 end;
end;

//****************************************************************************//
//                     New D0148 Populator for DC flows                       //
//****************************************************************************//
Procedure TFRM_Export.POP148DC(Example,MPANCORE,SSD:String);
var
FormattedSSD:string;
OLDDC:STRING;
Begin

 FormattedSSD:=Formatdatetime('YYYYMMDD',strtodate(SSD));
  // Group 271 P Q     S R  // DC Details
 if (example='P') or (Example='Q') or (Example='S') or (Example='R')then
  // Try and establish who the OLD DC is from the D0260
 begin
  with generalquery do
  begin
   close;
   sql.clear;
   sql.add('Select OLD_DC from edmgr.MPAN_STATUS where mpancore='''+MPANCORE+'''');
   open;
  end;
  olddc:=generalquery.fields[0].text;
  if olddc='' then
  begin
   // we dont know who old Dc was so do we guess?
   // Some assume same DC we are appointing
   with generalquery do
   begin
    close;
    sql.clear;
    sql.add('Select confirmed_DC_ID from edmgr.mpan_status where mpancore='''+MPANCORE+'''');
    open;
   end;
   oldDC:=generalquery.fields[0].text
  end;
  D0148Memo.lines.add('271|'+OLDDC+'|O|');
 end;

 // Group 273 P Q     S R  DC effective Date -1
 if (example='P') or (Example='Q') or (Example='S') or (Example='R')then
 begin
  ///////////////////////////////////////////////  effective from date -1
  if dc2.fields[3].text<>'' then
   D0148Memo.lines.add('273|'+formatdatetime('YYYYMMDD',strtodatetime(dc2.fields[3].text)-1)+'|')
   else D0148Memo.lines.add('273|'+formatdatetime('YYYYMMDD',strtodatetime(ssd)-1)+'|'); //d0012 no D0011
 end;

 // Group 274 P   B     R J New MO details
 if (example='P') or (Example='B') or (Example='R') or (Example='J') then
 begin
  with generalquery do
  begin
   close;
   sql.clear;
   sql.add('Select confirmed_MO_ID from edmgr.mpan_status where mpancore='''+MPANCORE+'''');
   open;
  end;
  D0148Memo.lines.add('274|'+Generalquery.fields[0].text+'|N|');
 end;

 // Group 275 P   B     R  J MO effective from date
 if (example='P') or (Example='B') or (Example='R') or (Example='J') then
 begin
  with generalquery do
  begin
   close;
   sql.clear;
   sql.add('Select SSD from edmgr.mpan_status where mpancore='''+MPANCORE+'''');
   open;
  end;
  ///////////////////////////////////////////////
  ///////effective from date ////////////////////
  ///////////////////////////////////////////////
  if moquery.fields[3].text<>'' then D0148Memo.lines.add('275|'+formatdatetime('YYYYMMDD',strtodatetime(moquery.fields[3].text))+'|')
  else D0148Memo.lines.add('275|'+formatdatetime('YYYYMMDD',strtodatetime(generalquery.fields[0].text))+'|');
 end;

 // Group 277 P     F S J // New Da details
 if (example='P') or (Example='F') or (Example='S') or (Example='J') then
 begin
  with generalquery do
  begin
   close;
   sql.clear;
   sql.add('Select confirmed_DA_ID from edmgr.mpan_status where mpancore='''+MPANCORE+'''');
   open;
  end;
  D0148Memo.lines.add('277|'+Generalquery.fields[0].text+'|N|');
 end;

 // Group 278 P     F S J // New Da effective from Date
 ////////////////////////////////////////////////////////
 ///// effective from date //////////////////////////////
 ////////////////////////////////////////////////////////
 if (example='P') or (Example='F') or (Example='S') or (Example='J') then
 begin
  D0148Memo.lines.add('278|'+formatdatetime('YYYYMMDD',strtodatetime(da2.fields[3].text))+'|');
 end;
end;

procedure TFRM_Export.newD0148btnClick(Sender: TObject);
begin
 COSD0148;
end;

{Procedure TFRM_Export.GenerateAgentAppointments;
begin
 NHHDAD0153;
 NHHDCD0155;
 NHHMOD0155;

 //HHDAD0153;
 //HHDCD0155;
 //HHMOD0155;
end;}

Procedure TFRM_Export.CreateAppointments(Agent,AgentRole,mtc,NHH,COAEFD,GSP,SMART_COS:String);
Var
AgentField:Integer;
Openfile: Boolean;
LastAgent, NewAgent,FlowType,readcycle:String;
Cefsd,pc:string;
Begin
 // ResetFileCounters etc
 OpenFile := False;
 LastAGENT:='';

 // Loop Through All MPANS
 FRM_File_Progress.progressbar.position:=0;

 while not MiscQuery.eof do
 begin
  // Get Agent to be Appointed
  if Agent='MO' then Agentfield:=29;
  if Agent='DC' then Agentfield:=35;
  if Agent='DA' then Agentfield:=41;
  NewAgent:=MiscQuery.fields[agentfield].text;
  pc:=MiscQuery.Fields[16].text;
  // if change of Agent then New Agent = Agent specified in Parameter
  if coaefd<>'' then newagent:=mtc;

  // Check if Agent is different to Last Agent
  if NewAgent<>LastAgent then
  Begin
   lastagent:=NewAgent;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   // Create File
   if Agent='MO' then FlowType:='D0155';
   if Agent='DC' then FlowType:='D0155';
   if Agent='DA' then FlowType:='D0153';
   CreateFlowHeader(Flowtype+'001',NewAgent,Agentrole);
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
  End;

  // Write Record to The File Check What the Flow Version is
  if FlowType='D0153' then
  Begin
   // Write D0153 Body
   S:='312'+'|';
   S:=S+MiscQuery.fields[0].text+ '|'; // Mpan Core

   //FRM_Main.statusbar.panels[1].text:='Creating D0153 to DA for MPANCore '+mpancore;
   //application.processmessages;

   cEFSD:=MiscQuery.fields[5].text;    // SSD
   s:=s+GETAPPDATE(coaefd,CEFSD);
   // Effective From Date DAA
   S:=S+formatdatetime('YYYYMMDD',strtodate(cEFSD))+'|';                 // Effective From Date Regi

   // Get contract Ref, Service Ref and Service Level Refs
   with queryrefcode do
   Begin
    close;
    sql.clear;
    sql.add('select * from edmgr.AGENT_SERVICE_CODES');
    sql.add('where Agent_iD = '+quotedstr(NEWAGENT));
    sql.add('AND AGENT_ROLE = '+quotedstr(AGENTROLE));
    sql.add('and hh_nhh_flag='+quotedstr(NHH));
    sql.add('and gsp_group_ID='+quotedstr(GSP));
    sql.Add('and profile_class='+quotedstr(PC));
    open;
   end;
   if queryrefcode.recordcount=0 then
   Begin
     // Get contract Ref, Service Ref and Service Level Refs
    with queryrefcode do
    Begin
     close;
     sql.clear;
     sql.add('select * from edmgr.AGENT_SERVICE_CODES');
     sql.add('where Agent_iD = '''+NEWAGENT+'''');
     sql.add('AND AGENT_ROLE = '''+AGENTROLE+'''');
     sql.add('and hh_nhh_flag='''+NHH+'''');
     open;
    end;
   end;

   S:=S+queryRefCode.fields[2].text+'|'; // Contract Reference
   S:=S+'|';                             // additional information
   frm_main.WriteLinetoFile(S);
   inc(OutputFlowLineCount);
   Inc(OutputFlowFlowCount);
   S:='313'+'|';
   S:=S+queryRefCode.fields[3].text+'|';  // Service Ref
   S:=S+queryRefCode.fields[4].text+'|';  // Service Level Ref
   frm_main.WriteLinetoFile(S);
   inc(OutputFlowLineCount);
  end;

  If Flowtype='D0155' then
  Begin
   // Write D0155 Body
   S:= '315'+'|';
   S:=S+MiscQuery.fields[0].text+'|'; // MPAN

   //FRM_Main.statusbar.panels[1].text:='Creating D0155 to '+agentrole+' for MPANCore '+mpancore;
   //application.processmessages;


   cEFSD:=MiscQuery.fields[5].text;     // SSD
   S:=S+formatdatetime('YYYYMMDD',strtodate(cEFSD))+ '|';
   // Obtain Address Details. Usually latest D0217 date
   // unless it's changed. See table EDMGR.MPAS_CURR_ADDR
   with addressquery do
   begin
    close;
    with sql do
    begin
     clear;
     add('select A.METERING_POINT_ADDRESS1,');
     add('A.METERING_POINT_ADDRESS2,');
     add('A.METERING_POINT_ADDRESS3,');
     add('A.METERING_POINT_ADDRESS4,');
     add('A.METERING_POINT_ADDRESS5,');
     add('A.METERING_POINT_ADDRESS6,');
     add('A.METERING_POINT_ADDRESS7,');
     add('A.METERING_POINT_ADDRESS8,');
     add('A.METERING_POINT_ADDRESS9,');
     add('A.METERING_POINT_POSTCODE');
     add('from edmgr.MPAS_CURRENT_ADDR A');
     add('where A.MPANCORE = '''+MiscQuery.fields[0].text+'''');
    end;
    open;
   end;
   S:=S+addressquery.fields[0].text+'|';
   S:=S+addressquery.fields[1].text+'|';
   S:=S+addressquery.fields[2].text+'|';
   S:=S+addressquery.fields[3].text+'|';
   S:=S+addressquery.fields[4].text+'|';
   S:=S+addressquery.fields[5].text+'|';
   S:=S+addressquery.fields[6].text+'|';
   S:=S+addressquery.fields[7].text+'|';
   S:=S+addressquery.fields[8].text+'|';
   S:=S+addressquery.fields[9].text+'|';
   // Get contract Ref, Service Ref and Serive Level Refs
   with queryrefcode do
   Begin
    close;
    sql.clear;
    sql.add('select * from edmgr.AGENT_SERVICE_CODES');
    sql.add('where Agent_iD = '+quotedstr(NEWAGENT));
    sql.add('AND AGENT_ROLE = '+quotedstr(AGENTROLE));
    sql.add('and hh_nhh_flag='+quotedstr(NHH));
    sql.add('and gsp_group_ID='+quotedstr(GSP));
    sql.Add('and profile_class='+quotedstr(PC));
    open;
   end;

   if queryrefcode.recordcount=0 then
   Begin
    with queryrefcode do
    Begin
     close;
     sql.clear;
     sql.add('select * from edmgr.AGENT_SERVICE_CODES');
     sql.add('where Agent_iD = '''+NEWAGENT+'''');
     sql.add('AND AGENT_ROLE = '''+AGENTROLE+'''');
     sql.add('and hh_nhh_flag='''+NHH+'''');
     open;
    End;
   end;
   // Check for MTC Code
   {if (MTC>='835') and (MTC<='843') then
   Begin
    with queryrefcode do
    Begin
     close;
     sql.clear;
     sql.add('select * from edmgr.AGENT_SERVICE_CODES_PPM');
     sql.add('where Agent_iD = '''+NEWAGENT+'''');
     sql.add('AND AGENT_ROLE = '''+AGENTROLE+'''');
     sql.add('and hh_nhh_flag='''+NHH+'''');
     open;
    end;
   end; }

   S:=S+queryRefCode.fields[2].text+'|'; // Contract Reference
   S:=S+'N' +'|';                        // Retreival Method
   //********
   s:=s+MiscQuery.fields[74].text+'|';          // Gsp Group Added DTC V77
   //********
   frm_main.WriteLinetoFile(S);
   Inc(OutputFlowLineCount);
   // Group 320 Reading Cycles
   // Only Do This if a Data Collector Flow
   if (AgentRole = 'D') or (AgentRole = 'C') then
   begin
    S:= '';
    S:=S+'320'+'|';

    //************************************************************************//
    // 154235139: Ability to set reading schedule in D0155 by Profile Class   //
    //************************************************************************//
    {
    if MiscQuery.fields[16].text='1' then readcycle:=queryrefcode.fields[8].text;
    if MiscQuery.fields[16].text='2' then readcycle:=queryrefcode.fields[9].text;
    if MiscQuery.fields[16].text='3' then readcycle:=queryrefcode.fields[10].text;
    if MiscQuery.fields[16].text='4' then readcycle:=queryrefcode.fields[11].text;
    if MiscQuery.fields[16].text='5' then readcycle:=queryrefcode.fields[12].text;
    if MiscQuery.fields[16].text='6' then readcycle:=queryrefcode.fields[13].text;
    if MiscQuery.fields[16].text='7' then readcycle:=queryrefcode.fields[14].text;
    if MiscQuery.fields[16].text='8' then readcycle:=queryrefcode.fields[15].text;
    }
    //************************************************************************//

    readcycle:=queryrefcode.fields[7].text;

    if readcycle='' then readcycle:='O';
    s:=s+readcycle+'|';

    S:=S+'|'; // Scheculed Read Date;
    frm_main.WriteLinetoFile(s);
    Inc(OutputFlowLineCount)
   end; // End Group 320

   // Group 316 Effective From Dates
   // Only Do This if a Data Collector Flow
   if (AgentRole = 'D') or (AgentRole = 'C') then
   begin
    S:='';
    S:=S+'316'+'|';
    s:=s+GETAPPDATE(coaefd,CEFSD);
    frm_main.WriteLinetoFile(s);
    Inc(OutputFlowLineCount)
   end; // End Group 316

   // Group 317 only if MO Flow
   if AgentRole = 'M' then // Meter Operator
   begin
    S:='';
    S:=S+'317'+'|';
    s:=s+GETAPPDATE(coaefd,CEFSD);
    frm_main.WriteLinetoFile(s);
    Inc(Outputflowlinecount);
   end;

   // Group 318 Agreed Service Details
   S:='';
   S:=S+'318'+'|';
   S:=S+queryRefcode.fields[3].text+'|';
   S:=S+queryRefCode.fields[4].text+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputflowFlowCount);
   Inc(OutputFlowLineCount);

   // DTC 11.6 Change
   if (AgentRole = 'D') or (AgentRole = 'C') then
   begin
    S:='';
    S:=S+'22L'+'|';
    s:=s+SMART_COS+'|';
    frm_main.WriteLinetoFile(s);
    Inc(OutputFlowLineCount)
   end; // End Group 22

     // DTC 11.6 Change
   if (AgentRole = 'M') then
   begin
    S:='';
    S:=S+'22L'+'|N|';     // Value is always 'N' when sending to MOP
    frm_main.WriteLinetoFile(s);
    Inc(OutputFlowLineCount)
   end; // End Group 22

  end;
  MiscQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;

{procedure TFRM_Export.D0153155BtnClick(Sender: TObject);
begin
 GenerateAgentAppointments;
end;}

Procedure TFRM_Export.TerminateAppointments;
Begin
 TerminateMOPLoss;
 TerminateDCloss;
 TerminateDALoss;
End;

Procedure TFRM_Export.CreateD0151flow(Reason:string);
Var
MPAN,AgentID,AgentROLE:String;
SSD,EFTSSD:Tdatetime;
Begin
 FRM_File_Progress.progressbar.position:=0;
 while not GeneralQuery.eof do
 Begin
  MPAN:=GeneralQuery.fields[0].text;
  EFTSSD:=strtodate(GeneralQuery.fields[1].text);
  AGENTID:=GeneralQuery.fields[2].text;
  AGENTROLE:=GeneralQuery.Fields[3].text;
  SSD:=strtodate(GeneralQuery.fields[4].text);
  if (agentrole<>oldagentrole) or (agentid<>oldagentid) then
  begin
   FRM_File_Progress.progressbar.position:=0;
   CreateFlowHeader('D0151001',Agentid,AgentRole);
   OutputFlowFlowcount:=0;
   OutPutFlowLineCount:=0;
  end;
  oldagentrole:=agentrole;
  oldagentid:=agentid;
  // Create Body
  frm_main.WriteLinetoFile('297|'+MPAN+'|'+FormatDateTime('YYYYMMDD',SSD)+'|'+Reason+'||'); // Termination will always be LC for Loss of Contract to Supply (COS)
  if agentrole='M' then frm_main.WriteLinetoFile('298|'+FormatDateTime('YYYYMMDD',EFTSSD)+'|');// If Mop Termination
  if (agentrole='C') or (agentrole='D') then frm_main.WriteLinetoFile('299|'+FormatDateTime('YYYYMMDD',EFTSSD)+'|');// If DC Termination
  if (agentrole='A') or (agentrole='B') then frm_main.WriteLinetoFile('300|'+FormatDateTime('YYYYMMDD',EFTSSD)+'|');// IF DA Termination
  inc(OutputFlowFlowCount);
  OutputFlowlinecount:=OutputFlowlinecount+2;

  if (reason='LC') and (AGENTROLE='M') then
  Begin
   frm_main.WriteLinetoFile('99H|'+GeneralQuery.fields[5].text+'|');// New SUpplier
   OutputFlowlinecount:=OutputFlowlinecount+1;
  end;


  GeneralQuery.next;
  AGENTID:=GeneralQuery.fields[2].text;
  AGENTROLE:=GeneralQuery.Fields[3].text;
  if (agentrole<>oldagentrole) or (oldagentid<>agentid) then cfooter:=true
  else cfooter:=false;
  if GeneralQuery.eof then cfooter:=true;
  if cfooter=true then CreateFlowFooter;
 end;
end;

procedure TFRM_Export.IssueD0005readRequests;
Var
V:Textfile;
agentrole,AgentID,datenow,newdir:String;
Mpan,REASON,ET,LT:string;
MPC:integer;
RDATE1:TDate;
begin
 MPC:=0;
 datenow:=FormatDateTime('YYYYMMDD',NOW);
 Finifile:=TReginiFile.Create(apptitle);
 NewDir:=FIniFile.ReadString('File Locations','OutgoingDflows','C:\');
 FInifile.Free;
// Newdir:=NEWDIR+'D0005\';
 with Generalquery do
 begin
  close;
  sql.clear;
  sql.add('select M.MPANCORE,M.Confirmed_DC_ID,M.Confirmed_DC_ROLE,A.READING_Date,A.appointment_slot,A.userid,A.requested_on,');
  sql.add('T.AM_start,T.AM_end,T.PM_start,T.PM_END,T.ALL_DAY_start,T.ALL_DAY_END');
  sql.add('from edmgr.mpan_status M, edmgr.D0005_requests A,edmgr.d0005_reading_times T');
  sql.add('Where M.MPANCORE=A.MPANCORE');
  sql.add('and M.Confirmed_DC_ID is not null');
  sql.add('and (M.Regstatus=''REGISTERED'' or M.REGSTATUS=''LOSS PENDING'')');
  sql.add('and M.MEASUREMENT_CLASS<''C''');
  sql.add('and M.confirmed_dc_id=T.DCID');
  sql.add('Order by M.Confirmed_DC_ID,M.Confirmed_DC_Role,M.MPANCORE,A.Reading_Date');
  open;
 end;
 if Generalquery.recordcount<>0 then
 Begin
  // Create a Log File
  Assignfile(V,newdir+'LOGS\LOGFILE_'+FormatDateTime('YYYYMMDDHHnnSS',NOW)+'.txt');
  rewrite(v);
  writeln(V,'D0005 Log. Created '+datetostr(Now));
  writeln(V,'Output files Created in '+h_outgoing);
  // Now Create Export File
  REASON:='Please Obtain Meter Reading';
  Agentrole:='';
  Agentid:='';
  FRM_File_Progress.progressbar.position:=0;
  while not Generalquery.eof do
  Begin
   MPAN:=GeneralQuery.fields[0].text;
   AGENTID:=GeneralQuery.fields[1].text;
   AGENTROLE:=GeneralQuery.Fields[2].text;
   RDATE1:=strtodate(GeneralQuery.fields[3].text);
   if (agentrole<>oldagentrole) or (agentid<>oldagentid) then
   begin
    ////////////////////////////////////////
    // Create Header Record
    ////////////////////////////////////////
    Writeln(V,'');
    Writeln(V,AgentID);
    Writeln(V,'MPAN,Reading Date,Appointment Time,Requested By, Requested on');
    CreateFlowHeader('D0005001',AgentID,Agentrole);
    OutputflowFlowcount:=0;
    OutputFlowLineCount:=0;
   end;
   oldagentrole:=agentrole;
   oldagentid:=agentid;
   if dayofweek(Rdate1)=1 then rdate1:=rdate1+1; // if request date is sunday then make it Monday
   if dayofweek(Rdate1)=7 then rdate1:=rdate1-1; // if request date is saturday then make it Friday
   // Now Obtain Reading Time For the DC
   ET:='08:00:00';
   LT:='17:00:00';
   if GeneralQuery.fields[4].text='AM' then ET:=GeneralQuery.fields[7].text;
   if GeneralQuery.fields[4].text='AM' then LT:=GeneralQuery.fields[8].text;
   if GeneralQuery.fields[4].text='PM' then ET:=GeneralQuery.fields[9].text;
   if GeneralQuery.fields[4].text='PM' then LT:=GeneralQuery.fields[10].text;
   if GeneralQuery.fields[4].text='DAY' then ET:=GeneralQuery.fields[11].text;
   if GeneralQuery.fields[4].text='DAY' then LT:=GeneralQuery.fields[12].text;
   // Create Body
   frm_main.WriteLinetoFile('017|'+MPAN+'|'+REASON+'|'+FormatDateTime('YYYYMMDD',RDATE1)+'|'+ET+'|'+LT+'|');
   frm_main.WriteLinetoFile('018|||');
   frm_main.WriteLinetoFile('019|01|');
   frm_main.WriteLinetoFile('020|'+FormatDateTime('YYYYMMDD',RDATE1)+'|');
   Writeln(V,GeneralQuery.fields[0].text+','+GeneralQuery.fields[3].text+','+GeneralQuery.fields[4].text+','+GeneralQuery.fields[5].text+','+GeneralQuery.fields[6].text);
   inc(MPC);
   inc(OutputFlowFlowCount);
   OutputflowLineCount:=OutputflowLineCount+4;
   GeneralQuery.next;
   AGENTID:=GeneralQuery.fields[1].text;
   AGENTROLE:=GeneralQuery.Fields[2].text;
   if (agentrole<>oldagentrole) or (oldagentid<>agentid) then cfooter:=true
   else cfooter:=false;
   if GeneralQuery.eof then cfooter:=true;
   if cfooter=true then CreateFlowFooter;
  end;
  Writeln(V,'');
  Writeln(V,'End of File: Number of MPANS = '+inttostr(MPC));
  closefile(V);
  with main_data_module.updatequery do
  begin
   close;
   sql.clear;
   sql.add('Truncate table edmgr.d0005_requests');
   execute;
  end;
 end;
end;

Function TFRM_Export.Readcommafield():string;
Var
Curfield:string;
x:string;
begin
 Curfield:='';
 repeat
  x:=s[Loop];
  if x<>',' then Curfield:=CurField+x;
  if x='''' then Curfield:=curfield+x;
  inc(loop);
 until x=',';
 Readcommafield:=curfield;
end;



procedure TFRM_Export.createD0297; // ONLY DOES SPOW HH DA
Var
BMUNIT,DAID,SSD,PESCODE:string;
sequencenumber,instructionnumber:integer;
Begin
 // Get List of MPANS
 with D0297list do
 begin
  close;
  open;
 end;
 if D0297list.recordcount=0 then exit;
 FRM_File_Progress.progressbar.position:=0;
 while not d0297list.eof do
 Begin
  // Create Header
  Createflowheader('D0297001',D0297list.fields[1].text,'A');
  // BODY
  // Get File Seq
  // Get instruction NO
  With SequenceQuery do
  begin
   close;
   sql.clear;
   sql.add('Select COUNTNUMBER from edmgr.file_seq');
   sql.add('Where FILEREF=''D0297FS''');
   open;
   SequenceNumber:=strtoint(Sequencequery.fields[0].text);
   inc(SequenceNumber);
   close;
   sql.clear;
   sql.add('Update Edmgr.File_Seq');
   sql.add('Set Countnumber='''+inttostr(sequencenumber)+'''');
   sql.add('Where FILEREF=''D0297FS''');
   execsql;
   FRM_login.mainsession.commit;
  end;
  s:='44C|'+inttostr(sequencenumber)+'|';
  frm_main.WriteLinetoFile(s);
  // Repeat for All MPANS
  With SequenceQuery do
  begin
   close;
   sql.clear;
   sql.add('Select COUNTNUMBER from edmgr.file_seq');
   sql.add('Where FILEREF=''D0297IN''');
   open;
   InstructionNumber:=strtoint(Sequencequery.fields[0].text);
   inc(InstructionNumber);
   close;
   sql.clear;
   sql.add('Update Edmgr.File_Seq');
   sql.add('Set Countnumber='''+inttostr(Instructionnumber)+'''');
   sql.add('Where FILEREF=''D0297IN''');
   execsql;
   FRM_LOGIN.mainsession.commit;
  end;
  with GeneralQuery do
  Begin
   close;
   sql.clear;
   sql.add('Select SSD,pes_code,Confirmed_da_ID from edmgr.mpan_status where mpancore='''+D0297list.fields[0].text+'''');
   open;
  end;
  DAID:=GeneralQuery.fields[2].text;
  ssd:=GeneralQuery.fields[0].text;
  MPAN:=D0297list.fields[0].text;
  pescode:=FRM_common.getgsp(mpan);
  if pescode='_A' then BMUNIT:='2__A'+X_MPID+'002';
  if pescode='_B' then BMUNIT:='2__B'+X_MPID+'002';
  if pescode='_C' then BMUNIT:='2__C'+X_MPID+'002';
  if pescode='_D' then BMUNIT:='2__D'+X_MPID+'002';
  if pescode='_E' then BMUNIT:='2__E'+X_MPID+'002';
  if pescode='_F' then BMUNIT:='2__F'+X_MPID+'002';
  if pescode='_G' then BMUNIT:='2__G'+X_MPID+'002';
  if pescode='_J' then BMUNIT:='2__J'+X_MPID+'002';
  if pescode='_H' then BMUNIT:='2__H'+X_MPID+'002';
  if pescode='_K' then BMUNIT:='2__K'+X_MPID+'002';
  if pescode='_L' then BMUNIT:='2__L'+X_MPID+'002';
  if pescode='_M' then BMUNIT:='2__M'+X_MPID+'002';
  // If SSD is in the Past or within the next 3 days then set date to NOW+3
  if (STRtoDate(SSD)<now+3) then SSD:=datetostr(now+3);
  s:='45C|'+inttostr(instructionnumber)+'|'+d0297list.fields[0].text+'|'+BMUNIT+'|'+FormatDateTime('YYYYMMDD',strtodate(SSD))+'|';
  frm_main.WriteLinetoFile(s);
  d0297list.next;
  inc(OutputflowLineCount);
  // Write Flow Footer
  inc(OutputflowLineCount);
  OutputflowFlowcount:=1;
  CreateFlowFooter;
  end;
End;

procedure TFRM_Export.createD0297toDefault; // ONLY DOES SPOW HH DA
Var
BMUNIT,DAID,PESCODE:string;
sequencenumber,instructionnumber:integer;
Begin
 // Get List of MPANS
 with D0297listdefault do
 begin
  close;
  open;
 end;
 if D0297listdefault.recordcount=0 then exit;
 screen.cursor:=crhourglass;
  // Repeat for All MPANS
 while not d0297listdefault.eof do
 Begin
 // Create Header
 Createflowheader('D0297001',D0297listdefault.fields[1].text,'A');
 // BODY
 // Get File Seq
 // Get instruction NO
 With SequenceQuery do
 begin
  close;
  sql.clear;
  sql.add('Select COUNTNUMBER from edmgr.file_seq');
  sql.add('Where FILEREF=''D0297FS''');
  open;
  SequenceNumber:=strtoint(Sequencequery.fields[0].text);
  inc(SequenceNumber);
  close;
  sql.clear;
  sql.add('Update Edmgr.File_Seq');
  sql.add('Set Countnumber='''+inttostr(sequencenumber)+'''');
  sql.add('Where FILEREF=''D0297FS''');
  execsql;
  FRM_LOGIN.mainsession.commit;
 end;
 s:='44C|'+inttostr(sequencenumber)+'|';
 frm_main.WriteLinetoFile(s);
 // Get First Instruction Number
 With SequenceQuery do
 begin
  close;
  sql.clear;
  sql.add('Select COUNTNUMBER from edmgr.file_seq');
  sql.add('Where FILEREF=''D0297IN''');
  open;
  InstructionNumber:=strtoint(Sequencequery.fields[0].text);
 end;

  inc(Instructionnumber);
  DAID:=D0297listdefault.fields[1].text;
  MPAN:=D0297listdefault.fields[0].text;
  pescode:=frm_common.getgsp(mpan);
  if pescode='_A' then BMUNIT:='2__A'+X_mpid+'000';
  if pescode='_B' then BMUNIT:='2__B'+X_mpid+'000';
  if pescode='_C' then BMUNIT:='2__C'+X_mpid+'000';
  if pescode='_D' then BMUNIT:='2__D'+X_mpid+'000';
  if pescode='_E' then BMUNIT:='2__E'+X_mpid+'000';
  if pescode='_F' then BMUNIT:='2__F'+X_mpid+'000';
  if pescode='_G' then BMUNIT:='2__G'+X_mpid+'000';
  if pescode='_J' then BMUNIT:='2__J'+X_mpid+'000';
  if pescode='_H' then BMUNIT:='2__H'+X_mpid+'000';
  if pescode='_K' then BMUNIT:='2__K'+X_mpid+'000';
  if pescode='_L' then BMUNIT:='2__L'+X_mpid+'000';
  if pescode='_M' then BMUNIT:='2__M'+X_mpid+'000';
  // If SSD is in the Past or within the next 3 days then set date to NOW+3
  s:='45C|'+inttostr(instructionnumber)+'|'+d0297listdefault.fields[0].text+'|'+BMUNIT+'|'+FormatDateTime('YYYYMMDD',strtodate('01/10/2002'))+'|';
  frm_main.WriteLinetoFile(s);
  application.processmessages;
  d0297listdefault.next;
  inc(OutputflowLineCount);
  // Write Flow Footer
  inc(OutputflowLineCount);
  OutputflowFlowcount:=1;
  CreateFlowFooter;
  With SequenceQuery do
  begin
   close;
   sql.clear;
   sql.add('Update Edmgr.File_Seq');
   sql.add('Set Countnumber='''+inttostr(Instructionnumber)+'''');
   sql.add('Where FILEREF=''D0297IN''');
   execsql;
   FRM_LOGIN.mainsession.commit;
  end;
 end;
 screen.cursor:=crdefault;
End;

procedure TFRM_Export.NewD0131s;
Var
PrevAgent,role:string;
openfile:boolean;
begin
 Prevagent:='';
 openfile:=false;
 while not agentsquery.eof do         //repeat until end of query
 begin
  toparty:=agentsquery.fields[1].text;
  role:=agentsquery.fields[2].text;
  if toparty<>prevagent then
  Begin   // Close the currently open file, as we are going to create a new one for the new agent/role;
   if openfile=true then CreateFlowFooter;
    // Create Header and File  for new agent/Role
   CreateFlowHeader('D0131001',ToParty,role);
   openfile:=true;
  end;
  // write body
  s:='253|'+agentsquery.fields[0].text+'||';
  frm_main.WriteLinetoFile(s);
  Inc(OutputflowLineCount);
  s:='71C|'+agentsquery.fields[3].text+'|'+agentsquery.fields[4].text+'|'+agentsquery.fields[5].text+'|'+
  agentsquery.fields[6].text+'|'+agentsquery.fields[7].text+'|'+agentsquery.fields[8].text+'|'+
  agentsquery.fields[9].text+'|'+agentsquery.fields[10].text+'|'+agentsquery.fields[11].text+'|'+agentsquery.fields[12].text+'|||';
  frm_main.WriteLinetoFile(s);
  INC(OutputflowLineCount);
  INC(OutputFlowFlowCount);
  Agentsquery.next;
  prevagent:=toparty;
  application.processmessages;
 end;
 // Write Flow Footer and Close the Last Open FIle
 CreateFlowFooter;
end;

procedure TFRM_Export.D0131NHHDC;
begin
 // Do NHH DC records
 with agentsquery do
 Begin
  close;
  sql.clear;
  sql.add('Select M.MPANCORE,M.CONFIRMED_DC_ID, M.CONFIRMED_DC_ROLE,');
  sql.add('A.METERING_POINT_ADDRESS1,');
  sql.add('A.METERING_POINT_ADDRESS2,');
  sql.add('A.METERING_POINT_ADDRESS3,');
  sql.add('A.METERING_POINT_ADDRESS4,');
  sql.add('A.METERING_POINT_ADDRESS5,');
  sql.add('A.METERING_POINT_ADDRESS6,');
  sql.add('A.METERING_POINT_ADDRESS7,');
  sql.add('A.METERING_POINT_ADDRESS8,');
  sql.add('A.METERING_POINT_ADDRESS9,');
  sql.add('A.METERING_POINT_POSTCODE');
  SQL.ADD('FROM EDMGR.MPAN_STATUS M, EDMGR.MPAS_CURRENT_ADDR A');
  sql.add('Where M.MPANCORE=A.MPANCORE');
  sql.add('and M.Confirmed_dc_role=''D''');
  sql.add('and (M.REGSTATUS=''REGISTERED''');
  sql.add('or M.REGSTATUS=''FUTURE LOSS''');
  sql.add('or M.REGSTATUS=''LOSS PENDING'')');
  sql.add('and D0131_dc is null');
  sql.add('order by m.confirmed_dc_id, m.mpancore');
  open
 end;
 if agentsquery.recordcount<>0 then NewD0131s;
end;

procedure TFRM_Export.D0131NHHDC_Addresschange;
begin
 // Do NHH DC records
 with agentsquery do
 Begin
  close;
  sql.clear;
  sql.add('Select M.MPANCORE,M.CONFIRMED_DC_ID, M.CONFIRMED_DC_ROLE,');
  sql.add('A.METERING_POINT_ADDRESS1,');
  sql.add('A.METERING_POINT_ADDRESS2,');
  sql.add('A.METERING_POINT_ADDRESS3,');
  sql.add('A.METERING_POINT_ADDRESS4,');
  sql.add('A.METERING_POINT_ADDRESS5,');
  sql.add('A.METERING_POINT_ADDRESS6,');
  sql.add('A.METERING_POINT_ADDRESS7,');
  sql.add('A.METERING_POINT_ADDRESS8,');
  sql.add('A.METERING_POINT_ADDRESS9,');
  sql.add('A.METERING_POINT_POSTCODE');
  SQL.ADD('FROM EDMGR.MPAN_STATUS M, EDMGR.MPAS_CURRENT_ADDR A');
  sql.add('Where M.MPANCORE=A.MPANCORE');
  sql.add('and M.Confirmed_dc_role=''D''');
  sql.add('and (M.REGSTATUS=''REGISTERED''');
  sql.add('or M.REGSTATUS=''FUTURE LOSS''');
  sql.add('or M.REGSTATUS=''LOSS PENDING'')');
  sql.add('and D0131_dc=to_date(''01/01/1950'',''DD/MM/YYYY'')');
  sql.add('order by m.confirmed_dc_id, m.mpancore');
  open
 end;
 if agentsquery.recordcount<>0 then NewD0131s;
end;

procedure TFRM_Export.D0131HHDC;
begin
 // Do HH DC records
 with agentsquery do
 Begin
  close;
  sql.clear;
  sql.add('Select M.MPANCORE,M.CONFIRMED_DC_ID, M.CONFIRMED_DC_ROLE,');
  sql.add('A.METERING_POINT_ADDRESS1,');
  sql.add('A.METERING_POINT_ADDRESS2,');
  sql.add('A.METERING_POINT_ADDRESS3,');
  sql.add('A.METERING_POINT_ADDRESS4,');
  sql.add('A.METERING_POINT_ADDRESS5,');
  sql.add('A.METERING_POINT_ADDRESS6,');
  sql.add('A.METERING_POINT_ADDRESS7,');
  sql.add('A.METERING_POINT_ADDRESS8,');
  sql.add('A.METERING_POINT_ADDRESS9,');
  sql.add('A.METERING_POINT_POSTCODE');
  SQL.ADD('FROM EDMGR.MPAN_STATUS M, EDMGR.MPAS_CURRENT_ADDR A');
  sql.add('Where M.MPANCORE=A.MPANCORE');
  sql.add('and M.Confirmed_dc_role=''C''');
  sql.add('and (M.REGSTATUS=''REGISTERED''');
  sql.add('or M.REGSTATUS=''FUTURE LOSS''');
  sql.add('or M.REGSTATUS=''LOSS PENDING'')');
  sql.add('and D0131_dc is null');
  sql.add('order by m.confirmed_dc_id, m.mpancore');
  open
 end;
 if agentsquery.recordcount<>0 then NewD0131s;
end;

procedure TFRM_Export.D0131HHDC_Addresschange;
begin
 // Do HH DC records
 with agentsquery do
 Begin
  close;
  sql.clear;
  sql.add('Select M.MPANCORE,M.CONFIRMED_DC_ID, M.CONFIRMED_DC_ROLE,');
  sql.add('A.METERING_POINT_ADDRESS1,');
  sql.add('A.METERING_POINT_ADDRESS2,');
  sql.add('A.METERING_POINT_ADDRESS3,');
  sql.add('A.METERING_POINT_ADDRESS4,');
  sql.add('A.METERING_POINT_ADDRESS5,');
  sql.add('A.METERING_POINT_ADDRESS6,');
  sql.add('A.METERING_POINT_ADDRESS7,');
  sql.add('A.METERING_POINT_ADDRESS8,');
  sql.add('A.METERING_POINT_ADDRESS9,');
  sql.add('A.METERING_POINT_POSTCODE');
  SQL.ADD('FROM EDMGR.MPAN_STATUS M, EDMGR.MPAS_CURRENT_ADDR A');
  sql.add('Where M.MPANCORE=A.MPANCORE');
  sql.add('and M.Confirmed_dc_role=''C''');
  sql.add('and (M.REGSTATUS=''REGISTERED''');
  sql.add('or M.REGSTATUS=''FUTURE LOSS''');
  sql.add('or M.REGSTATUS=''LOSS PENDING'')');
  sql.add('and D0131_dc=to_date(''01/01/1950'',''DD/MM/YYYY'')');
  sql.add('order by m.confirmed_dc_id, m.mpancore');
  open
 end;
 if agentsquery.recordcount<>0 then NewD0131s;
end;

procedure TFRM_Export.D0131MO;
begin
 // Do MOP records
 with agentsquery do
 Begin
  close;
  sql.clear;
  sql.add('Select M.MPANCORE,M.CONFIRMED_MO_ID, M.CONFIRMED_MO_ROLE,');
  sql.add('A.METERING_POINT_ADDRESS1,');
  sql.add('A.METERING_POINT_ADDRESS2,');
  sql.add('A.METERING_POINT_ADDRESS3,');
  sql.add('A.METERING_POINT_ADDRESS4,');
  sql.add('A.METERING_POINT_ADDRESS5,');
  sql.add('A.METERING_POINT_ADDRESS6,');
  sql.add('A.METERING_POINT_ADDRESS7,');
  sql.add('A.METERING_POINT_ADDRESS8,');
  sql.add('A.METERING_POINT_ADDRESS9,');
  sql.add('A.METERING_POINT_POSTCODE');
  SQL.ADD('FROM EDMGR.MPAN_STATUS M, EDMGR.MPAS_CURRENT_ADDR A');
  sql.add('Where M.MPANCORE=A.MPANCORE');
  sql.add('and M.Confirmed_MO_role=''M''');
  sql.add('and (M.REGSTATUS=''REGISTERED''');
  sql.add('or M.REGSTATUS=''FUTURE LOSS''');
  sql.add('or M.REGSTATUS=''LOSS PENDING'')');
  sql.add('and D0131_MO is null');
  sql.add('order by m.confirmed_MO_id, m.mpancore');
  open
 end;
 if agentsquery.recordcount<>0 then NewD0131s;
end;

procedure TFRM_Export.D0131MO_addresschange;
begin
 // Do MOP records
 with agentsquery do
 Begin
  close;
  sql.clear;
  sql.add('Select M.MPANCORE,M.CONFIRMED_MO_ID, M.CONFIRMED_MO_ROLE,');
  sql.add('A.METERING_POINT_ADDRESS1,');
  sql.add('A.METERING_POINT_ADDRESS2,');
  sql.add('A.METERING_POINT_ADDRESS3,');
  sql.add('A.METERING_POINT_ADDRESS4,');
  sql.add('A.METERING_POINT_ADDRESS5,');
  sql.add('A.METERING_POINT_ADDRESS6,');
  sql.add('A.METERING_POINT_ADDRESS7,');
  sql.add('A.METERING_POINT_ADDRESS8,');
  sql.add('A.METERING_POINT_ADDRESS9,');
  sql.add('A.METERING_POINT_POSTCODE');
  SQL.ADD('FROM EDMGR.MPAN_STATUS M, EDMGR.MPAS_CURRENT_ADDR A');
  sql.add('Where M.MPANCORE=A.MPANCORE');
  sql.add('and M.Confirmed_MO_role=''M''');
  sql.add('and (M.REGSTATUS=''REGISTERED''');
  sql.add('or M.REGSTATUS=''FUTURE LOSS''');
  sql.add('or M.REGSTATUS=''LOSS PENDING'')');
  sql.add('and D0131_MO=to_date(''01/01/1950'',''DD/MM/YYYY'')');
  sql.add('order by m.confirmed_MO_id, m.mpancore');
  open
 end;
 if agentsquery.recordcount<>0 then NewD0131s;
end;

procedure TFRM_Export.NHHDAD0153;
begin
  with MiscQuery do
  begin
   close;
   sql.clear;
   sql.add('select * from EDMGR.EXPORT_D0153_NHHDA order by Requested_DA_ID,MPANCORE');
   open;
   if MiscQuery.RecordCount > 0 then
   begin
     First;
   end;
  end;
  if MiscQuery.RecordCount > 0 then
  begin
    CreateAppointments('DA','B',MiscQuery.fields[18].text,'N','',Miscquery.fields[74].text,'');
  end;
end;

procedure TFRM_Export.NHHDAD0153COA;
var
efd:string;
begin
 with Coa do
 Begin
  close;
  sql.clear;
  sql.add('select * from EDMGR.EXPORT_D0153_NHHDA_COA');
  sql.add('order by 3,2,1');
  open;
 end;
 if coa.recordcount=0 then exit;

 while not coa.eof do
 Begin
  with MiscQuery do
  begin
   close;
   sql.clear;
   sql.add('select * from edmgr.mpan_status');
   sql.add('where (mpancore) in');
   sql.add('(select mpancore from edmgr.flowheaders where toid=''B'' and toname='''+coa.fields[1].text+'''');
   sql.add('and filename='''+coa.fields[0].text+''' and from_name=''REQU'')');
   open;
   if MiscQuery.RecordCount > 0 then
   begin
     First;
   end;
  end;
    if MiscQuery.RecordCount > 0 then
    begin
      efd:=copy(coa.fields[0].text,1,2)+'/'+copy(coa.fields[0].text,3,2)+'/'+copy(coa.fields[0].text,5,4);
      CreateAppointments('DA','B',coa.fields[1].text,'N',EFD,Miscquery.fields[74].text,'');
    end;

  // Update to Show Sent
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('update edmgr.flowheaders');
   sql.add('set from_name='''+X_MPID+'''');
   sql.add('where from_name=''REQU''');
   sql.add('and toid=''B'' and toname='''+coa.Fields[1].text+'''');
   sql.add('and filename='''+coa.fields[0].text+'''');
   execute;
  End;
  coa.next;
 end;
 frm_login.mainsession.commit;

end;

procedure TFRM_Export.MOD0155COA;
var
efd:string;
begin
 with Coa do
 Begin
  close;
  sql.clear;
  sql.add('select * from EDMGR.EXPORT_D0155_MO_COA');
  sql.add('order by 3,2,1');
  open;
 end;
 if coa.recordcount=0 then exit;

 while not coa.eof do
 Begin
  with MiscQuery do
  begin
   close;
   sql.clear;
   sql.add('select * from edmgr.mpan_status');
   sql.add('where (mpancore) in');
   sql.add('(select mpancore from edmgr.flowheaders where toid=''M'' and toname='''+coa.fields[1].text+'''');
   sql.add('and filename='''+coa.fields[0].text+''' and from_name=''REQU'')');
   open;
   if MiscQuery.RecordCount > 0 then
   begin
     First;
   end;
  end;
  if MiscQuery.RecordCount > 0 then
  begin
    efd:=copy(coa.fields[0].text,1,2)+'/'+copy(coa.fields[0].text,3,2)+'/'+copy(coa.fields[0].text,5,4);
    CreateAppointments('MO','M',coa.fields[1].text,'N',EFD,Miscquery.fields[74].text,'');
  end;
  // Update to Show Sent
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('update edmgr.flowheaders');
   sql.add('set from_name='''+X_MPID+'''');
   sql.add('where from_name=''REQU''');
   sql.add('and toid=''M'' and toname='''+coa.Fields[1].text+'''');
   sql.add('and filename='''+coa.fields[0].text+'''');
   execute;
  End;
  coa.next;
 end;
 frm_login.mainsession.commit;
end;


procedure TFRM_Export.NHHDCD0155;
begin
  with MiscQuery do
  begin
   close;
   sql.clear;
   sql.add('select * from EDMGR.EXPORT_D0155_NHHDC order by Requested_DC_ID,MPANCORE');
   open;
   if MiscQuery.RecordCount > 0 then
   begin
     First;
   end;
  end;
 //11.6  CreateAppointments('DC','D',MiscQuery.fields[18].text,'N','',Miscquery.fields[74].text,MiscQuery.fields[84].text);
  if MiscQuery.RecordCount > 0 then
  begin
    CreateAppointments('DC','D',MiscQuery.fields[18].text,'N','',Miscquery.fields[74].text,'N');
  end;
end;

procedure TFRM_Export.NHHDCD0155COA;
var
efd:string;
begin
 with Coa do
 Begin
  close;
  sql.clear;
  sql.add('select * from EDMGR.EXPORT_D0155_NHHDC_COA order by 3,2,1');
  open;
 end;
 if coa.recordcount=0 then exit;

 while not coa.eof do
 Begin
  with MiscQuery do
  begin
   close;
   sql.clear;
    sql.add('select * from edmgr.mpan_status');
   sql.add('where (mpancore) in');
   sql.add('(select mpancore from edmgr.flowheaders where toid=''D'' and toname='''+coa.fields[1].text+'''');
   sql.add('and filename='''+coa.fields[0].text+''' and from_name=''REQU'')');
   open;

   if MiscQuery.RecordCount > 0 then
   begin
     First;
   end;
  end;

  if MiscQuery.RecordCount > 0 then
  begin
    efd:=copy(coa.fields[0].text,1,2)+'/'+copy(coa.fields[0].text,3,2)+'/'+copy(coa.fields[0].text,5,4);
    //11.6 CreateAppointments('DC','D',coa.fields[1].text,'N',EFD,Miscquery.fields[74].text,MiscQuery.fields[84].text);
    CreateAppointments('DC','D',coa.fields[1].text,'N',EFD,Miscquery.fields[74].text,'N');
  end;

  // Update to Show Sent
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('update edmgr.flowheaders');
   sql.add('set from_name='''+X_MPID+'''');
   sql.add('where from_name=''REQU''');
   sql.add('and toid=''D'' and toname='''+coa.Fields[1].text+'''');
   sql.add('and filename='''+coa.fields[0].text+'''');
   execute;
  End;
  coa.next;
 end;
 frm_login.mainsession.commit;

end;

procedure TFRM_Export.NHHMOD0155;
begin
  with MiscQuery do
  begin
   close;
   sql.clear;
    sql.add('select * from EDMGR.EXPORT_D0155_MO order by Requested_MO_ID,MPANCORE');
   open;
   if MiscQuery.RecordCount > 0 then
   begin
     First;
   end;
  end;

  if MiscQuery.RecordCount > 0 then
  begin
    CreateAppointments('MO','M',MiscQuery.fields[18].text,'N','',Miscquery.fields[74].text,'');
  end;
end;

procedure TFRM_Export.HHDAD0153;
begin
  with MiscQuery do
  begin
   close;
   sql.clear;
   sql.add('select * from EDMGR.EXPORT_D0153_HHDA order by Requested_DA_ID,MPANCORE');
   open;

   if MiscQuery.RecordCount > 0 then
   begin
     First;
   end;
  end;

  if MiscQuery.RecordCount > 0 then
  begin
    CreateAppointments('DA','A',MiscQuery.fields[18].text,'H','',Miscquery.fields[74].text,'');
  end;
end;

procedure TFRM_Export.HHDCD0155;
begin
  with MiscQuery do
  begin
   close;
   sql.clear;
   sql.add('select * from EDMGR.EXPORT_D0155_HHDC order by Requested_DC_ID,MPANCORE');
   Open;
   if MiscQuery.RecordCount > 0 then
   begin
     First;
   end;
  end;
  //11.6 CreateAppointments('DC','C',MiscQuery.fields[18].text,'H','',Miscquery.fields[74].text,MiscQuery.fields[84].text);
  if MiscQuery.RecordCount > 0 then
  begin
    CreateAppointments('DC','C',MiscQuery.fields[18].text,'H','',Miscquery.fields[74].text,'N');
  end;
end;

procedure TFRM_Export.HHMOD0155;
begin
  with MiscQuery do
  begin
   close;
   sql.clear;
   sql.add('select * from EDMGR.EXPORT_D0155_HHMO order by Requested_MO_ID,MPANCORE');
   Open;
   if MiscQuery.RecordCount > 0 then
   begin
     First;
   end;
  end;

  if MiscQuery.RecordCount > 0 then
  begin
    CreateAppointments('MO','M',MiscQuery.fields[18].text,'H','',Miscquery.fields[74].text,'');
  end;
end;

procedure TFRM_Export.TerminateMOPLoss;
begin
 // Get List of MPANS that are LOST or FUTURE LOSS and MO needs Terminating
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('select * from edmgr.export_d0151_mop_loss order by confirmed_MO_ID,Confirmed_MO_Role,EFTSSD,MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('LC');
 end;
end;

procedure TFRM_Export.TerminateDCLoss;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('select * from edmgr.export_d0151_dc_loss order by confirmed_dc_ID,Confirmed_dc_Role,EFTSSD,MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('LC');
 end;
end;

procedure TFRM_Export.TerminateDALoss;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('select * from edmgr.export_d0151_da_loss order by confirmed_da_ID,Confirmed_da_Role,EFTSSD,MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('LC');
 end;
end;

procedure TFRM_Export.TerminateMOLossObj;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('SELECT * from edmgr.EXPORT_D0151_MOP_OBJ ORDER BY Confirmed_MO_id,MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('OB');
 end;
end;

procedure TFRM_Export.TerminateMOPSMI;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('SELECT MPANCORE,SSD-1,Confirmed_MO_ID, Confirmed_MO_Role,SSD');
  sql.add('FROM EDMGR.MPAN_STATUS MPAN_STATUS WHERE regstatus like ''PSM%''');
  sql.add('and confirmed_MO_id =''UMOL'' and D0151_MO is null ORDER BY Confirmed_MO_id,MPANCORE');  // only for UMOL
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('OB');
 end;
end;

procedure TFRM_Export.TerminateMOPWRONGSSD;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('SELECT E.MPANCORE,M.SSD-1,E.Confirmed_MO_ID, E.Confirmed_MO_Role,M.SSD');
  sql.add('FROM EDMGR.MPAN_STATUS E, MOPMGR.MPAN_STATUS M WHERE');
  sql.add('E.CONFIRMED_MO_ID is NOT NULL and E.MPANCORE=M.MPANCORE and E.SSD<>M.SSD and M.TERMINATION_REASON is NULL and m.ssd<sysdate');
  sql.add('ORDER BY E.Confirmed_MO_id,E.MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('OB');
 end;
end;

procedure TFRM_Export.TerminateDCLossObj;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('SELECT * from edmgr.EXPORT_D0151_DC_OBJ ORDER BY Confirmed_dc_id, confirmed_dc_role,MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('OB');
 end;
end;

procedure TFRM_Export.Terminate_ACCU_DC_WRONG_APP_DATE;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('select a.mpancore,a.eff_from_date-1,''ACCU'',''D'',a.efsd from edmgr.d0011 a,');
  sql.add('(select mpancore,max(filename) filename,efsd,max(eff_from_date) efd');
  sql.add('from');
  sql.add('edmgr.d0011');
  sql.add('where');
  sql.add('contract_ref=''ACCUGETWDC''');
  sql.add('group by mpancore,efsd) m');
  sql.add('where');
  sql.add('a.CONTRACT_REF = ''ACCUGETWDC''');
  sql.add('and a.eff_from_date<>a.efsd');
  sql.add('and a.mpancore=m.mpancore');
  sql.add('and a.filename=m.filename');
  sql.add('and a.efsd=m.efsd');
  sql.add('and a.eff_from_date=m.efd');
  sql.add('and a.efsd>to_date(''01/01/2009'',''DD/MM/YYYY'')');
  sql.add('order by 5,1');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('OB');
 end;
end;

procedure TFRM_Export.TerminateDALossObj;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('SELECT * from edmgr.EXPORT_D0151_DA_OBJ ORDER BY Confirmed_da_id,MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('OB');
 end;
end;


procedure TFRM_Export.TerminateDCLossCOA;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('SELECT MPANCORE,''27/11/2005'',Confirmed_MO_ID, Confirmed_DC_Role,SSD');
  sql.add('FROM EDMGR.MPAN_STATUS MPAN_STATUS');
  sql.add('where confirmed_dc_id=''ACCU''');
  sql.add('and (pes_code=''12''');
  sql.add('or pes_code=''19''');
  sql.add('or pes_code=''22'')');
  sql.add('ORDER BY Confirmed_mo_id, confirmed_dc_role,MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('CA');
 end;
end;

procedure TFRM_Export.TerminateDALossCOA;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('SELECT MPANCORE,''27/11/2005'',Confirmed_MO_ID, Confirmed_DA_Role,SSD');
  sql.add('FROM EDMGR.MPAN_STATUS MPAN_STATUS');
  sql.add('where confirmed_da_id=''ACCU''');
  sql.add('and (pes_code=''12''');
  sql.add('or pes_code=''19''');
  sql.add('or pes_code=''22'')');
  sql.add('ORDER BY Confirmed_mo_id, confirmed_da_role,MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('CA');
 end;
end;

procedure TFRM_Export.Createflowheader(FlowVersion,Rec_MPID,Rec_Role:string; bResetProgressBar: boolean = True);
begin
 FRM_File_Progress.setforExport(bResetProgressBar);
 frm_busy.Close;
 // Write header record
 OutputFlowIdentifier:=frm_common.Getnextfileid;
 OUTPUTfilename:=h_Outgoing+Copy(FlowVersion,1,5)+'_'+REC_MPID+'_'+REC_role+'_'+Outputflowidentifier+'.usr';
 S:='ZHV|';
 s:=s+Outputflowidentifier+'|';                // File ID
 S:=s+FLOWVERSION+'|';
 s:=s+X_MPIDROLE+'|';                          // FromRole
 s:=s+X_MPID+'|';                              // From ID
 s:=s+REC_ROLE+'|';                            // Recipient Role
 if H_Mode='TEST' then s:=s+H_REC+'|'        // Recipient Name
 else s:=s+REC_MPID+ '|';
 s:=s+formatdatetime('YYYYMMDDHHNNSS',now)+'|';// Create Date
 s:=s+H_APP+'|';                               // Application generation Flow
 s:=s+'|';
 s:=s+'|';
 s:=s+H_TESTFLAG+'|';                          // Live/Test Flag
//DISABLE Assignfile(Outputflow,OUTPUTFILENAME);
//DISABLE rewrite(Outputflow);
 FRM_File_Progress.setforExport(bResetProgressBar);
 frm_main.FileMEMO.Lines.Clear;
 frm_main.WriteLinetoFile(s);
 OutputflowLineCount:=0;
 OutputFlowFlowCount:=0;
 with FRM_File_Progress do
 Begin
  if bResetProgressBar then
  begin
    fileprogressbar.position:=0;
    fileprogressbar.max:=100;
  end;
  Caption:='Generating '+copy(flowversion,1,5)+'''s';
  l_filename.caption:=outputflowidentifier+'.usr';
  l_fileid.caption:=outputflowidentifier;
  l_flowversion.caption:=flowversion;
  l_FromRole.caption:=X_mpidrole;
  l_FromMPID.caption:=X_mpid;
  l_ToRole.caption:=rec_role;
  l_ToMPID.caption:=h_rec;
  l_filedatetime.caption:=formatdatetime('YYYYMMDDHHNNSS',now);

  if H_Mode='TEST' then l_tompid.caption:=H_REC
  else l_tompid.caption:=rec_mpid;

  l_testflag.caption:=h_testflag;

  l_testflag.caption:=h_testflag;
  application.processmessages;
 end;
end;

procedure TFRM_Export.CreateflowheaderMOP(FlowVersion,Rec_MPID,Rec_Role:string);
var
 bypassfolder:string;
begin
 FRM_File_Progress.setforExport;
 frm_busy.Close;
 // Write header record
 OutputFlowIdentifier:=frm_common.nextfileidMOP;
 if flowversion='D0313002' then
 begin
   bypassfolder:=frm_common.GETVALUE('FILE_ELEC_OUT_D0313');
   OUTPUTfilename:=bypassfolder+'R313'+Outputflowidentifier+'.usr';
 end
 else
 OUTPUTfilename:=h_Outgoing+Copy(FlowVersion,1,5)+'_'+REC_MPID+'_'+REC_role+'_'+Outputflowidentifier+'.usr';

 S:='ZHV|';
 s:=s+Outputflowidentifier+'|';                // File ID
 S:=s+FLOWVERSION+'|';
 s:=s+M_MPIDROLE+'|';                          // FromRole
 s:=s+M_MPID+'|';                              // From ID
 s:=s+REC_ROLE+'|';                            // Recipient Role
 if H_Mode='TEST' then s:=s+H_REC+'|'        // Recipient Name
 else s:=s+REC_MPID+ '|';
 s:=s+formatdatetime('YYYYMMDDHHNNSS',now)+'|';// Create Date
 s:=s+H_APP+'|';                               // Application generation Flow
 s:=s+'|';
 s:=s+'|';
 s:=s+H_TESTFLAG+'|';                          // Live/Test Flag

 //Assignfile(Outputflow,OUTPUTFILENAME);
 //rewrite(Outputflow);
 FRM_File_Progress.setforExport;
 frm_main.FileMEMO.Lines.Clear;

 frm_main.WriteLinetoFile(s);
 OutputflowLineCount:=0;
 OutputFlowFlowCount:=0;
 with FRM_File_Progress do
 Begin
  fileprogressbar.position:=0;
  fileprogressbar.max:=100;
  Caption:='Generating '+copy(flowversion,1,5)+'''s';
  l_filename.caption:=outputflowidentifier+'.usr';
  l_fileid.caption:=outputflowidentifier;
  l_flowversion.caption:=flowversion;
  l_FromRole.caption:=M_mpidrole;
  l_FromMPID.caption:=M_mpid;
  l_ToRole.caption:=rec_role;
  l_ToMPID.caption:=h_rec;
  l_filedatetime.caption:=formatdatetime('YYYYMMDDHHNNSS',now);

  if H_Mode='TEST' then l_tompid.caption:=H_REC
  else l_tompid.caption:=rec_mpid;

  l_testflag.caption:=h_testflag;

  l_testflag.caption:=h_testflag;
  application.processmessages;
 end;
end;


procedure TFRM_Export.Createflowfooter(bResetProgressBar: boolean = True);
begin
 s:='ZPT|';
 s:=s+OutPutFlowIdentifier+'|';
 s:=s+inttostr(frm_main.FileMEMO.Lines.count-1)+'|';
 s:=s+'|';
 s:=s+inttostr(OutputFlowFlowCount)+'|';
 s:=s+formatdatetime('YYYYMMDDHHNNSS',now)+'|';
 frm_main.WriteLinetoFile(s);
 //closefile(OutputFlow);
 frm_main.FileMEMO.Lines.SaveToFile(OutputFilename);
 if frm_main.FileMEMO.Lines.count<3 then deletefile(OutputFilename);

 FRM_File_Progress.clearlabels(bResetProgressBar);

 if bResetProgressBar then
 begin
   FRM_File_Progress.close;
 end;
end;

procedure TFRM_Export.CreateD0010(filter:string);
var
prevdc,prevmpan,prevmeter,prevregister,prevtype:string;
openfile:boolean;
datenow,newdir,prevdate:string;
G:Textfile;
rdate:TDateTime;
begin
 if filter='' then
 Begin
  With GeneralQuery do
  Begin
   close;
   sql.clear;
   sql.add('select * from edmgr.export_d0010');
   sql.add('order by Confirmed_DC_ID,MPANCORE,METERID,RDNGTYPE,REGISTERID,READDATE');
   open;
  end;
 end
 else
 Begin
  With GeneralQuery do
  Begin
   close;
   sql.clear;
   sql.add('select DISTINCT FLOWVERSION,MPANCORE,METERID,REGISTERID,REGISTERREADING,READDATE,RDNGTYPE,REQUESTOR,');
   sql.add('SENT_STATUS,dc_ID from edmgr.readings_to_send');
   sql.add('where ');
   sql.add('SENT_STATUS=''R'' and requestor=''SYSTEM'' and dc_id=''UDMS''');
   sql.add('order by DC_ID,MPANCORE,METERID,RDNGTYPE,REGISTERID,READDATE');
   open;
  end;
 end;
 PrevDC:='';
 Prevmpan:='';
 openfile:=false;
 if GeneralQuery.recordcount<>0 then
 Begin
 datenow:=FormatDateTime('YYYYMMDD',NOW);
 Finifile:=TReginiFile.Create(apptitle);
 NewDir:=FIniFile.ReadString('File Locations','OutgoingDflows','C:\');
 FInifile.Free;
// Newdir:=NEWDIR+'D0010\';
 try
   Tdirectory.CreateDirectory(newdir+'LOGS\');
 except
 end;
 Assignfile(G,newdir+'LOGS\LOGFILE_D0010_'+FormatDateTime('YYYYMMDDHHnnSS',NOW)+'.txt');
 Rewrite(g);
 writeln(G,'DFLOW,MPANCORE,METERID,REGISTER,REGISTER READING,READ DATE,READ TYPE,REQUESTOR,SENT STATUS,DC ID');
 writeln(g,'');
 While not GeneralQuery.eof do
 Begin
  rdate:=strtodate(copy(GeneralQuery.fields[5].text,1,10));
  if GeneralQuery.fields[9].text<>prevdc then
  begin
   // Close any open files
   if openfile=true then CreateFlowFooter;
   // Create Header for New File
   // Get Curent Date
   openfile:=true;
   TOPARTY:=GeneralQuery.fields[9].text;
   CreateFlowHeader('D0010002',ToParty,'D');
   OutputFlowLinecount:=0;
   OutputFlowFlowcount:=0;
  end;
  // Group 026
  if GeneralQuery.fields[1].text<>prevmpan then
  Begin
   frm_main.WriteLinetoFile('026|'+GeneralQuery.fields[1].text+'|U|');
   prevmpan:=GeneralQuery.fields[1].text;
   prevmeter:='';
   prevtype:='';
   inc(OutputFlowlinecount);
   inc(OutputFlowFlowcount);
  end;
  // Group 028
  if (GeneralQuery.fields[2].text<>prevmeter) or (GeneralQuery.fields[6].text<>PrevType) then
  Begin
   frm_main.WriteLinetoFile('028|'+GeneralQuery.fields[2].text+'|'+GeneralQuery.fields[6].text+'|');
   prevmeter:=GeneralQuery.fields[2].text;
   prevtype:=GeneralQuery.fields[6].text;
   prevregister:='';
   prevdate:='';
   inc(OutputFlowlinecount);
  end;
  // Group 030  on change of Register
  if GeneralQuery.fields[3].text<>prevregister then
  Begin
   frm_main.WriteLinetoFile('030|'+GeneralQuery.fields[3].text+'|'+FormatDateTime('YYYYMMDD',rdate)+'000000|'+GeneralQuery.fields[4].text+'.0|||T|'+Generalquery.fields[12].text+'|');
   //D%TC 11.1


   prevregister:=GeneralQuery.fields[3].text;
   prevdate:=datetostr(rdate);
   inc(OutputFlowlinecount);
   // write log record
   writeln(g,GeneralQuery.fields[0].text+','+GeneralQuery.fields[1].text+','+GeneralQuery.fields[2].text+','+
   GeneralQuery.fields[3].text+','+GeneralQuery.fields[4].text+','+GeneralQuery.fields[5].text+','+
   GeneralQuery.fields[6].text+','+GeneralQuery.fields[7].text+','+GeneralQuery.fields[8].text+','+GeneralQuery.fields[9].text);
  end;

  // if same mpan,meter,register and date is different.
  if (GeneralQuery.fields[1].text=prevmpan) and
     (GeneralQuery.fields[2].text=prevmeter) and
     (GeneralQuery.fields[3].text=prevregister) and
     (datetostr(rdate)<>prevdate) then
  Begin
   frm_main.WriteLinetoFile('030|'+GeneralQuery.fields[3].text+'|'+FormatDateTime('YYYYMMDD',rdate)+'000000|'+GeneralQuery.fields[4].text+'.0|||T|'+Generalquery.fields[12].text+'|');
   prevregister:=GeneralQuery.fields[3].text;
   inc(OutputFlowlinecount);
   prevdate:=datetostr(rdate);
   // write log record
   writeln(g,GeneralQuery.fields[0].text+','+GeneralQuery.fields[1].text+','+GeneralQuery.fields[2].text+','+
   GeneralQuery.fields[3].text+','+GeneralQuery.fields[4].text+','+GeneralQuery.fields[5].text+','+
   GeneralQuery.fields[6].text+','+GeneralQuery.fields[7].text+','+GeneralQuery.fields[8].text+','+GeneralQuery.fields[9].text);
  end;

  prevdc:=GeneralQuery.fields[9].text;
  GeneralQuery.next;
 end;
 writeln(G,'End Of File');
 closefile(g);

 // Write Flow Footer
 CreateFlowFooter;
 end;
end;

procedure TFRM_Export.CreateD0010fromD0188;
var
prevdc,prevmpan,prevmeter,prevregister,prevtype:string;
openfile:boolean;
datenow,newdir,RD:string;
begin
 exit; // No longer Required as D0188 reads are unreliable and can spanner the eacs.
 With GeneralQuery do
 Begin
  close;
  sql.clear;
  sql.add('select ');
  sql.add('distinct ');
  sql.add('p.mpancore, ');
  sql.add('p.reading_date_time, ');
  sql.add('p.meter_id, ');
  sql.add('p.meter_register_id,');
  sql.add('p.register_reading,');
  sql.add('m.confirmed_dc_id ');
  sql.add('from edmgr.d0188 p,');
  sql.add('edmgr.mpan_status m,');
  sql.add('(select mpancore,filename from edmgr.flowheaders where flow_version=''D0188'' and from_name<>''ACTA'') F, ');
  sql.add('(select distinct mpancore,meterid,meter_register_id from edmgr.mtds) mr');
  sql.add('where p.mpancore=m.mpancore ');
  sql.add('and m.confirmed_dc_id is not null ');
  sql.add('and p.reading_date_time>to_date(''01/02/2007'',''DD/MM/YYYY'')');
  sql.add('and p.date_sent_to_dc is null');
  sql.add('and p.mpancore=f.mpancore');
  sql.add('and p.filename=f.filename');
  sql.add('and p.mpancore=mr.mpancore');
  sql.add('and p.meter_id=mr.meterid');
  sql.add('and p.meter_register_id=mr.meter_register_id');
  sql.add('order by 6,1,2,3,4');
  open;
 end;
 PrevDC:='';
 Prevmpan:='';
 openfile:=false;
 if GeneralQuery.recordcount<>0 then
 Begin
 datenow:=FormatDateTime('YYYYMMDD',NOW);
 Finifile:=TReginiFile.Create(apptitle);
 NewDir:=FIniFile.ReadString('File Locations','OutgoingDflows','C:\');
 FInifile.Free;
 While not GeneralQuery.eof do
 Begin
  if GeneralQuery.fields[5].text<>prevdc then
  begin
   // Close any open files
   if openfile=true then CreateFlowFooter;
   // Create Header for New File
   // Get Curent Date
   openfile:=true;
   TOPARTY:=GeneralQuery.fields[5].text;
   CreateFlowHeader('D0010002',ToParty,'D');
   OutputFlowLinecount:=0;
   OutputFlowFlowcount:=0;
  end;
  // Group 026
  if GeneralQuery.fields[0].text<>prevmpan then
  Begin
   frm_main.WriteLinetoFile('026|'+GeneralQuery.fields[0].text+'|U|');
   prevmpan:=GeneralQuery.fields[0].text;
   prevmeter:='';
   prevtype:='';
   inc(OutputFlowlinecount);
   frm_main.WriteLinetoFile('027|03|READINGS TAKEN FROM PPMIP D0188 FLOW|');
   inc(OutputFlowlinecount);
   inc(OutputFlowFlowcount);
  end;
  // Group 028
  if (GeneralQuery.fields[2].text<>prevmeter) then
  Begin
   frm_main.WriteLinetoFile('028|'+GeneralQuery.fields[2].text+'|C|');
   prevmeter:=GeneralQuery.fields[2].text;
   prevregister:='';
   inc(OutputFlowlinecount);
  end;
  // Group 030
  if length(generalquery.Fields[1].text)<>10 then rd:=FormatDateTime('YYYYMMDDhhnnss',strtodatetime(GeneralQuery.fields[1].text))+'|'
  else
  rd:=FormatDateTime('YYYYMMDD',strtodate(GeneralQuery.fields[1].text))+'000000|';

  frm_main.WriteLinetoFile('030|'+GeneralQuery.fields[3].text+'|'+RD+GeneralQuery.fields[4].text+'.0|||T|N|');
  prevregister:=GeneralQuery.fields[3].text;
  inc(OutputFlowlinecount);

  prevdc:=GeneralQuery.fields[5].text;
  GeneralQuery.next;
 end;
 // Write Flow Footer
 CreateFlowFooter;
 end;
end;

procedure TFRM_Export.CreateD0010fromDials;
var
prevdc,prevmpan,prevmeter,prevregister,prevtype:string;
openfile:boolean;
datenow,newdir,RD:string;
begin
 With GeneralQuery do
 Begin
  close;
  sql.clear;
  sql.add('select * from edmgr.latest_dialled_read_to_dc where confirmed_dc_id in (''UDMS'',''LBSL'')');
  sql.add('order by 6,1,2,3,4');
  open;
 end;
 PrevDC:='';
 Prevmpan:='';
 openfile:=false;
 if GeneralQuery.recordcount<>0 then
 Begin
 datenow:=FormatDateTime('YYYYMMDD',NOW);
 Finifile:=TReginiFile.Create(apptitle);
 NewDir:=FIniFile.ReadString('File Locations','OutgoingDflows','C:\');
 FInifile.Free;
 While not GeneralQuery.eof do
 Begin
  if GeneralQuery.fields[5].text<>prevdc then
  begin
   // Close any open files
   if openfile=true then CreateFlowFooter;
   // Create Header for New File
   // Get Curent Date
   openfile:=true;
   TOPARTY:=GeneralQuery.fields[5].text;
   CreateFlowHeader('D0010002',ToParty,'D');
   OutputFlowLinecount:=0;
   OutputFlowFlowcount:=0;
  end;
  // Group 026
  if GeneralQuery.fields[0].text<>prevmpan then
  Begin
   frm_main.WriteLinetoFile('026|'+GeneralQuery.fields[0].text+'|U|');
   prevmpan:=GeneralQuery.fields[0].text;
   prevmeter:='';
   prevtype:='';
   inc(OutputFlowlinecount);
   inc(OutputFlowFlowcount);
  end;
  // Group 028
  if (GeneralQuery.fields[2].text<>prevmeter) then
  Begin
   frm_main.WriteLinetoFile('028|'+GeneralQuery.fields[2].text+'|P|');
   prevmeter:=GeneralQuery.fields[2].text;
   prevregister:='';
   inc(OutputFlowlinecount);
  end;
  // Group 030
  if length(generalquery.Fields[1].text)<>10 then rd:=FormatDateTime('YYYYMMDDhhnnss',strtodatetime(GeneralQuery.fields[1].text))+'|'
  else
  rd:=FormatDateTime('YYYYMMDD',strtodate(GeneralQuery.fields[1].text))+'000000|';

  frm_main.WriteLinetoFile('030|'+GeneralQuery.fields[3].text+'|'+RD+GeneralQuery.fields[4].text+'.0|||T|N|');
  prevregister:=GeneralQuery.fields[3].text;
  inc(OutputFlowlinecount);

  prevdc:=GeneralQuery.fields[5].text;
  GeneralQuery.next;
 end;
 // Write Flow Footer
 CreateFlowFooter;
 end;
end;


procedure TFRM_Export.CreateD0071;
var
prevdc,prevmpan,prevmeter,prevregister:string;
openfile:boolean;
datenow,newdir:string;
G:Textfile;
rdate:tdatetime;
begin
 With GeneralQuery do
 Begin
  close;
  sql.clear;
  sql.add('select * from edmgr.export_d0071');
  sql.add('order by confirmed_DC_ID,MPANCORE,METERID,REGISTERID,READDATE');
  open;
 end;
 PrevDC:='';
 Prevmpan:='';
 openfile:=false;
 if GeneralQuery.recordcount<>0 then
 Begin
 datenow:=FormatDateTime('YYYYMMDD',NOW);
 Finifile:=TReginiFile.Create(apptitle);
 NewDir:=FIniFile.ReadString('File Locations','OutgoingDflows','C:\');
 FInifile.Free;
 try
  Tdirectory.CreateDirectory(newdir+'LOGS\');
 except
 end;
 Assignfile(G,newdir+'LOGS\LOGFILE_D0071_'+FormatDateTime('YYYYMMDDHHnnSS',NOW)+'.txt');
 Rewrite(g);
 writeln(G,'DFLOW,MPANCORE,METERID,REGISTER,REGISTER READING,READ DATE,READ TYPE,REQUESTOR,SENT STATUS,DC ID');
 writeln(g,'');
 While not GeneralQuery.eof do
 Begin
  rdate:=strtodate(copy(GeneralQuery.fields[5].text,1,10));
  if GeneralQuery.fields[9].text<>prevdc then
  begin
   // Close any open files
   if openfile=true then CreateFlowFooter;
   // Create Header for New File
   // Get Curent Date
   openfile:=true;
   TOPARTY:=GeneralQuery.fields[9].text;
   CreateFlowHeader('D0071001',ToParty,'D');
   OutputFlowLinecount:=0;
   OutputFlowFlowcount:=0;
  end;
  // Group 026
  if GeneralQuery.fields[1].text<>prevmpan then
  Begin
   frm_main.WriteLinetoFile('147|'+GeneralQuery.fields[1].text+'|'+FormatDateTime('YYYYMMDD',rdate)+'000000|'+GeneralQuery.fields[6].text+'|');
   prevmpan:=GeneralQuery.fields[1].text;
   prevmeter:='';
   inc(OutputFlowlinecount);
   inc(OutputFlowFlowcount);
  end;
  // Group 028
  if GeneralQuery.fields[2].text<>prevmeter then
  Begin
   frm_main.WriteLinetoFile('148|'+GeneralQuery.fields[2].text+'|');
   prevmeter:=GeneralQuery.fields[2].text;
   prevregister:='';
   inc(OutputFlowlinecount);
  end;
  // Group 030
  if GeneralQuery.fields[3].text<>prevregister then
  Begin
   frm_main.WriteLinetoFile('149|'+GeneralQuery.fields[3].text+'|'+GeneralQuery.fields[4].text+'.0|');
   prevregister:=GeneralQuery.fields[3].text;
   inc(OutputFlowlinecount);
   // write log record
   writeln(g,'D0071'+','+GeneralQuery.fields[1].text+','+GeneralQuery.fields[2].text+','+
   GeneralQuery.fields[3].text+','+GeneralQuery.fields[4].text+','+GeneralQuery.fields[5].text+','+
   GeneralQuery.fields[6].text+','+GeneralQuery.fields[7].text+','+GeneralQuery.fields[8].text+','+GeneralQuery.fields[9].text);
  end;
  prevdc:=GeneralQuery.fields[9].text;
  GeneralQuery.next;
 end;
 writeln(G,'End Of File');
 closefile(g);
 // Write Flow Footer
 CreateFlowFooter;
 end;
 // End of D0010 Generation
end;

Procedure TFRM_Export.CreateD0005toMOP(AGENTID,MPAN,Reason,ReasonCode:string;ADATE:Tdatetime);
begin
 CreateFlowHeader('D0005001',AgentID,'M');
 OutputflowFlowcount:=0;
 OutputFlowLineCount:=0;
 frm_main.WriteLinetoFile('017|'+MPAN+'|'+REASON+'||||');
 frm_main.WriteLinetoFile('018|||');
 frm_main.WriteLinetoFile('019|'+reasoncode+'|');
 if reasoncode<>'02' then frm_main.WriteLinetoFile('020|'+FormatDateTime('YYYYMMDD',ADATE)+'|')
  else frm_main.WriteLinetoFile('021'+FormatDateTime('YYYYMMDD',ADATE)+'|');
 inc(OutputFlowFlowCount);
 OutputflowLineCount:=OutputflowLineCount+4;
 CreateFlowFooter;
end;

Procedure TFRM_Export.IdentifyD0302(AgentRole,MPANS,GROUPS:String);
Var
Openfile,populateme: Boolean;
LastAgent, NewAgent:String;
Cefsd:string;
Begin
 // NHH DC D0302
 if agentrole='C' then
 Begin
  with D0302Query do
  begin
   close;
   sql.clear;
   sql.add('select mpancore,ssd,confirmed_dc_id,new_connection,cot,comc from edmgr.mpan_status');
   sql.add('where regstatus = ''REGISTERED''');
   sql.add('and confirmed_dc_id is not null');
   sql.add('and confirmed_dc_role=''C''');
   if MPANS='ALL' then sql.add('and D0302_dc is null');
   if MPANS<>'ALL' then sql.add('and mpancore='''+mpans+'''');
   sql.add('order by confirmed_dc_id,MPANCORE');
   open;
   First;
  end;
 end;
 // HH DC D0302
 if agentrole='D' then
 Begin
  with D0302Query do
  begin
   close;
   sql.clear;
   sql.add('select mpancore,ssd,confirmed_dc_id,new_connection,cot,comc from edmgr.mpan_status');
   sql.add('where regstatus = ''REGISTERED''');
   sql.add('and confirmed_dc_id is not null');
   sql.add('and confirmed_dc_role=''D''');
   if MPANS='ALL' then sql.add('and D0302_dc is null');
   if MPANS<>'ALL' then sql.add('and mpancore='''+mpans+'''');
   sql.add('order by confirmed_dc_id,MPANCORE');
   open;
   First;
  end;
 end;
 // MOP D0302
 if agentrole='M' then
 Begin
  with D0302Query do
  begin
   close;
   sql.clear;
   sql.add('select mpancore,ssd,confirmed_mo_id,new_connection,cot,comc from edmgr.mpan_status');
   sql.add('where regstatus = ''REGISTERED''');
   sql.add('and confirmed_MO_id is not null');
   sql.add('and confirmed_MO_role=''M''');
   if MPANS='ALL' then sql.add('and D0302_MO is null');
   if MPANS<>'ALL' then sql.add('and mpancore='''+mpans+'''');
   sql.add('order by confirmed_mo_id,MPANCORE');
   open;
   First;
  end;
 end;
 // MOP D0302  Smart Meters UMOL
 if agentrole='S' then
 Begin
  with D0302Query do
  begin
   close;
   sql.clear;
   sql.add('select mpancore,ssd,confirmed_mo_id,new_connection,cot,comc from edmgr.mpan_status');
   sql.add('where regstatus like ''PSM%''');
   sql.add('and confirmed_MO_id =''UMOL''');
   sql.add('and confirmed_MO_role=''M''');
   if MPANS='ALL' then sql.add('and D0302_MO is null');
   if MPANS<>'ALL' then sql.add('and mpancore='''+mpans+'''');
   sql.add('order by confirmed_mo_id,MPANCORE');
   open;
   First;
  end;
 end;
 // D0302 Dist flows can go once agent appointments have been sent.
 if agentrole='R' then
 Begin
  with D0302Query do
  begin
   close;
   sql.clear;
   sql.add('select M.mpancore,M.ssd,A.agent_name,M.new_connection,M.cot,M.comc from edmgr.mpan_status M, edmgr.agent_reference_gsp_group A');
   sql.add('where M.regstatus = ''REGISTERED''');
   sql.add('and M.requested_MO_id is not null');
   if MPANS='ALL' then sql.add('and M.D0302_DIST is null');
   if MPANS<>'ALL' then sql.add('and mpancore='''+mpans+'''');
   sql.add('and a.pes_area=m.pes_code');
   sql.add('order by A.agent_name,M.MPANCORE');
   open;
   First;
  end;
 end;

 if agentrole='S' then agentrole:='M';
 while not D0302query.eof do
 Begin
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('Insert into edmgr.batch_flows_for_sending values(');
   sql.add('''D0302'','''+d0302query.fields[2].text+''','''+agentrole+''','''+d0302query.fields[0].text+''',');
   sql.add('to_date('''+d0302query.fields[1].text+''',''DD/MM/YYYY''))');
   try
    execute;
   except
   end;
  end;
  D0302query.next;
 End;
 FRM_Login.mainsession.commit;
end;

procedure TFRM_Export.TerminateMOPDisc;
begin
 // Get List of MPANS that are LOST or FUTURE LOSS and MO needs Terminating
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('select MPANCORE,EFTSSD,Confirmed_MO_ID, Confirmed_MO_Role,SSD from edmgr.MPAN_STATUS where');
  sql.add('(D0151_MO is null and confirmed_MO_ID is not null)');
  sql.add('and (regstatus=''DISCONNECTED'')');
  sql.add('order by confirmed_MO_ID,Confirmed_MO_Role,EFTSSD,MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('DE');
 end;
end;

procedure TFRM_Export.TerminateDCDisc;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('select MPANCORE,EFTSSD,Confirmed_DC_ID, Confirmed_DC_Role,SSD from edmgr.MPAN_STATUS where');
  sql.add('(D0151_DC is null and confirmed_DC_ID is not null)');
  sql.add('and (regstatus=''DISCONNECTED'')');
  sql.add('order by confirmed_DC_ID,Confirmed_DC_Role,EFTSSD,MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('DE');
 end;
end;

procedure TFRM_Export.TerminateDADisc;
begin
 with GeneralQuery do
 begin
  close;
  sql.clear;
  sql.add('select MPANCORE,EFTSSD,Confirmed_DA_ID, Confirmed_DA_Role,SSD from edmgr.MPAN_STATUS where');
  sql.add('(D0151_DA is null and confirmed_DA_ID is not null)');
  sql.add('and (regstatus=''DISCONNECTED'')');
  sql.add('order by confirmed_DA_ID,Confirmed_DA_Role,EFTSSD,MPANCORE');
  open;
  oldagentid:='';
  oldagentrole:='';
  CreateD0151flow('DE');
 end;
end;

procedure TFRM_Export.CreateSingleD0183(rcode:string);
Var
MPAN,toMPID:String;
begin
{ MPAN:=D0183form.mpanlookup.text; // Get MPAN
 if rcode='V' then ToMPID:=getPPMIP(mpan); // Get PPMIP
 if rcode='M' then ToMPID:=getMOP(mpan);   // Get MOP
 CreateFlowHeader('D0183001',toMPID,rcode);
 OutputFlowFlowCount:=1;
 OutputFlowLineCount:=0;
 // Now write Group 376
  frm_main.WriteLinetoFile('376|'+D0183form.m4.text+'|'+MPAN+'|'+D0183form.cn1.text+'|'+D0183form.s1.text+'|'+D0183form.s2.text+'|'+D0183form.s3.text+'|'+D0183form.s4.text+
 '|'+D0183form.s5.text+'|'+D0183form.s6.text+'|'+D0183form.s7.text+'|'+D0183form.s8.text+'|'+D0183form.s9.text+'|'+D0183form.sp.text+
 '|'+D0183form.mname.text+'|'+D0183form.a1.text+'|'+D0183form.a2.text+'|'+D0183form.a3.text+'|'+D0183form.a4.text+
 '|'+D0183form.a5.text+'|'+D0183form.a6.text+'|'+D0183form.a7.text+'|'+D0183form.a8.text+'|'+D0183form.a9.text+'|'+D0183form.ap.text+'|'+formatfloat('0.00',strtofloat(D0183form.m6.text))+'|'+D0183form.m2.text+'|');
 inc(OutputFlowLineCount);
 CreateFlowFooter;}
end;

procedure TFRM_Export.DoDailyExports;
Begin
 Do_E_SupplierFiles;
 updatestatusbar('','');
 Do_E_MOPFiles;
 updatestatusbar('','');
 Do_G_SUPPLIER_Files;
 updatestatusbar('','');
 Do_G_SHIPPER_Files;
 updatestatusbar('','');
end;



procedure TFRM_Export.QueryAgents(MPAN:String);
Var
NoRecords:boolean;
Begin
 requestedmo.close;
 requestedmo.setvariable('MPAN',MPAN);
 requestedmo.open;
 requestedmo.last;

 moquery.close;
 moquery.setvariable('MPAN',MPAN);
 moquery.open;
 moquery.last;

 moterm.close;
 moterm.setvariable('MPAN',MPAN);
 moterm.open;
 moterm.last;

 dc1.close;
 dc1.setvariable('MPAN',MPAN);
 dc1.open;
 dc1.last;

 dc2.close;
 dc2.setvariable('MPAN',MPAN);
 dc2.open;
 dc2.last;

 dc3.close;
 dc3.setvariable('MPAN',MPAN);
 dc3.open;
 dc3.last;

 da1.close;
 da1.setvariable('MPAN',MPAN);
 da1.open;
 da1.last;

 da2.close;
 da2.setvariable('MPAN',MPAN);
 da2.open;
 da2.last;

 da3.close;
 da3.setvariable('MPAN',MPAN);
 da3.open;
 da3.last;

end;

procedure TFRM_Export.CreateOutstandingD0302s;
Var
Openfile,populateme: Boolean;
LastAgent, NewAgent,lastrole,newrole:String;
Cefsd,tempfield:string;
begin

 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select recipient,recipient_role,mpancore,trunc(SSD) SSD from edmgr.batch_flows_for_sending');
  sql.add('where Dflow=''D0302''');
  sql.add('order by recipient,recipient_role,mpancore');
  open;
 end;
 frm_login.mainsession.commit;
 FRM_File_Progress.progressbar.position:=0;
 FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
 // ResetFileCounters etc
 OpenFile := False;
 LastAGENT:='non';
 LastRole:='non';
 // Loop Through All MPANS

 while not GeneralQuery.eof do
 begin

  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+Generalquery.fields[2].text;
  //FRM_Main.statusbar.panels[1].text:='Creating D0302 to '+generalquery.fields[1].text+' for MPAN '+Generalquery.fields[2].text;
  application.processmessages;

  NewAgent:=GeneralQuery.fields[0].text;
  NewRole:= GeneralQuery.fields[1].text;

  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) or (NewRole<>LastRole) then
  Begin
   lastagent:=NewAgent;
   lastrole:=newrole;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeader('D0302002',NewAgent,NewRole);
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
  End;

  With CustQuery do
  Begin
   Close;
   SetVariable('MPAN',GeneralQuery.fields[2].text);
   open;
  end;

  if custquery.fields[1].text<>'' then
  Begin
   // Write Group 68c
   S:= '68C'+'|';
   S:=S+GeneralQuery.fields[2].text+'|'; // MPAN
   cEFSD:=copy(GeneralQuery.fields[3].text,1,10);     // SSD
   S:=S+formatdatetime('YYYYMMDD',strtodate(cEFSD))+ '|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCount);
   s:='69C|';

   TempField:=frm_common.StripNonDtc(custquery.fields[1].text);

   s:=s+TempField+'|';
   s:=s+'|';   // Additonal Info
   TempField:=stringreplace(custquery.fields[4].text,'`','''',[rfreplaceall]);
   s:=s+tempfield+'|'; // Password

   if custquery.fields[5].text<>'' then
   S:=S+formatdatetime('YYYYMMDD',strtodate(custquery.fields[5].text))+ '|'
   else s:=s+custquery.fields[5].text+'|';   // Password Date

   s:=s+custquery.fields[20].text+'|';  // Special Access
   //DTC 11.2
   // TempField:=copy(stringreplace(custquery.fields[21].text,'`','''',[rfreplaceall]),1,50);
   // s:=s+tempfield+'|';
   // s:=s+copy(custquery.fields[22].text,1,14)+'|';
   // s:=s+custquery.fields[23].text+'|';
   s:=s+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCount);

   s:='15J|';
   TempField:=copy(stringreplace(custquery.fields[21].text,'`','''',[rfreplaceall]),1,50);
   s:=s+tempfield+'||';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCount);

   if CustQuery.FieldByName('TELEPHONE_NO_MOBILE').Text <> '' then
   begin
     s := '16J|';
     s := s + copy(CustQuery.FieldByName('TELEPHONE_NO_MOBILE').Text,1,14) + '|';
     s := s + CustQuery.FieldByName('FAX_1').Text + '|';
     Frm_Main.WriteLinetoFile(s);
     Inc(OutputFlowLineCount);
   end
   else if custquery.FieldByName('TELEPHONE_NO_DAY_1').text<>'' then
   begin
    s := '16J|';
    s := s + copy(CustQuery.FieldByName('TELEPHONE_NO_DAY_1').text,1,14)+'|';
    s := s + CustQuery.FieldByName('FAX_1').text+'|';
    frm_main.WriteLinetoFile(s);
    Inc(OutputFlowLineCount);
   end
   else if CustQuery.FieldByName('TELEPHONE_NO_DAY').Text <> '' then
   begin
     s := '16J|';
     s := s + copy(CustQuery.FieldByName('TELEPHONE_NO_DAY').Text,1,14) + '|';
     s := s + CustQuery.FieldByName('FAX_1').Text + '|';
     Frm_Main.WriteLinetoFile(s);
     Inc(OutputFlowLineCount);
   end;

   s:='70C|';
   s:=s+''+'|';
    TempField:=stringreplace(custquery.fields[6].text,'`','''',[rfreplaceall]);
   s:=s+tempfield+'|';
    TempField:=stringreplace(custquery.fields[7].text,'`','''',[rfreplaceall]);
   s:=s+tempfield+'|';
    TempField:=stringreplace(custquery.fields[8].text,'`','''',[rfreplaceall]);
   s:=s+tempfield+'|';
    TempField:=stringreplace(custquery.fields[9].text,'`','''',[rfreplaceall]);
   s:=s+tempfield+'|';
    TempField:=stringreplace(custquery.fields[10].text,'`','''',[rfreplaceall]);
   s:=s+tempfield+'|';
    TempField:=stringreplace(custquery.fields[11].text,'`','''',[rfreplaceall]);
   s:=s+tempfield+'|';
    TempField:=stringreplace(custquery.fields[12].text,'`','''',[rfreplaceall]);
   s:=s+tempfield+'|';
    TempField:=stringreplace(custquery.fields[13].text,'`','''',[rfreplaceall]);
   s:=s+tempfield+'|';
    TempField:=stringreplace(custquery.fields[14].text,'`','''',[rfreplaceall]);
   s:=s+tempfield+'|';
    TempField:=stringreplace(custquery.fields[15].text,'`','''',[rfreplaceall]);
   s:=s+tempfield+'|';
   frm_main.WriteLinetoFile(S);
   Inc(OutputFlowLineCount);
   Inc(OutputflowFlowCount);
  end;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
 DeleteBatches('D0302');
end;

procedure TFRM_Export.CreateOutstandingD0225s;
Var
Openfile,populateme: Boolean;
LastAgent, NewAgent,lastrole,newrole,ctel,cteltype,alt:String;
consenttosharing,expirydate,deleteall,Cefsd,tempfield,specneeds,contact,specneedscode,dtccode,actionindicator,additionalinfo:string;
begin
 FRM_File_Progress.progressbar.position:=0;
 try
   gSqlUtil.ExecProc('EDMGR.PR_CHECK_D0225', TRANSACTION_NO);
 Except
   on e:Exception do
     gLogger.Log('Error with EDMGR.PR_CHECK_D0225: %s', [e.Message], llError);
 end;
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select * from edmgr.export_d0225');
  sql.add('order by recipient,recipient_role,mpancore');
  open;
 end;
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
 // ResetFileCounters etc
 OpenFile := False;
 LastAGENT:='non';
 LastRole:='non';
 // Loop Through All MPANS

 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+Generalquery.fields[2].text;
  //FRM_Main.statusbar.panels[1].text:='Creating D0225 to '+generalquery.fields[1].text+' for MPAN '+Generalquery.fields[2].text;
  application.processmessages;

  NewAgent:=GeneralQuery.fields[0].text;
  NewRole:= GeneralQuery.fields[1].text;

  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) or (NewRole<>LastRole) then
  Begin
   lastagent:=NewAgent;
   lastrole:=newrole;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeader('D0225002',NewAgent,NewRole);
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
  End;

  s:= GeneralQuery.FieldByName('GROUP_510').Text;
  inc(OutputFlowLineCount);
  s:=frm_common.RemoveDodgyChars(s);
  frm_main.WriteLinetoFile(s);

  if GeneralQuery.FieldByName('GROUP_99C').Text <> '' then
  begin
    s:= GeneralQuery.FieldByName('GROUP_99C').Text;
    inc(OutputFlowLineCount);
    s:=frm_common.RemoveDodgyChars(s);
    frm_main.WriteLinetoFile(s);
  end;
  inc(OutputFlowFlowcount);
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
 DeleteBatches('D0225');
 DeleteBlankSpecialNeeds;
end;


procedure TFRM_Export.Deletebatches(Dflow:string);
begin
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from edmgr.batch_flows_for_sending');
  sql.add('where dflow='''+dflow+'''');
  execute;
  FRM_Login.mainsession.commit;
 end;
end;

procedure TFRM_Export.DeleteBlankSpecialNeeds;
begin
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from edmgr.special_needs');
  sql.add('where additional_info='''' or additional_info is null');
  execute;
  FRM_Login.mainsession.commit;
 end;
end;


procedure TFRM_Export.CreateMOPD0011s;
var
  Openfile, populateme: boolean;
  LastAgent, NewAgent, lastrole, newrole, Cefsd, tempfield, specneeds, contact,
  specneedscode, dtccode, actionindicator, additionalinfo: string;
  qryTemp: TOracleDataSet;
begin
  FRM_File_Progress.progressbar.position := 0;
  try
    qryTemp := gSqlUtil.CreateCursor
      ('MOPMGR.PK_MOP_APPOINTMENT.PR_GET_D0155_ACCEPTED(:p_dataset)',
      TRANSACTION_NO, [':p_dataset', otCursor, null]);
    try
      // ResetFileCounters etc
      Openfile := False;
      LastAgent := 'non';
      lastrole := 'non';

      while not qryTemp.eof do
      begin
        // get all the details for the selected MPAN
        MPAN := qryTemp.FieldByName('mpancore').AsString;

        FRM_File_Progress.d_file.caption := '';
        FRM_File_Progress.labelcount.caption := '';
        FRM_File_Progress.progressbar.position :=
          FRM_File_Progress.progressbar.position + 1;
        FRM_File_Progress.statusbar.panels[0].text := 'MPAN: ' + MPAN;
        FRM_File_Progress.statusbar.Update;

        NewAgent := qryTemp.FieldByName('supplier_mpid').AsString;

        // Check if Agent is different to Last Agent
        if (NewAgent <> LastAgent) then
        begin
          LastAgent := NewAgent;
          // Close any files that may be open
          if Openfile = True then
            CreateFlowFooter;
          // Now Create New File and Write Header Record
          CreateFlowHeaderMOP('D0011001', NewAgent, 'X');
          // Indicate There is an Open File
          OutputFlowFlowCount := 0;
          OutPutFlowLineCount := 0;
          Openfile := True;
        end;

        /// ////////////////////////////////////////////////////////////////////////////
        s := '034|';
        s := s + MPAN + '|';
        s := s + qryTemp.FieldByName('contract_reference').AsString + '|';
        s := s + Formatdatetime('YYYYMMDD', qryTemp.FieldByName('efsd')
          .AsDateTime) + '|';
        inc(OutPutFlowLineCount);
        frm_main.WriteLinetoFile(s);

        /// ////////////////////////////////////////////////////////////////////////////
        s := '036|';
        s := s + Formatdatetime('YYYYMMDD', qryTemp.FieldByName('efsd_agent')
          .AsDateTime) + '|';
        inc(OutPutFlowLineCount);
        frm_main.WriteLinetoFile(s);

        /// ////////////////////////////////////////////////////////////////////////////
        s := '038|';
        s := s + qryTemp.FieldByName('service_reference').AsString + '|';
        s := s + qryTemp.FieldByName('service_level_ref').AsString + '|';
        inc(OutputFlowFlowCount);
        inc(OutPutFlowLineCount);
        frm_main.WriteLinetoFile(s);
        LastAgent := NewAgent;
        qryTemp.next;
      end;
    finally
      FreeAndNil(qryTemp);
    end;

    // Close the Last File That may be Open
    if Openfile = True then
      CreateFlowFooter;

    try
      gSqlUtil.ExecProc
        ('MOPMGR.PK_MOP_APPOINTMENT.PR_UPDATE_D0155_RESPONSE_DATE(:p_flow_version)',
        TRANSACTION_YES, ['p_flow_version', otString, pdInput, FlowVersion]);
    except
      on E: EOracleError do
        if E.ErrorCode <> 1 then
          FRM_common.DisplayOracleError(E.Message, FlowVersion, 'Q', Filename);
    end;

  except
    on E: EOracleError do
    begin
      MessageDlg(E.Message, mtError, [mbOk], 0);
      exit;
    end;
  end;
end;

procedure TFRM_Export.CreateMOPD0261s;
var
  Openfile, populateme: boolean;
  LastAgent, NewAgent, lastrole, newrole: string;
  Cefsd, tempfield, specneeds, contact, specneedscode, dtccode, actionindicator,
  additionalinfo: string;
  p_dataset: Variant;
  qryTemp: TOracleDataSet;
begin
  FRM_File_Progress.ProgressBar.Position := 0;
  qryTemp := TOracleDataSet.Create(nil);
  try
    try
      qryTemp := gSqlUtil.CreateCursor
        ('MOPMGR.PK_MOP_APPOINTMENT.PR_GET_D0155_REJECTED(:p_dataset)',
        TRANSACTION_NO, [':p_dataset', otCursor, null]);
    except
      on E: EOracleError do
      begin
        MessageDlg(E.message, mtError, [mbOK], 0);
        exit;
      end;
    end;

    // ResetFileCounters etc
    Openfile := False;
    LastAgent := 'non';
    lastrole := 'non';
    FRM_File_Progress.ProgressBar.Max := qryTemp.RecordCount;

    with qryTemp do
    begin
      // Loop Through All MPANS
      while not Eof do
      begin
        // get all the details for the selected MPAN
        MPAN := FieldByName('mpancore').AsString;

        with FRM_File_Progress do
        begin
          D_File.Caption := '';
          LabelCount.Caption := '';
          ProgressBar.Position :=
          ProgressBar.Position + 1;
          Statusbar.Panels[0].Text := 'MPAN: ' + FieldByName('mpancore').AsString;
        end;
        // FRM_Main.statusbar.panels[1].text:='Creating D0261 to '+qryTemp.fields[0].text+' for MPAN '+qryTemp.fields[1].text;
        Application.ProcessMessages;

        NewAgent := FieldByName('supplier_mpid').AsString;

        // Check if Agent is different to Last Agent
        if (NewAgent <> LastAgent) then
        begin
          LastAgent := NewAgent;
          // Close any files that may be open
          if Openfile then
            CreateFlowFooter;
          // Now Create New File and Write Header Record
          CreateFlowHeaderMOP('D0261001', NewAgent, 'X');
          // Indicate There is an Open File
          OutputFlowFlowCount := 0;
          OutPutFlowLineCount := 0;
          Openfile := True;
        end;

        //////////////////////////////////////////////////////////////////////////////
        s := '761|';
        s := s + MPAN + '|';
        s := s + FieldByName('contract_reference').AsString + '|';
        s := s + Formatdatetime('YYYYMMDD', FieldByName('efsd').AsDateTime) + '|';
        s := s + FieldByName('response_type').AsString + '|';

        Inc(OutputFlowFlowCount);
        Inc(OutPutFlowLineCount);
        FRM_Main.WriteLinetoFile(s);
        LastAgent := NewAgent;
        Next;
      end; // Repeat for remaining MPANS
    end;
    // Close the Last File That may be Open
    if Openfile then
      CreateFlowFooter;

    try
      gSqlUtil.ExecProc('MOPMGR.PK_MOP_APPOINTMENT.PR_UPDATE_D0155_RESPONSE_DATE(:p_flow_version)',
                        TRANSACTION_YES, ['p_flow_version' , otString, pdInput,  FlowVersion]);
    except
      on E: EOracleError do
        if E.ErrorCode <> 1 then
          FRM_Common.DisplayOracleError(E.Message, FlowVersion, 'Q', filename);
    end;
  finally
    qryTemp.Free;
  end;
end;

procedure TFRM_Export.CreateMOPD0170s;
Var
Openfile,populateme: Boolean;
LastAgent, NewAgent,lastrole,newrole,darb,olddarb:String;
Cefsd,tempfield,specneeds,contact,specneedscode,dtccode,actionindicator,additionalinfo:string;
begin
 FRM_File_Progress.progressbar.position:=0;

 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select distinct D0148_OLD_MO,SSD,MPANCORE');
  sql.add('from mopmgr.MPAN_STATUS');
  sql.add('where D0170_SENT_OLD_MO is NULL and D0148_OLD_MO is NOT NULL and D0148_OLD_MO_TYPE=''O''');
  sql.add('and D0150_RECD_OLD_MO is null');
  sql.add('order by D0148_OLD_MO,SSD,MPANCORE');
  open;
 end;

 // ResetFileCounters etc
 OpenFile := False;
 LastAGENT:='non';
 LastRole:='non';
 olddarb:='non';
 // Loop Through All MPANS

 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+Generalquery.fields[2].text;
  //FRM_Main.statusbar.panels[1].text:='Creating D0170 to '+generalquery.fields[0].text+' for MPAN '+Generalquery.fields[2].text;
  application.processmessages;

  NewAgent:=GeneralQuery.fields[0].text;

  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) then
  Begin
   lastagent:=NewAgent;
   olddarb:='non';
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0170001',NewAgent,'M');
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
  End;
   // get all the details for the selected MPAN
  mpan:=generalquery.fields[2].text;
  DARB:=formatdatetime('YYYYMMDD',strtodate(generalquery.fields[1].text));
  //////////////////////////////////////////////////////////////////////////////
  if darb<>olddarb then
  Begin
   s:='350|'+DARB+'|06|PLEASE FORWARD MTDS|'+M_mpid+'||';
   inc(OutputFlowFlowcount);
   inc(OutputFlowLineCount);
   frm_main.WriteLinetoFile(s);
  end;
  s:='351|'+MPAN+'|';
  inc(OutputFlowLineCount);
  frm_main.WriteLinetoFile(s);
  olddarb:=darb;
  lastagent:=Newagent;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;

procedure TFRM_Export.CreateDISTD0170s;
Var
Openfile,populateme: Boolean;
LastAgent, NewAgent,lastrole,newrole,darb,olddarb:String;
Cefsd,tempfield,specneeds,contact,specneedscode,dtccode,actionindicator,additionalinfo:string;
begin
 FRM_File_Progress.progressbar.position:=0;

 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select distinct L.LDSO_MPID,M.SSD,M.MPANCORE');
  sql.add('from mopmgr.MPAN_STATUS M,MDDWORKING.LDSO L');
  sql.add('where L.MPANSTART=substr(M.MPANCORE,1,2)');
  sql.add(' and M.D0170_SENT_DIST is NULL and M.RESPONSE=''D0011''');
  sql.add('order by L.LDSO_MPID,M.MPANCORE');
  open;
 end;

 // ResetFileCounters etc
 OpenFile := False;
 LastAGENT:='non';
 LastRole:='non';
 olddarb:='non';
 // Loop Through All MPANS

 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+Generalquery.fields[2].text;
  //FRM_Main.statusbar.panels[1].text:='Creating D0170 to '+generalquery.fields[0].text+' for MPAN '+Generalquery.fields[2].text;
  application.processmessages;

  NewAgent:=GeneralQuery.fields[0].text;

  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) then
  Begin
   lastagent:=NewAgent;
   olddarb:='non';
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0170001',NewAgent,'R');
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
  End;
   // get all the details for the selected MPAN
  mpan:=generalquery.fields[2].text;
  DARB:=formatdatetime('YYYYMMDD',strtodate(generalquery.fields[1].text));
  //////////////////////////////////////////////////////////////////////////////
  if darb<>olddarb then
  Begin
   s:='350|'+DARB+'|21|PLEASE FORWARD SITE TECHNICAL DETAILS TO NEW MOP|'+M_mpid+'||';
   inc(OutputFlowFlowcount);
   inc(OutputFlowLineCount);
   frm_main.WriteLinetoFile(s);
  end;
  s:='351|'+MPAN+'|';
  inc(OutputFlowLineCount);
  frm_main.WriteLinetoFile(s);
  olddarb:=darb;
  lastagent:=Newagent;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;


procedure TFRM_Export.SelectAllXFlows;
Var
F:Integer;
begin
 For F:=0 to XChecklist.items.count-1 do
 Begin
  If XChecklist.itemenabled[F]=true then XChecklist.checked[F]:=true
  else XChecklist.checked[F]:=false;
 end;
end;

procedure TFRM_Export.SelectAllMFlows;
Var
F:Integer;
begin
 For f:=1 to MChecklist.items.count do
 Begin
  If MChecklist.itemenabled[f-1]=true then MChecklist.checked[f-1]:=true
  else MChecklist.checked[f-1]:=false;
 end;
end;

procedure TFRM_Export.SelectAllGFlows;
Var
F:Integer;
begin
 For f:=1 to Gas_Supplier_List.items.count do
 Begin
  If Gas_Supplier_List.itemenabled[f-1]=true then Gas_Supplier_List.checked[f-1]:=true
  else Gas_Supplier_List.checked[f-1]:=false;
 end;

 For f:=1 to Gas_Shipper_List.items.count do
 Begin
  If Gas_Shipper_List.itemenabled[f-1]=true then Gas_Shipper_List.checked[f-1]:=true
  else Gas_Shipper_List.checked[f-1]:=false;
 end;
end;


function TFRM_Export.EstimateFinalRead08D(MPRN, RegisterID: string; EndDate: TDateTime): string;
var RegisterReading : string;
begin
  result :=  EmptyStr;
   RegisterReading := GetRegisteredReading(MPAN, RegisterID,EndDate); //first check we don't already have a reading
   if RegisterReading = EmptyStr then
   begin
    RegisterReading := FRM_Common.EstimateFutureRead(MPAN, RegisterID,EndDate);
   end;
   result :=  RegisterReading;
end;

function TFRM_Export.GetRegisteredReading(MPRN, RegisterID: string;  EndDate: TDateTime): String;
var RegisterReading : string;
begin
  RegisterReading := EmptyStr;
   with CheckReading do
   Begin
    close;
    setvariable('MPAN',MPAN);
    setvariable('RegID',RegisterID);
    setvariable('DDATE',EndDate);
    open;
   End;
   RegisterReading := CheckReading.fields[0].text;
   result := RegisterReading;
end;

procedure TFRM_Export.ExportBtnClick(Sender: TObject);
var
msg:string;
begin
 msg:='Are you sure you wish to export selected dataflows?';
 if timetostr(now)<'17:00' then msg:='Are you sure you wish to export selected dataflows? Current time is before 5:00pm.';
 Exportbtn.enabled:=false;
 Pagecontrol1.enabled:=false;

 if MessageDlg(msg,
 mtConfirmation, [mbYes, mbNo], 0) = mryes then
 Begin
  if loadallX.checked=true then SelectAllXflows;
  if loadallM.checked=true then SelectAllMFLows;
  if loadallG.checked=true then SelectAllGFLows;
  FRM_Export.DoDailyExports;
 end;
 Pagecontrol1.enabled:=true;
// If Messagedlg('All flows created. Do you wish to run End of Day clean up Jobs?',mtconfirmation,[mbyes,mbno],0)=mryes then
// ////////////////////////////////////////////////////////////////////////////////
// // This block of code performs a one off daily batch run of outgoing files    //
// ////////////////////////////////////////////////////////////////////////////////
// Begin
//  FRM_Transfer_Aregi.show;
//  FRM_Transfer_Aregi.gobtn.click;
//  FRM_Transfer_Aregi.close;
//  //frm_import.resetownedenquiries;
// end;
 Exportbtn.enabled:=true;
 updatestatusbar('','');
end;

procedure TFRM_Export.CancelBtnClick(Sender: TObject);
begin
 close;
end;

procedure TFRM_Export.ExportD0149_X;
Begin
 // Get Supplier D0149s
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0149_X');
  sql.add('order by 1,3,4,7,6,8,9');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0149Query;
end;

procedure TFRM_Export.ExportD0149_D;
Begin
 // Get DC D0149s
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0149_D');
  sql.add('order by 1,3,4,7,6,8,9');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0149Query;
end;

procedure TFRM_Export.ExportD0149_R;
Begin
 // Get Dist D0149s
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0149_R');
  sql.add('order by 1,3,4,7,6,8,9');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0149Query;
end;

procedure TFRM_Export.ExportD0149_M;
Begin
 // Get Dist D0149s
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0149_M');
  sql.add('order by 1,3,4,7,6,8,9');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0149Query;
end;

procedure TFRM_Export.ExportThisD0149Query;
Var
Lastagent,OldMpan,OldNonSet,OldTpr,OldMeter,OldRegister:string;
Newagent,mpan,nonset,tpr,meter,registerid,efd,oldefd:string;
openfile:Boolean;
Begin

 // ResetFileCounters etc
 OpenFile := False;
 LastAGENT:='non';
 oldmpan:='non';
 oldefd:='non';
 oldnonset:='non';
 oldtpr:='non';
 oldmeter:='non';
 oldregister:='non';
 // Loop Through All Records

 while not GeneralQuery.eof do
 begin

  NewAgent:=GeneralQuery.fields[0].text;
  mpan:=Generalquery.fields[2].text;
  efd:=Generalquery.fields[3].text;
  NonSet:=Generalquery.fields[6].text;
  TPR:=Generalquery.fields[7].text;
  Meter:=Generalquery.fields[8].text;
  RegisterID:=Generalquery.fields[9].text;


  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MPAN;
  //FRM_Main.statusbar.panels[1].text:='Creating D0149 to '+MPAN+' for MPAN '+NEWAGENT;
  application.processmessages;



  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) then
  Begin
   lastagent:=NewAgent;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0149001',NewAgent,Generalquery.fields[1].text);
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
  End;

  // Mpan Details
  if (MPAN<>OLDMPAN) or (EFD<>OLDEFD) then
  Begin
   S:='280|'+MPAN+'|'+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[3].text))+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   inc(outputflowFlowcount);
   oldnonset:='non';
  End;

  // SSC Details
  if OldNonSet<>NonSet then
  Begin
   if NONSET='N' then
   Begin
    if generalquery.fields[4].text<>'' then
    Begin
     S:='281|'+generalquery.fields[4].text+'|'+formatdatetime('YYYYMMDD',strtodate(generalquery.fields[5].text))+'|';
     frm_main.WriteLinetoFile(s);
     Inc(OutputFlowLineCOunt);
    end;
    OldTpr:='non';
   end
   else
   Begin
    if generalquery.fields[4].text<>'' then
    Begin
     S:='23A|'+generalquery.fields[10].text+'|';
     if generalquery.fields[11].text<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(generalquery.fields[11].text));
     s:=s+'|';
     frm_main.WriteLinetoFile(s);
     Inc(OutputFlowLineCOunt);
    end;
    OldTpr:='non';
   End;
  end;

  // TPR Details
  if (TPR<>OLDTPR) then
  Begin
   if NONSET='N' then s:='778|'+TPR+'|'
   else S:='24A|'+TPR+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   OldMeter:='Non';
   OldRegister:='Non';
  end;

  // Meter Details
  if (Meter<>OLDMeter) then
  Begin
   if NONSET='N' then s:='283|'+Meter+'|'
   else S:='25A|'+meter+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   oldregister:='non';
  end;

  // Register Details
  if (registerid<>OLDRegister) then
  Begin
   if NONSET='N' then s:='284|'+registerid+'|'+generalquery.fields[12].text+'|'
   else S:='26A|'+registerid+'|'+generalquery.fields[12].text+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
  end;

  lastagent:=Newagent;
  Oldmpan:=mpan;
  oldefd:=efd;
  OldNonSet:=NonSet;
  OldTpr:=Tpr;
  OldMeter:=Meter;
  oldregister:=registerid;

  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
End;

procedure TFRM_Export.ExportD0150_X;
Begin
 // Get Supplier D0150s
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0150_X');
  sql.add('order by 1,3,4,10,23');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0150Query;
end;

procedure TFRM_Export.ExportD0150_D;
Begin
 // Get DC D0150s
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0150_D');
  sql.add('order by 1,3,4,10,23');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0150Query;
end;

procedure TFRM_Export.ExportD0150_R;
Begin
 // Get Dist D0150s
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0150_R');
  sql.add('order by 1,3,4,10,23');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0150Query;
end;


procedure TFRM_Export.ExportD0150_M;
Begin
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0150_M');
  sql.add('order by 1,3,4,10,23');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0150Query;
end;

procedure TFRM_Export.ExportThisD0150Query;
Var
oldmap,udmstext,Lastagent,OldMpan,OldNonSet,OldTpr,OldMeter,OldRegister:string;
Newagent,agentrole,mpan,nonset,tpr,meter,registerid,filename,removed,manmake,efd,oldefd:string;
openfile,error:Boolean;
efsdmsmtd:tdatetime;
Begin

 // ResetFileCounters etc
 OpenFile := False;
 LastAGENT:='non';
 oldmpan:='non';
 oldefd:='non';
 oldmeter:='non';
 oldregister:='non';

 FRM_File_Progress.progressbar.position:=0;
 FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
 FRM_File_Progress.d_file.caption:='';
 FRM_File_Progress.labelcount.caption:='';
  // Loop Through All Records

 while not GeneralQuery.eof do
 begin
  NewAgent:=GeneralQuery.fields[0].text;
  agentrole:=Generalquery.fields[1].text;
  mpan:=Generalquery.fields[2].text;
  efd:=Generalquery.fields[3].text;
  Meter:=Generalquery.fields[9].text;
  RegisterID:=Generalquery.fields[22].text;
  Filename:=Generalquery.fields[28].text;
  Removed:=GeneralQuery.fields[27].text;
  manmake:=GeneralQuery.fields[12].text;

  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MPAN;
  //FRM_Main.statusbar.panels[1].text:='Creating D0150 to '+MPAN+' for MPAN '+NEWAGENT;
  application.processmessages;



  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) then
  Begin
   lastagent:=NewAgent;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0150002',NewAgent,agentrole);
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
  End;

  EFSDMSMTD:=strtodate(Generalquery.fields[3].text);

  // Mpan Details
  if (MPAN<>OLDMPAN) or (EFD<>OLDEFD) then
  Begin
   S:='288|'+MPAN+'|'+formatdatetime('YYYYMMDD',EFSDMSMTD)+'||'+Generalquery.fields[4].text+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   inc(outputflowFlowcount);
   // If Meters Present
   if generalquery.fields[9].text<>'' then
   Begin
    S:='289|'+Generalquery.fields[5].text+'|';
    if Generalquery.fields[6].text<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[6].text));
    s:=s+'|';
    s:=s+Generalquery.fields[7].text+'|';
    if Generalquery.fields[8].text<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[8].text));
    s:=s+'|';
    frm_main.WriteLinetoFile(s);
    Inc(OutputFlowLineCOunt);
   end;


   // if new agent=UDMS and its a Liberty meter then include
   // group 762 - Metering Point maintenance
   // include comms address, gas y/n mprn gas meter id etc.
    if (newagent='UDMS') and
       (Generalquery.fields[1].text='D') and
       (uppercase(copy(manmake,1,11))='PRI LIBERTY') then
    Begin
     with main_data_module.tempquery do
     Begin
      close;
      sql.clear;
      sql.add('select mpancore,to_char(install_date,''YYYYMMDD''),new_a_hub_telephone||'';''||meter_point_reference||'';''||new_g_serial||'';''||actual_serial||'';'' from mopmgr.udms_meter_view_new');
      sql.add('where mpancore='''+mpan+'''');
      sql.add('and e_meter_id='''+uppercase(generalquery.fields[9].text)+'''');
      sql.Add('order by install_date desc');
      open;
      if recordcount=0 then
      begin
       close;
       sql.clear;
       sql.add('select mpancore,to_char(install_date,''YYYYMMDD''),new_a_hub_telephone||'';''||meter_point_reference||'';''||new_g_serial||'';''||actual_serial||'';'' from mopmgr.udms_meter_view');
       sql.add('where mpancore='''+mpan+'''');
       sql.add('and e_meter_id='''+uppercase(generalquery.fields[9].text)+'''');
       sql.Add('order by install_date desc');
       open;
      end;
     end;
     if main_data_module.tempquery.recordcount<>0 then
     Begin
      S:='762|'+main_data_module.tempquery.fields[1].text+'|'+main_data_module.tempquery.fields[2].text+'|';
      frm_main.WriteLinetoFile(s);
      Inc(OutputFlowLineCOunt);
     end;
    end;

    // Check for Group 762 based on Filename, only 1 record allowed
    With D0150Info do
    Begin
     close;
     sql.clear;
     sql.add('Select additional_info,visit_date');
     sql.add('from mopmgr.site_visit_info');
     sql.add('where DFLOW=''D0150'' and Meterid is null');
     sql.add('and mpancore='''+mpan+''' and filename='''+filename+'''');
     sql.add('order by visit_date desc');
     open;
     if d0150info.recordcount<>0 then
     begin
      s:='762|'+formatdatetime('YYYYMMDD',strtodate(D0150info.fields[1].text))+'|'+D0150info.Fields[0].text+'|';
      frm_main.WriteLinetoFile(s);
      Inc(OutputFlowLineCOunt);
     end;
    End;
   OldMeter:='non';
  End;

  // Meter Details
  if (Meter<>OLDMeter) and (meter<>'') then
  Begin
   S:='290|'+meter+'|||'+generalquery.fields[10].text+'|';
   s:=s+Generalquery.fields[11].text+'|';
   s:=s+Generalquery.fields[12].text+'|';
   s:=s+Generalquery.fields[13].text+'||||||||||';
   s:=s+Generalquery.fields[14].text+'|';
   s:=s+Generalquery.fields[15].text+'|';
   if Generalquery.fields[16].text<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[16].text));
   s:=s+'|';
   if Generalquery.fields[17].text<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[17].text));
   s:=s+'|';
   if Generalquery.fields[18].text<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[18].text));
   s:=s+'|';
   s:=s+Generalquery.fields[19].text;
   s:=s+'||';
   s:=s+Generalquery.fields[20].text+'|';
   if Generalquery.fields[21].text<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[21].text));
   s:=s+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   oldregister:='non';
  end;
  // Register Details
  if (registerid<>OLDRegister) and (Registerid<>'') then
  Begin
   S:='293|'+registerid+'|'+generalquery.fields[23].text+'|';
   s:=s+Generalquery.fields[24].text+'|';
   if Generalquery.fields[25].text='0.1' then s:=s+'0.10';
   if Generalquery.fields[25].text='1' then s:=s+'1.00';
   if Generalquery.fields[25].text='10' then s:=s+'10.00';
   s:=s+'||';
   s:=s+Generalquery.fields[26].text+'|||';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
  end;
  Oldmpan:=mpan;
  lastagent:=Newagent;
  OldMeter:=Meter;
  oldefd:=efd;
  oldregister:=registerid;

  GeneralQuery.Next;

  // If End of Meter Record of EOF then Write any Maintenance details
  if (GeneralQuery.fields[9].text<>oldmeter) or (Generalquery.eof=true) then
  Begin
  // Check for Group 296 based on Filename, only 1 record allowed
   With D0150Info do
   Begin
    close;
    sql.clear;
    sql.add('Select additional_info,visit_date');
    sql.add('from mopmgr.site_visit_info');
    sql.add('where DFLOW=''D0150'' and Meterid is not null');
    sql.add('and additional_info<>''** METER REMOVED **''');
    sql.add('and mpancore='''+mpan+''' and filename='''+filename+'''');
    sql.add('order by visit_date desc');
    open;
    while not d0150info.eof do
    begin
     s:='296|'+formatdatetime('YYYYMMDD',strtodate(D0150info.fields[1].text))+'|'+D0150info.Fields[0].text+'|';
     frm_main.WriteLinetoFile(s);
     Inc(OutputFlowLineCOunt);
     d0150info.next;
    end;
   End;
  end;


  // If End of MPAN record and there are Meters Removed
  // Then Create a Group 08A Record);
  // Dont include this group when sending to MOP   Wrike 164606267: D0150 Issue
  if agentrole<>'M' then
  begin
    if (Generalquery.fields[2].text<>oldMPAN) or (Generalquery.eof=true) then
    Begin
     if Removed<>'' then
     Begin
      with getremovedmeters do
      Begin
       Close;
       SetVariable('MPAN',MPAN);
       SetVariable('FILENAME',FILENAME);
       open;
       if getremovedmeters.recordcount<>0 then
       Begin
        // Have we already notified this agent role of meter change
        // if so dont include again.
        // Removed meters only sent once, after that just current MTDS
        with main_data_module.tempquery do
        Begin
         close;
         sql.clear;
         sql.add('select * from mopmgr.agents_removed_meters');
         sql.add('where mpancore='''+mpan+'''');
         sql.add('and efsdmsmtd=to_date('''+datetostr(efsdmsmtd)+''',''DD/MM/YYYY'')');
         sql.Add('and toid='''+agentrole+'''');
         open;
        end;
        if main_data_module.tempquery.recordcount=0 then
        Begin
         While not getremovedmeters.eof do
         Begin
          oldMAP:=Getremovedmeters.fields[4].text;
          if( oldmap='') and (copy(mpan,1,2)='20') then oldmap:='SOUT';
          s:='08A|'+GetRemovedMeters.fields[2].text+'|'+formatdatetime('YYYYMMDD',strtodate(getremovedmeters.fields[3].text))+'|'+OLDMAP+'|';
          frm_main.WriteLinetoFile(s);
          Inc(OutputFlowLineCOunt);
          GetRemovedMeters.Next;
         End;
        end;
       End;

      End;
     End;
    End;
  end;

 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
End;


procedure TFRM_Export.ExportD0313_X;
Begin
 // Get Supplier D0150s
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0313_X');
  sql.add('order by 1,2,3');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0313Query;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Export.ExportD0386;
var
  q               : TOracleDataSet;
  lastDistribName : string;
  recordData      : string;
  isNewRecord     : boolean;
begin
  q := TOracleDataSet.Create(nil);
  try
    q.Session := FRM_Login.MainSession;

    q.SQL.Text :=
      'select distributor_id, flowline, rnum '+
        'from edmgr.vw_exp_D0386';
    q.Open;
    if q.RecordCount = 0 then
      exit;

    // using EDMGR.VW_EXP_D0386
    FRM_File_Progress.ProgressBar.Position := 0;
    FRM_File_Progress.ProgressBar.Max      := q.RecordCount;

    lastDistribName := q.FieldByName('distributor_id').AsString;
    isNewRecord     := true;

    while not q.Eof do
    begin
      if isNewRecord then
      begin
        // Create Flow Header  - NEW FILE
        CreateFlowHeader('D0386001', q.FieldByName('distributor_id').AsString, 'P', false);
        isNewRecord := false;
      end;

      recordData := q.FieldByName('flowline').AsString;
      if Copy(recordData, 1, 3) = '18M' then
        recordData := StringReplace(recordData, 'INSTNO', FRM_Common.NextInstructionNumber, [rfReplaceAll]);

      FRM_Main.WriteLinetoFile(recordData);

      Inc(OutPutFlowLineCount);

      if q.FieldByName('rnum').AsInteger = 1 then
        Inc(OutputFlowFlowCount);

      FRM_File_Progress.ProgressBar.Position     := FRM_File_Progress.ProgressBar.Position + 1;
      FRM_File_Progress.Statusbar.Panels[0].Text := 'Distributor: ' + q.FieldByName('distributor_id').AsString;
      FRM_File_Progress.Statusbar.Update;

      q.Next;

      if not q.Eof then
      begin
        if q.FieldByName('distributor_id').AsString <> lastDistribName then
          isNewRecord := true;

        // write Footer (previous record)
        CreateFlowFooter(false);
        lastDistribName := q.FieldByName('distributor_id').AsString;
      end
      else
      begin
        // write last footer.
        Inc(OutputFlowFlowCount);
        CreateFlowFooter;
        OutputFlowFlowCount := 0;
        OutPutFlowLineCount := 0;
      end;
    end;
  finally
    FreeAndNil(q);
  end;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Export.ExportD0386Unrelated;
var
  q               : TOracleDataSet;
  lastDistribName : string;
  isNewRecord     : boolean;
  recordData      : string;
begin
  q := TOracleDataSet.Create(nil);
  try
    q.Session := FRM_Login.MainSession;

    q.SQL.Text :=
      'select distributor, flowline '+
        'from edmgr.vw_exp_D0386_unrelated '+
        'order by distributor';
    q.Open;
    if q.RecordCount = 0 then
      exit;

    FRM_File_Progress.ProgressBar.Position := 0;
    FRM_File_Progress.ProgressBar.Max      := q.RecordCount;

    lastDistribName := q.FieldByName('distributor').AsString;
    isNewRecord     := true;

    while not q.Eof do
    begin
      if isNewRecord then
      begin
        // Create Flow Header  - NEW FILE
        CreateFlowHeader('D0386001', q.FieldByName('distributor').AsString, 'P', false);
        isNewRecord := false;
      end;

      recordData := q.FieldByName('flowline').AsString;

      if Copy(recordData, 1, 3) = '18M' then
        recordData := StringReplace(recordData, 'INSTNO', frm_common.NextInstructionNumber, [rfReplaceAll]);

      FRM_Main.WriteLinetoFile(recordData);

      CRLFPos := 1;
      while CRLFPos <> 0 do
      begin
        CRLFPos := PosEx(#13#10, S, CRLFPos);
        if (CRLFPos > 0) then
        begin
          if CRLFPos < Length(S) then
          begin
            Inc(OutputFlowLineCount);
            Inc(CRLFPos);
          end
          else
            CRLFPos := 0;
        end;
      end;

      Inc(OutputFlowFlowCount);

      // Move to next record
      q.Next;

      if not q.Eof then
      begin
        if UpperCase(q.FieldByName('distributor').AsString) <> UpperCase(lastDistribName) then
        begin
          isNewRecord := true;
          CreateFlowFooter(false);
          lastDistribName := q.FieldByName('distributor').AsString;
        end;
      end
      else
      begin
        // write last footer.
        CreateFlowFooter;
        OutputFlowFlowCount := 0;
        OutPutFlowLineCount := 0;
      end;
    end;
  finally
    FreeAndNil(q);
  end;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Export.ExportD0313_D;
Begin
 // Get DC D0150s
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0313_D');
  sql.add('order by 1,2,3');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0313Query;
end;

procedure TFRM_Export.ExportD0313_R;
Begin
 // Get Dist D0150s
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0313_R');
  sql.add('order by 1,2,3');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0313Query;
end;


procedure TFRM_Export.ExportD0313_M;
Begin
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.Add('select * from mopmgr.export_D0313_M');
  sql.add('order by 1,2,3');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0313Query;
end;

procedure TFRM_Export.ExportThisD0313Query;
Var
oldmap,udmstext,Lastagent,OldMpan,OldNonSet,OldTpr,OldMeter,OldRegister:string;
Newagent,agentrole,mpan,nonset,tpr,meter,registerid,filename,removed,manmake,efd,oldefd:string;
openfile,error:Boolean;
efsdmsmtd:tdatetime;
Begin

 // ResetFileCounters etc
 OpenFile := False;
 LastAGENT:='non';
 oldmpan:='non';
 oldefd:='non';
 oldmeter:='non';
 oldregister:='non';

 FRM_File_Progress.progressbar.position:=0;
 FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
 FRM_File_Progress.d_file.caption:='';
 FRM_File_Progress.labelcount.caption:='';
  // Loop Through All Records

 while not GeneralQuery.eof do
 begin
  NewAgent:=GeneralQuery.fields[0].text;
  agentrole:=Generalquery.fields[1].text;
  mpan:=Generalquery.fields[2].text;
  efd:=Generalquery.fields[3].text;

  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MPAN;
  //FRM_Main.statusbar.panels[1].text:='Creating D0150 to '+MPAN+' for MPAN '+NEWAGENT;
  application.processmessages;



  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) then
  Begin
   lastagent:=NewAgent;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0313002',NewAgent,agentrole);
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
  End;

  // Mpan Details
  if (EFD<>OLDEFD) then
  Begin
   S:=efd;
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   inc(outputflowFlowcount);
  end;
  Oldmpan:=mpan;
  lastagent:=Newagent;
  OldMeter:=Meter;
  oldefd:=efd;
  oldregister:=registerid;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
End;


procedure TFRM_Export.FormShow(Sender: TObject);
begin
  PageControl1.ActivePage := Tabsheet_Export_E;
  PageControl2.ActivePage := TabSheet_EXPORT_EX;
  PageControl3.ActivePage := TabSheet_EXPORT_GS;
end;

procedure TFRM_Export.ExportD0312_P;
var
  LastRecipientName, CurrRecipientName: String;
  bNewRecord: Boolean;
begin
  try
    gSqlUtil.ExecProc('MOPMGR.pr_process_bulk_d0312', TRANSACTION_NO);
  except
    on e:Exception do
     gLogger.Log('Error with mopmgr.pr_process_bulk_d0312: %s', [e.Message], llError);
  end;
  // Create D0312 for MPAS
  with Generalquery do
  begin
    Close;
    SQL.Clear;
    SQL.Add('Select * from MOPMGR.EXPORT_D0312_P');
    Open;
  end;

  if GeneralQuery.recordcount=0 then exit;

  // Create Flow
  LastRecipientName := GeneralQuery.Fields[0].AsString;
  bNewRecord := True;
  FRM_File_Progress.progressbar.position:=0;
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;


  while not GeneralQuery.eof do
  begin
    CurrRecipientName := GeneralQuery.Fields[0].AsString;

    //headers
    CurrRecipientName := GeneralQuery.Fields[0].AsString;

    if bNewRecord then
    begin
      CreateFlowHeaderMOP('D0312003',Generalquery.fields[0].text,Generalquery.fields[1].text);
      bNewRecord := False;
      OutputFlowFlowcount:=0;
      OutputFlowLineCount:=0;

    end;

    //details
    MPAN := Generalquery.fields[2].text;
    FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
    FRM_File_Progress.d_file.caption:='';
    FRM_File_Progress.labelcount.caption:='';
    FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
    FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MPAN;
    application.processmessages;

    // Mpan Details
    S:=  Generalquery.fields[3].text;
    frm_main.WriteLinetoFile(s);
    Inc(OutputFlowLineCOunt);
    inc(outputflowFlowcount);

    // Meter Details
    if    Generalquery.fields[4].text <> '' then
    Begin
       S:=  Generalquery.fields[4].text;
       frm_main.WriteLinetoFile(s);
       Inc(OutputFlowLineCOunt);
    end;

    //
    if Generalquery.fields[5].text <> '' then  // Meter Removed Value is not empty  DEL 64
    Begin
     S:= Generalquery.fields[5].text;
     frm_main.WriteLinetoFile(s);
     Inc(OutputFlowLineCOunt);
    end;

    GeneralQuery.Next;

    if not GeneralQuery.Eof then
    begin
      // DISTRIBUTOR ID
      CurrRecipientName := GeneralQuery.Fields[0].AsString;

      if CurrRecipientName <> LastRecipientName then
      begin
        bNewRecord := True;
        // write Footer (previous record)
        CreateFlowFooter(False);
        LastRecipientName := CurrRecipientName;
      end
      else
      begin
        bNewRecord := False;
      end;
    end
    else
    begin
      // write last footer.
      inc(OutputFlowFlowCount);
      CreateFlowFooter;
      OutputFlowFlowCount := 0;
      OutPutFlowLineCount := 0;
    end;
  end;
end;

procedure TFRM_Export.ExportD0303_APPOINT;
Begin
 // Create D0303
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select * from mopmgr.export_d0303_appointment');
  sql.add('order by 1,3,4,8,2');

  open;
 End;
 if GeneralQuery.recordcount=0 then exit;
 ExportThisD0303;
end;

procedure TFRM_Export.ExportD0303_DEAPPOINT;
Begin
 // Create D0303
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select * from mopmgr.export_d0303_termination');
  sql.add('order by 1,3,4,8,2');
  open;
 End;
 if GeneralQuery.recordcount=0 then exit;
 ExportThisD0303;
end;

Procedure TFRM_Export.ExportThisD0303;
Var
Lastagent,OldMpan,OldNonSet,OldTpr,OldMeter,OldRegister,olddoa:string;
group,date_of_action,Newagent,mpan,nonset,tpr,meter,registerid,supp,date_removed,meter_removed:string;
openfile:Boolean;
begin
 // Create Flow
 // ResetFileCounters etc
 OpenFile := False;
 LastAGENT:='non';
 oldmpan:='non';
 oldmeter:='non';
 olddoa:='non';
 // Loop Through All Records

 while not GeneralQuery.eof do
 begin

  NewAgent:=GeneralQuery.fields[0].text;
  Group:=GeneralQuery.fields[1].text;
  mpan:=Generalquery.fields[2].text;
  SUPP:=Generalquery.fields[11].text;
  Meter:=Generalquery.fields[7].text;
  date_OF_ACTION:=Generalquery.fields[9].text;

  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MPAN;
  //FRM_Main.statusbar.panels[1].text:='Creating D0303 to '+NEWAGENT+' for MPAN '+MPAN;
  application.processmessages;



  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) then
  Begin
   lastagent:=NewAgent;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0303001',NewAgent,'8');
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   s:='72C|'+M_MPID+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   Openfile:=true;
   LastAGENT:='non';
   oldmpan:='non';
   oldmeter:='non';
   olddoa:='non';
  End;

  // Mpan Details
  if MPAN<>OLDMPAN then
  Begin
   S:='73C|'+MPAN+'|';
   if Generalquery.fields[3].text<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[3].text));
   s:=s+'|';
   if Generalquery.fields[4].text<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[4].text));
   s:=s+'|';
   s:=s+'|';
   // Write Address
   With MopSiteAddress Do
   Begin
    close;
    setvariable('MPAN',MPAN);
    setvariable('SSD',Generalquery.fields[6].text);
    open;
   End;

 // Exclude Address Details for Dataflows to NPML - Wrike 225279572 Feb 2019
   if NewAgent='NPML' then
   begin
    s:=s+'|';
    s:=s+'|';
    s:=s+'|';
    s:=s+'|';
    s:=s+'|';
    s:=s+'|';
    s:=s+'|';
    s:=s+'|';
    s:=s+'|';
    s:=s+'||';
   end
   else
   begin
    s:=s+MopSiteAddress.fields[1].text+'|';
    s:=s+MopSiteAddress.fields[2].text+'|';
    s:=s+MopSiteAddress.fields[3].text+'|';
    s:=s+MopSiteAddress.fields[4].text+'|';
    s:=s+MopSiteAddress.fields[5].text+'|';
    s:=s+MopSiteAddress.fields[6].text+'|';
    s:=s+MopSiteAddress.fields[7].text+'|';
    s:=s+MopSiteAddress.fields[8].text+'|';
    s:=s+MopSiteAddress.fields[9].text+'|';
    s:=s+MopSiteAddress.fields[10].text+'||';
   end;

   s:=s+generalquery.Fields[5].text+'|';
   if Generalquery.fields[6].text<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[6].text));
   s:=s+'|';
   s:=s+'|||'; // contract details
   s:=s+SUPP+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   inc(outputflowFlowcount);
   oldmeter:='non';
  End;

   // Meter Details
  if (group='74C') and (meter<>oldmeter) then
  Begin
   S:='74C|'+meter+'|'+Generalquery.fields[8].text+'|';
   if date_of_Action<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(date_of_Action));
   s:=s+'|';
   s:=s+generalquery.fields[10].text+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   oldregister:='non';
  end
  else
  if (group='76C') and (olddoa<>date_OF_ACTION) then
  Begin
   s:='76C|'+formatdatetime('YYYYMMDD',strtodate(date_OF_ACTION))+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
  end;



  lastagent:=Newagent;
  Oldmpan:=mpan;
  OldMeter:=Meter;
  olddoa:=date_of_action;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;

procedure TFRM_Export.CreateMOPD0139;
Begin
 // Create D0139
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select A.RECIPIENT_MPID,A.RECIPIENT_ROLE,');
  sql.add('A.MPANCORE,A.DATE_OF_ACTION,R.METERID,A.REASON_FOR_SENDING,R.REGISTERID,R.RDNGTYPE,');
  sql.add('R.REGISTERREADING,R.SITE_VISIT_CHECK_CODE');
  sql.add('from mopmgr.FLOWS_TO_SEND A,MOPMGR.READINGS R');
  sql.add('where ');
  sql.add('A.FLOWVERSION=''D0139'' and A.sent_status=''R''');
  sql.add('and A.MPANCORE=R.MPANCORE (+)');
  sql.add('and A.DATE_OF_ACTION=R.READDATE (+)');
  sql.add('order by 1,2,3,5,4');
  open;
 End;
 if GeneralQuery.recordcount=0 then exit;
 ExportThisD0139;
 With main_data_module.UpdateQuery do
 Begin
  close;
  sql.clear;
  sql.add('delete from mopmgr.flows_to_send where flowversion=''D0139''');
  execute;
 End;
 frm_login.mainsession.commit;
end;

procedure TFRM_Export.CreateMOPD0135;
Begin
 // Create D0135
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select * from SMIFF.WMOL_VW_EXPORT_TO_SFIC_FAULTS');
  sql.add('order by 2,3');
  open;
 End;
 if GeneralQuery.recordcount=0 then exit;
 ExportThisD0135;
 frm_login.mainsession.commit;
end;


Procedure TFRM_Export.ExportThisD0139;
Var
lastrole,newrole,Lastagent,OldMpan,OldNonSet,OldTpr,OldMeter,OldRegister:string;
Newagent,mpan,nonset,tpr,meter,registerid:string;
openfile:Boolean;
begin
 // Create Flow
 // ResetFileCounters etc
 OpenFile := False;
 lastrole:='non';
 LastAGENT:='non';
 oldmpan:='non';
 oldmeter:='non';
 oldregister:='non';
 // Loop Through All Records

 while not GeneralQuery.eof do
 begin

  NewAgent:=GeneralQuery.fields[0].text;
  NewRole:=GeneralQuery.fields[1].text;
  mpan:=Generalquery.fields[2].text;
  Meter:=Generalquery.fields[4].text;
  registerid:=Generalquery.fields[6].text;

  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MPAN;
  //FRM_Main.statusbar.panels[1].text:='Creating D0139 to '+NEWAGENT+' for MPAN '+MPAN;
  application.processmessages;

  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) or (NewRole<>LastRole) then
  Begin
   lastagent:=NewAgent;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0139002',NewAgent,Generalquery.fields[1].text);
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
   lastrole:='non';
   LastAGENT:='non';
   oldmpan:='non';
   oldmeter:='non';
   oldregister:='non';
  End;

  // Mpan Details
  if MPAN<>OLDMPAN then
  Begin
   s:='261|'+MPAN+'|'+Generalquery.fields[5].text[1]+'|'+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[3].text))+'||'+generalquery.fields[9].text+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   inc(outputflowFlowcount);
   oldmeter:='non';
  End;

   // Meter Details
  if (Meter<>OLDMeter) and (meter<>'') then
  Begin
   S:='262|'+meter+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   oldregister:='non';
  end;

  if (registerid<>OLDRegister) and (Registerid<>'') then
  Begin
   S:='263|'+registerid+'|'+Generalquery.fields[7].text+'|'+Generalquery.fields[8].text+'.0|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
  end;

  lastagent:=Newagent;
  Oldmpan:=mpan;
  OldMeter:=Meter;
  oldregister:=registerid;
  lastrole:=newrole;

  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;

Procedure TFRM_Export.ExportThisD0135;
Var
lastrole,newrole,Lastagent,OldMpan,OldNonSet,OldTpr,OldMeter,OldRegister:string;
Newagent,mpan,nonset,tpr,meter,registerid:string;
openfile:Boolean;
begin
 // Create Flow
 // ResetFileCounters etc
 OpenFile := False;
 lastrole:='non';
 LastAGENT:='non';
 oldmpan:='non';
 oldmeter:='non';
 oldregister:='non';
 // Loop Through All Records

 while not GeneralQuery.eof do
 begin

  NewAgent:=GeneralQuery.fields[1].text;
  NewRole:=GeneralQuery.fields[0].text;
  mpan:=Generalquery.fields[2].text;
  Meter:=Generalquery.fields[16].text;
  registerid:=Generalquery.fields[20].text;

  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MPAN;
  //FRM_Main.statusbar.panels[1].text:='Creating D0135 to '+NEWAGENT+' for MPAN '+MPAN;
  application.processmessages;

  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) or (NewRole<>LastRole) then
  Begin
   lastagent:=NewAgent;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0135002',NewAgent,Generalquery.fields[0].text);
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
   lastrole:='non';
   LastAGENT:='non';
   oldmpan:='non';
   oldmeter:='non';
   oldregister:='non';
  End;

  // Mpan Details
  if MPAN<>OLDMPAN then
  Begin
   s:='257|'+MPAN+'|'+Generalquery.fields[3].text+'|'+
   Generalquery.fields[4].text+'|'+
   Generalquery.fields[5].text+'|'+
   Generalquery.fields[6].text+'|'+
   Generalquery.fields[7].text+'|'+
   Generalquery.fields[8].text+'|'+
   Generalquery.fields[9].text+'|'+
   Generalquery.fields[10].text+'|'+
   Generalquery.fields[11].text+'|'+
   Generalquery.fields[12].text+'|'+
   Generalquery.fields[13].text+'|'+
   Generalquery.fields[14].text+'|'+
   Generalquery.fields[15].text+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   inc(outputflowFlowcount);
   oldmeter:='non';
  End;

   // Meter Details
  if (Meter<>OLDMeter) and (meter<>'') then
  Begin
   S:='258|'+meter+'|'+Generalquery.fields[17].text+'||';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   oldregister:='non';
  end;

  if (registerid<>OLDRegister) and (Registerid<>'') then
  Begin
   S:='60H|T|'+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[19].text))+'|'+registerid+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
  end;

  lastagent:=Newagent;
  Oldmpan:=mpan;
  OldMeter:=Meter;
  oldregister:=registerid;
  lastrole:=newrole;

  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;


procedure TFRM_Export.CreateMOPD0010;
Begin
 // Create D0010
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select A.RECIPIENT_MPID,A.RECIPIENT_ROLE,');
  sql.add('A.MPANCORE,A.DATE_OF_ACTION,R.METERID,A.REASON_FOR_SENDING,R.REGISTERID,R.RDNGTYPE,');
  sql.add('R.REGISTERREADING,R.SITE_VISIT_CHECK_CODE,R.READINGFLAG,R.REASONCODE,R.READINGSTATUS,NVL(R.READING_METHOD,''N'')');
  sql.add('from mopmgr.FLOWS_TO_SEND A,MOPMGR.READINGS R');
  sql.add('where ');
  sql.add('A.FLOWVERSION=''D0010'' and A.sent_status=''R''');
  sql.add('and A.MPANCORE=R.MPANCORE');
  sql.add('and A.DATE_OF_ACTION=R.READDATE');
  sql.add('and R.CURRENT_STATUS<>''D''');
  sql.add('and R.RDNGTYPE=''F''');
  sql.add('union all');
  sql.add('(Select A.DC_ID,''D'',');
  sql.add('A.MPANCORE,A.READDATE,A.METERID,NULL,A.REGISTERID,A.RDNGTYPE,');
  sql.add('A.REGISTERREADING,A.SITE_VISIT_CHECK_CODE,A.READING_FLAG,A.REASON_CODE,A.READING_FLAG,''N''');
  sql.add('from mopmgr.READINGS_TO_SEND A where A.RDNGTYPE=''F'')');
  sql.add('order by 1,2,3,5,8,7');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0010;

 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select A.RECIPIENT_MPID,A.RECIPIENT_ROLE,');
  sql.add('A.MPANCORE,A.DATE_OF_ACTION,R.METERID,A.REASON_FOR_SENDING,R.REGISTERID,R.RDNGTYPE,');
  sql.add('R.REGISTERREADING,R.SITE_VISIT_CHECK_CODE,R.READINGFLAG,R.REASONCODE,R.READINGSTATUS,NVL(R.READING_METHOD,''N'')');
  sql.add('from mopmgr.FLOWS_TO_SEND A,MOPMGR.READINGS R');
  sql.add('where ');
  sql.add('A.FLOWVERSION=''D0010'' and A.sent_status=''R''');
  sql.add('and A.MPANCORE=R.MPANCORE');
  sql.add('and A.DATE_OF_ACTION=R.READDATE');
  sql.add('and R.CURRENT_STATUS<>''D''');
  sql.add('and R.RDNGTYPE<>''F''');
  sql.add('union all');
  sql.add('(Select A.DC_ID,''D'',');
  sql.add('A.MPANCORE,A.READDATE,A.METERID,NULL,A.REGISTERID,A.RDNGTYPE,');
  sql.add('A.REGISTERREADING,A.SITE_VISIT_CHECK_CODE,A.READING_FLAG,A.REASON_CODE,A.READING_FLAG,''N''');
  sql.add('from mopmgr.READINGS_TO_SEND A where A.RDNGTYPE<>''F'')');
  sql.add('order by 1,2,3,5,8,7');
  open;
 End;
 if GeneralQuery.recordcount<>0 then ExportThisD0010;
end;

Procedure TFRM_Export.ExportThisD0010;
Var
lastrole,newrole,Lastagent,OldMpan,OldNonSet,OldTpr,OldMeter,OldRegister:string;
Newagent,mpan,nonset,tpr,meter,registerid,rtype,oldrtype:string;
openfile:Boolean;
begin
 // Create Flow
 // ResetFileCounters etc
 OpenFile := False;
 lastrole:='non';
 LastAGENT:='non';
 oldmpan:='non';
 oldmeter:='non';
 oldregister:='non';
 oldrtype:='non';
 // Loop Through All Records

 while not GeneralQuery.eof do
 begin

  NewAgent:=GeneralQuery.fields[0].text;
  NewRole:=GeneralQuery.fields[1].text;

  IF (NEWROLE='D') AND (NEWAGENT='') THEN NEWAGENT:='UDMS';

  mpan:=Generalquery.fields[2].text;
  Meter:=Generalquery.fields[4].text;
  registerid:=Generalquery.fields[6].text;
  rtype:=Generalquery.fields[7].text;

  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MPAN;
  //FRM_Main.statusbar.panels[1].text:='Creating D0010 to '+NEWAGENT+' for MPAN '+MPAN;
  application.processmessages;

  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) or (NewRole<>LastRole) then
  Begin
   lastagent:=NewAgent;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0010002',NewAgent,Generalquery.fields[1].text);
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
   lastrole:='non';
   LastAGENT:='non';
   oldmpan:='non';
   oldmeter:='non';
   oldregister:='non';
  End;

  // Mpan Details
  if MPAN<>OLDMPAN then
  Begin
   s:='026|'+MPAN+'|U|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   if generalquery.Fields[9].text<>'' then
   Begin;
    s:='027|'+generalquery.Fields[9].text+'||';
    frm_main.WriteLinetoFile(s);
    Inc(OutputFlowLineCOunt);
   end;
   inc(outputflowFlowcount);
   oldmeter:='non';
  End;

   // Meter Details
  if ((Meter<>OLDMeter) and (meter<>'')) or (rtype<>oldrtype) then
  Begin
   S:='028|'+meter+'|'+rtype+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   oldregister:='non';
  end;

  if (registerid<>OLDRegister) and (Registerid<>'') then
  Begin
   S:='030|'+registerid+'|'+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[3].text))+'000000|'+Generalquery.fields[8].text+'.0|||'+generalquery.fields[10].text+'|'+generalquery.fields[13].text+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   if generalquery.fields[10].text='F' then
   Begin
    S:='032|'+Generalquery.fields[11].text+'|'+generalquery.fields[12].text+'|';
    frm_main.WriteLinetoFile(s);
    Inc(OutputFlowLineCount);
   End;
  end;

  lastagent:=Newagent;
  Oldmpan:=mpan;
  OldMeter:=Meter;
  oldrtype:=rtype;
  oldregister:=registerid;
  lastrole:=newrole;

  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;

Procedure TFRM_Export.ExportEmails;
Var
Emailadd,mpid,role,oldmpid,oldrole,oldmpan,oldmeter,oldregister,registerid,mpan,meter,msg,oldtype,addinfo:string;
openemail:boolean;
Begin
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select A.RECIPIENT_MPID,A.RECIPIENT_ROLE,');
  sql.add('A.MPANCORE,A.DATE_OF_ACTION,R.METERID,A.REASON_FOR_SENDING,R.REGISTERID,R.RDNGTYPE,');
  sql.add('R.REGISTERREADING,R.SITE_VISIT_CHECK_CODE,R.READINGFLAG,R.REASONCODE,R.READINGSTATUS,A.ADDITIONAL_INFO');
  sql.add('from mopmgr.FLOWS_TO_SEND A,MOPMGR.READINGS R');
  sql.add('where ');
  sql.add('A.FLOWVERSION=''EMAIL'' and A.sent_status=''R''');
  sql.add('and A.MPANCORE=R.MPANCORE');
  sql.add('and A.DATE_OF_ACTION=R.READDATE ');
  sql.add('and R.CURRENT_STATUS<>''D''');
  sql.add('order by 1,2,3,4,5,8,7');
  open;
 End;
 if GeneralQuery.recordcount=0 then exit;
 oldmpan:='non';
 oldmeter:='non';
 msg:='';
 oldmpid:='Lee';
 oldrole:='LEE';
 oldtype:='lee';
 addinfo:='';
 openemail:=false;

 While not Generalquery.eof do
 Begin
  mpid:=generalquery.fields[0].text;
  role:=generalquery.fields[1].text;
  mpan:=generalquery.fields[2].text;
  meter:=generalquery.fields[4].text;
  RegisterID:=generalquery.fields[6].text;
  addinfo:=generalquery.fields[13].text;
  if (oldmpid<>mpid) or (oldrole<>role) then
  Begin
   if openemail=true then
   Begin
    msg:=msg+#13+#13+'** End Of Message. **';
    frm_email_comp.mailmessage.body.text:=msg;
    try
     try
      frm_email_comp.smtp.port:=2525;
      frm_email_comp.smtp.connect;
      frm_email_comp.smtp.Send(frm_email_comp.mailmessage); // Send Email
     except
     end;
    finally
     if frm_email_comp.smtp.connected then frm_email_comp.smtp.Disconnect;
    end;

     Openemail:=true;
    oldrole:='non';
    oldmpid:='non';
    oldmpan:='non';
    oldmeter:='non';
    oldregister:='non';
   end;
  end;

  if (mpan<>oldmpan) then
  Begin
   with main_data_module.generalquery do
   Begin
    close;
    sql.clear;
    sql.add('select email from mddworking.market_participant_email');
    sql.add('where MPID='''+GENERALQUERY.FIELDS[0].text+''' and role='''+Generalquery.fields[1].text+'''');
    open;
    if recordcount=0 then emailadd:='not known' else emailadd:=main_data_module.generalquery.fields[0].text;
   End;

   msg:='Please Forward to ('+generalquery.fields[0].text+'-'+generalquery.fields[1].text+') - ('+emailadd+')';
   msg:=msg+#13+#13+addinfo;
   msg:=msg+#13+#13+Generalquery.fields[5].text+' - MPANCORE:-'+generalquery.fields[2].text+' Date Of Action:-'+generalquery.fields[3].text;
   msg:=msg+#13;
   openemail:=true;
   oldmeter:='non';
   oldregister:='non';
   // Get Recipient Email Address


   if frm_email_comp.smtp.connected then frm_email_comp.smtp.Disconnect;
   frm_email_comp.mailmessage.clear;
   frm_email_comp.mailmessage.recipients.emailaddresses:=customerservices;
   // Set Sender Details
   frm_email_comp.mailmessage.From.Address := moenquiries;
   frm_email_comp.mailmessage.Subject :=Generalquery.fields[5].text+' - MPANCORE:-'+generalquery.fields[2].text;// Reason
   try
    try
     frm_email_comp.smtp.port:=2525;
     frm_email_comp.smtp.connect;
     frm_email_comp.smtp.Send(frm_email_comp.mailmessage);
     except
     end;
   finally
    if frm_email_comp.smtp.connected then frm_email_comp.smtp.Disconnect;
   end;


  end;

  if (oldmeter<>meter) or (oldregister<>registerid) or (generalquery.fields[7].text<>oldtype) then
  begin
   msg:=msg+#13+'Meter Serial: '+Meter+' Register ID: '+Registerid+' Reading: '+generalquery.fields[8].text;
   msg:=msg+' Read Type: '+generalquery.fields[7].text;
  End;
  oldmpid:=mpid;
  oldrole:=role;
  oldmpan:=mpan;
  oldmeter:=meter;
  oldregister:=registerid;
  oldtype:=generalquery.fields[7].text;
  generalquery.Next;
 end;

 msg:=msg+#13+#13+'** End Of Message. **';
 frm_email_comp.mailmessage.body.text:=msg;
 try
  try
   frm_email_comp.smtp.port:=2525;
   frm_email_comp.smtp.connect;
   frm_email_comp.smtp.Send(frm_email_comp.mailmessage);
   except
   end;
 finally
  if frm_email_comp.smtp.connected then frm_email_comp.smtp.Disconnect;
 end;

 if frm_email_comp.smtp.connected then frm_email_comp.smtp.Disconnect;
 With main_data_module.UpdateQuery do
 Begin
  close;
  sql.clear;
  sql.add('delete from mopmgr.flows_to_send where flowversion=''EMAIL''');
  execute;
 End;
 frm_login.mainsession.commit;
end;

Procedure TFRM_Export.CreateMOPD0002;
begin
// Create D0002
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('Select A.RECIPIENT_MPID,A.RECIPIENT_ROLE,');
  sql.add('A.MPANCORE,A.DATE_OF_ACTION,A.METERID,A.REASON_FOR_SENDING,A.REGISTERID,A.ADDITIONAL_INFO');
  sql.add('from mopmgr.FLOWS_TO_SEND A');
  sql.add('where ');
  sql.add('A.FLOWVERSION=''D0002'' and A.sent_status=''R''');
  sql.add('order by 1,2,3');
  open;
 End;
 if GeneralQuery.recordcount=0 then exit;
 ExportThisD0002;
 With main_data_module.UpdateQuery do
 Begin
  close;
  sql.clear;
  sql.add('delete from mopmgr.flows_to_send where flowversion=''D0002''');
  execute;
 End;
 frm_login.mainsession.commit;
end;

Procedure TFRM_Export.ExportThisD0002;
Var
lastrole,newrole,Lastagent,OldMpan,OldNonSet,OldTpr,OldMeter,OldRegister:string;
Newagent,mpan,nonset,tpr,meter,registerid,oldreason,reason:string;
openfile:Boolean;
begin
 // Create Flow
 // ResetFileCounters etc
 OpenFile := False;
 lastrole:='non';
 LastAGENT:='non';
 oldmpan:='non';
 oldmeter:='non';
 oldregister:='non';
 oldreason:='non';
 // Loop Through All Records

 while not GeneralQuery.eof do
 begin

  NewAgent:=GeneralQuery.fields[0].text;
  NewRole:=GeneralQuery.fields[1].text;
  mpan:=Generalquery.fields[2].text;
  Meter:=Generalquery.fields[4].text;
  registerid:=Generalquery.fields[6].text;
  reason:=Generalquery.fields[5].text;
  if meter='' then meter:='-';
  if registerid='' then registerid:='-';

  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MPAN;
  //FRM_Main.statusbar.panels[1].text:='Creating D0002 to '+NEWAGENT+' for MPAN '+MPAN;
  application.processmessages;

  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) or (NewRole<>LastRole) then
  Begin
   lastagent:=NewAgent;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0002001',NewAgent,Generalquery.fields[1].text);
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
   lastrole:='non';
   LastAGENT:='non';
   oldmpan:='non';
   oldmeter:='non';
   oldregister:='non';
  End;

  // Mpan Details
  if (MPAN<>OLDMPAN) or (oldreason<>Reason )then
  Begin
   s:='004|'+MPAN+'|'+Generalquery.fields[5].text+'|'+formatdatetime('YYYYMMDD',strtodate(Generalquery.fields[3].text))+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   inc(outputflowFlowcount);
   oldmeter:='non';
  End;

   // Meter Details
  if (Meter<>OLDMeter) and (meter<>'') then
  Begin
   S:='005|'+meter+'|'+generalquery.fields[7].text+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   oldregister:='non';
  end;

  if (registerid<>OLDRegister) and (Registerid<>'') then
  Begin
   S:='006|'+registerid+'|||';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
  end;

  lastagent:=Newagent;
  Oldmpan:=mpan;
  oldreason:=reason;
  OldMeter:=Meter;
  oldregister:=registerid;
  lastrole:=newrole;

  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;

procedure  TFRM_Export.CreateMopD0224;
Var
Openfile,populateme: Boolean;
LastAgent, NewAgent,lastrole,newrole,newfrom,lastfrom:String;
Cefsd,tempfield,specneeds,contact,specneedscode,dtccode,actionindicator,additionalinfo:string;
begin
 FRM_File_Progress.progressbar.position:=0;

 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select * from mopmgr.export_d0224 order by 1,3,2');
  open;
 end;

 // ResetFileCounters etc
 OpenFile := False;
 LastAGENT:='non';
 LastRole:='non';
 lastfrom:='non';
 // Loop Through All MPANS

 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+Generalquery.fields[1].text;
  //FRM_Main.statusbar.panels[1].text:='Creating D0224 to '+generalquery.fields[0].text+' for MPAN '+Generalquery.fields[1].text;
  application.processmessages;

  NewAgent:=GeneralQuery.fields[0].text;
  NewFrom:=GeneralQuery.fields[2].text;
  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) then
  Begin
   lastagent:=NewAgent;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0224001',NewAgent,'X');
   lastfrom:='non';
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
  End;

  if newfrom<>lastfrom then
  Begin
   s:='508|'+Formatdatetime('YYYYMMDD',strtodate(copy(generalquery.fields[2].text,1,10)))+'|'+Formatdatetime('YYYYMMDD',strtodate(copy(generalquery.fields[3].text,1,10)))+'|';
   inc(OutputFlowLineCount);
   frm_main.WriteLinetoFile(s);
  End;
  mpan:=generalquery.fields[1].text;
  s:='509|'+mpan+'|';
  inc(OutputFlowLineCount);
  inc(OutputFlowFlowcount);
  frm_main.WriteLinetoFile(s);
  lastagent:=newagent;
  lastfrom:=newfrom;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;

procedure TFRM_Export.CreateD0311s;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DTC Version Created V8.3
//
// Go Live Date: 23-FEB-2006
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Code Creation Date:  01/05/2006
// Prgrammer Lee Kitchen
// Version 1.0
// Impact Notes: The Following Changes have been indentified
//  Add 2 New Fields to EDMGR.MPAN_STATUS table
//  Field 1 = D0311_OLD_SUPPLIER (Used to identify if we have recevied a D0311 from OLD Supplier on GAIN)
//  Field 2 = D0311_NEW_SUPPLIER (Used to Identify if we have sent a D0311 to New Supplier on LOSS)
//
// Add New Module to Dataflow Manager to Generate D0311 on LOSS                   (TFRM_Export.CreateD0311s)
// Add New Module to Dataflow Manager to Load a D0311 file (Sent or received)     (TFRM_Import.D0311Import)
// Add New Module to CRM to allow users to view contents of D0311                 (TFRM_MPAD_Details.NOSIE FLOW}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// Flow Notes:
// The Flow is sent in order for the New Supplier to have early visibility of key information
// and may be used at the New Supplier's discretion as an aid to identify potential problems when
// data is compared to that from other sources. It does not supersede industry process flow data.
//
// Address details in this flow (Address Lines 1-9 and Postcode) should refer to the Site Address
// held within the billing system, rather than the Billing/ Contact Address.
//
// Groups 09D and 10D refer to Estimated Annual Consumption and should be completed with information
// received from the NHHDC only,  e.g. groups EAH and EAD from D0019.
//
// All other fields should be populated with the Old Supplier's view of the data items from their
// billing processes. For example, Meter Type should be populated with the Meter Type that the customer
// is being billed to rather than the latest view of Meter Type from the MOp. This is to enable the
// New Supplier to identify potential mismatches between historic customer billing and the MPAS/ Agent view.
//
// The Old Supplier Estimated CoS Reading should be populated with an Estimated Reading produced from
// the Old Supplier Billing system only where the Customer has been billed at least once by that Supplier.
//
// The receiving Supplier should assume that the Reading Type for Old Supplier's Estimated CoS Reading is "O"
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Var
Openfile,populateme: Boolean;
LastAgent, NewAgent,lastrole,newrole,newfrom,OLDMPAN,oldmeter,oldregister,meterid,registerid:String;
Cefsd,tempfield,specneeds,contact,specneedscode,dtccode,actionindicator,additionalinfo, data08D, AllData08D:string;
lossdate: TDateTime;
begin
 FRM_File_Progress.progressbar.position:=0;
 ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 // Identify list of MPANS, along with associated data, which are idenified as LOST or Going to be LOST
 // and where a D0311 has not been not yet been sent to NEW Supplier
 // And where Effective Date is after 01/05/2006 - Date this change implemented in system
 ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select distinct');
  sql.add('case when M.regstatus=''LOST'' then substr(M.comments,9,4)');
  sql.add('else substr(M.comments,16,4)');
  sql.add('End NSID,');
  sql.add('M.mpancore,');
  sql.add('M.EFTSSD,');
  sql.add('M.ENERGISATION_STATUS,');
  sql.add('M.PROFILE_CLASS,');
  sql.add('M.MTC,');
  sql.add('M.LLF,');
  sql.add('A.METERING_POINT_ADDRESS1,');
  sql.add('A.METERING_POINT_ADDRESS2,');
  sql.add('A.METERING_POINT_ADDRESS3,');
  sql.add('A.METERING_POINT_ADDRESS4,');
  sql.add('A.METERING_POINT_ADDRESS5,');
  sql.add('A.METERING_POINT_ADDRESS6,');
  sql.add('A.METERING_POINT_ADDRESS7,');
  sql.add('A.METERING_POINT_ADDRESS8,');
  sql.add('A.METERING_POINT_ADDRESS9,');
  sql.add('A.METERING_POINT_POSTCODE');
  sql.add('from');
  sql.add('EDMGR.mpan_status M,');
  sql.add('EDMGR.MPAS_CURRENT_ADDR A');
  sql.add('where a.mpancore=m.mpancore');
  sql.add('and m.d0311_NEW_SUPPLIER is null');
  sql.add('and (M.REGSTATUS=''LOST'' or M.REGSTATUS=''FUTURE LOSS'')');
  sql.add('and eftssd>=to_date(''01/05/2006'',''DD/MM/YYYY'')');
  sql.add('order by 1,2,3');
  open;
 end;

 // Reset File Counters etc
 OpenFile := False;
 LastAGENT:='non';
 LastRole:='non';
 oldmpan:='non';

 // Loop Through All MPANS idenified in above Query
 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+Generalquery.fields[1].text;
  //FRM_Main.statusbar.panels[1].text:='Creating D0311 to '+generalquery.fields[0].text+' for MPAN '+Generalquery.fields[1].text;
  application.processmessages;

  NewAgent:=GeneralQuery.fields[0].text;
  mpan:=Generalquery.fields[1].text;
  lossdate:= StrToDateTime(Generalquery.fields[2].text);
  // Check if NEW SUPPLIER is different, this would indicate that a new file is required.
  if (NewAgent<>LastAgent) then
  Begin
   lastagent:=NewAgent;
   // If an existing SUPPLIER file is open, then close it.
   if openfile = true then CreateFlowFooter;
   // Now Create New File going to NEW supplier and Write Header Record
   CreateFlowHeader('D0311001',NewAgent,'X');
   // Indicate there is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
  End;

  // On Change of MPANCore, Create Group '05'. Group is MANDATORY as per DTC Spec
  if MPAN<>OLDMPAN then
  Begin
   mpan:=generalquery.fields[1].text;
   s:='05D|'+mpan+'|'+Formatdatetime('YYYYMMDD',strtodate(generalquery.fields[2].text))+'|';  // Effective to OLD SUPP
   s:=s+generalquery.fields[3].text+'|';     // Energisation Status
   s:=s+generalquery.fields[4].text+'|';     // Profile Class ID
   s:=s+generalquery.fields[5].text+'|';     // Meter Timeswitch Code
   s:=s+generalquery.fields[6].text+'|';     // LLF
   s:=s+generalquery.fields[7].text+'|';     // Address Line 1
   s:=s+generalquery.fields[8].text+'|';     // Address Line 2
   s:=s+generalquery.fields[9].text+'|';     // Address Line 3
   s:=s+generalquery.fields[10].text+'|';    // Address Line 4
   s:=s+generalquery.fields[11].text+'|';    // Address Line 5
   s:=s+generalquery.fields[12].text+'|';    // Address Line 6
   s:=s+generalquery.fields[13].text+'|';    // Address Line 7
   s:=s+generalquery.fields[14].text+'|';    // Address Line 8
   s:=s+generalquery.fields[15].text+'|';    // Address Line 9
   s:=s+generalquery.fields[16].text+'|';    // Address Line Postcode
   inc(OutputFlowLineCount);
   inc(OutputFlowFlowcount);
   frm_main.WriteLinetoFile(s);  // Write Record to File
   lastagent:=newagent;
   oldmpan:=mpan;

   /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   // Now Check for any Existing Meter Details
   // Along with Last Actual or COR
   /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   with MetersQuery do
   Begin
    close;
    setvariable('MPAN',MPAN);
    open;
   End;
   oldmeter:='NON';
   oldregister:='NON';
   AllData08D := EmptyStr;
   // Repeat if Meters Exists
   while not metersquery.eof do
   Begin
    MeterID:=Metersquery.fields[1].text;
    RegisterID:=MetersQuery.fields[4].text;
    // Check for Change Of Meter
    if meterid<>oldmeter then
    Begin
     s:='06D|'+METERID+'|'+Metersquery.fields[2].text+'||'; // Meterid and Meter Type
     frm_main.WriteLinetoFile(s); // Write Meter Record to File
     oldmeter:=meterid;
    End;
    // Check for Change of Register
    if registerid<>oldregister then
    Begin
     s:='07D|'+REGISTERID+'|'+formatdatetime('YYYYMMDDHHNNSS',strtodatetime(Metersquery.fields[7].text))+'|';
     s:=s+frm_common.asdecimal(Metersquery.fields[5].text)+'|'+Metersquery.fields[6].text+'|';
     inc(OutputFlowLineCount);
     frm_main.WriteLinetoFile(s);   // Write Register Records and Readings
      data08D:= EmptyStr;
      data08D:= EstimateFinalRead08D(MPAN, REGISTERID, lossdate);   //create 08D line with estimate
     if (data08D <> EmptyStr) then
     begin
        if AllData08D <> EmptyStr then
        AllData08D := AllData08D + #13#10;
        AllData08D := AllData08D + '08D|'+REGISTERID+'|'+FRM_Common.AsDecimal(data08D)+'|';
     end;
      oldregister:=registerid;
    End;
    metersquery.next;
   End;

   if AllData08D <> EmptyStr then
    frm_main.WriteLinetoFile(AllData08D);

   // Obtain Last EAC from D0019 Dataflow
   with EACDetails do
   Begin
    close;
    setvariable('MPAN',MPAN);
    open;
   End;
   // Write Any Existing EAC to File
   if eacdetails.recordcount<>0 then
   Begin
    s:='09D|'+formatdatetime('YYYYMMDD',strtodate(copy(eacdetails.fields[1].text,1,10)))+'|';
    inc(OutputFlowLineCount);
    frm_main.WriteLinetoFile(s);
    while not eacdetails.eof do
    Begin
     s:='10D|'+eacdetails.fields[2].text+'|'+frm_common.asdecimal(eacdetails.fields[3].text)+'|';
     inc(OutputFlowLineCount);
     frm_main.WriteLinetoFile(s);
     eacdetails.Next;
    End;
   end;
   GeneralQuery.Next;
  end;
 end; // Repeat Above process for remaining MPANS

 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;

{------------------------------------------------------------------------------}

procedure TFrm_Export.Do_E_SupplierFiles;
  {----------------------------------------------------------------------------}
  procedure COAFChangeOfAgentBatchedFile; // item index = 0
  var
    q : TOracleQuery;
  begin
    UpdateStatusBar('Generating Change Of Agent Batched Files - D0151', 'Y');
    CreateD0151COA;
    UpdateStatusBar('Generating Change Of Agent Batched Files - D0148', 'Y');
    CreateD0148COA;
    UpdateStatusBar('Generating Change Of Agent Batched Files - D0205', 'Y');
    CreateD0205COA;
    UpdateStatusBar('Generating Change Of Agent Batched Files - D0170', 'Y');
    CreateD0170COA;
    UpdateStatusBar('Finalising Change Of Agent Batched Files','Y');

    // Update Sent Status of all Blank Status.s
    try
      q := TOracleQuery.Create(nil);
      try
        q.Session  := FRM_Login.MainSession;
        q.SQL.Text :=
          'update edmgr.batch_flows_for_sending_coa '+
            'set date_generated = sysdate, '+
                'status         = ' + QuotedStr('S')+' '+
            'where date_generated is null';
        q.Execute;
      finally
        FreeAndNil(q);
      end;

      FRM_Login.MainSession.Commit;
    except
      on e:Exception do
      begin
        gLogger.Log('Error in "Change of Agent Batched Files" generating: %s', [e.Message], llError);
        FRM_Login.MainSession.Rollback;
      end;
    end;
  end;
  {----------------------------------------------------------------------------}
  procedure D0005ReadingRequests; // item index = 1
  begin
    UpdateStatusBar('Generating D0005 Reading Requests','Y');
    IssueD0005readRequests;
  end;
  {----------------------------------------------------------------------------}
  procedure D0010_D0071ReadingsToDCAndSupplier; // item index = 2
  var
    q : TOracleQuery;
  begin
    UpdateStatusBar('Generating D0010 Readings to DC - D0010', 'Y');
    CreateD0010('');
    //created0010('UDMSALL'); manual job for UDMS resends
    UpdateStatusBar('Generating D0010 Readings to DC - D0010 from D0188', 'Y');
    CreateD0010fromD0188;
    UpdateStatusBar('Generating D0010 Readings to DC - D0010 from Dials', 'Y');
    CreateD0010fromDIALS;
    UpdateStatusBar('Generating D0071 Customer Own Reads', 'Y');
    CreateD0071;

    UpdateStatusBar('Finalising D0010 to DC', 'Y');
    // DC is always null;

    try
      q := TOracleQuery.Create(nil);
      try
        q.Session  := FRM_Login.MainSession;
        q.SQL.Text :=
          'update edmgr.readings_to_send ' +
            'set sent_status = ' + QuotedStr('S') + ' ' +
            'where sent_status = ' + QuotedStr('R') + ' and readdate <= trunc(sysdate) and ' +
                  'mpancore in (select mpancore '+
                                 'from edmgr.mpan_status '+
                                 'where confirmed_dc_id is not null)';
        q.Execute;
      finally
        FreeAndNil(q);
      end;

      FRM_Login.MainSession.Commit;
    except
      on e:Exception do
      begin
        gLogger.Log('Error in "D0010 - D0071" generating: %s', [e.Message], llError);
        FRM_Login.MainSession.Rollback;
      end;
    end;
  end;
  {----------------------------------------------------------------------------}
  procedure D0052AffirmationOfMeteringSystemDetails; // item index = 3
  begin
    UpdateStatusBar('Generating D0052s','Y');
    FRM_Common.Export_ELEC_Supplier_file('D0052');
  end;
  {----------------------------------------------------------------------------}
  procedure D0064ObjectionRequests; // item index = 4
  begin
    // NOT IMPLEMENTED
  end;
  {----------------------------------------------------------------------------}
  procedure D0131AddressChangeHHDC; // item index = 5
  begin
    UpdateStatusBar('Generating D0131s HHDC Change of Address','Y');
    D0131nhhdc_addresschange;
  end;
  {----------------------------------------------------------------------------}
  procedure D0131AddressChangeMOP;  // item index = 6
  begin
    UpdateStatusBar('Generating D0131s Change of Address MOP','Y');
    D0131MO_addresschange;
  end;
  {----------------------------------------------------------------------------}
  procedure D0131AddressChangeNHHDC;   // item index = 7
  begin
    UpdateStatusBar('Generating D0131s NHHDC Change of Address','Y');
    D0131nhhdc_addresschange;
  end;
  {----------------------------------------------------------------------------}
  procedure D0132SupplyDisconnectionDeatails;   // item index = 8
  begin
    UpdateStatusBar('Generating D0132''s', 'Y');
    FRM_Common.Export_ELEC_Supplier_file('D0132');
  end;
  {----------------------------------------------------------------------------}
  procedure D0142RequestToChangeInstallMetering;   // item index = 9
  begin
    UpdateStatusBar('Generating D0142s', 'Y');
    DoD0142s;
  end;
  {----------------------------------------------------------------------------}
  procedure D0148AgentDetails;   // item index = 10
  begin
    UpdateStatusBar('Generating D0148s', 'Y');
    cosD0148;
  end;
  {----------------------------------------------------------------------------}
  procedure D0151DisconnectionDA;   // item index = 11
  begin
    UpdateStatusBar('Generating D0151 to DA (Disconnection)', 'Y');
    TerminateDAdisc;
  end;
  {----------------------------------------------------------------------------}
  procedure D0151DisconnectionDC;   // item index = 12
  begin
    UpdateStatusBar('Generating D0151 to DC (Disconnection)', 'Y');
    TerminateDCdisc;
  end;
  {----------------------------------------------------------------------------}
  procedure D0151DisconnectionMO;   // item index = 13
  begin
    UpdateStatusBar('Generating D0151 to MOP (Disconnection)', 'Y');
    TerminateMOPdisc;
  end;
  {----------------------------------------------------------------------------}
  procedure D0151LossDA;   // item index = 14
  begin
    UpdateStatusBar('Generating D0151 to DA (Loss)', 'Y');
    TerminateDALoss;
  end;
  {----------------------------------------------------------------------------}
  procedure D0151LossDC;   // item index = 15
  begin
    UpdateStatusBar('Generating D0151 to DC (Loss)', 'Y');
    TerminateDCLoss;
  end;
  {----------------------------------------------------------------------------}
  procedure D0151LossMO;   // item index = 16
  begin
    UpdateStatusBar('Generating D0151 to MOP (Loss)', 'Y');
    TerminateMOPLoss;
  end;
  {----------------------------------------------------------------------------}
  procedure D0151ObjectionDA;   // item index = 17
  begin
    UpdateStatusBar('Generating D0151 to DA (Objection)', 'Y');
    TerminateDALossOBJ;
  end;
  {----------------------------------------------------------------------------}
  procedure D0151ObjectionDC;   // item index = 18
  begin
    UpdateStatusBar('Generating D0151 to DC (Objection)', 'Y');
    TerminateDCLossOBJ;
  end;
  {----------------------------------------------------------------------------}
  procedure D0151ObjectionMO;   // item index = 19
  begin
    UpdateStatusBar('Generating D0151 to MOP (Objection)', 'Y');
    TerminateMOLossOBJ;
    UpdateStatusBar('Generating D0151 to MOP (PSMI)', 'Y');
    TerminateMOPSMI;
  end;
  {----------------------------------------------------------------------------}
  procedure D0153AppointmentDA;   // item index = 20
  begin
    UpdateStatusBar('Generating D0153 to DA (COA)', 'Y');
    NHHDAD0153COA;
    UpdateStatusBar('Generating D0153 to DA (Appointment)', 'Y');
    NHHDAD0153;
    HHDAD0153;
  end;
  {----------------------------------------------------------------------------}
  procedure D0155AppointmentDC;   // item index = 21
  begin
    UpdateStatusBar('Generating D0155 to DC (COA)', 'Y');
    NHHDCD0155COA;
    UpdateStatusBar('Generating D0155 to DC (Appointment)', 'Y');
    NHHDCD0155;
    HHDCD0155;
  end;
  {----------------------------------------------------------------------------}
  procedure D0155AppointmentMO;   // item index = 22
  begin
    UpdateStatusBar('Generating D0155 to MOP (COA)', 'Y');
    MOD0155COA;
    UpdateStatusBar('Generating D0155 to MOP (Appointment)', 'Y');
    NHHMOD0155;
    HHMOD0155;
  end;
  {----------------------------------------------------------------------------}
  procedure D0190PrepayKeyRequests;   // item index = 23
  var
    q : TOracleQuery;
  begin
    UpdateStatusBar('Generating D0190 PrePay Key Requests', 'Y');
    CreateD0190;
    UpdateStatusBar('Finalising D0190','Y');

    // Update Sent Status of all Blank Status.s
    try
      q := TOracleQuery.Create(nil);
      try
        q.Session  := FRM_Login.MainSession;
        q.SQL.Text :=
          'update edmgr.batch_flows_for_sending_all '+
            'set date_generated = sysdate, '+
                'status         = '+QuotedStr('S')+' '+
            'where flowversion = '+QuotedStr('D0190') + ' and date_generated is null';
        q.Execute;
      finally
        FreeAndNil(q);
      end;

      FRM_Login.MainSession.Commit;
    except
      on e:Exception do
      begin
        gLogger.Log('Error in "D0190 PrePay Key Requests" generating', [e.Message], llError);
        FRM_Login.MainSession.Rollback;
      end;
    end;
  end;
  {----------------------------------------------------------------------------}
  procedure D0205MPASUpdates;   // item index = 24
  begin
    UpdateStatusBar('Generating D0205 MPAS Updates', 'Y');
    Do_E_D0205s;
  end;
  {----------------------------------------------------------------------------}
  procedure D0225SpecialNeeds;   // item index = 25
  begin
    UpdateStatusBar('Refreshing D0225 Snapshot', 'Y');
    FRM_COMMON.Execute_Oracle_Procedure('CRM.PR_REFRESH_SPECIAL_NEEDS_SNAP');
    UpdateStatusBar('Generating D0225 Special Needs', 'Y');
    CreateOutstandingD0225s;
  end;
  {----------------------------------------------------------------------------}
  procedure D0301ErroneousTransfers;   // item index = 26
  begin
    UpdateStatusBar('Generating D0301 ETs', 'Y');
    FRM_COMMON.Export_ELEC_Supplier_file('D0301');
  end;
  {----------------------------------------------------------------------------}
  procedure D0302CustomerDetailsDC;   // item index = 27
  begin
    UpdateStatusBar('Generating D0302 to DC', 'Y');
    IdentifyD0302('C', 'ALL', 'CHECK'); // HH DC
    IdentifyD0302('D', 'ALL', 'CHECK'); // NHH DC
    CreateOutstandingD0302s;
  end;
  {----------------------------------------------------------------------------}
  procedure D0302CustomerDetailsDist;   // item index = 28
  begin
    UpdateStatusBar('Generating D0302 to Distributor', 'Y');
    IdentifyD0302('R', 'ALL', 'CHECK'); // DIST
    CreateOutstandingD0302s;
  end;
  {----------------------------------------------------------------------------}
  procedure D0302CustomerDetailsMO;   // item index = 29
  begin
    UpdateStatusBar('Generating D0302 to MO', 'Y');
    IdentifyD0302('M', 'ALL', 'CHECK'); // MOP
    IdentifyD0302('S', 'ALL', 'CHECK'); // MOP UMOL SMART METERS
    CreateOutstandingD0302s;
  end;
  {----------------------------------------------------------------------------}
  procedure D0306RequestForDebtInformation;   // item index = 30
  begin
    UpdateStatusBar('Generating D0306', 'Y');
    FRM_Common.Export_ELEC_Supplier_file('D0306');
  end;
  {----------------------------------------------------------------------------}
  procedure D0307DebtInformation;   // item index = 31
  begin
    UpdateStatusBar('Generating D0307', 'Y');
    FRM_Common.Export_ELEC_Supplier_file('D0307');
  end;
  {----------------------------------------------------------------------------}
  procedure D0308ConfirmationOfCustomerDebtTransfer;   // item index = 32
  begin
    UpdateStatusBar('Generating D0308', 'Y');
    FRM_Common.Export_ELEC_Supplier_file('D0308');
  end;
  {----------------------------------------------------------------------------}
  procedure D0309ConfirmationOfDebtAssigned;   // item index = 33
  begin
    UpdateStatusBar('Generating D0309', 'Y');
    FRM_Common.Export_ELEC_Supplier_file('D0309');
  end;
  {----------------------------------------------------------------------------}
  procedure D0311NosieFlow;   // item index = 34
  begin
    UpdateStatusBar('Generating D0311 to Old Supplier', 'Y');
    CreateD0311s;
  end;
  {----------------------------------------------------------------------------}
  procedure D0358RegistrationWithdrawlRequest;   // item index = 35
  begin
    // NOT IMPLEMENTED
  end;
  {----------------------------------------------------------------------------}
  procedure D0381MeteringPointAddressUpdates;   // item index = 36
  begin
    UpdateStatusBar('Generating D0381''s', 'Y');
    FRM_Common.Export_ELEC_Supplier_file('D0381');
  end;
  {----------------------------------------------------------------------------}
  procedure D0386ManageMeteringPointRelationships;   // item index = 37
  var
    q : TOracleDataSet;
  begin
    UpdateStatusBar('Generating D0386''s', 'Y');

    ExportD0386;
    ExportD0386Unrelated;
  end;
  {----------------------------------------------------------------------------}
  procedure D2026DUOSRemittanceAdvice;   // item index = 38
  begin
    UpdateStatusBar('Generating D2026''s', 'Y');
    FRM_Common.Export_ELEC_Supplier_file('D2026');
  end;
  {----------------------------------------------------------------------------}
  procedure PARMSSupplierPARMSReports;
  var
    q           : TOracleDataSet;
    updateQ     : TOracleQuery;
    xd          : string;
    year        : string;
    month       : string;
    mon         : word;
    fileName135 : string;
    fileName142 : string;
  begin
    UpdateStatusBar('Generating PARMS to POOL', 'Y');

    q := TOracleDataSet.Create(nil);
    try
      q.Session := FRM_Login.MainSession;
      q.SQL.Text :=
        'select reporting_period '+
          'from crm.parms_calender '+
          'where due_date <= sysdate and sup_p0135_issued is null';
      q.Open;

      while not q.Eof do
      begin
        xd := q.FieldByName('reporting_period').AsString;

        if Trim(xd) = '' then
        begin
          gLogger.Log('Error in "PARMS to POOL export flow" generating: reporting period in CRM.PARMS_CALENDER is empty.', llError);
          exit;
        end;

        year  := Copy(xd, 1, 4);
        month := Copy(xd, 6, 2);

        try
          mon := StrToInt(month);
        except
          gLogger.Log('Error in "PARMS to POOL export flow" generating: invalid reporting period in CRM.PARMS_CALENDER (%s).', [xd], llError);
          exit;
        end;

        DoSupplierParms_P0135(year, month);
        DoSupplierParms_P0142(year, month);

        updateQ := TOracleQuery.Create(nil);
        try
          updateQ.Session := FRM_Login.MainSession;

          updateQ.SQL.Text :=
            'update crm.parms_calender '+
              'set sup_p0135_issued = sysdate, '+
                  'sup_p0142_issued = sysdate '+
              'where reporting_period = '+QuotedStr(xd);
           updateQ.Execute;
        finally
          FreeAndNil(updateQ);
        end;

        // Now do the email
        if Frm_Email_Comp.SMTP.Connected then
          Frm_Email_Comp.SMTP.Disconnect;
        Frm_Email_Comp.MailMessage.Clear;

        fileName135 := h_Outgoing + 'PARMS\' + X_MPID + '135' + Copy(year, 4, 1) + '.' + UpperCase(FormatSettings.ShortMonthNames[mon]) + '.txt';
        fileName142 := h_Outgoing + 'PARMS\' + X_MPID + '142' + Copy(year, 4, 1) + '.' + UpperCase(FormatSettings.ShortMonthNames[mon]) + '.txt';

        TIdAttachmentFile.Create(Frm_Email_Comp.MailMessage.MessageParts, fileName135);
        TIdAttachmentFile.Create(Frm_Email_Comp.MailMessage.MessageParts, fileName142);


        Frm_Email_Comp.MailMessage.Body.Text                 := 'Find attached PARMS serials for Supplier ' + X_MPID + ', Period: ' + xd;
        Frm_Email_Comp.MailMessage.Recipients.EMailAddresses := FRM_Common.GetValue('INTERNAL_PARMS_TO');
        Frm_Email_Comp.MailMessage.From.Address              := Frm_Common.GetValue('INTERNAL_PARMS_FROM');
        Frm_Email_Comp.MailMessage.Subject                   := 'PARMS serials for Supplier ' + X_MPID + ', Period: ' + xd;

        try
          try
            Frm_Email_Comp.SMTP.Port := 2525;
            Frm_Email_Comp.SMTP.Connect;
            Frm_Email_Comp.SMTP.Send(Frm_Email_Comp.MailMessage);

            // Delete the files
            DeleteFile(fileName135);
            DeleteFile(fileName142);
          except
            on e:Exception do
              gLogger.Log('Email Error PARMS(135, 142) for Supplier: %s; Location: %sPARMS\', [e.Message, h_Outgoing], llError);
          end;
        finally
          if Frm_Email_Comp.SMTP.Connected then
            Frm_Email_Comp.SMTP.Disconnect;
        end;

        q.Next;
      end;

      FRM_Login.MainSession.Commit;
    finally
      FreeAndNil(q);
    end;
  end;
  {----------------------------------------------------------------------------}
var
  iniFile : TRegIniFile;
  q       : TOracleQuery;
  i       : TXCheckListItems;
begin
  PageControl1.ActivePage := TabSheet_EXPORT_EX;

  iniFile := TReginiFile.Create(AppTitle);
  try
    H_outgoing := iniFile.Readstring('File Locations', 'outgoingDflows', 'C:\OUT\');
  finally
    FreeAndNil(iniFile);
  end;

  FRM_FILE_PROGRESS.Caption := 'Exporting Dataflows';
  FRM_FILE_PROGRESS.Update;

  try
  q := TOracleQuery.Create(nil);
    try
      q.Session := FRM_Login.MainSession;

      q.SQL.Text :=
        'update edmgr.flowheaders '+
          'set toname = '+QuotedStr('UMOL') + ' '+
          'where toname <> '+QuotedStr('UMOL')+' and flow_version like '+QuotedStr('PSM%');
      q.Execute;

      q.SQL.Clear;
      q.SQL.Text :=
        'update edmgr.mpan_status '+
          'set requested_mo_id = '+QuotedStr('UMOL')+' '+
          'where regstatus like '+QuotedStr('PSM%');
      q.Execute;
    finally
      FreeAndNil(q);
    end;

    FRM_Login.MainSession.Commit;
  except
    on e:Exception do
    begin
      FRM_Login.MainSession.Rollback;
      gLogger.Log('Error in exporting dataflows: %s', [e.Message], llError);
      MessageDlg(e.Message, mtError, [mbOk], 0);
      exit;
    end;
  end;

  for i := Low(TXCheckListItems) to High(TXCheckListItems) do
  begin
    if XCheckList.Checked[Ord(i)] then
    begin
      case i of
        xchliCOAFChangeOfAgentBatchedFile:            COAFChangeOfAgentBatchedFile;
        xchliD0005ReadingRequests:                    D0005ReadingRequests;
        xchliD0010_D0071ReadingsToDCAndSupplier:      D0010_D0071ReadingsToDCAndSupplier;
        xchliD0052AffirmationOfMeteringSystemDetails: D0052AffirmationOfMeteringSystemDetails;
        xchliD0064ObjectionRequests:                  D0064ObjectionRequests;
        xchliD0131AddressChangeHHDC:                  D0131AddressChangeHHDC;
        xchliD0131AddressChangeMOP:                   D0131AddressChangeMOP;
        xchliD0131AddressChangeNHHDC:                 D0131AddressChangeNHHDC;
        xchliD0132SupplyDisconnectionDeatails:        D0132SupplyDisconnectionDeatails;
        xchliD0142RequestToChangeInstallMetering:     D0142RequestToChangeInstallMetering;
        xchliD0148AgentDetails:                       D0148AgentDetails;
        xchliD0151DisconnectionDA:                    D0151DisconnectionDA;
        xchliD0151DisconnectionDC:                    D0151DisconnectionDC;
        xchliD0151DisconnectionMO:                    D0151DisconnectionMO;
        xchliD0151LossDA:                             D0151LossDA;
        xchliD0151LossDC:                             D0151LossDC;
        xchliD0151LossMO:                             D0151LossMO;
        xchliD0151ObjectionDA:                        D0151ObjectionDA;
        xchliD0151ObjectionDC:                        D0151ObjectionDC;
        xchliD0151ObjectionMO:                        D0151ObjectionMO;
        xchliD0153AppointmentDA:                      D0153AppointmentDA;
        xchliD0155AppointmentDC:                      D0155AppointmentDC;
        xchliD0155AppointmentMO:                      D0155AppointmentMO;
        xchliD0190PrepayKeyRequests:                  D0190PrepayKeyRequests;
        xchliD0205MPASUpdates:                        D0205MPASUpdates;
        xchliD0225SpecialNeeds:                       D0225SpecialNeeds;
        xchliD0301ErroneousTransfers:                 D0301ErroneousTransfers;
        xchliD0302CustomerDetailsDC:                  D0302CustomerDetailsDC;
        xchliD0302CustomerDetailsDist:                D0302CustomerDetailsDist;
        xchliD0302CustomerDetailsMO:                  D0302CustomerDetailsMO;
        xchliD0306RequestForDebtInformation:          D0306RequestForDebtInformation;
        xchliD0307DebtInformation:                    D0307DebtInformation;
        xchliD0308ConfirmationOfCustomerDebtTransfer: D0308ConfirmationOfCustomerDebtTransfer;
        xchliD0309ConfirmationOfDebtAssigned:         D0309ConfirmationOfDebtAssigned;
        xchliD0311NosieFlow:                          D0311NosieFlow;
        xchliD0358RegistrationWithdrawlRequest:       D0358RegistrationWithdrawlRequest;
        xchliD0381MeteringPointAddressUpdates:        D0381MeteringPointAddressUpdates;
        xchliD0386ManageMeteringPointRelationships:   D0386ManageMeteringPointRelationships;
        xchliD2026DUOSRemittanceAdvice:               D2026DUOSRemittanceAdvice;
        xchliPARMSSupplierPARMSReports:               PARMSSupplierPARMSReports;
      end;
    end;
  end;

  PageControl1.ActivePage := TabSheet_EXPORT_EX;
  Caption                 := 'Dataflow Export';
  UpdateStatusBar('', '');
  FRM_File_Progress.Close;
end;

{------------------------------------------------------------------------------}

procedure TFrm_Export.Do_E_MopFiles;
var
Year,month,xd:string;
begin
 pagecontrol1.activepage:=tabsheet_Export_EM;
 Finifile:=TReginiFile.Create(apptitle);
 H_outgoing:=FIniFile.Readstring('File Locations','outgoingDflows','C:\OUT\');
 FiniFile.Free;

 FRM_FILE_PROGRESS.caption:='Exporting Dataflows';


 application.processmessages;


{ If mchecklist.checked[0]=true then
 Begin
  updatestatusbar('Generating EMAILS');
  ExportEmails;
 end;  }

 If mchecklist.checked[1]=true then
 Begin
  UpdateStatusBar('Generating MOP D0002 Responses','Y');
  CreateMOPD0002;
 end;

 If mchecklist.checked[2]=true then
 Begin
  UpdateStatusBar('Generating MOP D0010 Responses','Y');
  CreateMOPD0010;
 end;

 If mchecklist.checked[3]=true then
 Begin
  UpdateStatusBar('Generating MOP D0011 Responses','Y');
  CreateMOPD0011s;
 end;

 If mchecklist.checked[4]=true then
 Begin
  UpdateStatusBar('Generating MOP D0135s','Y');
  CreateMOPD0135;
 end;

  If mchecklist.checked[5]=true then
 Begin
  UpdateStatusBar('Generating MOP D0139s','Y');
  CreateMOPD0139;
 end;

  If mchecklist.checked[6]=true then
 Begin
  UpdateStatusBar('Generating D0149s','Y');
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('delete from mopmgr.flows_to_send where recipient_mpid=''GETW'' and Recipient_Role=''M''');
   execute;
   close;
   sql.clear;
   sql.add('delete from mopmgr.flows_to_send where recipient_mpid is null');
   execute;
  End;
  UpdateStatusBar('Generating D0149s_X','Y');
  ExportD0149_X;
  UpdateStatusBar('Generating D0149s_D','Y');
  ExportD0149_D;
  UpdateStatusBar('Generating D0149s_R','Y');
  ExportD0149_R;
  UpdateStatusBar('Generating D0149s_M','Y');
  ExportD0149_M;
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('delete from mopmgr.flows_to_send where flowversion=''D0149'' and additional_info is not null');
   execute;
   close;
   sql.clear;
   sql.add('delete from mopmgr.flows_to_send where flowversion=''D0149'' and recipient_mpid=''UMOL''');
   execute;
  End;
 end;

 If mchecklist.checked[7]=true then
 Begin
  UpdateStatusBar('Generating D0150s_X','Y');
  ExportD0150_X;
  UpdateStatusBar('Generating D0150s_D','Y');
  ExportD0150_D;
  UpdateStatusBar('Generating D0150s_R','Y');
  ExportD0150_R;
  UpdateStatusBar('Generating D0150s_M','Y');
  ExportD0150_M;

  UpdateStatusBar('Generating D0313s_X','Y');
  ExportD0313_X;
  UpdateStatusBar('Generating D0313s_D','Y');
  ExportD0313_D;
  UpdateStatusBar('Generating D0313s_R','Y');
  ExportD0313_R;
  UpdateStatusBar('Generating D0313s_M','Y');
  ExportD0313_M;

  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('delete from mopmgr.flows_to_send where flowversion=''D0150'' and additional_info is not null');
   execute;
   close;
   sql.clear;
   sql.add('delete from mopmgr.flows_to_send where flowversion=''D0150'' and recipient_mpid=''UMOL''');
   execute;
  End;
 end;

  If mchecklist.checked[8]=true then
 Begin
  UpdateStatusBar('Generating MOP D0170 Requests_M','Y');
  CreateMOPD0170s;
  UpdateStatusBar('Generating MOP D0170 Requests_R','Y');
  CreateDistD0170s;
 end;

   If mchecklist.checked[9]=true then
 Begin
  UpdateStatusBar('Generating MOP D0224','Y');
  CreateMOPD0224;
 end;

 If mchecklist.checked[10]=true then
 Begin
  UpdateStatusBar('Generating MOP D0261 Responses','Y');
  CreateMOPD0261s;
 end;

 If mchecklist.checked[11]=true then
 Begin
  UpdateStatusBar('Rebuilding Data - For D0303','Y');
  FRM_COMMON.Execute_Oracle_Procedure('MOPMGR.PR_REFRESH_MTDS_D0303');
  UpdateStatusBar('Generating D0303 to MAP - Appointment','Y');
  ExportD0303_Appoint;
  UpdateStatusBar('Generating D0303 to MAP - Deappointment','Y');
  ExportD0303_DeAppoint;
 end;

 If mchecklist.checked[12]=true then
 Begin
  UpdateStatusBar('Generating D0304 to MAP (Distributor)','Y');
  ExportD0304_DIST;
  UpdateStatusBar('Generating D0304 to MAP (Supplier)','Y');
  ExportD0304_SUPP;
  UpdateStatusBar('Generating D0304 to MPAS','Y');
  ExportD0304_MPAS;
 end;

 If mchecklist.checked[13]=true then
 Begin
  UpdateStatusBar('Generating D0312 to MPAS','Y');
  ExportD0312_P;
 end;

 If mchecklist.checked[14]=true then
 Begin
  UpdateStatusBar('Generating PARMS to POOL','Y');
  With MOPPARMS Do
  Begin
   close;
   sql.clear;
   sql.add('Select reporting_period from crm.parms_calender where due_date<=sysdate');
   sql.add('and mo_nm12_issued is null');
   open;
   While Not MOPPARMS.eof do
   Begin
    XD:=MOPPARMS.fields[0].text;
    year:=copy(xd,1,4);
    month:=copy(xd,6,2);
    DoMOPParms(copy(xd,1,4),copy(xd,6,2));
    with main_data_module.updatequery do
    Begin
     close;
     sql.clear;
     sql.add('Update crm.parms_calender set mo_NM12_issued=sysdate,mo_sp11_issued=sysdate,mo_sp14_issued=sysdate,mo_sp15_issued=sysdate');
     sql.add('where reporting_period='''+xd+'''');
     execute;
    End;
    // Now Do The Email
    if frm_email_comp.smtp.connected then frm_email_comp.smtp.Disconnect;
    frm_email_comp.mailmessage.clear;
    TIdAttachmentfile.Create(frm_email_comp.mailmessage.MessageParts,h_Outgoing+M_MPID+'234'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt');
    TIdAttachmentfile.Create(frm_email_comp.mailmessage.MessageParts,h_Outgoing+M_MPID+'228'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt');
    TIdAttachmentfile.Create(frm_email_comp.mailmessage.MessageParts,h_Outgoing+M_MPID+'224'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt');
    TIdAttachmentfile.Create(frm_email_comp.mailmessage.MessageParts,h_Outgoing+M_MPID+'227'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt');
    frm_email_comp.mailmessage.body.text:='Find attached PARMS serials for NHHMO '+M_MPID+', Period: '+XD;
    //frm_import.mailmessage.recipients.emailaddresses:=customerservices;
    frm_email_comp.mailmessage.recipients.emailaddresses:=frm_common.GETVALUE('INTERNAL_PARMS_TO');
    frm_email_comp.mailmessage.From.Address := frm_common.GETVALUE('INTERNAL_PARMS_FROM');
    frm_email_comp.mailmessage.Subject :='PARMS serials for NHHMO '+M_MPID+', Period: '+XD;//
    try
     try
      frm_email_comp.smtp.port:=2525;
      frm_email_comp.smtp.connect;
      frm_email_comp.smtp.Send(frm_email_comp.mailmessage);
      except
      end;
    finally
     if frm_email_comp.smtp.connected then frm_email_comp.smtp.Disconnect;
    end;
    // Now Delete Files
    deletefile(h_Outgoing+M_MPID+'234'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt'); //nm12
    deletefile(h_Outgoing+M_MPID+'228'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt'); //sp15
    deletefile(h_Outgoing+M_MPID+'224'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt'); //sp11
    deletefile(h_Outgoing+M_MPID+'227'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt'); //sp14
    MopParms.next
   End;
  End;
 end;

 pagecontrol1.activepage:=tabsheet_Export_EX;
 Caption:='Dataflow Export';
 UpdateStatusBar('','');
 FRM_File_Progress.close;
end;


procedure TFrm_Export.Do_G_SUPPLIER_Files;
Var
supid:string;
begin
 pagecontrol1.activepage:=tabsheet_Export_G;
 Finifile:=TReginiFile.Create(apptitle);
 H_outgoing:=FIniFile.Readstring('File Locations','outgoingDflows','C:\OUT\');
 FiniFile.Free;

 FRM_FILE_PROGRESS.caption:='Exporting Gas Dataflows';

 application.processmessages;

 {If shipperlist.checked[0]=true then
 Begin
  updatestatusbar('Generating 102-ACCU Files','Y');
  FRM_EXPORT_GAS.CreateGas102ACCU;  // Appointment FLows
 end;}

 If gas_supplier_list.checked[1]=true then
 Begin
  updatestatusbar('Generating 105-ACCU Files','Y');
  FRM_EXPORT_GAS.CreateGas105ACCU;
  //FRM_EXPORT_GAS.CreateGas105ACCU_SMART_METER_INSTALLED; // knackered
 end;

 {If shipperlist.checked[2]=true then
 Begin
  FRM_Main.Statusbar.panels[1].text:='Generating 106-ACCU Files';
  FRM_EXPORT_GAS.CreateGas106ACCU;      //Access Instructions
 end; }

{  If shipperlist.checked[3]=true then
 Begin
  FRM_Main.Statusbar.panels[1].text:='Generating 111-ACCU Files';
  FRM_EXPORT_GAS.CreateGas111ACCU;    // Must Reads
 end;   }

 {If shipperlist.checked[4]=true then
 Begin
  FRM_Main.Statusbar.panels[1].text:='Generating 125-ACCU Files';
  FRM_EXPORT_GAS.CreateGas125ACCU;   // Estimated Monthly File Reads
 end;  }
  If gas_supplier_list.checked[5]=true then
 Begin
  updatestatusbar('Generating G0806 Files','Y');
  FRM_COMMON.Export_Gas_Supplier_file('G0806');   //
 end;

  If gas_supplier_list.checked[6]=true then
 Begin
  updatestatusbar('Generating G0807 Files','Y');
  FRM_COMMON.Export_Gas_Supplier_file('G0807');   //
 end;

  If gas_supplier_list.checked[7]=true then
 Begin
  updatestatusbar('Generating G0808 Files','Y');
  FRM_COMMON.Export_Gas_Supplier_file('G0808');   //
 end;

  If gas_supplier_list.checked[8]=true then
 Begin
  updatestatusbar('Generating G0809 Files','Y');
  FRM_COMMON.Export_Gas_Supplier_file('G0809');   //
 end;


 If gas_supplier_list.checked[9]=true then
 Begin
  updatestatusbar('Generating NOSI Files','Y');
  FRM_COMMON.Export_Gas_Supplier_file('NOS');   // NOS FIles
 end;


 If (gas_supplier_list.checked[10]=true) then
 Begin
  updatestatusbar('Generating ONAGE Files','Y');
  SUPID:=frm_common.GETVALUE('GAS_SUPPLIER_ID');
  if SUPID='UEL' then
  Begin
     { Wrike MAM Phase1}
     UpdateStatusbar('Generating MAM Appointment records into the export table','Y');

     with main_data_module.updatequery do
     Begin
      close;
      sql.clear;
      sql.add(' begin');
      sql.add(' GDMGR.PKG_EXPORT_ONAGE.PR_EXPORT_MAM_APPOINTMENT;');
      sql.add(' end;');
      execute;
     End;

     FRM_EXPORT_GAS.CreateGasONAS('GTM');
     FRM_EXPORT_GAS.CreateGasONAS('SGM');
     FRM_EXPORT_GAS.CreateGasONAS('WML');

     FRM_EXPORT_GAS.CreateGasONASdeappointLoss('GTM');
     FRM_EXPORT_GAS.CreateGasONASdeappointLoss('SGM');
     FRM_EXPORT_GAS.CreateGasONASdeappointLoss('WML');
     FRM_EXPORT_GAS.CreateGasONASdeappointLoss('EPM');

     FRM_EXPORT_GAS.CreateGasONASdeappointCASmart('GTM');
     FRM_EXPORT_GAS.CreateGasONASdeappointCASmart('SGM');
     FRM_EXPORT_GAS.CreateGasONASdeappointCASmart('EPM');

     FRM_EXPORT_GAS.CreateGasONASSmart('WML');
  end;
 end;

 If gas_supplier_list.checked[11]=true then
 Begin
  UpdateStatusbar('Generating ORD Files','Y');
 // FRM_EXPORT_GAS.Create_MAM_ORDETS;
 end;

 If gas_supplier_list.checked[12]=true then
 Begin
  Updatestatusbar('Generating RET ET files','Y');
  FRM_EXPORT_GAS.Create_Gas_ETS;
 end;

 If gas_supplier_list.checked[13]=true then
 Begin
  Updatestatusbar('Generating MAM RNA files','Y');
  FRM_EXPORT_GAS.Create_MAM_RNA('UEL');
 end;

 If gas_supplier_list.checked[14]=true then
 Begin
  // First Check If Any files have been created today.
  // Can Only Run this once per day
  with main_data_module.tempquery do
  Begin
   close;
   sql.clear;
   sql.add('select disputes_issued from gdmgr.bis_sar_generated_log where disputes_issued=trunc(sysdate)');
   open;
  End;
  if main_data_module.tempquery.recordcount=1 then
  Begin
  // exit;
  end;
  Updatestatusbar('Cleaning Up Disputes Data','Y');
  FRM_EXPORT_GAS.cleanupDisputesdata;

  Updatestatusbar('Generating GAS Disputes','Y');  // Initial Request from Utilita
  FRM_EXPORT_GAS.Create_Gas_Disputes;
  Updatestatusbar('Generating GAS Returns','Y');   // Responses by Utilita
  FRM_EXPORT_GAS.Create_Gas_Responses;

  // Escalastions no Longer Required, managed via Spreadsheet.
 // Updatestatusbar('Generating GAS Escalations','Y');
 // FRM_EXPORT_GAS.Create_Gas_Escalations;

 end;


 If gas_supplier_list.checked[15]=true then
 Begin
  Updatestatusbar('Generating SQ01 Files','Y');
  FRM_EXPORT_GAS.CreateGasSQ01;
 end;

 If gas_supplier_list.checked[16]=true then
 Begin
  UpdateStatusbar('Generating SQ05 Files','Y');
  FRM_EXPORT_GAS.CreateGasSQ05;
 end;

  If gas_supplier_list.checked[17]=true then
 Begin
  UpdateStatusbar('Generating SQ08 Files','Y');
  FRM_EXPORT_GAS.CreateGasSQ08;
 end;

 If gas_supplier_list.checked[18]=true then
 Begin
  UpdateStatusBar('Generating SQ09 Files','Y');
  FRM_EXPORT_GAS.CreateGasSQ09;
 end;

  If gas_supplier_list.checked[19]=true then
 Begin
  UpdateStatusBar('Generating G0209, G0210 and G0211 Files','Y');
  FRM_EXPORT_GAS.Create_ONJOB_MAM_TO_CDSP;
 end;

 If gas_supplier_list.checked[20]=true then
 Begin
  UpdateStatusBar('Generating G0609 & G0610 Files','Y');
  FRM_EXPORT_GAS.Create_ONUPD_MAM_TO_CDSP;
 end;

 frm_login.mainsession.commit;
 pagecontrol1.activepage:=tab_Exp_Shipper;
 Caption:='Dataflow Export';
 UpdateStatusbar('','');
 FRM_File_Progress.close;
end;

procedure TFRM_Export.Do_G_Shipper_Files;
Var
  supid: string;
begin
  PageControl1.activepage := TabSheet_Export_G;
  Finifile := TRegIniFile.Create(apptitle);
  h_outgoing := Finifile.ReadString('File Locations', 'outgoingDflows',
    'C:\OUT\');
  Finifile.free;

  FRM_File_Progress.caption := 'Exporting Gas Dataflows';

  application.ProcessMessages;

  ////////////////////////////////////////////////////////////
  //  FSP-2218 Merge code.                                  //
  ////////////////////////////////////////////////////////////

  // If gas_shipper_list.checked[0]=true then
  If Gas_Shipper_list.checked[IndexOf('AQI')] then
  Begin
    UpdateStatusBar('Generating AQI Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_AQI;
  end;

  If Gas_Shipper_list.checked[IndexOf('BRN')] then
  Begin
    UpdateStatusBar('Generating BRN Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_BRN_T87;
    FRM_EXPORT_GAS.Create_GAS_BRN_T90;
    FRM_EXPORT_GAS.Create_GAS_BRN_T91;
  end;

  // If gas_shipper_list.checked[1]=true then
  If Gas_Shipper_list.checked[IndexOf('CNC')] then
  Begin
    UpdateStatusBar('Generating CNC Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_CNC_S81;
    FRM_EXPORT_GAS.Create_GAS_CNC_S82;
  end;

  // If gas_shipper_list.checked[2]=true then
  If Gas_Shipper_list.checked[IndexOf('CNF')] then
  Begin
    UpdateStatusBar('Generating CNF Files', 'Y');
    // FRM_EXPORT_GAS.Create_Gas_CNF_S38;
    FRM_EXPORT_GAS.Create_GAS_CNF_T05;
    FRM_EXPORT_GAS.Create_GAS_CNF_SSP;
    FRM_Export_GAS.Create_GAS_CNF_LSP;
  end;

  // If gas_shipper_list.checked[3]=true then
  If Gas_Shipper_list.checked[IndexOf('EMC')] then
  Begin
    UpdateStatusBar('Generating EMC Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_EMC_S51;
  end;

  // If gas_shipper_list.checked[4]=true then
  If Gas_Shipper_list.checked[IndexOf('GEA')] then
  Begin
    UpdateStatusBar('Generating GEA Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_GEA_S96;
  end;

  // If gas_shipper_list.checked[5]=true then
  If Gas_Shipper_list.checked[IndexOf('MAI')] then
  Begin
    UpdateStatusBar('Generating MAI Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_MAI_S93;
  end;

  // If gas_shipper_list.checked[6]=true then
  If Gas_Shipper_list.checked[IndexOf('MAM')] then
  Begin
    UpdateStatusBar('Generating MAM Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_MAM_K08; // K08 FIles
  end;

  // If gas_shipper_list.checked[7]=true then
  If Gas_Shipper_list.checked[IndexOf('MID')] then
  Begin
    UpdateStatusBar('Generating MID Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_MID_N44;
  end;

  // If gas_shipper_list.checked[8]=true then
  If Gas_Shipper_list.checked[IndexOf('MSI')] then
  Begin
    UpdateStatusBar('Generating MSI Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_MSI_T73;
  end;

  if Gas_Shipper_list.checked[IndexOf('NOM')] then
  begin
    UpdateStatusBar('Generating Enquiries Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_NOM_ENQ;
    UpdateStatusBar('Generating Nominations Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_NOM_NOM;
  end;

  // If gas_shipper_list.checked[9]=true then
  If Gas_Shipper_list.checked[IndexOf('ONJ')] then
  Begin
    UpdateStatusBar('Generating ONJ Files', 'Y');
    // FRM_EXPORT_GAS.CreateLibertyExchanges;
    FRM_EXPORT_GAS.Create_Gas_ONJ;
  end;

  // If gas_shipper_list.checked[10]=true then
  if Gas_Shipper_list.checked[IndexOf('ONU')] then
  Begin
    UpdateStatusBar('Generating ONU Files', 'Y');
    // FRM_EXPORT_GAS.CreateLibertyExchanges;
    FRM_EXPORT_GAS.Create_Gas_ONU;
  end;

  // If gas_shipper_list.checked[11]=true then
  if Gas_Shipper_list.checked[IndexOf('ORJ')] then
  Begin
    UpdateStatusBar('Generating ORJ Files', 'Y');
    FRM_EXPORT_GAS.CreateGasORJS;
  end;

  // If gas_shipper_list.checked[12]=true then
  if Gas_Shipper_list.checked[IndexOf('RFA')] then
  Begin
    UpdateStatusBar('Generating RFA files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_RFA_S89;
  end;

  if Gas_Shipper_list.checked[IndexOf('RRP')] then
  Begin
    UpdateStatusBar('Generating RRP files', 'Y');
    FRM_EXPORT_GAS.CreateRRP;
  end;

  // If gas_shipper_list.checked[13]=true then
  if Gas_Shipper_list.checked[IndexOf('SFN')] then
  Begin
    UpdateStatusBar('Generating SFN O15 Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_SFN_O15;
    UpdateStatusBar('Generating SFN O17 Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_SFN_O17;
  end;

//  If Gas_Shipper_list.checked[14] = True then
  if Gas_Shipper_list.checked[IndexOf('SPC')] then
  Begin
    UpdateStatusBar('Generating SPC C38 Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_SPC_C38;
    UpdateStatusBar('Generating SPC C39 Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_SPC_C39;
    UpdateStatusBar('Generating SPC S34 Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_SPC_S34;
    UpdateStatusBar('Generating SPC S35 Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_SPC_S35;
  end;

//  If Gas_Shipper_list.checked[15] = True then
  if Gas_Shipper_list.checked[IndexOf('SPI')] then
  Begin
    UpdateStatusBar('Generating SPI Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_SPI_X99;
  end;

//  If Gas_Shipper_list.checked[16] = True then
  if Gas_Shipper_list.checked[IndexOf('UBR')] then
  Begin
    UpdateStatusBar('Generating UBR Files', 'Y');
    FRM_EXPORT_GAS.Create_Gas_UBR;
  end;

//  If Gas_Shipper_list.checked[17] = True then
  if Gas_Shipper_list.checked[IndexOf('UDR')] then
  Begin
    UpdateStatusBar('Generating UDR Files', 'Y');
    FRM_EXPORT_GAS.Create_Gas_UDR;
  end;

//  If Gas_Shipper_list.checked[18] = True then
  if Gas_Shipper_list.checked[IndexOf('UMR')] then
  Begin
    UpdateStatusBar('Generating UMR Files', 'Y');
    // FRM_EXPORT_GAS.CreateGasReads;
    FRM_EXPORT_GAS.Create_Gas_UMR;
  end;

//  If Gas_Shipper_list.checked[19] = True then
  if Gas_Shipper_list.checked[IndexOf('WAO')] then
  Begin
    UpdateStatusBar('Generating WAO S40 Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_WAO_S40; // Objection
    UpdateStatusBar('Generating WAO S41 Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_WAO_S41; // Objection Removal;
    UpdateStatusBar('Generating WAO S73 Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_WAO_S73; // Objection Cancellation and Withdrawl;
    UpdateStatusBar('Generating WAO S39 Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_WAO_S39; // Withdrawal;
    UpdateStatusBar('Generating WAO S54 Files', 'Y');
    FRM_EXPORT_GAS.Create_GAS_WAO_S54; // Withdrawal;
  end;

 frm_login.mainsession.commit;
 pagecontrol1.activepage:=tabsheet_Export_EX;
 Caption:='Dataflow Export';
 UpdateStatusBar('','Y');
 FRM_File_Progress.close;
end;


procedure TFrm_Export.DoMOPPARMS(YEAR,MONTH:string);
begin
 //DOMOPPARMS_SP05(YEAR,MONTH);
 //DOMOPPARMS_SP06(YEAR,MONTH);
 //DOMOPPARMS_NM03(YEAR,MONTH);
 //DOMOPPARMS_NM04(YEAR,MONTH);
 DOMOPPARMS_NM12(YEAR,MONTH);
 DOMOPPARMS_SP11(YEAR,MONTH);
 DOMOPPARMS_SP14(YEAR,MONTH);
 DOMOPPARMS_SP15(YEAR,MONTH);
End;

procedure TFrm_Export.DoMOPPARMS_SP05(YEAR,MONTH:string);
Var
Oldsupplier,supplier,monthend:string;
begin
 with generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select B.MON REP_MONTH,B.FROM_NAME SUPPLIER,B.TOT RECEIVED,nvl(A.LATE_MPANS,0)LATE ,nvl(A.ADATE,0) AVG_DAYS_LATE');
  sql.add('from');
  sql.add('(select To_CHAR(F.FILE_DATE_TIME,''YYYY-MM'') MON, F.FROM_NAME,count(*) LATE_MPANS, avg(F.FILE_DATE_TIME-M.EFSD) ADATE from MOPMGR.d0155 M, MOPMGR.FLOWHEADERS F');
  sql.add('where M.response_type=''D0011''');
  sql.add('and F.FILENAME=M.FILENAME');
  sql.add('and F.MPANCORE=M.MPANCORE');
  sql.add('and F.FLOW_VERSION=''D0155''');
  sql.add('and f.file_date_time>M.EFSD');
  sql.add('Group by To_CHAR(F.FILE_DATE_TIME,''YYYY-MM''), F.FROM_NAME) A,');
  sql.add('(select To_CHAR(F.FILE_DATE_TIME,''YYYY-MM'') MON,F.FROM_NAME,count(*) TOT from mopmgr.flowheaders F');
  sql.add('where f.flow_version=''D0155'' group by To_CHAR(F.FILE_DATE_TIME,''YYYY-MM''), F.FROM_NAME) B');
  sql.add('where b.from_name=a.from_name (+)');
  sql.add('and b.mon=a.mon (+)');
  sql.add('and B.MON='''+YEAR+'-'+MONTH+'''');
  sql.add('order by 1,2');
  open;
 End;
 oldsupplier:='';
 outputfilename:=h_Outgoing+M_MPID+'143'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)]+'.txt');
 CreateFlowHeaderMOPPARMS('P0143001',M_MPIDROLE,M_MPID,generalquery.fields[0].text);
 OutputFlowFlowcount:=0;
 OutputFlowLineCount:=0;
 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='Supplier: '+Generalquery.fields[1].text;
  FRM_Main.statusbar.panels[1].text:='Creating SP05 Parms Report ';
  application.processmessages;
  supplier:= generalquery.fields[1].text;
  if oldsupplier<>supplier then
  Begin
   s:='SUB|B|X|'+SUPPLIER+'|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)))+'|M';
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
  End;
  s:='SP5|'+Generalquery.fields[3].text+'|'+floattostr(roundto(strtofloat(Generalquery.fields[4].text),-1))	; // Change to .x
  inc(OutputFlowLineCount);
  inc(OutputFlowFlowcount);
  parmsfile.lines.add(s);
  oldsupplier:=supplier;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 CreateFlowFooterPARMS;
end;

procedure TFrm_Export.DoMOPPARMS_SP06(YEAR,MONTH:string);
Var
Oldsupplier,supplier,monthend:string;
begin
 with generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select B.MON REP_MONTH,B.FROM_NAME SUPPLIER,B.GSP_GROUP,B.TOT D0148_RECEIVED,nvl(A.LATE_MPANS,0)LATE');
  sql.add('from');
  sql.add('(select To_CHAR(F.FILE_DATE_TIME,''YYYY-MM'') MON, F.FROM_NAME,count(*) LATE_MPANS');
  sql.add('from MOPMGR.D0148_272 M,MOPMGR.D0148_271 G, MOPMGR.FLOWHEADERS F');
  sql.add('where F.FILEID=M.FILEID');
  sql.add('and F.FILENAME=M.FILENAME');
  sql.add('and F.MPANCORE=M.MPANCORE ');
  sql.add('and F.FLOW_VERSION=''D0148''');
  sql.add('and G.AGENT_STATUS=''N''');
  sql.add('and G.FILEID=M.FILEID');
  sql.add('and G.FILENAME=M.FILENAME');
  sql.add('and G.MPANCORE=M.MPANCORE');
  sql.add('and f.file_date_time>M.EFD_DCA');
  sql.add('Group by To_CHAR(F.FILE_DATE_TIME,''YYYY-MM''), F.FROM_NAME) A,');
  sql.add('(select To_CHAR(F.FILE_DATE_TIME,''YYYY-MM'') MON,F.FROM_NAME,L.GSP_GROUP,count(*) TOT');
  sql.add('from mopmgr.flowheaders F,MOPMGR.D0148_270 H,MOPMGR.D0155 L');
  sql.add('where f.flow_version=''D0148''');
  sql.add('and F.FILENAME=H.FILENAME');
  sql.add('and F.FILEID=H.FILEID');
  sql.add('and F.MPANCORE=H.MPANCORE');
  sql.add('and L.SUPPLIER_MPID=F.FROM_NAME');
  sql.add('and L.EFSD=H.EFSD');
  sql.add('and L.MPANCORE=H.MPANCORE');
  sql.add('group by To_CHAR(F.FILE_DATE_TIME,''YYYY-MM''), F.FROM_NAME,L.GSP_GROUP) B');
  sql.add('where b.from_name=a.from_name (+)');
  sql.add('and b.mon=a.mon (+)');
  sql.add('and B.MON='''+YEAR+'-'+MONTH+'''');
  sql.add('order by 1,2,3');
  open;
 End;
 oldsupplier:='';
 outputfilename:=h_Outgoing+M_MPID+'144'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt';
 CreateFlowHeaderMOPPARMS('P0144001',M_MPIDROLE,M_MPID,generalquery.fields[0].text);
 OutputFlowFlowcount:=0;
 OutputFlowLineCount:=0;
 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='Supplier: '+Generalquery.fields[1].text;
  FRM_Main.statusbar.panels[1].text:='Creating SP06 Parms Report ';
  application.processmessages;
  supplier:= generalquery.fields[1].text;
  if oldsupplier<>supplier then
  Begin
   s:='SUB|B|X|'+SUPPLIER+'|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)))+'|M';
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
  End;
  s:='SP6|'+Generalquery.fields[2].text+'|'+Generalquery.fields[4].text; // Change to .x
  inc(OutputFlowLineCount);
  inc(OutputFlowFlowcount);
  parmsfile.lines.add(s);
  oldsupplier:=supplier;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 CreateFlowFooterPARMS;
end;

procedure TFrm_Export.DoMOPPARMS_NM03(YEAR,MONTH:string);
Var
Oldsupplier,supplier,monthend:string;
begin
 // Populate Parms Data
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('insert into mopmgr.parms_temp_NM03');
  sql.add('select  G.DC_ID DC_ID, M.EFD_DCA EFD_DCA,F.FROM_NAME SUPPLIER,G.MPANCORE,NULL');
  sql.add('from MOPMGR.D0148_272 M, MOPMGR.D0148_271 G, MOPMGR.D0148_270 A, MOPMGR.FLOWHEADERS F');
  sql.add('where F.FILEID=M.FILEID');
  sql.add('and F.FILENAME=M.FILENAME');
  sql.add('and F.MPANCORE=M.MPANCORE');
  sql.add('and F.FLOW_VERSION=''D0148''');
  sql.add('and F.FROMID=''X''');
  sql.add('and G.AGENT_STATUS=''N''');
  sql.add('and G.FILEID=M.FILEID');
  sql.add('and G.FILENAME=M.FILENAME');
  sql.add('and G.MPANCORE=M.MPANCORE');
  sql.add('and A.FILEID=M.FILEID');
  sql.add('and A.FILENAME=M.FILENAME');
  sql.add('and A.MPANCORE=M.MPANCORE');
  sql.add('and (G.DC_ID,M.EFD_DCA,F.FROM_NAME,G.MPANCORE)');
  sql.add('not in');
  sql.add('(select NEW_DC_ID,EFD_DCA,SUPPLIER_ID,MPAN');
  sql.add('from mopmgr.parms_temp_NM03)');
  execute;
   // update Date D0150 sent
  close;
  sql.clear;
  sql.add('UPDATE mopmgr.parms_temp_nm03 F');
  sql.add('SET (DATE_D0150_SENT_TO_DC)=');
  sql.add('(SELECT DISTINCT D.FILE_DATE_TIME');
  sql.add('from MOPMGR.FLOWHEADERS D');
  sql.add('WHERE');
  sql.add(' F.MPAN=D.MPANCORE');
  sql.add('and F.NEW_DC_ID=D.TONAME');
  sql.add('and D.TOID=''D''');
  sql.add('and D.FROMID=''M''');
  sql.add('and d.FLOW_VERSION=''D0150''');
  sql.add(' and D.FILE_DATE_TIME<=F.EFD_DCA+14');
  sql.add('and (d.file_date_time,d.mpancore,d.toid,d.fromid) in ');
  sql.add('(select min(file_date_time),mpancore,toid,fromid');
  sql.add('from mopmgr.flowheaders d where flow_version=''D0150''');
  sql.add('group by mpancore,toid,fromid)');
  sql.add(')');
  sql.add('Where F.DATE_D0150_SENT_TO_DC is null');
  execute;
 end;

 with generalquery do
 Begin
  close;
  sql.clear;
  //get query
  sql.add('select');
  sql.add('to_char(EFD_DCA,''YYYY-MM'') PERIOD,');
  sql.add('SUPPLIER_ID,');
  sql.add('COUNT(EFD_DCA) D0148_IN,');
  sql.add('COUNT(EFD_DCA)-COUNT(DATE_D0150_SENT_TO_DC)  PENDING,');
  sql.add('100-((COUNT(EFD_DCA)-COUNT(DATE_D0150_SENT_TO_DC))*100/COUNT(EFD_DCA)) PC ');
  sql.add('from mopmgr.parms_temp_nm03');
  sql.add('where to_char(EFD_DCA,''YYYY-MM'')='''+YEAR+'-'+MONTH+'''');
  sql.add('group by');
  sql.add('to_char(EFD_DCA,''YYYY-MM''),');
  sql.add('NEW_DC_ID,');
  sql.add('SUPPLIER_ID');
  sql.add('order by 1,2');
  open;
 End;
 oldsupplier:='';
 outputfilename:=h_Outgoing+M_MPID+'156'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt';
 CreateFlowHeaderMOPPARMS('P0156001',M_MPIDROLE,M_MPID,generalquery.fields[0].text);
 OutputFlowFlowcount:=0;
 OutputFlowLineCount:=0;
 s:='SUB|N|M|'+M_MPID+'|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)))+'|M';
 inc(OutputFlowLineCount);
 parmsfile.lines.add(s);
 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='Supplier: '+Generalquery.fields[1].text;
  FRM_Main.statusbar.panels[1].text:='Creating NM03 Parms Report ';
  application.processmessages;
  supplier:= generalquery.fields[1].text;
  if oldsupplier<>supplier then
  Begin
   s:='NM3|'+Generalquery.fields[1].text+'|'+Generalquery.fields[2].text+'|'+Generalquery.fields[3].text+'|';
   s:=s+floattostr(roundto(strtofloat(Generalquery.fields[4].text),-2));
   inc(OutputFlowLineCount);
   inc(OutputFlowFlowcount);
   parmsfile.lines.add(s);
  end;
   oldsupplier:=supplier;
   GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 CreateFlowFooterPARMS;

end;

Procedure TFRM_Export.DoSupplierParms_P0135(YEAR,MONTH:string);
var
oldmo:string;
ParmsDir: string;
begin
 with generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select distinct gsp_group_id from mddworking.gsp_group');
  sql.add('order by 1');
  open;
 end;
 //oldsupplier:='';
 ParmsDir := H_OUTGOING+'PARMS\';
 if not DirectoryExists(ParmsDir) then
  CreateDir(ParmsDir);
 outputfilename:=ParmsDir+X_MPID+'135'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt';
 CreateFlowHeaderMOPPARMS('P0135001',X_MPIDROLE,X_MPID,generalquery.fields[0].text);
 OutputFlowFlowcount:=0;
 OutputFlowLineCount:=0;
 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='GSP: '+Generalquery.fields[0].text;
  FRM_Main.statusbar.panels[1].text:='Creating DPI Parms Report ';
  //////////////////////////////////////////////////////////////////////////////
  // Changes for July 11 Implementation
  // http://www.elexon.co.uk/ELEXON%20Documents/bscp533_v17.0.pdf
  //////////////////////////////////////////////////////////////////////////////

  // Supplier Serials
  s:='DPI|'+generalquery.fields[0].text+'|SP04|'+X_MPID+'|X|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
  inc(OutputFlowLineCount);
  parmsfile.lines.add(s);
  application.processmessages;

  // NHH DC Serials
  with Tempquery do
  Begin
   close;
   sql.clear;
   sql.add('select distinct confirmed_dc_id from edmgr.mpan_status');
   sql.add('where ssd<add_months(to_date(''01/'+month+'/'+year+''',''DD/MM/YYYY''),1) and confirmed_dc_role=''D'' and gsp_group ='''+generalquery.fields[0].text+''' and confirmed_dc_id is not null and (regstatus in (''REGISTERED'',''FUTURE_LOSS'',''LOSS PENDING'')');
   sql.add('or (regstatus=''LOST'' and EFTSSD>=to_date(''01/'+month+'/'+year+''',''DD/MM/YYYY'')))');
   sql.add('order by 1');
   open;
  end;
  while not Tempquery.eof do
  Begin
   s:='DPI|'+generalquery.fields[0].text+'|SP11|'+tempquery.fields[0].text+'|D|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|SP12|'+tempquery.fields[0].text+'|D|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|SP13|'+tempquery.fields[0].text+'|D|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|SP15|'+tempquery.fields[0].text+'|D|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|NM11|'+tempquery.fields[0].text+'|D|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|NM12|'+tempquery.fields[0].text+'|D|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|NC11|'+tempquery.fields[0].text+'|D|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   Tempquery.next;
  END;

   // HH DC Serials
  with Tempquery do
  Begin
   close;
   sql.clear;
   sql.add('select distinct confirmed_dc_id from edmgr.mpan_status');
   sql.add('where ssd<add_months(to_date(''01/'+month+'/'+year+''',''DD/MM/YYYY''),1) and confirmed_dc_role=''C'' and gsp_group ='''+generalquery.fields[0].text+''' and confirmed_dc_id is not null and (regstatus in (''REGISTERED'',''FUTURE_LOSS'',''LOSS PENDING'')');
   sql.add('or (regstatus=''LOST'' and EFTSSD>=to_date(''01/'+month+'/'+year+''',''DD/MM/YYYY'')))');
   sql.add('order by 1');
   open;
  end;
  while not Tempquery.eof do
  Begin
   s:='DPI|'+generalquery.fields[0].text+'|SP11|'+tempquery.fields[0].text+'|C|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|SP12|'+tempquery.fields[0].text+'|C|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|SP13|'+tempquery.fields[0].text+'|C|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|SP15|'+tempquery.fields[0].text+'|C|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|HM11|'+tempquery.fields[0].text+'|C|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|HM12|'+tempquery.fields[0].text+'|C|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|HM13|'+tempquery.fields[0].text+'|C|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   s:='DPI|'+generalquery.fields[0].text+'|HM14|'+tempquery.fields[0].text+'|C|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
   Tempquery.next;
  END;

   // Meter Operator Details
  with Tempquery do
  Begin
   close;
   sql.clear;
   sql.add('select distinct confirmed_mo_id,case when measurement_class in(''C'',''E'',''G'',''F'') then ''HM12'' else ''NM12'' end typecode from edmgr.mpan_status');
   sql.add('where ssd<add_months(to_date(''01/'+month+'/'+year+''',''DD/MM/YYYY''),1) and gsp_group ='''+generalquery.fields[0].text+''' and confirmed_mo_id is not null and (regstatus in (''REGISTERED'',''FUTURE_LOSS'',''LOSS PENDING'')');
   sql.add('or (regstatus=''LOST'' and EFTSSD>=to_date(''01/'+month+'/'+year+''',''DD/MM/YYYY'')))');

   sql.add('and confirmed_mo_id<>''SWEB'''); // Added 03/09/2012 Exclude SWEB no longer ACTIVE

   sql.add('order by 1,2');
   open;
  end;

  oldmo:='';
  while not Tempquery.eof do
   Begin
     if tempquery.fields[0].text<>oldmo then
     begin
       s:='DPI|'+generalquery.fields[0].text+'|SP11|'+tempquery.fields[0].text+'|M|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
      inc(OutputFlowLineCount);
      parmsfile.lines.add(s);
      s:='DPI|'+generalquery.fields[0].text+'|SP14|'+tempquery.fields[0].text+'|M|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
      inc(OutputFlowLineCount);
      parmsfile.lines.add(s);
      s:='DPI|'+generalquery.fields[0].text+'|SP15|'+tempquery.fields[0].text+'|M|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
      inc(OutputFlowLineCount);
      parmsfile.lines.add(s);
     end;
     oldmo:=tempquery.fields[0].text;
     s:='DPI|'+generalquery.fields[0].text+'|'+tempquery.fields[1].text+'|'+tempquery.fields[0].text+'|M|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);

   parmsfile.lines.add(s);
   {s:='DPI|'+generalquery.fields[0].text+'|HM12|'+tempquery.fields[0].text+'|M|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)));
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);}
   Tempquery.next;
  end;
 GeneralQuery.Next;
 end; // Repeat for remaining GSPS
 CreateFlowFooterPARMS;
end;

Procedure TFRM_Export.DoSupplierParms_P0142(YEAR,MONTH:string);
var
  ParmsDir: string;
begin
 with generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select distinct gsp_group_id from mddworking.gsp_group');
  sql.add('order by 1');
  open;
 end;
 //oldsupplier:='';
 ParmsDir := H_OUTGOING+'PARMS\';
 if not DirectoryExists(ParmsDir) then
  CreateDir(ParmsDir);
 outputfilename:=ParmsDir+X_MPID+'142'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt';
 CreateFlowHeaderMOPPARMS('P0142001',X_MPIDROLE,X_MPID,generalquery.fields[0].text);
 OutputFlowFlowcount:=0;
 OutputFlowLineCount:=0;
 s:='SUB|H|X|'+X_MPID+'|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)))+'|M';
 inc(OutputFlowLineCount);
 parmsfile.lines.add(s);
 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='GSP: '+Generalquery.fields[0].text;
  FRM_Main.statusbar.panels[1].text:='Creating P0142 Parms Report ';
  s:='SP4|'+generalquery.fields[0].text+'||||';
  inc(OutputFlowLineCount);
  parmsfile.lines.add(s);
  application.processmessages;
  GeneralQuery.Next;
 end; // Repeat for remaining GSPS
 CreateFlowFooterPARMS;
end;


procedure TFrm_Export.DoMOPPARMS_NM04(YEAR,MONTH:string);
Var
Oldsupplier,supplier,monthend:string;
begin
 // Populate Parms Data
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  // Get D0170s from NEW MOP where action code='06'
  sql.add('insert into mopmgr.parms_temp_NM04');
  sql.add('select distinct F.FROM_NAME,H.DATE_ACTION_REQUIRED_BY MON, NULL,H.MPANCORE,null');
  sql.add('from');
  sql.add('mopmgr.flowheaders F,');
  sql.add('MOPMGR.D0170 H');
  sql.add('where f.flow_version=''D0170''');
  sql.add('and F.FILENAME=H.FILENAME');
  sql.add('and F.MPANCORE=H.MPANCORE');
  sql.add('and H.REQUESTED_ACTION_CODE=''06''');
  sql.add('and H.NEW_MO=F.FROM_NAME');
  sql.add('and F.FROMID=''M''');
  // and date requested is before cut off period
  sql.add('and (F.FROM_name,H.DATE_ACTION_REQUIRED_BY,H.MPANCORE)');
  sql.add('not in');
  sql.add('(select NEW_MOP_ID,DATE_ACTION_REQUIRED_BY,MPAN from');
  sql.add('mopmgr.parms_temp_nm04)');
  execute;
  // Update Last Known Supplier ID
  close;
  sql.clear;
  sql.add('UPDATE mopmgr.parms_temp_nm04 F');
  sql.add('SET (SUPPLIER_ID)=');
  sql.add('(SELECT DISTINCT max(D.SUPPLIER_MPID)');
  sql.add('from MOPMGR.D0155 D');
  sql.add('WHERE');
  sql.add(' F.MPAN=D.MPANCORE');
  sql.add(' and f.DATE_ACTION_REQUIRED_BY>D.EFSD');
  sql.add(')');
  sql.add('Where F.supplier_id is null');
  execute;
  // update last known supplier is above is null
  close;
  sql.clear;
  sql.add('UPDATE mopmgr.parms_temp_nm04 F');
  sql.add('SET (SUPPLIER_ID)=');
  sql.add('(SELECT DISTINCT max(D.SUPPLIER_MPID)');
  sql.add('from MOPMGR.D0155 D');
  sql.add('WHERE');
  sql.add(' F.MPAN=D.MPANCORE');
  sql.add(')');
  sql.add('Where F.supplier_id is null');
  execute;
  close;
  sql.clear;
  sql.Add('update mopmgr.parms_temp_nm04 set supplier_id=''GETW'' where supplier_id is null');
  execute;

  // update Date D0150 sent
  close;
  sql.clear;
  sql.add('UPDATE mopmgr.parms_temp_nm04 F');
  sql.add('SET (DATE_D0150_SENT_TO_MOP)=');
  sql.add('(SELECT DISTINCT D.FILE_DATE_TIME');
  sql.add('from MOPMGR.FLOWHEADERS D');
  sql.add('WHERE');
  sql.add(' F.MPAN=D.MPANCORE');
  sql.add('and F.NEW_MOP_ID=D.TONAME');
  sql.add('and D.TOID=''M''');
  sql.add('and D.FROMID=''M''');
  sql.add('and D.FLOW_VERSION=''D0150''');
  sql.add(' and D.FILE_DATE_TIME<=F.DATE_ACTION_REQUIRED_BY+14');
  sql.add('and (d.file_date_time,d.mpancore,d.toid,d.fromid) in ');
  sql.add('(select min(file_date_time),mpancore,toid,fromid');
  sql.add('from mopmgr.flowheaders d where flow_version=''D0150''');
  sql.add('group by mpancore,toid,fromid)');
  sql.add(')');
  sql.add('Where F.DATE_D0150_SENT_TO_MOP is null');
  execute;
 end;

 with generalquery do
 Begin
  close;
  sql.clear;
  //get query
  sql.add('select');
  sql.add('to_char(DATE_ACTION_REQUIRED_BY,''YYYY-MM'') PERIOD,');
  sql.add('SUPPLIER_ID,');
  sql.add('COUNT(DATE_ACTION_REQUIRED_BY) D0170_IN,');
  sql.add('COUNT(DATE_ACTION_REQUIRED_BY)-COUNT(DATE_D0150_SENT_TO_MOP)  PENDING,');
  sql.add('100-((COUNT(DATE_ACTION_REQUIRED_BY)-COUNT(DATE_D0150_SENT_TO_MOP))*100/COUNT(DATE_ACTION_REQUIRED_BY)) PC ');
  sql.add('from mopmgr.parms_temp_nm04');
  sql.add('where to_char(DATE_ACTION_REQUIRED_BY,''YYYY-MM'')='''+YEAR+'-'+MONTH+'''');
  sql.add('group by');
  sql.add('to_char(DATE_ACTION_REQUIRED_BY,''YYYY-MM''),');
  sql.add('NEW_MOP_ID,');
  sql.add('SUPPLIER_ID');
  sql.add('order by 1,2');
  open;
 End;
 oldsupplier:='';
 outputfilename:=h_Outgoing+M_MPID+'157'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt';
 CreateFlowHeaderMOPPARMS('P0157001',M_MPIDROLE,M_MPID,generalquery.fields[0].text);
 OutputFlowFlowcount:=0;
 OutputFlowLineCount:=0;
 s:='SUB|N|M|'+M_MPID+'|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)))+'|M';
 inc(OutputFlowLineCount);
 parmsfile.lines.add(s);
 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='Supplier: '+Generalquery.fields[1].text;
  FRM_Main.statusbar.panels[1].text:='Creating NM04 Parms Report ';
  application.processmessages;
  supplier:= generalquery.fields[1].text;
  if oldsupplier<>supplier then
  begin
   s:='NM4|'+Generalquery.fields[1].text+'|'+Generalquery.fields[2].text+'|'+Generalquery.fields[3].text+'|';
   s:=s+floattostr(roundto(strtofloat(Generalquery.fields[4].text),-2));
   inc(OutputFlowLineCount);
   inc(OutputFlowFlowcount);
   parmsfile.lines.add(s);
  end;
  oldsupplier:=supplier;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 CreateFlowFooterPARMS;
end;




procedure TFRM_Export.CreateflowheaderMOPPARMS(FlowVersion,role,mpid,MON:string);
begin
 FRM_File_Progress.setforExport;
 frm_busy.Close;
 // Write header record
 OutputFlowIdentifier:=frm_common.nextfileidMOP;
 S:='ZHD|';
 S:=s+FLOWVERSION+'|';
 s:=s+ROLE+'|';                          // FromRole
 s:=s+MPID+'|';                              // From ID
 s:=s+'Z|';                                    // Recipient Role Z
 if H_Mode='TEST' then s:=s+H_REC+'|'
 else s:=s+ 'POOL|';                           // Recipient Name POOL
 s:=s+formatdatetime('YYYYMMDDHHNNSS',now);   // Create Date
 parmsfile.lines.clear;
 parmsfile.lines.add(s);
 OutputflowLineCount:=0;
 OutputFlowFlowCount:=0;

 with FRM_File_Progress do
 Begin
  fileprogressbar.position:=0;
  fileprogressbar.max:=100;
  Caption:='Generating '+copy(flowversion,1,5)+'''s';
  l_filename.caption:=outputflowidentifier+'.usr';
  l_fileid.caption:=outputflowidentifier;
  l_flowversion.caption:=flowversion;
  l_FromRole.caption:=M_mpidrole;
  l_FromMPID.caption:=M_mpid;
  l_ToRole.caption:='Z';
  l_ToMPID.caption:=h_rec;
  l_filedatetime.caption:=formatdatetime('YYYYMMDDHHNNSS',now);
  if H_Mode='TEST' then l_tompid.caption:=H_REC
  else l_tompid.caption:='POOL';
  l_testflag.caption:=h_testflag;
  l_testflag.caption:=h_testflag;
  application.processmessages;
 end;
end;


procedure TFRM_Export.CreateFlowFooterPARMS;
begin
 inc(OutPutFlowLineCount); // Include Header
 inc(OutPutFlowLineCount); // Include Footer
 s:='ZPT|';
 s:=s+inttostr(OutPutFlowLineCount)+'|';
 s:=s+'CHECKSUM|';
 parmsfile.lines.add(s);
 FRM_File_Progress.clearlabels;
 DOPARMSCHECKSUM;
 PARMSFILE.Lines.SaveToFile(OUTPUTfilename);
end;

procedure TFRM_Export.DOPARMSCHECKSUM;
Var
z,i,j,a,b,c,d,a1,b1,c1,d1,value,NoLines:integer;
s1,s,filechk:string;
chk:integer;
begin
 chk:=0;
 a:=0;
 b:=0;
 c:=0;
 d:=0;
 a1:=0;
 b1:=0;
 c1:=0;
 d1:=0;
 z:=0;
 NoLines:=parmsfile.lines.count;
 for z :=0 to nolines-2 do
 Begin
  s:=parmsfile.lines[z];
  i:=1;
  While i<= length(s) do
  Begin
   if i+0<=length(s) then a1:=ord(s[i+0]) else a1:=0;
   if i+1<=length(s) then B1:=ord(s[i+1]) else b1:=0;
   if i+2<=length(s) then C1:=ord(s[i+2]) else c1:=0;
   if i+3<=length(s) then D1:=ord(s[i+3]) else d1:=0;
   i:=i+4;
   a:=a xor a1;
   b:=b xor b1;
   c:=c xor c1;
   d:=d xor d1;
  End;

 end;

 chk:=chk+(16777216*a);
 chk:=chk+(65536*b);
 chk:=chk+(256*c);
 chk:=chk+(1*d);

 s:=parmsfile.lines[z];
 filechk:=copy(s,length(s)-8,9);

 OUTPUTMEMO.lines.clear;
 for z :=0 to nolines-2 do
 Begin
   OUTPUTMEMO.lines.Add(PARMSFILE.Lines[z])
 End;
 s1:=stringreplace(s,filechk,inttostr(chk),[rfreplaceall]);
 OUTPUTMEMO.lines.add(S1);
 PARMSFILE.Lines:=OUTPUTMEMO.lines;
end;

procedure TFRM_Export.ExportD0304_DIST;
Begin
 // Create D0304
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select distinct');
  sql.add('''R'',');
  sql.add('l.ldso_mpid,');
  sql.add('c.new_map,');
  sql.add('c.date_of_change,');
  sql.add('c.mpancore,');
  sql.add('c.meterid,');
  sql.add('upper(m.man_make_type),');
  sql.add('m.timing_device_id');
  sql.add('from mopmgr.d0150_map_change C,');
  sql.add('mopmgr.d0150_290 M,');
  sql.add('mopmgr.mpan_status R,');
  sql.add('mddworking.ldso L');
  sql.add('where');
  sql.add('substr(c.mpancore,1,2)=l.mpanstart and');
  sql.add('c.mpancore=r.mpancore and');
  sql.add('c.mpancore=m.mpancore');
  sql.add('and m.meter_asset_provider_id=c.new_map');
  sql.add('and (c.mpancore,c.date_of_change,c.new_map) not in');
  sql.add('(select mpancore,efd_mapa,map_id from mopmgr.D0304');
  sql.add('where filename like ''R%'')');
  sql.add('order by 2,3,4,5');
  open;
 End;
 if GeneralQuery.recordcount=0 then exit;
 ExportThisD0304;
end;

procedure TFRM_Export.ExportD0304_SUPP;
Begin
 // Create D0304
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select distinct');
  sql.add('''X'',');
  sql.add('R.Supplier_mpid,');
  sql.add('c.new_map,');
  sql.add('c.date_of_change,');
  sql.add('c.mpancore,');
  sql.add('c.meterid,');
  sql.add('upper(m.man_make_type),');
  sql.add('m.timing_device_id');
  sql.add('from mopmgr.d0150_map_change C,');
  sql.add('mopmgr.d0150_290 M,');
  sql.add('mopmgr.mpan_status R');
  sql.add('where');
  sql.add('c.mpancore=r.mpancore and');
  sql.add('c.mpancore=m.mpancore');
  sql.add('and m.meter_asset_provider_id=c.new_map');
  sql.add('and (c.mpancore,c.date_of_change,c.new_map) not in');
  sql.add('(select mpancore,efd_mapa,map_id from mopmgr.D0304');
  sql.add('where filename like ''R%'')');
  sql.add('order by 2,3,4,5');
  open;
 End;
 if GeneralQuery.recordcount=0 then exit;
 ExportThisD0304;
end;

procedure TFRM_Export.ExportD0304_MPAS;
Begin
 // Create D0304  // BASED upon DISTRIBUTOR 'R' method
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select distinct');
  sql.add('''P'',');
  sql.add('l.ldso_mpid,');
  sql.add('c.new_map,');
  sql.add('c.date_of_change,');
  sql.add('c.mpancore,');
  sql.add('c.meterid,');
  sql.add('upper(m.man_make_type),');
  sql.add('m.timing_device_id');
  sql.add('from mopmgr.d0150_map_change C,');
  sql.add('mopmgr.d0150_290 M,');
  sql.add('mopmgr.mpan_status R,');
  sql.add('mddworking.ldso L');
  sql.add('where');
  sql.add('substr(c.mpancore,1,2)=l.mpanstart and');
  sql.add('c.mpancore=r.mpancore and');
  sql.add('c.mpancore=m.mpancore');
  sql.add('and m.meter_asset_provider_id=c.new_map');
  sql.add('and (c.mpancore,c.date_of_change,c.new_map) not in');
  sql.add('(select mpancore,efd_mapa,map_id from mopmgr.D0304');
  sql.add('where filename like ''R%'')');
  sql.add('order by 2,3,4,5');
  open;

 End;
 if GeneralQuery.recordcount=0 then exit;
 ExportThisD0304;
end;


Procedure TFRM_Export.ExportThisD0304;
Var
newmap,oldmap,agentrole,Lastagent,OldMpan,OldNonSet,OldTpr,OldMeter,OldRegister:string;
group,date_of_action,Newagent,mpan,nonset,tpr,meter,registerid,date_removed,meter_removed:string;
openfile:Boolean;
begin
 // Create Flow
 // ResetFileCounters etc
 OpenFile := False;
 LastAGENT:='non';
 oldmpan:='non';
 oldmeter:='non';
 // Loop Through All Records

 while not GeneralQuery.eof do
 begin

  agentrole:=generalquery.fields[0].text;
  NewAgent:=GeneralQuery.fields[1].text;
  newMap:=GeneralQuery.fields[2].text;
  mpan:=Generalquery.fields[4].text;
  Meter:=Generalquery.fields[5].text;

  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+MPAN;
  //FRM_Main.statusbar.panels[1].text:='Creating D0304 to '+NEWAGENT+' for MPAN '+MPAN;
  application.processmessages;

  // Check if Agent is different to Last Agent
  if (NewAgent<>LastAgent) or (newmap<>oldmap) or (date_of_action <> generalquery.fields[3].text) then
  Begin
   lastagent:=NewAgent;
   date_OF_ACTION:=Generalquery.fields[3].text;
   // Close any files that may be open
   if openfile = true then CreateFlowFooter;
   // Now Create New File and Write Header Record
   CreateFlowHeaderMOP('D0304001',NewAgent,''+agentrole+'');
   // Indicate There is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
   LastAGENT:='non';
   oldmpan:='non';
   oldmeter:='non';
   oldmap:='non';
  End;

  // Mpan Details
  if newMap<>OLDMap then
  Begin
   S:='77C|'+newMAP+'|';
   if date_of_action<>'' then s:=s+formatdatetime('YYYYMMDD',strtodate(date_of_action));
   s:=s+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   inc(outputflowFlowcount);
   oldmap:='non';
   oldmeter:='non';
  End;

  // Mpan Details
  if MPAN<>OLDMPAN then
  Begin
   S:='78C|'+MPAN+'|||||';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
   oldmeter:='non';
  End;

   // Meter Details
  if (meter<>oldmeter) then
  Begin
   S:='79C|'+meter+'|'+Generalquery.fields[6].text+'|';
   s:=s+generalquery.fields[7].text+'|';
   frm_main.WriteLinetoFile(s);
   Inc(OutputFlowLineCOunt);
  end;

  oldmap:=newmap;
  lastagent:=Newagent;
  Oldmpan:=mpan;
  OldMeter:=Meter;

  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;

Procedure TFRM_Export.dod0142s;
Var
Openfile,populateme: Boolean;
LastAgent, NewAgent,lastrole,newrole,newfrom,OLDMPAN,oldmeter,oldregister,meterid,registerid:String;
Cefsd,tempfield,specneeds,contact,specneedscode,dtccode,actionindicator,addinfo,e7ssc,tfrom,tto,aptim,pre:string;
begin
 FRM_File_Progress.progressbar.position:=0;
 With Generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select distinct');
  sql.add('M.CONFIRMED_MO_ID,M.MPANCORE,M.SSD,');
  sql.add('A.METERING_POINT_ADDRESS1,');
  sql.add('A.METERING_POINT_ADDRESS2,');
  sql.add('A.METERING_POINT_ADDRESS3,');
  sql.add('A.METERING_POINT_ADDRESS4,');
  sql.add('A.METERING_POINT_ADDRESS5,');
  sql.add('A.METERING_POINT_ADDRESS6,');
  sql.add('A.METERING_POINT_ADDRESS7,');
  sql.add('A.METERING_POINT_ADDRESS8,');
  sql.add('A.METERING_POINT_ADDRESS9,');
  sql.add('A.METERING_POINT_POSTCODE,');
  sql.add('m.regstatus,m.gsp_group,m.new_connection, f.install_date');
  sql.add('from');
  sql.add('EDMGR.mpan_status M,');
  sql.add('EDMGR.MPAS_CURRENT_ADDR A,');
  sql.add('(select mpancore,substr(filename,1,2)||''/''||substr(filename,3,2)||''/20''||substr(filename,5,2) install_date from EDMGR.FLOWHEADERS');
  sql.add('where flow_version like ''PSMI%'' and to_date(substr(filename,1,2)||''/''||substr(filename,3,2)||''/20''||substr(filename,5,2),''DD/MM/YYYY'')>=sysdate) F');
  sql.add('where a.mpancore=m.mpancore');
  sql.add('and f.mpancore=m.mpancore');
  //  sql.add('and M.confirmed_mo_id =''UMOL''');
 sql.add('and m.confirmed_mo_id is not null');
//  sql.add('and (m.mpancore,M.confirmed_mo_id) not in (select distinct c.mpancore,f.toname from edmgr.d0142 c,edmgr.flowheaders f where f.filename=c.filename and f.flow_version=''D0142'')');
  sql.add('and (m.mpancore) not in (select distinct c.mpancore from edmgr.d0142 c,edmgr.flowheaders f where f.filename=c.filename and f.flow_version=''D0142'' and c.mpancore is not null and f.mpancore is not null)');
  sql.add('order by 1,2,3');
  open;
 end;

 // Reset File Counters etc
 OpenFile := False;
 LastAGENT:='non';
 LastRole:='non';
 oldmpan:='non';

 // Loop Through All MPANS idenified in above Query
 while not GeneralQuery.eof do
 begin
  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MPAN: '+Generalquery.fields[1].text;
  //FRM_Main.statusbar.panels[1].text:='Creating D0142 to '+generalquery.fields[0].text+' for MPAN '+Generalquery.fields[1].text;
  application.processmessages;

  NewAgent:=GeneralQuery.fields[0].text;
  mpan:=Generalquery.fields[1].text;
  // Check if NEW SUPPLIER is different, this would indicate that a new file is required.
  if (NewAgent<>LastAgent) then
  Begin
   lastagent:=NewAgent;
   // If an existing SUPPLIER file is open, then close it.
   if openfile = true then CreateFlowFooter;
   // Now Create New File going to NEW supplier and Write Header Record
   CreateFlowHeader('D0142001',NewAgent,'M');
   // Indicate there is an Open File
   OutputFlowFlowcount:=0;
   OutputFlowLineCount:=0;
   Openfile:=true;
  End;

  // On Change of MPANCore, Create Group '05'. Group is MANDATORY as per DTC Spec
  if MPAN<>OLDMPAN then
  Begin
   tfrom:='080000';
   tto:='180000';

   with main_data_module.tempquery do
   Begin
    close;
    sql.Clear;
    sql.add('select filename from edmgr.flowheaders');
    sql.add('where mpancore='''+mpan+''' and flow_version like ''PSMI%''');
    sql.add('order by file_Date_time desc');
    open;
    aptim:=copy(main_data_module.tempquery.fields[0].text,7,2);
   End;

   if aptim='PM' then
   Begin
    tfrom:='120000';
    tto:='180000';
   end;
   if aptim='AM' then
   Begin
    tfrom:='080000';
    tto:='130000';
   end;

   mpan:=generalquery.fields[1].text;
   s:='267|'+mpan+'|';  //
   s:=s+generalquery.fields[3].text+'|';     // Address Line 1
   s:=s+generalquery.fields[4].text+'|';     // Address Line 2
   s:=s+generalquery.fields[5].text+'|';     // Address Line 3
   s:=s+generalquery.fields[6].text+'|';    // Address Line 4
   s:=s+generalquery.fields[7].text+'|';    // Address Line 5
   s:=s+generalquery.fields[8].text+'|';    // Address Line 6
   s:=s+generalquery.fields[9].text+'|';    // Address Line 7
   s:=s+generalquery.fields[10].text+'|';    // Address Line 8
   s:=s+generalquery.fields[11].text+'|';    // Address Line 9
   s:=s+generalquery.fields[12].text+'|';    // Address Line Postcode
   s:=s+Formatdatetime('YYYYMMDD',strtodate(generalquery.fields[16].text));
   s:=s+'|'+tfrom+'|'+tto+'|N|GS|E|';

   if generalquery.fields[14].text='_A' then E7SSC:='0152';
   if generalquery.fields[14].text='_B' then E7SSC:='0151';
   if generalquery.fields[14].text='_C' then E7SSC:='0349';
   if generalquery.fields[14].text='_D' then E7SSC:='0151';
   if generalquery.fields[14].text='_E' then E7SSC:='0151';
   if generalquery.fields[14].text='_F' then E7SSC:='0244';
   if generalquery.fields[14].text='_G' then E7SSC:='0244';
   if generalquery.fields[14].text='_H' then E7SSC:='0151';
   if generalquery.fields[14].text='_J' then E7SSC:='0151';
   if generalquery.fields[14].text='_K' then E7SSC:='0151';
   if generalquery.fields[14].text='_L' then E7SSC:='0151';
   if generalquery.fields[14].text='_M' then E7SSC:='0151';
   if generalquery.fields[14].text='_N' then E7SSC:='0723';
   if generalquery.fields[14].text='_P' then E7SSC:='0151';

   pre:='';
   if NEWAGENT='UMOL' Then pre:='*** SGN JOB *** ';

   if copy(generalquery.Fields[13].text,5,1)='0' then
   Begin
    s:=s+'0393|';
    addinfo:=pre+'PLS INSTALL 1 PH 1 RT DOM LIBERTY METER WITH FREEDOM UNIT IN CREDIT MODE';
   end
   else
   if copy(generalquery.Fields[13].text,5,1)='1' then
   Begin
    s:=s+'0393|';
    addinfo:=pre+'PLS INSTALL 1 PH 1 RT DOM LIBERTY METER WITH FREEDOM UNIT IN PREPAYMENT MODE';
   end
   else
   if copy(generalquery.Fields[13].text,5,1)='2' then
   Begin
    addinfo:=pre+'PLS INSTALL 1 PH 2 RT DOM LIBERTY METER WITH FREEDOM UNIT IN CREDIT MODE';
    s:=s+E7SSC+'|';
   end
   else
   if copy(generalquery.Fields[13].text,5,1)='3' then
   Begin
    addinfo:=pre+'PLS INSTALL 1 PH 2 RT DOM LIBERTY METER WITH FREEDOM UNIT IN PREPAYMENT MODE';
    s:=s+E7SSC+'|';
   end
   else
   // Catch All
   Begin
    s:=s+'0393|';
    addinfo:=pre+'PLS INSTALL 1 PH 1 RT DOM LIBERTY METER WITH FREEDOM UNIT IN PREPAYMENT MODE';
   end;


   if generalquery.Fields[15].text='T' then
   Begin
    addinfo:=addinfo+' (NEW CONNECTION)';
    s:=s+addinfo+'|F||';
   end
   else s:=s+addinfo+'|T||';

   inc(OutputFlowLineCount);
   inc(OutputFlowFlowcount);
   frm_main.WriteLinetoFile(s);  // Write Record to File
  end;
  lastagent:=newagent;
  oldmpan:=mpan;
  GeneralQuery.Next;
 end; // Repeat Above process for remaining MPANS
 // Close the Last File That may be Open
 if openfile = true then CreateFlowFooter;
end;

Procedure TFRM_Export.Do_E_D0205s;
Var
MPAN,AgentID,AgentROLE,news:String;
SSD,EFTSSD:Tdatetime;
Begin
 FRM_File_Progress.progressbar.position:=0;

 // Identify any D0205s as a result of SSC changes etc and add them to the list
 try
  FRM_COMMON.Execute_Oracle_Procedure('EDMGR.PR_EXP_D0205_POST_SSD');
 except
 end;

 // Identify Any MTC /SMSO ID Corrections
 try
  FRM_COMMON.Execute_Oracle_Procedure('EDMGR.PR_ADD_D0205_SMSO_MTC');
 except
 end;

 with generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select * from edmgr.batch_flows_for_sending_all where flowversion=''D0205''');
  sql.add('and to_mpid is not null and to_role is not null and date_generated is null and status=''R''');
  sql.add('order by TO_ROLE,TO_MPID,LINE_2');
  open;
 End;
 oldagentid:='';
 oldagentrole:='';
 while not GeneralQuery.eof do
 Begin
  MPAN:=GeneralQuery.fields[0].text;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  if (agentrole<>oldagentrole) or (agentid<>oldagentid) then
  begin
   FRM_File_Progress.progressbar.position:=0;
   CreateFlowHeader('D0205001',Agentid,AgentRole);
   OutputFlowFlowcount:=0;
   OutPutFlowLineCount:=0;
   S:= '752'+'|'+frm_common.smrssequenceno(AGENTID)+'|';
   frm_main.WriteLinetoFile(s);
   inc(OutputFlowLineCount);
   inc(OutputFlowFlowCount);
  end;
  oldagentrole:=agentrole;
  oldagentid:=agentid;
  // Create Body

  s:=GeneralQuery.Fields[5].text;
  if copy(s,5,6)='INSTNO' then
  Begin
   ins:=frm_common.NextInstructionNumber;
   news:=stringreplace(s,'INSTNO',ins,[rfreplaceall]);
   s:=news;
  end;

  frm_main.WriteLinetoFile( s);
  OutputFlowlinecount:=OutputFlowlinecount+1;
  GeneralQuery.next;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  if (agentrole<>oldagentrole) or (oldagentid<>agentid) then cfooter:=true
  else cfooter:=false;
  if GeneralQuery.eof then cfooter:=true;
  if cfooter=true then CreateFlowFooter;
 end;

  //Update Sent Status of all Blank Status.s
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('update edmgr.batch_flows_for_sending_all');
  sql.add('set date_generated=sysdate,status=''S'' where flowversion=''D0205'' and date_generated is null');
  execute;
 End;
 frm_login.mainsession.commit;

end;


Procedure TFRM_Export.CreateD0205COA;
Var
MPAN,AgentID,AgentROLE,news:String;
SSD,EFTSSD:Tdatetime;
Begin
 FRM_File_Progress.progressbar.position:=0;

 // Identify any D0205s as a result of SSC changes etc and add them to the list
 FRM_COMMON.Execute_Oracle_Procedure('EDMGR.PR_EXP_D0205_POST_SSD');

 with generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select * from edmgr.batch_flows_for_sending_coa where flowversion=''D0205''');
  sql.add('and to_mpid is not null and to_role is not null and date_generated is null and status=''R''');
  sql.add('order by TO_ROLE,TO_MPID,LINE_2');
  open;
 End;
 oldagentid:='';
 oldagentrole:='';
 while not GeneralQuery.eof do
 Begin
  MPAN:=GeneralQuery.fields[0].text;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  if (agentrole<>oldagentrole) or (agentid<>oldagentid) then
  begin
   FRM_File_Progress.progressbar.position:=0;
   CreateFlowHeader('D0205001',Agentid,AgentRole);
   OutputFlowFlowcount:=0;
   OutPutFlowLineCount:=0;
   S:= '752'+'|'+frm_common.smrssequenceno(AGENTID)+'|';
   frm_main.WriteLinetoFile(s);
   inc(OutputFlowLineCount);
   inc(OutputFlowFlowCount);
  end;
  oldagentrole:=agentrole;
  oldagentid:=agentid;
  // Create Body

  s:=GeneralQuery.Fields[5].text;
  if copy(s,5,6)='INSTNO' then
  Begin
   ins:=frm_common.NextInstructionNumber;
   news:=stringreplace(s,'INSTNO',ins,[rfreplaceall]);
   s:=news;
  end;

  frm_main.WriteLinetoFile( s);
  OutputFlowlinecount:=OutputFlowlinecount+1;
  GeneralQuery.next;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  if (agentrole<>oldagentrole) or (oldagentid<>agentid) then cfooter:=true
  else cfooter:=false;
  if GeneralQuery.eof then cfooter:=true;
  if cfooter=true then CreateFlowFooter;
 end;
end;


//Procedure TFRM_Export.CreatebATCHd0055;
//Var
//AgentID,AgentROLE,SMRS:String;
//SSD,EFTSSD:Tdatetime;
//Begin
// FRM_File_Progress.progressbar.position:=0;
// with generalquery do
// Begin
//  close;
//  sql.clear;
//  sql.add('select * from EDMGR.BATCH_FLOWS_FOR_SENDING_D0055 WHERE DATE_OUTPUT IS NULL order by 1,2');
//  open;
// End;
// oldagentid:='';
// while not GeneralQuery.eof do
// Begin
//  AGENTID:=GeneralQuery.fields[0].text;
//  if (agentid<>oldagentid) then
//  begin
//   FRM_File_Progress.progressbar.position:=0;
//   CreateFlowHeader('D0055001',Agentid,'P');
//   SMRS:=frm_common.smrssequenceno(agentid);
//   s:='733|'+SMRS+'|';
//   frm_main.WriteLinetoFile(s);      // Now write Group 733 File Sequence Number
//   OutputFlowFlowcount:=1;
//   OutPutFlowLineCount:=1;
//  end;
//  oldagentid:=agentid;
//  // Create Body
//  frm_main.WriteLinetoFile( GeneralQuery.Fields[1].text);
//  OutputFlowlinecount:=OutputFlowlinecount+1;
//  GeneralQuery.next;
//  AGENTID:=GeneralQuery.fields[0].text;
//  if (oldagentid<>agentid) then cfooter:=true
//  else cfooter:=false;
//  if GeneralQuery.eof then cfooter:=true;
//  if cfooter=true then CreateFlowFooter;
// end;
// with main_data_module.updatequery do
// Begin
//  close;
//  sql.clear;
//  sql.add('update EDMGR.BATCH_FLOWS_FOR_SENDING_D0055 set date_output=sysdate where date_output is null');
//  execute;
// End;
// frm_login.mainsession.commit;
//
//end;


Procedure TFRM_Export.CreateD0151COA;
Var
MPAN,AgentID,AgentROLE:String;
SSD,EFTSSD:Tdatetime;
Begin
 FRM_File_Progress.progressbar.position:=0;
 with generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select * from edmgr.batch_flows_for_sending_coa where flowversion=''D0151''');
  sql.add('and date_generated is null and status=''R''');
  sql.add('order by TO_ROLE,TO_MPID,LINE_2');
  open;
 End;
 oldagentid:='';
 oldagentrole:='';
 while not GeneralQuery.eof do
 Begin
  MPAN:=GeneralQuery.fields[0].text;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  if (agentrole<>oldagentrole) or (agentid<>oldagentid) then
  begin
   FRM_File_Progress.progressbar.position:=0;
   CreateFlowHeader('D0151001',Agentid,AgentRole);
   OutputFlowFlowcount:=0;
   OutPutFlowLineCount:=0;
  end;
  oldagentrole:=agentrole;
  oldagentid:=agentid;
  // Create Body
  frm_main.WriteLinetoFile( GeneralQuery.Fields[5].text);
  inc(OutputFlowFlowCount);
  OutputFlowlinecount:=OutputFlowlinecount+2;
  GeneralQuery.next;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  if (agentrole<>oldagentrole) or (oldagentid<>agentid) then cfooter:=true
  else cfooter:=false;
  if GeneralQuery.eof then cfooter:=true;
  if cfooter=true then CreateFlowFooter;
 end;
end;

Procedure TFRM_Export.CreateD0148COA;
Var
MPAN,AgentID,AgentROLE,ins:String;
SSD,EFTSSD:Tdatetime;
Begin
 FRM_File_Progress.progressbar.position:=0;
 with generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select * from edmgr.batch_flows_for_sending_coa where flowversion=''D0148''');
  sql.add('and date_generated is null and status=''R''');
  sql.add('order by TO_ROLE,TO_MPID,LINE_2');
  open;
 End;
 oldagentid:='';
 oldagentrole:='';
 while not GeneralQuery.eof do
 Begin
  MPAN:=GeneralQuery.fields[0].text;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  if (agentrole<>oldagentrole) or (agentid<>oldagentid) then
  begin
   FRM_File_Progress.progressbar.position:=0;
   CreateFlowHeader('D0148001',Agentid,AgentRole);
   OutputFlowFlowcount:=0;
   OutPutFlowLineCount:=0;
  end;
  oldagentrole:=agentrole;
  oldagentid:=agentid;
  // Create Body
  frm_main.WriteLinetoFile( GeneralQuery.Fields[5].text);
  inc(OutputFlowFlowCount);
  OutputFlowlinecount:=OutputFlowlinecount+3;
  GeneralQuery.next;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  if (agentrole<>oldagentrole) or (oldagentid<>agentid) then cfooter:=true
  else cfooter:=false;
  if GeneralQuery.eof then cfooter:=true;
  if cfooter=true then CreateFlowFooter;
 end;
end;


Procedure TFRM_Export.CreateD0170COA;
Var
MPAN,AgentID,AgentROLE,LINEDATE,OLDLINE:String;
SSD,EFTSSD:Tdatetime;
Begin
 FRM_File_Progress.progressbar.position:=0;
 with generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select * from edmgr.batch_flows_for_sending_coa where flowversion=''D0170''');
  sql.add('and date_generated is null and status=''R''');
  sql.add('order by TO_ROLE,TO_MPID,LINE_1,LINE_2');
  open;
 End;
 oldagentid:='';
 oldagentrole:='';
 OLDLINE:='';
 while not GeneralQuery.eof do
 Begin
  MPAN:=GeneralQuery.fields[0].text;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  LINEDATE:=GeneralQuery.Fields[4].text;
  if (agentrole<>oldagentrole) or (agentid<>oldagentid) or (linedate<>oldline) then
  begin
   FRM_File_Progress.progressbar.position:=0;
   CreateFlowHeader('D0170001',Agentid,AgentRole);
   OutputFlowFlowcount:=0;
   OutPutFlowLineCount:=0;
   S:=GeneralQuery.Fields[4].text;
   frm_main.WriteLinetoFile(s);
   inc(OutputFlowLineCount);
  end;
  oldagentrole:=agentrole;
  oldagentid:=agentid;
  oldline:=linedate;
  // Create Body
  frm_main.WriteLinetoFile( GeneralQuery.Fields[5].text);
  inc(OutputFlowFlowCount);
  OutputFlowlinecount:=OutputFlowlinecount+1;
  GeneralQuery.next;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  LINEDATE:=GeneralQuery.fields[4].text;
  if (agentrole<>oldagentrole) or (oldagentid<>agentid) or (oldline<>linedate) then cfooter:=true
  else cfooter:=false;
  if GeneralQuery.eof then cfooter:=true;
  if cfooter=true then CreateFlowFooter;
 end;
end;

Procedure TFRM_Export.CreateD0190;
Var
MPAN,AgentID,AgentROLE,LINEDATE,OLDLINE:String;
SSD,EFTSSD:Tdatetime;
Begin
 FRM_File_Progress.progressbar.position:=0;
 with generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select * from edmgr.batch_flows_for_sending_all where flowversion=''D0190''');
  sql.add('and date_generated is null and status=''R''');
  sql.add('order by TO_ROLE,TO_MPID,LINE_1,LINE_2');
  open;
 End;
 oldagentid:='';
 oldagentrole:='';
 OLDLINE:='';
 while not GeneralQuery.eof do
 Begin
  MPAN:=GeneralQuery.fields[0].text;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  LINEDATE:=GeneralQuery.Fields[4].text;
  if (agentrole<>oldagentrole) or (agentid<>oldagentid) or (linedate<>oldline) then
  begin
   FRM_File_Progress.progressbar.position:=0;
   CreateFlowHeader('D0190001',Agentid,AgentRole);
   OutputFlowFlowcount:=0;
   OutPutFlowLineCount:=0;
  end;
  oldagentrole:=agentrole;
  oldagentid:=agentid;
  oldline:=linedate;
  // Create Body
  frm_main.WriteLinetoFile( GeneralQuery.Fields[5].text);
  inc(OutputFlowFlowCount);
  OutputFlowlinecount:=OutputFlowlinecount+1;
  GeneralQuery.next;
  AGENTID:=GeneralQuery.fields[3].text;
  AGENTROLE:=GeneralQuery.Fields[2].text;
  LINEDATE:=GeneralQuery.fields[4].text;
  if (agentrole<>oldagentrole) or (oldagentid<>agentid) or (oldline<>linedate) then cfooter:=true
  else cfooter:=false;
  if GeneralQuery.eof then cfooter:=true;
  if cfooter=true then CreateFlowFooter;
 end;
end;

function TFRM_Export.GetAppDate(COADATE,SSD:string):string;
begin
 // If COAEFD is null then agenet effective date is SSD
 if coadate='' then GetAppdate:=formatdatetime('YYYYMMDD',strtodate(ssd))+'|'
 else
 // if COAEFD is not null then, if COAEFD>=ssd then use the COAEFD
 // Else COAEFD is before SSD then use SSD
 if strtodate(coadate)>=strtodate(ssd) then GetAppdate:=formatdatetime('YYYYMMDD',strtodate(coadate))+'|'
 else GetAppdate:=formatdatetime('YYYYMMDD',strtodate(ssd))+'|';
end;

procedure TFrm_Export.DoMOPPARMS_NM12(YEAR,MONTH:string);
Var
OLD_MO,PREVIOUS_MO,monthend:string;
s:string;
z:integer;
begin

 main_data_module.parmstemp.lines.loadfromfile(mopparmsdir+'PARMS_NM12.txt');
 main_data_module.parms.lines.clear;
 for z:=1 to main_data_module.parmstemp.lines.count do
 begin
  s:=stringreplace(main_data_module.parmstemp.lines[z-1],'01/07/2011','01/'+month+'/'+year,[rfreplaceall]);
  main_data_module.parms.lines.add(s);
 End;
 main_data_module.parms.execute;
 // Get Data
 with generalquery do
 Begin
  close;
  sql.clear;
  //get query
  sql.add('select OLD_MOP_ID,SUPPLIER_ID,PERIOD,GSP_GROUP,');
  sql.add('sum(UNIQUE_REGISTRATIONS),');
  sql.add('sum(TOTAL_NO_D0150),');
  sql.add('sum(MISSING_BEFORE_R1),');
  sql.add('sum(MISSING_BEFORE_R2),');
  sql.add('sum(MISSING_BEFORE_R3),');
  sql.add('sum(MISSING_BEFORE_RF),');
  sql.add('sum(MISSING_AFTER_RF)');
  sql.add('from MOPMGR.PARMS_TEMP_NM12');
  sql.add('group by ');
  sql.add('OLD_MOP_ID,');
  sql.add('SUPPLIER_ID, ');
  sql.add('PERIOD,');
  sql.add('GSP_GROUP ');
  sql.add('order by 1,2,3,4');
  open;
 End;
 PREVIOUS_MO:='';
 outputfilename:=h_Outgoing+M_MPID+'234'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt';
 CreateFlowHeaderMOPPARMS('P0234001',M_MPIDROLE,M_MPID,'');
 OutputFlowFlowcount:=0;
 OutputFlowLineCount:=0;

 while not GeneralQuery.eof do
 begin
  old_MO:=generalquery.fields[0].text;
  if previous_mo<>old_mo then
  Begin
   s:='SUB|N|M|'+old_mo+'|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)))+'|M';
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
  end;

  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='MOP: '+OLD_MO;
  FRM_Main.statusbar.panels[1].text:='Creating NM12 Parms Report ';
  application.processmessages;

  s:='2NM|'+Generalquery.fields[3].text+'|'+Generalquery.fields[1].text+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[4].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[5].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[6].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[7].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[8].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[9].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[10].text),-2))+'';
  inc(OutputFlowLineCount);
  inc(OutputFlowFlowcount);
  parmsfile.lines.add(s);

  previous_mo:=old_mo;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 CreateFlowFooterPARMS;
end;


procedure TFrm_Export.DoMOPPARMS_SP11(YEAR,MONTH:string);
Var
OLD_supp,PREVIOUS_supp,monthend:string;
s:string;
z:integer;
begin

 main_data_module.parmstemp.lines.loadfromfile(mopparmsdir+'PARMS_SP11.txt');
 main_data_module.parms.lines.clear;
 for z:=1 to main_data_module.parmstemp.lines.count do
 begin
  s:=stringreplace(main_data_module.parmstemp.lines[z-1],'01/07/2011','01/'+month+'/'+year,[rfreplaceall]);
  main_data_module.parms.lines.add(s);
 End;
 main_data_module.parms.execute;
 // Get Data
 with generalquery do
 Begin
  close;
  sql.clear;
  //get query
  sql.add('select * FROM MOPMGR.PARMS_TEMP_SP11');
  sql.add('order by 1,3');
  open;
 End;
 PREVIOUS_supp:='';
 outputfilename:=h_Outgoing+M_MPID+'224'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt';
 CreateFlowHeaderMOPPARMS('P0224001',M_MPIDROLE,M_MPID,'');
 OutputFlowFlowcount:=0;
 OutputFlowLineCount:=0;

 while not GeneralQuery.eof do
 begin
  old_supp:=generalquery.fields[0].text;
  if previous_supp<>old_supp then
  Begin
   s:='SUB|N|X|'+old_supp+'|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)))+'|M';
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
  end;

  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='SUP: '+OLD_supp;
  FRM_Main.statusbar.panels[1].text:='Creating SP11 Parms Report ';
  application.processmessages;

  s:='X11|'+Generalquery.fields[2].text+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[3].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[4].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[5].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[6].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[7].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[8].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[9].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[10].text),-2))+'';
  inc(OutputFlowLineCount);
  inc(OutputFlowFlowcount);
  parmsfile.lines.add(s);

  previous_supp:=old_supp;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 CreateFlowFooterPARMS;
end;

procedure TFrm_Export.DoMOPPARMS_SP14(YEAR,MONTH:string);
Var
OLD_supp,PREVIOUS_supp,monthend:string;
s:string;
z:integer;
begin

 main_data_module.parmstemp.lines.loadfromfile(mopparmsdir+'PARMS_SP14.txt');
 main_data_module.parms.lines.clear;
 for z:=1 to main_data_module.parmstemp.lines.count do
 begin
  s:=stringreplace(main_data_module.parmstemp.lines[z-1],'01/07/2011','01/'+month+'/'+year,[rfreplaceall]);
  main_data_module.parms.lines.add(s);
 End;
 main_data_module.parms.execute;
 // Get Data
 with generalquery do
 Begin
  close;
  sql.clear;
  //get query
  sql.add('select * FROM MOPMGR.PARMS_TEMP_SP14');
  sql.add('order by 1,3');
  open;
 End;
 PREVIOUS_supp:='';
 outputfilename:=h_Outgoing+M_MPID+'227'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt';
 CreateFlowHeaderMOPPARMS('P0227001',M_MPIDROLE,M_MPID,'');
 OutputFlowFlowcount:=0;
 OutputFlowLineCount:=0;

 while not GeneralQuery.eof do
 begin
  old_supp:=generalquery.fields[0].text;
  if previous_supp<>old_supp then
  Begin
   s:='SUB|N|X|'+old_supp+'|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)))+'|M';
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
  end;

  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='SUP: '+OLD_supp;
  FRM_Main.statusbar.panels[1].text:='Creating SP14 Parms Report ';
  application.processmessages;

  s:='X14|'+Generalquery.fields[2].text+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[3].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[4].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[5].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[6].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[7].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[8].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[9].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[10].text),-2))+'';
  inc(OutputFlowLineCount);
  inc(OutputFlowFlowcount);
  parmsfile.lines.add(s);

  previous_supp:=old_supp;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 CreateFlowFooterPARMS;
end;

procedure TFRM_EXPORT.UpdateStatusBar(MSg,banner:string);
begin
 FRM_Main.Statusbar.panels[1].text:=msg;
 if banner<>'' then
 begin
  frm_busy.labelbusy.caption:=msg;
  Frm_busy.show
 end
 else
 begin
  frm_busy.labelbusy.caption:='';
  Frm_busy.close;
 end;

 application.processmessages;
end;

procedure TFrm_Export.DoMOPPARMS_SP15(YEAR,MONTH:string);
Var
OLD_supp,PREVIOUS_supp,monthend:string;
s:string;
z:integer;
begin

 main_data_module.parmstemp.lines.loadfromfile(mopparmsdir+'PARMS_SP15.txt');
 main_data_module.parms.lines.clear;
 for z:=1 to main_data_module.parmstemp.lines.count do
 begin
  s:=stringreplace(main_data_module.parmstemp.lines[z-1],'01/07/2011','01/'+month+'/'+year,[rfreplaceall]);
  main_data_module.parms.lines.add(s);
 End;
 main_data_module.parms.execute;
 // Get Data

 with generalquery do
 Begin
  close;
  sql.clear;
  //get query
  sql.add('select * FROM MOPMGR.PARMS_TEMP_SP15');
  sql.add('order by 1,3');
  open;
 End;
 PREVIOUS_supp:='';
 outputfilename:=h_Outgoing+M_MPID+'228'+copy(YEAR,4,1)+'.'+UPPERCASE(formatsettings.shortmonthnames[strtoint(MONTH)])+'.txt';
 CreateFlowHeaderMOPPARMS('P0228001',M_MPIDROLE,M_MPID,'');
 OutputFlowFlowcount:=0;
 OutputFlowLineCount:=0;

 while not GeneralQuery.eof do
 begin
  old_supp:=generalquery.fields[0].text;
  if previous_supp<>old_supp then
  Begin
   s:='SUB|N|X|'+old_supp+'|'+formatdatetime('YYYYMMDD',frm_common.LDOM(strtodate('01/'+MONTH+'/'+YEAR)))+'|M';
   inc(OutputFlowLineCount);
   parmsfile.lines.add(s);
  end;

  FRM_File_Progress.progressbar.max:=Generalquery.recordcount;
  FRM_File_Progress.d_file.caption:='';
  FRM_File_Progress.labelcount.caption:='';
  FRM_File_Progress.progressbar.position:=FRM_File_Progress.progressbar.position+1;
  FRM_File_Progress.statusbar.panels[0].text:='SUP: '+OLD_supp;
  FRM_Main.statusbar.panels[1].text:='Creating SP15 Parms Report ';
  application.processmessages;

  s:='X15|'+Generalquery.fields[2].text+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[3].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[4].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[5].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[6].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[7].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[8].text),-2))+'|';
  s:=s+floattostr(roundto(strtofloat(Generalquery.fields[9].text),-2))+'';
  inc(OutputFlowLineCount);
  inc(OutputFlowFlowcount);
  parmsfile.lines.add(s);

  previous_supp:=old_supp;
  GeneralQuery.Next;
 end; // Repeat for remaining MPANS
 CreateFlowFooterPARMS;
end;


function TFRM_Export.IndexOf(s: string): Integer;
begin
  if s='AQI' then
    result := Ord(TGasShipper.gsAQI)
  else if s='BRN' then
    result := Ord(TGasShipper.gsBRN)
  else if s='CNC' then
    result := Ord(TGasShipper.gsCNC)
  else if s='CNF' then
    result := Ord(TGasShipper.gsCNF)
  else if s='EMC' then
    result := Ord(TGasShipper.gsEMC)
  else if s='GEA' then
    result := Ord(TGasShipper.gsGEA)
  else if s='MAI' then
    result := Ord(TGasShipper.gsMAI)
  else if s='MAM' then
    result := Ord(TGasShipper.gsMAM)
  else if s='MID' then
    result := Ord(TGasShipper.gsMID)
  else if s='MSI' then
    result := Ord(TGasShipper.gsMSI)
  else if s='NOM' then
    result := Ord(TGasShipper.gsNOM)
  else if s='ONJ' then
    result := Ord(TGasShipper.gsONJ)
  else if s='ONU' then
    result := Ord(TGasShipper.gsONU)
  else if s='ORJ' then
    result := Ord(TGasShipper.gsORJ)
  else if s='RFA' then
    result := Ord(TGasShipper.gsRFA)
  else if s='RRP' then
    result := Ord(TGasShipper.gsRRP)
  else if s='SFN' then
    result := Ord(TGasShipper.gsSFN)
  else if s='SPC' then
    result := Ord(TGasShipper.gsSPC)
  else if s='SPI' then
    result := Ord(TGasShipper.gsSPI)
  else if s='UBR' then
    result := Ord(TGasShipper.gsUBR)
  else if s='UDR' then
    result := Ord(TGasShipper.gsUDR)
  else if s='UMR' then
    result := Ord(TGasShipper.gsUMR)
  else if s='WAO' then
    result := Ord(TGasShipper.gsWAO);
end;

end.