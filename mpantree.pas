unit MPANTREE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, OracleData,oracle, RxLookup, Grids, StdCtrls, ComCtrls, ExtCtrls, Menus, ImgList,
  ToolWin, Buttons, OleServer, ShellAPI, fmxutils, smets_updates, Vcl.ActnList, Vcl.Themes,
  VirtualTrees, System.Actions, AgreementCheckedNodeManager, CRMTreeViewData, DateUtils,
  TaskDialog, TaskDialogEx;

type
  TFRM_Tree = class(TForm)
    GeneralQuery: TOracleDataSet;
    PopUpElectric: TPopupMenu;
    S_Dflows: TMenuItem;
    MainMenu1: TMainMenu;
    M_View: TMenuItem;
    Enquiries1: TMenuItem;
    PopUpCust: TPopupMenu;
    AddCustomerNote1: TMenuItem;
    AddCustomerEnquiry1: TMenuItem;
    V_FullNotes: TMenuItem;
    N3: TMenuItem;
    V_Non: TMenuItem;
    N4: TMenuItem;
    ShowAllEnquiriesNotes1: TMenuItem;
    N6: TMenuItem;
    ViewCustomerTree1: TMenuItem;
    PopUpPremise: TPopupMenu;
    P1: TMenuItem;
    PopUpEnquiry: TPopupMenu;
    E_TakeOwnership: TMenuItem;
    Historic1: TMenuItem;
    N1: TMenuItem;
    MaintainSpecialNeeds1: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    MaintainCustomerDetails: TMenuItem;
    MeterDetailsReadings1: TMenuItem;
    N9: TMenuItem;
    Panel1: TPanel;
    PopUpAgreements: TPopupMenu;
    MenuItem1: TMenuItem;
    N2: TMenuItem;
    MaintainAccountHolders: TMenuItem;
    N5: TMenuItem;
    AddAgreement1: TMenuItem;
    AddService1: TMenuItem;
    P3: TMenuItem;
    ViewAggreement1: TMenuItem;
    productquery: TOracleDataSet;
    SpanQuery: TOracleDataSet;
    N10: TMenuItem;
    A_ProdWizard: TMenuItem;
    RegisterItem: TMenuItem;
    N11: TMenuItem;
    PopUpAccount: TPopupMenu;
    EditAccountHolders: TMenuItem;
    PopUpTelecom: TPopupMenu;
    CallDataRecords1: TMenuItem;
    PopUpGas: TPopupMenu;
    MenuItem2: TMenuItem;
    N12: TMenuItem;
    A_dec: TMenuItem;
    A_rev: TMenuItem;
    RemoveAccountHolder1: TMenuItem;
    N13: TMenuItem;
      NEWBTN: TBitBtn;
    MaintainPremiseDetails1: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    G_ReOrder: TMenuItem;
    N16: TMenuItem;
    E_Re_Order: TMenuItem;
    N17: TMenuItem;
    T_ReOrder: TMenuItem;
    MaintainServicesNetworkFeatures1: TMenuItem;
    N18: TMenuItem;
    FeatureQuery: TOracleDataSet;
    FriendsFamily1: TMenuItem;
    N19: TMenuItem;
    QuarterlyStatement1: TMenuItem;
    QuarterlyStatementItemised1: TMenuItem;
    N20: TMenuItem;
    N21: TMenuItem;
    CustomerLetters1: TMenuItem;
    Electric1: TMenuItem;
    Electricity1: TMenuItem;
    Telecoms1: TMenuItem;
    N23: TMenuItem;
    Customer1: TMenuItem;
    AddScannedDocument1: TMenuItem;
    Select_File_To_Attach: TOpenDialog;
    RatedUsage: TOracleDataSet;
    MTDS: TOracleDataSet;
    ScheduleQuery: TOracleDataSet;
    MeterDetailsReadings2: TMenuItem;
    ViewRefreshMPRN1: TMenuItem;
    N24: TMenuItem;
    N26: TMenuItem;
    ViewRefreshCallerLineID1: TMenuItem;
    PopUpRated: TPopupMenu;
    MenuItem15: TMenuItem;
    AddProspectDetails1: TMenuItem;
    PopUpProspect: TPopupMenu;
    MenuItem3: TMenuItem;
    PopUpBroadBand: TPopupMenu;
    MenuItem4: TMenuItem;
    BroadbandReorder: TMenuItem;
    N27: TMenuItem;
    N28: TMenuItem;
    Letters1: TMenuItem;
    WebSignupLetter1: TMenuItem;
    InformationRequest11: TMenuItem;
    CustomerCredits1: TMenuItem;
    SAC: TMenuItem;
    ShowFinancialHistory1: TMenuItem;
    ViewSPANSetupDetails1: TMenuItem;
    ViewSPANSetupDetails2: TMenuItem;
    View1: TMenuItem;
    ViewSPANSetupDetails3: TMenuItem;
    WelcomeLetter1: TMenuItem;
    CancelAllSubOrders1: TMenuItem;
    N29: TMenuItem;
    C1: TMenuItem;
    N30: TMenuItem;
    COT1: TMenuItem;
    N31: TMenuItem;
    M_Single: TMenuItem;
    N32: TMenuItem;
    A_C_1: TMenuItem;
    N33: TMenuItem;
    MakeLIVeCOTmoveIN1: TMenuItem;
    UndofromOrderReadytoOrderPlaced1: TMenuItem;
    Panel2: TPanel;
    Button1: TButton;
    AquireAttachDocument1: TMenuItem;
    N22: TMenuItem;
    Cancellation1: TMenuItem;
    N35: TMenuItem;
    Letters2: TMenuItem;
    elecomsApplicationForm1: TMenuItem;
    N36: TMenuItem;
    RequestMissingMPRN1: TMenuItem;
    RequestMissingMPAN1: TMenuItem;
    RequestMissingMPRNMPAN1: TMenuItem;
    L_G_O: TMenuItem;
    ObjectionReceived1: TMenuItem;
    L_E_O: TMenuItem;
    ObjectionReveived1: TMenuItem;
    N38: TMenuItem;
    IGTMENU: TMenuItem;
    N39: TMenuItem;
    ETUtilitaFault1: TMenuItem;
    N40: TMenuItem;
    ElectricETLetter1: TMenuItem;
    N41: TMenuItem;
    DUALETUtilitaFault1: TMenuItem;
    N42: TMenuItem;
    GroupBox4: TGroupBox;
    N43: TMenuItem;
    ManuallySetSPANStatusSSD1: TMenuItem;
    N44: TMenuItem;
    RatedIssues: TMenuItem;
    RatedIssuesQuery: TOracleDataSet;
    N46: TMenuItem;
    miReRateAgreement: TMenuItem;
    StatusBar: TStatusBar;
    ShowSiteVisitInformation1: TMenuItem;
    GasMeter: TOracleDataSet;
    G_SET: TMenuItem;
    N45: TMenuItem;
    E_SET: TMenuItem;
    N48: TMenuItem;
    elecomsSignupLetter1: TMenuItem;
    SignupLetter1: TMenuItem;
    FriendsFamilyForm1: TMenuItem;
    InfoPackApplicationForm1: TMenuItem;
    N34: TMenuItem;
    N47: TMenuItem;
    d1: TMenuItem;
    PopUpReviewer: TPopupMenu;
    MenuItem5: TMenuItem;
    MovetoNewPPMAgreement1: TMenuItem;
    MovetoNewPPMAgreemen1: TMenuItem;
    N51: TMenuItem;
    MeterReadings1: TMenuItem;
    GasReadRequired1: TMenuItem;
    ElectricReadRequired1: TMenuItem;
    DualReadsRequired1: TMenuItem;
    SPANOverride1: TMenuItem;
    N52: TMenuItem;
    OnHoldPopUp: TPopupMenu;
    MenuItem6: TMenuItem;
    RegistrationHistory1: TMenuItem;
    G_RET: TMenuItem;
    E_RET: TMenuItem;
    N54: TMenuItem;
    CurrentTariffSheet1: TMenuItem;
    N55: TMenuItem;
    VacantPremiseProgrammedforDisconnection1: TMenuItem;
    N56: TMenuItem;
    raceExecutors1: TMenuItem;
    raceExecutors2: TMenuItem;
    N57: TMenuItem;
    EnquiryRecevied1: TMenuItem;
    raceExecutorsFollowUp1: TMenuItem;
    ShowAnnualUsagekWh1: TMenuItem;
    ShowAnnualUsagekWh2: TMenuItem;
    DoNotBillDisconnectedDeEnergised1: TMenuItem;
    DoNotBillDisconnectedDeEnergised2: TMenuItem;
    SAC1: TMenuItem;
    ChequeReceived1: TMenuItem;
    Disputes: TOracleDataSet;
    PopUpDispute: TPopupMenu;
    MenuItem7: TMenuItem;
    AddAccountDispute1: TMenuItem;
    Gas1: TMenuItem;
    Elec1: TMenuItem;
    Dual1: TMenuItem;
    Surcharge1: TMenuItem;
    ET1: TMenuItem;
    N59: TMenuItem;
    Correspondance1: TMenuItem;
    PageControl1: TPageControl;
    TABSUPP: TTabSheet;
    TABMOP: TTabSheet;
    MOPD0155: TOracleDataSet;
    MOP302: TOracleDataSet;
    MTDSMOP: TOracleDataSet;
    PopupMenuMop: TPopupMenu;
    ViewDataflows1: TMenuItem;
    Panel3: TPanel;
    PopUpElectricMop: TPopupMenu;
    MenuItem8: TMenuItem;
    MenuItem10: TMenuItem;
    MopHistory: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem19: TMenuItem;
    MOPSITE: TOracleDataSet;
    N60: TMenuItem;
    AddEnquiry1: TMenuItem;
    Add1: TMenuItem;
    AquireAttachDocument2: TMenuItem;
    AddServiceOrder1: TMenuItem;
    N61: TMenuItem;
    N62: TMenuItem;
    ShowAllEnquiriesNotesDocs1: TMenuItem;
    ShowAllServiceOrders1: TMenuItem;
    N63: TMenuItem;
    M_AUDIT: TMenuItem;
    N64: TMenuItem;
    IGTAdminCharge1: TMenuItem;
    RegistrationHistory2: TMenuItem;
    RemoveDonNotBillAllowBilling1: TMenuItem;
    RemoveDoNotBillAllowBilling1: TMenuItem;
    AddOneOffCharge1: TMenuItem;
    AddOneOffCharge2: TMenuItem;
    N69: TMenuItem;
    PopUpNote: TPopupMenu;
    MenuItem9: TMenuItem;
    N70: TMenuItem;
    COTMoveIn1: TMenuItem;
    SPANTypeChange1: TMenuItem;
    SUnrestricted1: TMenuItem;
    EEconomy71: TMenuItem;
    N71: TMenuItem;
    PopUpSO: TPopupMenu;
    FeedbackForm1: TMenuItem;
    gsmiletter: TMenuItem;
    ObjectionReceivedgetSmart1: TMenuItem;
    ObjectionReceivedgetSmart2: TMenuItem;
    getSmart1: TMenuItem;
    N72: TMenuItem;
    N73: TMenuItem;
    RejectionReceivedgetSmart1: TMenuItem;
    getSmart2: TMenuItem;
    N74: TMenuItem;
    RejectionReceivedgetSmart2: TMenuItem;
    N75: TMenuItem;
    AddCustomerFLAG1: TMenuItem;
    BillingTools1: TMenuItem;
    N53: TMenuItem;
    N66: TMenuItem;
    N76: TMenuItem;
    N77: TMenuItem;
    N78: TMenuItem;
    N49: TMenuItem;
    ShowLossNotifications1: TMenuItem;
    BillingTools2: TMenuItem;
    N79: TMenuItem;
    N80: TMenuItem;
    N37: TMenuItem;
    N65: TMenuItem;
    ShowLossNotifications2: TMenuItem;
    N50: TMenuItem;
    N81: TMenuItem;
    MeterExchangeQuery1: TMenuItem;
    N82: TMenuItem;
    MeterExchangeQuery2: TMenuItem;
    N83: TMenuItem;
    E7Cancellation1: TMenuItem;
    ShowLibertyVends1: TMenuItem;
    N84: TMenuItem;
    PaypointAgencyLocator1: TMenuItem;
    FinancialTools1: TMenuItem;
    N85: TMenuItem;
    N86: TMenuItem;
    DirectSignupLetterGetSmart1: TMenuItem;
    GetSmartRenewalLetter1: TMenuItem;
    ShowLegacyPrePayVends1: TMenuItem;
    N87: TMenuItem;
    N88: TMenuItem;
    SuperCustomers1: TMenuItem;
    AggretagewithSuperCustomer1: TMenuItem;
    DisagregatefromSuperCustomer1: TMenuItem;
    PopUpSuperCust: TPopupMenu;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem21: TMenuItem;
    MenuItem23: TMenuItem;
    MenuItem24: TMenuItem;
    MenuItem25: TMenuItem;
    MenuItem27: TMenuItem;
    MenuItem28: TMenuItem;
    MenuItem29: TMenuItem;
    MenuItem31: TMenuItem;
    MenuItem33: TMenuItem;
    MenuItem49: TMenuItem;
    MenuItem50: TMenuItem;
    MenuItem51: TMenuItem;
    N58: TMenuItem;
    SuperCustomerStatement1: TMenuItem;
    ChangeDefaultSpanType1: TMenuItem;
    RejectIGT1: TMenuItem;
    SetStartDate1: TMenuItem;
    SetBillingProfile1: TMenuItem;
    PC11: TMenuItem;
    N021: TMenuItem;
    N031: TMenuItem;
    N041: TMenuItem;
    N051: TMenuItem;
    N061: TMenuItem;
    N071: TMenuItem;
    N081: TMenuItem;
    AddAdditionalCharges1: TMenuItem;
    N90: TMenuItem;
    DeleteAgreementfromCRM1: TMenuItem;
    LegacyPrepay1: TMenuItem;
    IssueReplacementQuantumCard1: TMenuItem;
    Losses1: TMenuItem;
    N67: TMenuItem;
    G_REL: TMenuItem;
    losses: TMenuItem;
    E_REL: TMenuItem;
    E_RELO: TMenuItem;
    N68: TMenuItem;
    G_RELO: TMenuItem;
    SuppressCatchUpDDs1: TMenuItem;
    PopUpSuppress: TPopupMenu;
    RemoveSuppressDDMarker1: TMenuItem;
    PopUpCustom: TPopupMenu;
    Custom_SPAN: TMenuItem;
    N91: TMenuItem;
    MTDSCustom: TOracleDataSet;
    SetCharge1: TMenuItem;
    COT2: TMenuItem;
    RemoveCOTAsifNeverVacated1: TMenuItem;
    ChangeDateMovedOut1: TMenuItem;
    COTTools1: TMenuItem;
    ChangeCOTMovedfinDate1: TMenuItem;
    N92: TMenuItem;
    M_AL: TMenuItem;
    ools1: TMenuItem;
    N93: TMenuItem;
    SmartMop: TMenuItem;
    N94: TMenuItem;
    custom_meter: TMenuItem;
    N95: TMenuItem;
    Custom_Billing: TMenuItem;
    ransfertoNeworExistingAgreement1: TMenuItem;
    N96: TMenuItem;
    AdditionalCharges1: TMenuItem;
    SPANOverride2: TMenuItem;
    N97: TMenuItem;
    T_CC: TMenuItem;
    N100: TMenuItem;
    FMS: TMenuItem;
    N98: TMenuItem;
    N99: TMenuItem;
    CustomerLetters2: TMenuItem;
    Letters1N: TMenuItem;
    Letters2N: TMenuItem;
    L_G_N: TMenuItem;
    L_E_N: TMenuItem;
    Custom_letters: TMenuItem;
    N102: TMenuItem;
    N103: TMenuItem;
    G_Debt: TMenuItem;
    G_DebtR: TMenuItem;
    E_DebtR: TMenuItem;
    E_Debt: TMenuItem;
    PopUpSmets: TPopupMenu;
    Smets_Vend: TMenuItem;
    Smets_vend_add: TMenuItem;
    smets_vend_deduct: TMenuItem;
    Smets_vend_set: TMenuItem;
    Smets_debt: TMenuItem;
    smets_debt_add: TMenuItem;
    smets_debt_deduct: TMenuItem;
    Smets_debt_Set: TMenuItem;
    N101: TMenuItem;
    Smets_Read: TMenuItem;
    Smets_Read_Bar: TMenuItem;
    Smets_txt: TMenuItem;
    Smets_txt_bar: TMenuItem;
    Smets_IHD: TMenuItem;
    smets_ihd_replace: TMenuItem;
    smets_ihd_add: TMenuItem;
    smets_ihd_remove: TMenuItem;
    smets_ihd_bar: TMenuItem;
    Smets_View: TMenuItem;
    DataflowHisopry1: TMenuItem;
    VendHistory1: TMenuItem;
    CommsData1: TMenuItem;
    extMEssageHistory1: TMenuItem;
    MeterReadings2: TMenuItem;
    Smets_Admin: TMenuItem;
    Smets_admin_bar: TMenuItem;
    Alarms1: TMenuItem;
    N104: TMenuItem;
    ProfileData1: TMenuItem;
    COTWizard1: TMenuItem;
    Treeview1: TVirtualStringTree;
    MopTree: TVirtualStringTree;
    Smets_Loan: TMenuItem;
    SPANOverride3: TMenuItem;
    PopUpMobile: TPopupMenu;
    SendSMSMessage1: TMenuItem;
    CopyNotestoAnotherCustomer1: TMenuItem;
    N105: TMenuItem;
    NoInstallTariffManagment1: TMenuItem;
    Popup_Losses: TPopupMenu;
    ShowScript1: TMenuItem;
    ShowSMSLog1: TMenuItem;
    N106: TMenuItem;
    SMETS_COS_ELECI: TMenuItem;
    InitiateCOSRequesttoUtilita1: TMenuItem;
    N107: TMenuItem;
    SMETS_COS_GASI: TMenuItem;
    InitiateCOSRequestGAIN1: TMenuItem;
    N108: TMenuItem;
    UnAllocateLibertyCustomerID1: TMenuItem;
    N110: TMenuItem;
    GER: TMenuItem;
    N111: TMenuItem;
    RiaseErroneousTransferRequest1: TMenuItem;
    ErroneousTrnasfers1: TMenuItem;
    N112: TMenuItem;
    CustomerDIsSatiisfied1: TMenuItem;
    aclMPANTree: TActionList;
    acAmendJob: TAction;
    acRescheduling: TAction;
    acCancel: TAction;
    acArrangeJobDtls: TAction;
    pupRescheduling: TPopupMenu;
    AmendJobDetails1: TMenuItem;
    CancelJobBooking1: TMenuItem;
    CancelJobBooking2: TMenuItem;
    ArrangeJobDetails1: TMenuItem;
    N113: TMenuItem;
   ChangeJobType1: TMenuItem;
    acDualFuelInstall: TAction;
    acElecOnlyInstall: TAction;
    acGasOnlyInstall: TAction;
    acFaults: TAction;
    DualFuelInstall1: TMenuItem;
    ElectricOnlyInstall1: TMenuItem;
    GasOnlyInstall1: TMenuItem;
    Faults1: TMenuItem;
    acChangeJobType: TAction;
    N114: TMenuItem;
    acFuelDirect: TAction;
    acRemoveFuelDirect: TAction;
    acAddFuelDirect: TAction;
    qrAgrFuelDir: TOracleDataSet;
    qrAgrFuelDirCUSTOMER_ID: TStringField;
    qrAgrFuelDirSTATUS: TStringField;
    qrAgrFuelDirLAST_UPDATED_BY: TStringField;
    qrAgrFuelDirLAST_UPDATED_DATE: TDateTimeField;
    qrAgrFuelDirPAYMENTS: TStringField;
    qrAgrFuelDirDESCRIPTION: TStringField;
    FueldirectThirdpartypayments1: TMenuItem;
    AddFuelDirect1: TMenuItem;
    N116: TMenuItem;
    RemoveFuelDirect3: TMenuItem;
    smets_ihd_pin: TMenuItem;
    CheckCVStatus1: TMenuItem;
    N25: TMenuItem;
    LmitNotesonTreeCount1: TMenuItem;
    N115: TMenuItem;
    N510: TMenuItem;
    N1010: TMenuItem;
    RefereAFriend1: TMenuItem;
    HideCheck: TCheckBox;
    AddComplaint1: TMenuItem;
    AddNoteEnquiryFlagDocComplaint1: TMenuItem;
    MarketingConsent1: TMenuItem;
    BillpayCard1: TMenuItem;
    N89: TMenuItem;
    LegacyPrepay2: TMenuItem;
    IssueD0190Key1: TMenuItem;
    ChangeJobPriority1: TMenuItem;
    AcPriorityChange: TAction;
    m_OrderNewStockItem_Customer: TMenuItem;
    m_ShowStockOrders: TMenuItem;
    m_OrderNewStockItem_Premise: TMenuItem;
    m_OrderNewStockItem_agreement: TMenuItem;
    m_OrderNewStockItem_Gas: TMenuItem;
    m_OrderNewStockItem_Elec: TMenuItem;
    N109: TMenuItem;
    N117: TMenuItem;
    N118: TMenuItem;
    N120: TMenuItem;
    N121: TMenuItem;
    S_TransferCredit: TMenuItem;
    N122: TMenuItem;
    S_COSLOSSE: TMenuItem;
    S_COSLOSSG: TMenuItem;
    N123: TMenuItem;
    ErroneousTransfers1: TMenuItem;
    N124: TMenuItem;
    Raiseerroneoustransferrequest1: TMenuItem;
    PopUpEmail: TPopupMenu;
    MenuItem32: TMenuItem;
    AcView: TAction;
    ViewJobInformation1: TMenuItem;
    N125: TMenuItem;
    M_Reassign: TMenuItem;
    grpbxReassignAgreements: TGroupBox;
    Label1: TLabel;
    edtAssignee: TEdit;
    btnReassign: TBitBtn;
    btnHideReassignAgreementPanel: TBitBtn;
    lblAssigneeName: TLabel;
    N126: TMenuItem;
    Action1: TAction;
    acHistoricCustDtl: TAction;
    N127: TMenuItem;
    mnuDUoSInvoicing: TMenuItem;
    N129: TMenuItem;
    VATExceptions1: TMenuItem;
    acVATExemption: TAction;
    N130: TMenuItem;
    VATExceptions2: TMenuItem;
    ShowAllCustomerComplaints1: TMenuItem;
 	est1: TMenuItem;
    ACWhereEngineer: TAction;
    BACS1: TMenuItem;
    DAPe: TMenuItem;
    DAPg: TMenuItem;
    SendD03062: TMenuItem;
    SendD0308: TMenuItem;
    SendD0309: TMenuItem;
    SendD03071: TMenuItem;
    ShowMyUtilitaData1: TMenuItem;
    divider131: TMenuItem;
    CreateAgreementttoPay1: TMenuItem;
    NATPDivider: TMenuItem;
    HistoricPayPlan: TMenuItem;
    ATPQuery: TOracleDataSet;
    SM_E: TMenuItem;
    SM_G: TMenuItem;
    SM_B: TMenuItem;
    N132: TMenuItem;
    N131: TMenuItem;
    CopyBill: TMenuItem;
    UAdivider: TMenuItem;
    UAandFPs1: TMenuItem;
    N133: TMenuItem;
    RefundAction: TMenuItem;
    MyUtilita1: TMenuItem;
    MyUtilitaWebFeatures: TMenuItem;
    PopUpSmets_DCC: TPopupMenu;
    Change1: TMenuItem;
    RelationshipRating1: TMenuItem;
    N134: TMenuItem;
    N135: TMenuItem;
    WANMatrix1: TMenuItem;
    WANQuery: TOracleDataSet;
    smets_vend_DCC: TMenuItem;
    smets_vend_add_DCC: TMenuItem;
    Smets_vend_deduct_DCC: TMenuItem;
    smets_vend_set_DCC: TMenuItem;
    smets_loan_DCC: TMenuItem;
    View2: TMenuItem;
    VendHistory2: TMenuItem;
    qrGetCustIcon: TOracleDataSet;
    qrGetCustIconCUSTOMER_TYPE_ID: TFloatField;
    qrGetCustIconDESCRIPTION: TStringField;
    qrGetCustIconICON_INDEX: TFloatField;
    PPMIDMessages1: TMenuItem;
    AlertsDCC: TMenuItem;
    DataflowHistory2: TMenuItem;
    Smets_Debt_DCC: TMenuItem;
    Smets_Debt_Add_DCC: TMenuItem;
    Smets_Debt_Deduct_DCC: TMenuItem;
    mnuRaiseFlag: TMenuItem;
    mnuDynamic: TMenuItem;
    ShowContractInformation1: TMenuItem;
    CancelServices1: TMenuItem;
    N136: TMenuItem;
    mnuDCCMeterReadings: TMenuItem;
    mnuPriorityNotification: TMenuItem;
    mnuViewPriorityNotification: TMenuItem;
    mnuRefreshPriorityNotification: TMenuItem;
    mnuProfileDataDCC: TMenuItem;
    PronounQuery: TOracleDataSet;
    EBSSPayment: TMenuItem;
    mnuIncomeAndExpenditureWebForm: TMenuItem;
    N137: TMenuItem;
    miRe_RateNow: TMenuItem;
    miRe_RateOvernight: TMenuItem;
    ReRateHistory1: TMenuItem;
    N138: TMenuItem;
    mnuCreditCheck: TMenuItem;
    SendExtracareletter1: TMenuItem;
    mniWinterWarmerFinancialAssistancePayment: TMenuItem;
    mnuGovernmentSchemes: TMenuItem;
    SendfurtherNOI1: TMenuItem;
    N140: TMenuItem;
    CheckComms: TMenuItem;
    PopUpSmartPay: TPopupMenu;
    SavingsTransMenuItem: TMenuItem;
    PowerPayEligibility: TMenuItem;
    N139: TMenuItem;
    UsageCalculatorAutomatic1: TMenuItem;
    mniWarmHomeDiscount: TMenuItem;
    mniEnergyBillsSupportScheme: TMenuItem;
    mniWinterWarmerPayment: TMenuItem;
    mniAlternativeFuelPayment: TMenuItem;
    mniTransferCreditFromSavingsToAgreement: TMenuItem;
    mniTransferCreditFromSavingsToMeter: TMenuItem;
    mniTransferDebitFromSavingsToAgreement: TMenuItem;
    N141: TMenuItem;
    mniChangeOfTenancy: TMenuItem;
    procedure BuildTree(Selection, DisplayOrder: String);
    procedure ResGridDblClick(Sender: TObject);
    // SJ-BSL - 02/05/2021 - Replacing constant assignment by Global Variable.
    Procedure ShowCustomerOnly(XNode: PVirtualNode; CustId, CustName, CustAddress, PremiseCount, CustomerDeceased, IsProspect, {: String; aCustIcon: Integer;} Debt, Lib_Id, IsLive: String);
    Procedure ShowSuperCustomerOnly(XNode:PVirtualNode; In_Id_Is_Cust, Full, CustId, CustName, CustAddress, PremiseCount, CustomerDeceased, IsProspect: String;
                                    {aCustIcon1: Integer;} Debt: String);
    Procedure BuildpremiseNode(premisename,premiseid,premisetype:string);
    Procedure BuildMPANNode(MPAN,regstatus,energisation_status,new_connection:String);
    procedure BuildElectricMeterNode(Xnode:Pvirtualnode);
    procedure Button1Click(Sender: TObject);
    procedure treeview1DblClick(Sender: TObject);
    procedure BuildEnquiriesNode(MPANNODE:PvirtualNode);
    procedure Historic1Click(Sender: TObject);
    procedure S_DflowsClick(Sender: TObject);
    procedure Enquiries1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    Procedure BuildCustomerFuelDirect(C_Id: String);
    procedure BuildCustomerNotes(C_id:string);
    procedure BuildCustomerNotifications(C_id:string);
    procedure BuildCustomerLosses(C_id:string);
    procedure AddCustomerNote1Click(Sender: TObject);
    procedure AddCustomerEnquiry1Click(Sender: TObject);
    procedure V_FullNotesClick(Sender: TObject);
    procedure V_NonClick(Sender: TObject);
    procedure ShowAllEnquiriesNotes1Click(Sender: TObject);
    procedure ShowAllCustomerComplaints1Click(Sender: TObject);
    procedure E_TakeOwnershipClick(Sender: TObject);
    procedure MeterDetailsReadings1Click(Sender: TObject);
    procedure MaintainCustomerDetailsClick(Sender: TObject);
    Procedure RefreshCustomerNode(XNode: PVirtualNode);
    procedure refreshAgreementnode(mynodeagreement:Pvirtualnode);
    procedure ExpandSpanNode(aSpanNode: PVirtualNode);
    procedure RefreshJBSPushBackNode(mynodeJOB:Pvirtualnode);
    procedure MenuItem1Click(Sender: TObject);
    procedure MaintainAccountHolder1Click(Sender: TObject);
    procedure MaintainAccountHoldersClick(Sender: TObject);
    procedure AddAgreement1Click(Sender: TObject);
    procedure P1Click(Sender: TObject);
    procedure P3Click(Sender: TObject);
    procedure ViewAggreement1Click(Sender: TObject);
    procedure A_ProdWizardClick(Sender: TObject);
    procedure RegisterItemClick(Sender: TObject);
    procedure EditAccountHoldersClick(Sender: TObject);
    procedure CallDataRecords1Click(Sender: TObject);
    procedure A_decClick(Sender: TObject);
    procedure A_revClick(Sender: TObject);
    procedure RemoveAccountHolder1Click(Sender: TObject);
    procedure MopHistoryClick(Sender: TObject);
    procedure NEWBTNClick(Sender: TObject);
    procedure BuildGasMeterNode(MPANNODE:Pvirtualnode);
    procedure BuildTelecomMeterNode(MPANNODE:PVirtualnode);
    procedure BuildCustomMeterNode(MPANNODE:PVirtualnode);
    procedure MaintainPremiseDetails1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure G_ReOrderClick(Sender: TObject);
    procedure ReOrder(OStatus:string);
    procedure E_Re_OrderClick(Sender: TObject);
    procedure T_ReOrderClick(Sender: TObject);
    procedure MaintainServicesNetworkFeatures1Click(Sender: TObject);
    procedure AddFeature(code,span:string);
    procedure RefreshAgentPremiseNode;
    procedure FriendsFamily1Click(Sender: TObject);
    procedure ShowFriend(No,IType,Bf:string);
    procedure QuarterlyStatement1Click(Sender: TObject);
    procedure QuarterlyStatementItemised1Click(Sender: TObject);
    procedure ErroneousTransfer1Click(Sender: TObject);
    Procedure ShowProduct(mynodeagreement:Pvirtualnode;Agreement_id:string;ShowBD:boolean;Fstatus:string);
    procedure ShowAgreement(mynodeagreement:Pvirtualnode;Iconindex:integer;AID,CID,Astatus,Astart,Aend,Aperiod:String;Sc:boolean;renewaldate:string);
    procedure ShowSpan(spannode:Pvirtualnode;spandesc,status,SSD,span,spantype:string;spanindex:integer;regid,servicetype,btacno,btssd,IGT,SPANEND,SpanEndReason,agid,locked,premid,fl1,fl2,fl3,fl4,fl5,fl6,fl7,fl8,fl9:string; custtype: Integer);
    procedure ShowLatestRatedUsage(xnode:Pvirtualnode;Agreement_id:string);
    procedure ShowAnyDisputes(xnode:Pvirtualnode;Agreement_id:string);
    procedure ShowLatestAccountReview(xnode:Pvirtualnode;Agreement_id:string);
    procedure ShowRatedUsage(Agreement_id:string);
     procedure Showsites(xnode:pvirtualnode;SPAN:string);
    procedure MeterDetailsReadings2Click(Sender: TObject);
    procedure ViewRefreshMPRN1Click(Sender: TObject);
    procedure ViewRefreshMPAN1Click(Sender: TObject);
    procedure ViewRefreshCallerLineID1Click(Sender: TObject);
    procedure MenuItem15Click(Sender: TObject);
    procedure AddProspectDetails1Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure BroadbandReorderClick(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure WebSignupLetter1Click(Sender: TObject);
    procedure InformationRequest11Click(Sender: TObject);
    procedure CustomerCredits1Click(Sender: TObject);
    procedure SACClick(Sender: TObject);
    procedure ShowFinancialHistory1Click(Sender: TObject);
    procedure ViewSPANSetupDetails1Click(Sender: TObject);
    procedure ViewSPANSetupDetails2Click(Sender: TObject);
    procedure View1Click(Sender: TObject);
    procedure ViewSPANSetupDetails3Click(Sender: TObject);
    procedure WelcomeLetter1Click(Sender: TObject);
    procedure CancelAllSubOrders1Click(Sender: TObject);
    procedure C1Click(Sender: TObject);
    procedure COT1Click(Sender: TObject);
    procedure M_SingleClick(Sender: TObject);
    procedure A_C_1Click(Sender: TObject);
    procedure MakeLIVeCOTmoveIN1Click(Sender: TObject);
    procedure UndofromOrderReadytoOrderPlaced1Click(Sender: TObject);
    procedure ViewCustomerTree1Click(Sender: TObject);
    procedure OnTesting(Sender: TObject);
    procedure ChangeSuplyStartDate1Click(Sender: TObject);
    procedure AquireAttachDocument1Click(Sender: TObject);
    procedure AddScannedDoc(CUSTID,Filename,role:string);
    procedure AddScannedDocument1Click(Sender: TObject);
    procedure Cancellation1Click(Sender: TObject);
    procedure elecomsApplicationForm1Click(Sender: TObject);
    procedure RequestMissingMPRN1Click(Sender: TObject);
    procedure RequestMissingMPAN1Click(Sender: TObject);
    procedure RequestMissingMPRNMPAN1Click(Sender: TObject);
    procedure ObjectionReceived1Click(Sender: TObject);
    procedure ObjectionReveived1Click(Sender: TObject);
    procedure IGTMENUClick(Sender: TObject);
    procedure ETUtilitaFault1Click(Sender: TObject);
    procedure ElectricETLetter1Click(Sender: TObject);
    procedure DUALETUtilitaFault1Click(Sender: TObject);
    procedure ManuallySetSPANStatusSSD1Click(Sender: TObject);
    procedure RatedIssuesClick(Sender: TObject);
    procedure ReRateMPAN1Click(Sender: TObject);
    procedure ShowSiteVisitInformation1Click(Sender: TObject);
    procedure Actioning(ActionText:string);
    procedure G_SETClick(Sender: TObject);
    procedure E_SETClick(Sender: TObject);
    procedure SetasET(Regid:string);
    procedure SetasDNB(Regid:string);
    procedure SetasBILL(Regid:string);
    procedure RemoveET(Regid:string);
    procedure SignupLetter1Click(Sender: TObject);
    procedure FriendsFamilyForm1Click(Sender: TObject);
    procedure InfoPackApplicationForm1Click(Sender: TObject);
    procedure d1Click(Sender: TObject);
    procedure MenuItem5Click(Sender: TObject);
    procedure MovetoNewPPMAgreement1Click(Sender: TObject);
    procedure GasReadRequired1Click(Sender: TObject);
    procedure ElectricReadRequired1Click(Sender: TObject);
    procedure DualReadsRequired1Click(Sender: TObject);
    procedure MenuItem6Click(Sender: TObject);
    procedure RegistrationHistory1Click(Sender: TObject);
    procedure G_RETClick(Sender: TObject);
    procedure E_RETClick(Sender: TObject);
    procedure CurrentTariffSheet1Click(Sender: TObject);
    procedure VacantPremiseProgrammedforDisconnection1Click(
      Sender: TObject);
    procedure raceExecutors2Click(Sender: TObject);
    procedure EnquiryRecevied1Click(Sender: TObject);
    procedure raceExecutorsFollowUp1Click(Sender: TObject);
    procedure SignupLetterVerbal1Click(Sender: TObject);
    procedure ShowAnnualUsagekWh2Click(Sender: TObject);
    procedure DoNotBillDisconnectedDeEnergised1Click(Sender: TObject);
    procedure DoNotBillDisconnectedDeEnergised2Click(Sender: TObject);
    procedure ChequeReceived1Click(Sender: TObject);
    procedure MenuItem7Click(Sender: TObject);
    procedure Gas1Click(Sender: TObject);
    procedure Elec1Click(Sender: TObject);
    procedure Dual1Click(Sender: TObject);
    procedure Surcharge1Click(Sender: TObject);
    procedure ET1Click(Sender: TObject);
    procedure AddDispute(Rcode,Rdesc:string);
    procedure Correspondance1Click(Sender: TObject);
    procedure ShowMopTree(Span:String);
    procedure RefreshMopPremiseNode(xnode:pvirtualnode);
    procedure BuildElectricMeterNodeMOP(Mpan,enddate:string);
    procedure ViewDataflows1Click(Sender: TObject);
    procedure BuildMOPNOTES(MPAN:string);
    procedure MenuItem13Click(Sender: TObject);
    procedure MenuItem8Click(Sender: TObject);
    procedure MopTreeDblClick(Sender: TObject);
    procedure BuildMopSiteAddress(Mpancore,filename:string;contracterror:boolean);
    procedure N60Click(Sender: TObject);
    procedure AddEnquiry1Click(Sender: TObject);
    procedure Add1Click(Sender: TObject);
    procedure AquireAttachDocument2Click(Sender: TObject);
    procedure ShowAllEnquiriesNotesDocs1Click(Sender: TObject);
    procedure AddServiceOrder1Click(Sender: TObject);
    procedure ShowAllServiceOrders1Click(Sender: TObject);
    procedure M_AUDITClick(Sender: TObject);
    procedure IGTAdminCharge1Click(Sender: TObject);
    procedure RegistrationHistory2Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure RemoveDonNotBillAllowBilling1Click(Sender: TObject);
    procedure RemoveDoNotBillAllowBilling1Click(Sender: TObject);
    procedure AddOneOffCharge1Click(Sender: TObject);
    procedure AddOneOffCharge2Click(Sender: TObject);
    procedure AddOneOffCharge3Click(Sender: TObject);
    procedure AddOneOffCharge4Click(Sender: TObject);
    procedure RefreshAccountOrder(Custid:string);
    procedure PrepareInstall1Click(Sender: TObject);
    procedure MenuItem9Click(Sender: TObject);
    procedure COTMoveIn1Click(Sender: TObject);
    procedure SUnrestricted1Click(Sender: TObject);
    procedure EEconomy71Click(Sender: TObject);
    procedure DoSpanChange(sType,span:string);
    procedure ReviveAgreement1Click(Sender: TObject);
    procedure PrepareInstall2Click(Sender: TObject);
    procedure FeedbackForm1Click(Sender: TObject);
    procedure gsmiletterClick(Sender: TObject);
    procedure ObjectionReceivedgetSmart1Click(Sender: TObject);
    procedure ObjectionReceivedgetSmart2Click(Sender: TObject);
    procedure RejectionReceivedgetSmart1Click(Sender: TObject);
    procedure RejectionReceivedgetSmart2Click(Sender: TObject);
    procedure AddCustomerFLAG1Click(Sender: TObject);
    procedure ShowLossNotifications1Click(Sender: TObject);
    procedure ShowLossNotifications2Click(Sender: TObject);
    procedure MeterExchangeQuery1Click(Sender: TObject);
    procedure MeterExchangeQuery2Click(Sender: TObject);
    procedure E7Cancellation1Click(Sender: TObject);
    procedure ShowLibertyVends1Click(Sender: TObject);
    procedure PaypointAgencyLocator1Click(Sender: TObject);
    procedure DirectSignupLetterGetSmart1Click(Sender: TObject);
    procedure GetSmartRenewalLetter1Click(Sender: TObject);
    procedure ShowLegacyPrePayVends1Click(Sender: TObject);
    procedure MenuItem11Click(Sender: TObject);
    procedure DisagregatefromSuperCustomer1Click(Sender: TObject);
    procedure AggretagewithSuperCustomer1Click(Sender: TObject);
    procedure SuperCustomerStatement1Click(Sender: TObject);
    procedure ChangeDefaultSpanType1Click(Sender: TObject);
    procedure RejectIGT1Click(Sender: TObject);
    procedure SetStartDate1Click(Sender: TObject);
    procedure PC11Click(Sender: TObject);
    procedure N021Click(Sender: TObject);
    procedure N031Click(Sender: TObject);
    procedure N041Click(Sender: TObject);
    procedure doBillPCchange(PC,span:string);
    procedure N051Click(Sender: TObject);
    procedure N061Click(Sender: TObject);
    procedure N071Click(Sender: TObject);
    procedure N081Click(Sender: TObject);
    procedure AddAdditionalCharges1Click(Sender: TObject);
    procedure DeleteAgreementfromCRM1Click(Sender: TObject);
    procedure IssueReplacementQuantumCard1Click(Sender: TObject);
    procedure G_RELClick(Sender: TObject);
    procedure DoNotObject(Span,AGid:string);
    procedure RemoveRelease(Span,custid:string);
    procedure E_RELClick(Sender: TObject);
    procedure E_RELOClick(Sender: TObject);
    procedure G_RELOClick(Sender: TObject);
    procedure RemoveSuppressDDMarker1Click(Sender: TObject);
    procedure SuppressCatchUpDDs1Click(Sender: TObject);
    procedure Custom_SPANClick(Sender: TObject);
    procedure SetCharge1Click(Sender: TObject);
    procedure SetGreenDeal(REGID:string);
    procedure RemoveCOTAsifNeverVacated1Click(Sender: TObject);
    procedure ChangeDateMovedOut1Click(Sender: TObject);
    procedure ChangeCOTMovedfinDate1Click(Sender: TObject);
    procedure M_ALClick(Sender: TObject);
    procedure SmartMopClick(Sender: TObject);
    procedure custom_meterClick(Sender: TObject);
    procedure TransferSpan;
    procedure ransfertoNeworExistingAgreement1Click(Sender: TObject);
    procedure AdditionalCharges1Click(Sender: TObject);
    procedure BEN_TEMP_SPANS_SPLIT;
    procedure ransferSPANS1Click(Sender: TObject);
    procedure T_CCClick(Sender: TObject);
    procedure FMSClick(Sender: TObject);
    procedure CustomerLetters2Click(Sender: TObject);
    procedure Letters1NClick(Sender: TObject);
    procedure Letters2NClick(Sender: TObject);
    procedure L_G_NClick(Sender: TObject);
    procedure L_E_NClick(Sender: TObject);
    procedure VendPaymentStatementReport1Click(Sender: TObject);
    procedure Custom_lettersClick(Sender: TObject);
    procedure AddDapMarker(Span,Agid:string);
    procedure RemoveDAPMarker(Span,Agid:string);
    procedure G_DebtClick(Sender: TObject);
    procedure E_DebtClick(Sender: TObject);
    procedure G_DebtRClick(Sender: TObject);
    procedure E_DebtRClick(Sender: TObject);
    procedure Smets_AdminClick(Sender: TObject);
    procedure Smets_vend_addClick(Sender: TObject);
    procedure smets_vend_deductClick(Sender: TObject);
    procedure Smets_vend_setClick(Sender: TObject);
    procedure smets_debt_addClick(Sender: TObject);
    procedure smets_debt_deductClick(Sender: TObject);
    procedure Smets_debt_SetClick(Sender: TObject);
    procedure Smets_ReadClick(Sender: TObject);
    procedure Smets_txtClick(Sender: TObject);
    procedure smets_ihd_replaceClick(Sender: TObject);
    procedure smets_ihd_removeClick(Sender: TObject);
    procedure smets_ihd_addClick(Sender: TObject);
    procedure DataflowHisopry1Click(Sender: TObject);
    procedure VendHistory1Click(Sender: TObject);
    procedure CommsData1Click(Sender: TObject);
    procedure MeterReadings2Click(Sender: TObject);
    procedure Alarms1Click(Sender: TObject);
    procedure extMEssageHistory1Click(Sender: TObject);
    procedure ShowSmetsMeterCommsSupplier(xnode:pvirtualnode;SPAN,MeterID,service,dateremoved,Role:string);
    procedure ShowSmetsMeterCommsMop(xnode:pvirtualnode;SPAN,MeterID,service,dateremoved,Role:string);
    procedure ProfileData1Click(Sender: TObject);
    procedure COTWizard1Click(Sender: TObject);
    procedure DoSPanOverride;

    procedure Treeview1GetText(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
         var CellText: String);


    procedure Treeview1PaintText(Sender: TBaseVirtualTree;
      const TargetCanvas: TCanvas; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType);
    procedure Treeview1Change(Sender: TBaseVirtualTree;
      Node: PVirtualNode);
    procedure FormCreate(Sender: TObject);
    procedure Treeview1Click(Sender: TObject);
    procedure Treeview1Expanding(Sender: TBaseVirtualTree; Node: PVirtualNode; var Allowed: Boolean);
    procedure MopTreePaintText(Sender: TBaseVirtualTree;
      const TargetCanvas: TCanvas; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType);
    procedure MopTreeInitNode(Sender: TBaseVirtualTree; ParentNode,
      Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);

    procedure MopTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType;
         var CellText: String);


    procedure MopTreeExpanding(Sender: TBaseVirtualTree;
      Node: PVirtualNode; var Allowed: Boolean);
    procedure MopTreeChange(Sender: TBaseVirtualTree; Node: PVirtualNode);
    procedure Treeview1InitNode(Sender: TBaseVirtualTree; ParentNode,
      Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
    procedure Smets_LoanClick(Sender: TObject);
    procedure UnAllocateLibertyCUstomerID1Click(Sender: TObject);
    procedure SPANOverride3Click(Sender: TObject);

    procedure CopyNotestoAnotherCustomer1Click(Sender: TObject);
    procedure NoInstallTariffManagment1Click(Sender: TObject);
    procedure ShowScript1Click(Sender: TObject);
    procedure ShowSMSLog1Click(Sender: TObject);
    procedure InitiateCOSRequestGAIN1Click(Sender: TObject);
    procedure HH1Click(Sender: TObject);
    procedure StartComplaintsProcess(Customerid,method:string);
    procedure CustomerDIsSatiisfied1Click(Sender: TObject);
    procedure acAmendJobExecute(Sender: TObject);
    procedure acReschedulingExecute(Sender: TObject);
    procedure acCancelExecute(Sender: TObject);
    procedure acArrangeJobDtlsExecute(Sender: TObject);
    procedure acReschedulingUpdate(Sender: TObject);
     procedure acChangeJobTypeUpdate(Sender: TObject);
    procedure acDualFuelInstallUpdate(Sender: TObject);
    procedure acDualFuelInstallExecute(Sender: TObject);
    procedure acChangeJobTypeExecute(Sender: TObject);
    procedure acRemoveFuelDirectExecute(Sender: TObject);
    procedure acAddFuelDirectExecute(Sender: TObject);
    procedure acFuelDirectExecute(Sender: TObject);
    procedure acAddFuelDirectUpdate(Sender: TObject);
    procedure acRemoveFuelDirectUpdate(Sender: TObject);
	  procedure smets_ihd_pinClick(Sender: TObject);
    procedure CheckCVStatus1Click(Sender: TObject);
    procedure N115Click(Sender: TObject);
    procedure N510Click(Sender: TObject);
    procedure N1010Click(Sender: TObject);
    procedure RefereAFriend1Click(Sender: TObject);
    procedure HideCheckClick(Sender: TObject);
    procedure AddComplaint1Click(Sender: TObject);
    procedure E_ErroneousClick(Sender: TObject);
    procedure MarketingConsent1Click(Sender: TObject);
    procedure BillpayCard1Click(Sender: TObject);

    procedure IssueD0190Key1Click(Sender: TObject);
    procedure AcPriorityChangeUpdate(Sender: TObject);
    procedure AcPriorityChangeExecute(Sender: TObject);
    procedure m_OrderNewStockItem_CustomerClick(Sender: TObject);
    procedure m_ShowStockOrdersClick(Sender: TObject);
    procedure m_OrderNewStockItem_PremiseClick(Sender: TObject);
    procedure m_OrderNewStockItem_GasClick(Sender: TObject);
    procedure m_OrderNewStockItem_agreementClick(Sender: TObject);
    procedure m_OrderNewStockItem_ElecClick(Sender: TObject);
    Procedure Raise_Hot_Note(aNType: integer; aCustomerId: string);
    procedure S_TransferCreditClick(Sender: TObject);
    Procedure Show_Hot_Note(aNID: string);
    function CheckHotNoteOpen(aNID: string): boolean;
    procedure S_COSLOSSEClick(Sender: TObject);
    procedure RiaseErroneousTransferRequest1Click(Sender: TObject);
    procedure Raiseerroneoustransferrequest1Click(Sender: TObject);
    procedure MenuItem32Click(Sender: TObject);
    procedure MopTreeGetImageIndex(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Kind: TVTImageKind; Column: TColumnIndex; var Ghosted: Boolean;
      var ImageIndex: TImageIndex);
    procedure Treeview1GetImageIndex(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Kind: TVTImageKind; Column: TColumnIndex;
      var Ghosted: Boolean; var ImageIndex: TImageIndex);
    procedure AcViewExecute(Sender: TObject);
    procedure btnHideReassignAgreementPanelClick(Sender: TObject);
    procedure M_ReassignClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnReassignClick(Sender: TObject);
    procedure edtAssigneeChange(Sender: TObject);
    procedure Treeview1Collapsing(Sender: TBaseVirtualTree; Node: PVirtualNode;
      var Allowed: Boolean);
    procedure mnuDUoSInvoicingClick(Sender: TObject);
    procedure acHistoricCustDtlExecute(Sender: TObject);
    procedure acVATExemptionExecute(Sender: TObject);
    procedure acVATExemptionUpdate(Sender: TObject);
     procedure ACWhereEngineerExecute(Sender: TObject);
    procedure ACWhereEngineerUpdate(Sender: TObject);

    procedure BuildSMSMenu;
    procedure SendSMSMessage1Click(Sender: TObject);
    procedure BACS1Click(Sender: TObject);
    procedure PopUpCustPopup(Sender: TObject);
    procedure SendD03061Click(Sender: TObject);
    procedure SendD0308Click(Sender: TObject);
    procedure SendD0309Click(Sender: TObject);
    procedure SendD03071Click(Sender: TObject);
    procedure SendG08061Click(Sender: TObject);
    procedure SendG08071Click(Sender: TObject);
    procedure SendG08081Click(Sender: TObject);
    procedure SendG08091Click(Sender: TObject);
    procedure ShowMyUtilitaData1Click(Sender: TObject);
    procedure SM_EClick(Sender: TObject);
    procedure SM_GClick(Sender: TObject);
    procedure CreateAgreementttoPay1Click(Sender: TObject);
    procedure HistoricPayPlanClick(Sender: TObject);
    procedure PopUpAgreementsPopup(Sender: TObject);
    procedure PopUpAccountPopup(Sender: TObject);
    procedure UAandFPs1Click(Sender: TObject);
    procedure CopyBillClick(Sender: TObject);
    procedure RefundActionClick(Sender: TObject);
    procedure MyUtilitaWebFeaturesClick(Sender: TObject);
    procedure RelationshipRating1Click(Sender: TObject);
    procedure AnnulGasContractClick(Sender: TObject);
    procedure AnnulElectricityContract1Click(Sender: TObject);
    procedure ShowContractInformation1Click(Sender: TObject);
    procedure CancelServices1Click(Sender: TObject);
    procedure WANMatrix1Click(Sender: TObject);
    procedure VendHistory2Click(Sender: TObject);
    procedure PPMIDMessages1Click(Sender: TObject);
    procedure LoanAmount1Click(Sender: TObject);
    procedure AlertsDCCClick(Sender: TObject);
    procedure DataflowHistory2Click(Sender: TObject);
    procedure Smets_Debt_Add_DCCClick(Sender: TObject);
    procedure Smets_Debt_Deduct_DCCClick(Sender: TObject);
    procedure mnuDCCMeterReadingsClick(Sender: TObject);
    procedure BuildS1EnrolledNode(mpxn:string;spannode:PVirtualNode);
    procedure mnuViewPriorityNotificationClick(Sender: TObject);
    procedure mnuRefreshPriorityNotificationClick(Sender: TObject);
    procedure mnuProfileDataDCCClick(Sender: TObject);
    procedure EBSSPaymentClick(Sender: TObject);
    procedure PopUpElectricPopup(Sender: TObject);
    procedure PopUpGasPopup(Sender: TObject);
    procedure PopUpSmets_DCCPopup(Sender: TObject);
    procedure mnuIncomeAndExpenditureWebFormClick(Sender: TObject);
    procedure miRe_RateNowClick(Sender: TObject);
    procedure miRe_RateOvernightClick(Sender: TObject);
    procedure ReRateHistory1Click(Sender: TObject);
    procedure mnuCreditCheckClick(Sender: TObject);
    procedure SendExtracareletter1Click(Sender: TObject);
    procedure mniWinterWarmerFinancialAssistancePaymentClick(Sender: TObject);
    procedure SendfurtherNOI1Click(Sender: TObject);
    procedure CheckCommsClick(Sender: TObject);
    procedure SavingsTransMenuItemClick(Sender: TObject);
    procedure PowerPayEligibilityClick(Sender: TObject);
    procedure PopUpPremisePopup(Sender: TObject);
    procedure UsageCalculatorAutomatic1Click(Sender: TObject);
    procedure mniWarmHomeDiscountClick(Sender: TObject);
    procedure mniEnergyBillsSupportSchemeClick(Sender: TObject);
    procedure mniWinterWarmerPaymentClick(Sender: TObject);
    procedure mniAlternativeFuelPaymentClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure mniTransferCreditFromSavingsToAgreementClick(Sender: TObject);
    procedure mniTransferDebitFromSavingsToAgreementClick(Sender: TObject);
	  procedure mniTransferCreditFromSavingsToMeterClick(Sender: TObject);
    procedure mniChangeOfTenancyClick(Sender: TObject);
  {$IFNDEF CRMTEST}
  private
  {$ELSE}
  protected
  {$ENDIF}
    fAgreementsToReassign: TAgreementCheckedNodeManager;  //Re-assign agreements between Customers Wrike Ref 151273193
    fAssignFromCustomerId: string;
    fSuspenseAGID: string;
    fSuspenseCustID : string;
    fPremiseDcc: Boolean;
    fCustomerNote: String;
    fCustomerAccountHolder: String;
    fCustomerPriorityNotification: String;

    FFuelDirect: Char;  // A: Added - R: Removed - N: Nothing yet.
    FCustIcon: Integer;
    FSuperCustIcon: Integer;

    fIsExpanding : boolean;

    { Private declarations }
    Function fOpenTwoWayAppointment(Const aNoCancel: Boolean): Boolean;
    Function fCloseTwoWayAppointment(Const aNoCancel: Boolean): Boolean;
    Property FuelDirect: Char Read FFuelDirect Write FFuelDirect Default 'N';
    Function isDCCMeter: Boolean;
    function GetCustomerPronoun(ContactID, CustomerID: string): string;
    // BSL-SJ - 11/08/2021 - CRM-511 - Control over Super Customer - Fixing Bug.
    Function  GetSuperCustIcon: Integer;
    Procedure SetSuperCustIcon(Const Value: Integer);

    Function isS1Enrolled(mpxn:string): Boolean;
    function isCosGainEnrolled(mpxn: string): Boolean;

    procedure DoSmets2Credit(aMode: integer);
    procedure DoSmets2Debt(aMode: integer);

    Function  fGetCustIcon(aCustTypeId: Integer): Integer;
    function IsValidFinancial: Boolean;
    function IsPrepayAndLive : Boolean;
    function CanSendFurtherNOI: Boolean;
    procedure RaiseFlag(Sender: TObject);
    procedure AddOneOffCharge(desc: String);

    // BSL-SJ - 11/08/2021 - CRM-511 - Control over Super Customer - Fixing Bug.
    Property  SuperCustIcon: Integer Read GetSuperCustIcon Write SetSuperCustIcon;
    { SMETS2 }
    procedure DoSmetsCreditDCC(AMode: Integer);
    procedure DoSmetsDebt(aMode: integer);
    procedure DoWarrentJobLock(Sender: TObject);
    procedure MeterTypeSwitch(span, ssd: String);
    procedure InsertCustomerAccessedAudit(const aUserID: string; const aCustomerID, aAgreementID, aPremiseID, aSpanID: variant);
    procedure AddCustomerNote(pTreeView: TVirtualStringTree);
    procedure GovernmentEnergySchemesScreen(xSchemes: String);
    procedure ExpandCustomerNode(aSender: TBaseVirtualTree; aCustomerNode: PVirtualNode; var oDataProtectionOK: boolean);
    procedure PersistStringValue(var aValue: string; aNewValue: string);
    function ShowInvoluntaryModeChangePopUp(const aCustomerId: Int64): Integer;
    function CustomerHasInvoluntaryModeChangeFlag(const aCustomerId: Int64): Boolean;

    procedure RefreshpremiseNode(MyNodePremise: PVirtualNode);
  protected
    //Re-assign agreements between Customers Wrike Ref 151273193
    procedure doHideAgreementReassignment; virtual;
    procedure doShowAgreementReassignment; virtual;

    function  validCustomerId(id,existingid: string): Boolean; virtual;

    procedure processAgreements; virtual;
    procedure reassignAgreements; virtual;

    procedure refreshCustomerTree(node: PVirtualNode; nag: Boolean; expand: Boolean); virtual;

    function  getSQLMoveAgreement(customer: string; agreement: string): string; virtual;

    procedure CosGainStatus (SPAN : String);

    procedure doCopyNotestoAnotherCustomer(copyTo: string); virtual;
    procedure doExpandAgreements(node: PVirtualNode); virtual;
    procedure doExpandNode(node: PVirtualNode); virtual;
    function  isAgreementNode(node: PVirtualNode): Boolean; virtual;
    //END Re-assign agreements between Customers Wrike Ref 151273193

    //CSLC-106 - DCC Verification Corrected
    function getMeterDCC(mpxn:string): Boolean;

    // PT-757
    function FundAvailab: Boolean;

    procedure ShowSmartPayScreen(aCustId: Int64);
    function IsSmartPay(aIdentifier: string): Boolean;
  public
    { Public declarations }
    procedure HideAgreementReassignment;
    property  SuspenseAGID : string Read fSuspenseAGID Write fSuspenseAGID;
    property  SuspenseCustID : string Read fSuspenseCustID Write fSuspenseCustID;
  end;

   PMyRec = ^TMyRec;
   TMyRec = Record
   D_Tel,
   D_Email : String;

   D_ContractRef : string;
   D_ServiceRef : string;
   D_ServiceLevelRef : string;
   D_GSP : string;
   D_EffectiveFrom : string;
   D_SupplierMPID : string;
   D_SupplierName : string;

   D_ACTIONED: string;
   D_Customer_Id:string;
   D_LIB_CUST_ID:STRING;
   D_CDEBT:string;
   D_Customer_Name: string;
   D_PremiseCount:string;
   D_Account_holder_id:string;
   D_Contact_ID:string;

   D_Agreement_ID:string;
   D_Reason:string;
   D_ET:boolean;
   D_Agreement_End_date:String;
   D_Agreement_Start_date:String;
   D_Period_id:String;
   D_Period_type:String;
   D_Sales_Ref:String;
   D_Premise_Id:string;
   D_Premise_Name:string;
   D_Premise_postcode:string;
   D_Service_ID:string;

   D_Span:string;

   D_SpanE:string;
   D_SpanG:string;
   D_SpanDCC_E:boolean;
   D_SpanDCC_G:boolean;

       D_SPAN_NC:string;
       D_FILENAME:string;
       D_ETDMOA:string;
       D_ContractError:Boolean;
       D_SpanEnd:string;
       D_regid:string;
       D_Status:string;
       D_spantype:string;
       D_spandesc:string;
       D_desc:string;
       D_SSD:string;
       D_Refno:String;
       D_Priority:string;
       D_Pdesc:string;
       D_InObjPeriod:string;
       D_Order:Integer;
   C_Data:String;
   C_Date_Raised:String;
   C_Record_id:String;
   C_Record_Status:String;
   C_Owner:String;
   C_Raised_By:String;
   c_postcode:string;
   C_firstline:string;
   M_SERVICE:string;
  // D_SPAN :String;
   M_EFSDMSMTD :String;
   M_METERID :String;
   M_REGISTERID :String;
   M_HH_Register :String;
   Caption: WideString;
   Index: Integer;
   FontColor: TColor;
   FontBold: Boolean;
   FontUnderline: Boolean;
   FontName: String;
   CheckType: TCheckType;
   D_Cust_Type: Integer;

   Metertype : string;
  end;

const
  FBalanceAccountIconIndex = 254;
  FPowerPayIconIndex = 316;

var
  FRM_Tree: TFRM_Tree;
  loop, cust_type:integer;
  x:char;
  Desc, Cust, mcust,mpremise,mmpan, Mailing,MPAS,status,mtype,
  oldregister,mregister,
  config,msid,efsdmsmtd,oldefsdmsmtd,oldmeterid,meterid,
  dateremoved,regstatus, CustName: String;
  ReviewNode, MyNodeCustomer, MyNodeSuperCustomer, CustomerDetailsNode,
  servicenode,spannode,metercommsnode,metersupplynode,
  mynodeagreement,mynodeagreementdetails,mynodeagreementitem,
  premisenode,premiseitemnode,disputesnode,
  mynodepremise,premisecontactitemnode,MPANnode,
  mynodeMopAgreement,MyNodeMopCustomer,mynodemopspan,mynodetop,mynodemopSite,RatedNode,mynodeprosp,RatedSubnode,mpasaddressnode,specialneedsnode,mynode1,mynodes,premiseContactNode,MeterConfigNode,pcconfignode,
  MeterNode,featurenode,featureitemnode,FeatureItemSubNode,MeterRegisterNode,ServiceOrderNode,EnquiryNode, MyUtilitaWalletPowerPayNode,
  notenode,mynodesuppress,fmsnode,fmssubnode,jbsnode,notificationsnode,notificationnode,jbssubnode, jbssubnode1,jbspushnode,notesubnode,notetopnode,gasnode,premdetailsnode,telecomsnode,mynodepaymentplan,mynodeproduct,FuelDirectNode,atpnode:PVirtualNode;
  treeupdating:boolean;
  mpanonly:boolean;
  CustCount, PremiseCount, MPANCount: LongInt;
  Comments1,commsstatus,supplystatus:string;
  TREEENQUIRYResolved,ShowReassign{,IS_EXPANDING}:BOOLEAN;
  FirstLine:String;
  mpan:string;

  // comeent no eeffect

  CustId, Password, PasswordDate, PremiseName, PremiseId, CustomerDeceased, PremiseType,
  Ah, PrevAh, Relationship, SpecialNeeds, CustomerType, AH_Order, AH_TYPE, AH_DOB, AH_Contact_Method,
  AH_ID,AH_CONTACT_TITLE_ID,AH_INITIALS,AH_SURNAME,AH_FORENAME,AH_DISPLAY_NAME,AH_ADDITIONAL_INFORMATION,
  AH_SPECIAL_NEEDS_INFORMATION,AH_TELEPHONE_NO_DAY,AH_TELEPHONE_NO_EVE,AH_TELEPHONE_NO_MOBILE,AH_EMAIL,AH_FAX,
  agreement_start_date,agreement_end_date,agreement_id,contdet,telno,
  agreement_status,agreement_status_id,oldagreement,oldpremise,premaddr,
  service_id,spantype,span_type_id,
  ssd,eftssd,comments,en_status,new_conn,osid,oldspan,spanendreason,
  IGT,Cdebt: String;
  no_of_holders,

  PaymentIcon,
  super,
  hi, si, treelimit: Integer;

  TreeData: PMyRec;
  XNode   : PVirtualNode;
  NodeData: PMyRec;


implementation

uses
  loginunit, Main, MainSearch, AccountHolders, AddAccountHolder,
  AddNewAgreement, BankDetails, PremiseandSupplies, AddnewCustomerWizard,
  Enquiries, EnqSummary, EnquirySearchUnit, MaintainCustomer, Telephone,
  paymentSchedule, DFLOWS, NHH_Metering, GasStatus, BatchCalendar,
  TServices, FriendsFamily, ReportsUnit, AttaatchDoc, GasMetering,
  RatedUsage, InitialProspect, datepick, Credits_Awarded,
  Financial_History, SpanDetails, MoveOutPremise, Authority,
  Spanoverride, RatingErrors, RateElectric, Rating_Gas, LedgerUtils,
  singledate, FAFOPTIONS, account_reviews, AccountReview, Charge_Or_Credit,
  DflowsMop, NHH_Metering_MOP, Processing, MOServiceOrders, Custom,
  CustAccesslog, Prefs, DataModule, Common, oneoffcharge, prepareinstall,
  GSMI, LossSupplier, liberty_vend_codes, legacy_vends,
  selectSuperCust, NextFileSeq, About, DefaultSpanType,
  gas_quantum_requests, custom_meters, nhhmeterinstall, greendeal,
  liberty_customerid, custom_metering, wmol_book_job, ratingcustom,
  span_transfer, credit_control, FMS_Fault, addtojbs, letter_templates,
  vend_payment_summary, smets, smets_manage_credit, smets_debt, Smets_txt,
  smets_removedevice, smets_flow_history, smets_vend_hist, smets_comms,
  SMETS_READINGS, smets_alarms,smets_alarms_DCC, smets_txt_history, SMETS_PROFILE,
  cot_wizard,send_sms, copynotes, no_install_tariff, Loss_Outcome, helper,
  sms_message_history, smets_cos_gain, DMJBS, OneWayAppointment, CancelJob, TwoWayAppAmend, Data,
  Base, IHD_PIN,GER, ELEC_ET, cverrors, DMIMAGES,HHAuditLog, mpanlookup,
  OrderNewStockItem,ShowStockOrders, CreditCheck,
  Refer_Friend, Customer_Consent, Billpay, D0190, job_priority,
  DUoSPassCharges , //added by maryam on wrike ticket 171744089
  TransferCredit, //added by maryam on 23/11/2016
  jbsweb, GAS_ET, rateaccounts, System.UITypes,
  TreeViewIterator, TreeViewIteratorInterface, HistoricCustDtls, VAT_Exemption, variants, DDMain,MyUtilita,
  DAP306Requests, DAP307Requests, DAP308Requests, DAP309Requests, PkDebtInfo,
  CancelContract, CollectionsStatusChange, AgreePaySchedule, SuperCustBill, Refunds, DPACheck, RelationshipRatingChange,
  smets_DCC, WAN_Matrix, smets_manage_credit_dcc, smets_message_history_dcc, smets_debt_DCC, System.StrUtils, PriorityNotifications,
  AnnulContract,  ContractInformation, EBSSPayment, CrmCommon, Reratehistory, UELSqlUtils, WinterWarmerFinancialAssistancePayment,
  smets_check_comms_DCC, power_pay_history, UsageCalcAuto, DataCapture, Smets_Readings_Dcc,
  Smets_Flow_History_Dcc, Smets_Vend_Hist_Dcc, SmetsCommon, BillingCommon, RateAccountsWrapper, CustomerCommon, uChangeOfTenancy,
  TransferCreditFromSavingsToAgreement, SavingsTransactionsToAgreement, uTransferCreditFromSavingsToMeter;  

{$R *.DFM}

procedure TFRM_Tree.InsertCustomerAccessedAudit(const aUserID: string; const aCustomerID, aAgreementID, aPremiseID, aSpanID: Variant);
var
  oldCursor: TCursor;
begin
  oldCursor := Screen.Cursor;
  Screen.Cursor := crHourGlass;
  try
    try
      gSQLUtil.InsertRecord('crm.customer_accessed_audit', TRANSACTION_YES,
        ['date_time_accessed', otDate, Now,
         'user_id', otString, aUserId,
         'customer_id', otLong, aCustomerId,
         'agreement_id', otLong, aAgreementId,
         'premise_id', otLong, aPremiseId,
         'span', otString, aSpanId]);
    except
      on e:Exception do
        MessageDlg(Format('Error logging customer account access: %s', [e.Message]), mtError, [mbOk], 0);
    end;
  finally
    Screen.Cursor := oldCursor;
  end;
end;

Procedure TFRM_Tree.BuildTree(Selection, DisplayOrder: String);
Const cCommaSpc = ', ';
Var
   x, SpanIndex: Integer;
   Span, RegId, ServiceType, bTacNo, bTssd, Spanend, Agid, Debt, Locked, Premid,
   Fl1, Fl2, Fl3, Fl4, Fl5, Fl6, Fl7, F18, F19: String;
Begin
  FRM_Tree.Tag   := StrToInt(DisplayOrder);
  Selection      := 'ALL';
  TreeUpdating   := True;
  TreeView1.Clear;
  mCust          := 'dfsdLEEOK';
  mPremise       := 'dfsdLEEOK';
  mMPAN          := 'dfsdLEEOK';
  CustCount      := 0;
  PremiseCount   := 0;
  MPANCount      := 0;
  ////////////////////////////////////////////////////////////
  // Start to Build Tree                                    //
  // Tag 1,2,3,4 indicates which oder the tree is formatted //
  // 0= Display super customer  First
  // 1= Display Customer first                              //
  // 2= Display agreement First                             //
  // 3=Display Premise first                                //
  // 4=Display Supply point first                           //
  ////////////////////////////////////////////////////////////

  ////////////////////////////////////////////////////////////
  // 1 CUSTOMER ONLY LEVEL                                  //
  ////////////////////////////////////////////////////////////
  HideCheck.Visible := False;
  Screen.Cursor     := crhourglass;
  Try
   FRM_Main_Search.Enabled := False;
  Except
  End;

  With GeneralQuery do
    Begin
      Close;
      Sql.Clear;
      DeleteVariables;
      Sql.Add('SELECT SUSPENSE_ACCOUNT_NO FROM SALESLEDGER.BACS_SUSPENSE_ACCOUNT_CONFIG WHERE ACTIVE LIKE ''Y''');
      Open;
      DeleteVariables;
    End;

  SuperCustIcon   := 0;       // BSL-SJ - 11/08/2021 - CRM-511 - Control over Super Customer - Fixing Bug.
  fSuspenseAGID   := GeneralQuery.Fields[0].AsString;
  fSuspenseCustID := FRM_Common.GetCustomerIdfromAgreementid(SuspenseAGID);
  GeneralQuery.SQL.Clear;

  If (FRM_Tree.Tag < 2) or (FRM_Tree.Tag = 5) then
    Begin
      FRM_MAIN_SEARCH.CustomerQuery.First;
      mpanonly := False;
      Super    := 0;

      If FRM_Tree.Tag = 0 then
        Begin
         Super    := 1;
         Cust     := FRM_Main_Search.CustomerQuery.Fields[0].Text;
         CustName := FRM_Main_Search.CustomerQuery.Fields[1].Text;
         //mynodeSuperCustomer:=treeview1.items.Add(nil, 'Super Customer '+cust+' - '+Custname);

         MyNodeSuperCustomer := Treeview1.AddChild(Nil);
         NodeData            := Treeview1.GetNodeData(MyNodeSuperCustomer);
         NodeData.Caption    := 'Super Customer ' + Cust + ' - ' + CustName;
         NodeData.Index      := -1;
         Treeview1.Selected[MyNodeSuperCustomer] := True;

         ShowSuperCustomerOnly(MyNodeSuperCustomer, 'N', 'Y', Cust, CustName, '0', '0', '', 'N', {214,} '0');
        End;

      For x := 1 to FRM_Main_Search.CustomerQuery.RecordCount do
        Begin
           Caption := 'Customer View';
           With FRM_Main_Search.CustomerQuery do
             Begin
               Cust      := Fields[0].Text;
               CustName  := Fields[1].Text;
               FCustIcon := Fields[2].Value;
               Mailing   := EmptyStr;

               if Fields[3].Text  <> EmptyStr then Mailing := Mailing+Fields[3].Text  + cCommaSpc;
               if Fields[4].Text  <> EmptyStr then Mailing := Mailing+Fields[4].Text  + cCommaSpc;
               if Fields[5].Text  <> EmptyStr then Mailing := Mailing+Fields[5].Text  + cCommaSpc;
               if Fields[6].Text  <> EmptyStr then Mailing := Mailing+Fields[6].Text  + cCommaSpc;
               if Fields[7].Text  <> EmptyStr then Mailing := Mailing+Fields[7].Text  + cCommaSpc;
               if Fields[8].Text  <> EmptyStr then Mailing := Mailing+Fields[8].Text  + cCommaSpc;
               if Fields[9].Text  <> EmptyStr then Mailing := Mailing+Fields[9].Text  + cCommaSpc;
               if Fields[10].Text <> EmptyStr then Mailing := Mailing+Fields[10].Text + cCommaSpc;
               if Fields[11].Text <> EmptyStr then Mailing := Mailing+Fields[11].Text + cCommaSpc;

               Mailing := Mailing + Fields[12].Text + EmptyStr;
               Debt    := Fields[63].Text;
              End;

          If Cust = EmptyStr then
            Cust := '0';

          CustomerDeceased := EmptyStr;
          PremiseCount     := 0;  // premise Count

          If Cust <> mCust then
            Begin
              If Super = 0 then
                Begin
                 // mynodeCustomer:=treeview1.items.Add(nil, 'Customer '+cust+' - '+Custname)
                  MyNodeCustomer   := Treeview1.AddChild(Nil);
                  NodeData         := Treeview1.GetNodeData(MyNodeCustomer);
          //        NodeData.Caption := 'Customer ' + Cust + ' - ' + CustName;
                  NodeData.Index   := 3;
          //        Treeview1.Selected[MyNodeCustomer] := True;
          //        ShowCustomerOnly(mynodeCustomer,cust,custname,mailing,inttostr(premisecount),customerdeceased,FRM_MAIN_SEARCH.CustomerQuery.fields[53].text, {FCustIcon,} debt,FRM_MAIN_SEARCH.CustomerQuery.fields[69].text,FRM_MAIN_SEARCH.CustomerQuery.fields[77].text);
                End
              Else
                Begin
                 // mynodeCustomer:=treeview1.items.Addchild(mynodeSuperCustomer,'Customer '+cust+' - '+Custname);
                  MyNodeCustomer   := Treeview1.AddChild(MyNodeSuperCustomer);
                  NodeData         := Treeview1.GetNodeData(MyNodeCustomer);
          //        NodeData.Caption := 'Customer ' + Cust + ' - ' + CustName;
                  NodeData.Index   := -1;
          //        Treeview1.Selected[MyNodeCustomer] := True;
          //        ShowCustomerOnly(MynodeCustomer,cust,custname,mailing,inttostr(premisecount),customerdeceased,FRM_MAIN_SEARCH.CustomerQuery.fields[53].text, {FCustIcon,} debt,FRM_MAIN_SEARCH.CustomerQuery.fields[69].text,FRM_MAIN_SEARCH.CustomerQuery.fields[77].text);
                End;

              // BSL - 22/06/2021 - Moved for Optimization.
              NodeData.Caption                   := 'Customer ' + Cust + ' - ' + CustName + GetCustomerPronoun(EmptyStr, Cust);
              Treeview1.Selected[MyNodeCustomer] := True;
              ShowCustomerOnly(MyNodeCustomer, Cust, CustName, Mailing, IntToStr(PremiseCount), CustomerDeceased, FRM_MAIN_SEARCH.CustomerQuery.Fields[53].Text, {FCustIcon,} Debt, FRM_MAIN_SEARCH.CustomerQuery.Fields[69].Text, FRM_MAIN_SEARCH.CustomerQuery.Fields[77].Text);
            End;

          mCust := Cust;
          FRM_Main_Search.CustomerQuery.Next;
          HideCheck.Visible := True;
        End;
    End

 ////////////////////////////////////////////////////////////
 // 2 Agreement                                            //
 ////////////////////////////////////////////////////////////
 Else
 If frm_tree.tag=2 then
 Begin
  with frm_main_search.customerquery do
  Begin
   oldagreement:='Lee';
   while not eof do
   Begin
    agreement_status_id:=fields[30].text;
    agreement_status:=fields[27].text;
    Agreement_Start_Date:=fields[24].text;
    Agreement_End_Date:=fields[25].text;
    Agreement_id:=fields[23].text;
      if (agreement_id<>oldagreement) and (agreement_id<>'') then
    Begin
     desc:='';
     //mynodeAgreement:=Treeview1.items.AddChild(nil,desc);

     mynodeagreement:=Treeview1.Addchild(Nil);
     nodeData := Treeview1.GetNodeData(mynodeagreement);
     NodeData.caption := 'Agreement';
     NodeData.Index := -1;
     ShowAgreement(mynodeagreement,71,Agreement_id,fields[0].text,Agreement_status,agreement_start_Date,Agreement_end_date,fields[29].text,true,fields[66].text);
     //Status:=FRM_Financial_History.GetStatus(Agreement_id);
     Status := TFinancialHistoryInfo.GetFinancialStatusText(StrToInt64(Agreement_id));
     showproduct(mynodeagreement,agreement_id,false,status);
     treeview1.selected[mynodeagreement]:=true;
    end;
    oldagreement:=agreement_id;
    next;
   end;
  end;
 end

 ////////////////////////////////////////////////////////////
 // 3 Premise                                              //
 ////////////////////////////////////////////////////////////
 else
 If frm_tree.tag=3 then
 Begin
  with frm_main_search.customerquery do
  Begin
   oldPremise:='Lee';
   while not eof do
   Begin
    if (fields[32].text<>oldpremise) and (fields[32].text<>'') then
    Begin
     premaddr:='';
     if fields[33].text<>'' then premaddr:=premaddr+fields[33].text+',';
     if fields[34].text<>'' then premaddr:=premaddr+fields[34].text+',';
     if fields[35].text<>'' then premaddr:=premaddr+fields[35].text+',';
     if fields[36].text<>'' then premaddr:=premaddr+fields[36].text+',';
     if fields[37].text<>'' then premaddr:=premaddr+fields[37].text+',';
     if fields[38].text<>'' then premaddr:=premaddr+fields[38].text+',';
     if fields[39].text<>'' then premaddr:=premaddr+fields[39].text+',';
     if fields[40].text<>'' then premaddr:=premaddr+fields[40].text+',';
     if fields[41].text<>'' then premaddr:=premaddr+fields[41].text+',';
     premaddr:=premaddr+fields[42].text;
     premaddr:=premaddr+' - ['+fields[32].text+']';
     //mynodepremise:=Treeview1.items.AddChild(nil,'Premises - '+premaddr);

     mynodepremise:=Treeview1.Addchild(Nil);
     nodeData := Treeview1.GetNodeData(mynodepremise);
     NodeData.caption := 'Premises - '+premaddr;
     NodeData.Index := fields[43].value;

     treeview1.selected[mynodepremise]:=true;

     cot1.visible:=true;  // Move Out option
     cot2.visible:=false;  // Tools Option
     // Check if Moving Out

    { mynodepremise.imageindex:=fields[43].value;
     mynodepremise.selectedindex:=mynodepremise.imageindex; }

     // Check if Moving Out
     if fields[59].text<>'' then
     Begin
      cot1.visible:=false; // Dont Show COT option if already vavacted
      cot2.visible:=true;  // Show COT Tools if Vacated
      if strtodate(fields[59].text)<date then
      Begin
       NodeData.caption:=NodeData.caption+' - Vacated on '+fields[59].text;
       NodeData.index:=142;
       //mynodepremise.selectedindex:=mynodepremise.imageindex;
       NodeData.Fontcolor:=clmaroon;
      end
      else
      Begin
       NodeData.caption:=NodeData.caption+' - Vacating on '+fields[59].text;
       NodeData.index:=139;
      // mynodepremise.selectedindex:=mynodepremise.imageindex;
      End;
     end;

     NodeData.D_premise_Id :=   fields[32].text;
     NodeData.D_agreement_Id := fields[23].text;
     mynode1:=Treeview1.AddChild(mynodepremise);
    end;
    oldpremise:=fields[32].text;
    next;
   end;
  end;
 end

 ////////////////////////////////////////////////////////////
 // 4 SPANS                                                //
 ////////////////////////////////////////////////////////////
 else
 If frm_tree.tag=4 then
 Begin
  // Get Customer Type
   with GeneralQuery Do
   begin
    close;
    DeleteVariables;
    DeclareVariable(':CUSTID',otLong);
    sql.clear;
    sql.add('SELECT');
    sql.add('C.CUSTOMER_TYPE_ID ');
    sql.add('FROM');
    sql.add('  CRM.CUSTOMER C ');
    sql.add('WHERE');
    sql.add('  C.CUSTOMER_ID =:CUSTID');
    SetVariable('CUSTID',FRM_MAIN_SEARCH.CustomerQuery.Fields[0].AsLargeInt);
    open;
    deletevariables;
   end;
   if Generalquery.recordcount<>0 then
   begin
     cust_type := generalquery.fields[0].AsInteger;
   end;


  with frm_main_search.customerquery do
  Begin
   oldspan:='Lee';
   while not eof do
   Begin
    SPAN:=fields[46].text;
    if SPAN<>oldspan then
    Begin
     if fields[50].text<>'' then spanindex:=fields[50].value
     else spanindex:=1;

     try
      strtodate(copy(fields[51].text,11, 10));
      BTSSD:=copy(fields[51].text,11, 10)
     except
      BTSSD:='';
     end;
     btacno:=fields[55].text;
     SSD:=fields[48].text;
     Status:=fields[49].text;
     regid:=fields[54].text;
     servicetype:=fields[45].text;
     spantype:=fields[44].text;
     IGT:=fields[57].text;
     SPANEND:=fields[58].text;
     SpanEndReason:=fields[60].text;
     locked:=fields[67].text;
     agid:=fields[23].text;
     premid:=fields[32].text;
     // Span Flags
     fl1:=fields[70].text;
     fl2:=fields[71].text;
     fl3:=fields[72].text;
     fl4:=fields[73].text;
     fl5:=fields[74].text;
     fl6:=fields[75].text;
     fl7:=fields[76].text;
     // field 77 is live customer
     try
     f18:=fields[78].text;  // 3 Phase Marker
     f19:=fields[79].text;  // Related MPAN marker
     except
      f18:='';
      f19:='';
     end;
     //showmessage(fl1+'-'+fl2+'-'+fl3+'-'+fl4+'-'+fl5+'-'+fl6+'-'+fl7);
     if span<>'' then
     Begin
      //spannode:=treeview1.items.addchild(nil,'test');
      spannode:=Treeview1.Addchild(Nil);
      nodeData := Treeview1.GetNodeData(spannode);
      NodeData.caption := 'test';
      NodeData.Index := -1;
      ShowSpan(spannode,fields[47].text,status,SSD,span,spantype,spanindex,regid,servicetype,btacno,btssd,IGT,SPANEND,SpanEndReason,agid,locked,premid,fl1,fl2,fl3,fl4,fl5,fl6,fl7,f18,f19, cust_type);
      treeview1.selected[spannode]:=true;
     end;
    end;
    oldspan:=span;
    next;
   End;
  end;
 end;

// BSL - 29/06/2021-CRM-510-Super Customer menu pops up for the last commercial account under the Super customer profile.
                         // Happens when menu is accessed first time after that shows correct menu
//  if frm_tree.tag=0 then
// Begin
//  treeview1.popupmenu:=popupsupercust;
// end;

 treeupdating:=false;
  try
  frm_main_search.enabled:=true;
 except
 end;
 screen.Cursor:=crdefault;
 FRM_Common.TrimAppMemorySize;
end;

procedure TFRM_Tree.ResGridDblClick(Sender: TObject);
begin
 Buildtree('One','3');
end;

procedure TFRM_Tree.ShowContractInformation1Click(Sender: TObject);
var
  Customer_Id: string;
begin
  xnode := treeview1.FocusedNode;
  TreeData := treeview1.GetNodeData(xnode);
  Customer_Id := TreeData.D_Customer_Id;

  if not assigned(FRM_Contract_Info) then
    Application.CreateForm(TFRM_Contract_Info, FRM_Contract_Info);
  FRM_Contract_Info.CustomerId := Customer_Id;
  FRM_Contract_Info.ShowModal;
end;

Procedure TFRM_Tree.ShowCustomerOnly(XNode: PVirtualNode; CustId, CustName, CustAddress, PremiseCount, CustomerDeceased, IsProspect, {: String; aCustIcon: Integer;} Debt, Lib_Id, IsLive: String);
//////////////////////////////////////////////////////////////////////////////////////////
// Just Display Top Level Customer Details Only, from Search Query                      //
//////////////////////////////////////////////////////////////////////////////////////////
Var
//MyRecPtr   : PMyRec;
  DebtMessage: String;
  bDebt, k   : Integer;
Begin
  NodeData                  := Treeview1.GetNodeData(XNode);
  Treeview1.Selected[XNode] := True;

  // check if its a super customer, if so show super customer node
  // SJ-BSL - 30/04/2021 - Change 214 Icon by 309..311 (Changed back 07/05/2021)
  // SJ-BSL - 02/05/2021 - Replacing constant assignment by Global Variable.
  If FCustIcon = 214 then
    Begin
      // BSL-SJ - 11/08/2021 - CRM-511 - Control over Super Customer - Fixing Bug.
      // BSL - 22/06/2021 - CRM-511 - Control over Super Customer.
      //FSuperCustIcon      := FCustIcon;
      MyNodeSuperCustomer := XNode;
      ShowSuperCustomerOnly(MyNodeSuperCustomer, 'Y', 'Y', CustId, CustName, CustAddress, PremiseCount, CustomerDeceased, IsProspect, {aCustIcon,} Debt);
      // BSL - 11/05/2021 - Add FCustomIcon to avoid show a wrong PopUp Menu.
      Treeview1.PopupMenu := PopUpSuperCust;
    End // If
  Else
    Begin
      // BSL - 12/05/2021 - Code Optimization.
      bDebt := 0;
      Val(Debt, bDebt, k);

      Case bDebt of
      1: DebtMessage :=  '** Credit Control - Follow Up Letter Issued **';
      2: DebtMessage :=  '** Credit Control - Final Reminder Letter Issued **';
      3: DebtMessage :=  '** Credit Control - Prepay Install Letter Issued **';
      4: DebtMessage :=  '** Credit Control - Central Recoveries Letter Issued **';
      5: DebtMessage :=  '** Credit Control - Warrant Letter Issued **';
      6: DebtMessage :=  '** Credit Control - Debt Collector Letter Issued **';
      Else // '7' or ''
        DebtMessage := EmptyStr;
      End; // Case

//      If Debt = '1' then DebtMessage:='** Credit Control - Follow Up Letter Issued **';
//      if Debt='2' then DebtMessage:='** Credit Control - Final Reminder Letter Issued **';
//      if Debt='3' then DebtMessage:='** Credit Control - Prepay Install Letter Issued **';
//      if Debt='4' then DebtMessage:='** Credit Control - Central Recoveries Letter Issued **';
//      if Debt='5' then DebtMessage:='** Credit Control - Warrant Letter Issued **';
//      if Debt='6' then DebtMessage:='** Credit Control - Debt Collector Letter Issued **';
//      if Debt='7' then DebtMessage:='';
//      if Debt='' then DebtMessage:='';

      //New(MyRecPtr);
      NodeData.D_Customer_ID   := CustId;
      NodeData.D_LIB_CUST_ID   := Lib_Id;
      NodeData.D_Customer_Name := CustName;
      NodeData.D_PremiseCount  := PremiseCount;
      NodeData.D_CDEBT         := Debtmessage;
      //mynodecustomer.data:=MyRecPtr;

      Inc(CustCount);
      // SJ-BSL - 30/04/2021 - Replacing constant assignment by Global Variable.
      NodeData.Index := FCustIcon; //303; // Customer
      //mynodeCustomer.selectedindex:=mynodecustomer.imageindex;
      NodeData.Caption := EmptyStr;
      NodeData.Caption := 'Customer ' + Cust + ' - ' + CustName + GetCustomerPronoun(EmptyStr, Cust);
      // Show Multi premise Customer
      If StrToInt(PremiseCount) > 1 then
        Begin
          NodeData.Caption := EmptyStr;
          NodeData.Caption := 'Customer ' + Cust + ' - ' + CustName + ' - (' + PremiseCount + ' premises)';
          NodeData.Index   := 40;
        //  mynodeCustomer.Imageindex:=40;
        End;

      If Lib_Id <> EmptyStr then
        NodeData.Caption := NodeData.Caption + ' -(use ' + Lib_Id + ' in Liberty)';

      If DebtMessage <> EmptyStr then
        Begin
          NodeData.FontColor := clMaroon;
          NodeData.FontBold  := True;
          NodeData.Caption   := NodeData.Caption + #10 + DebtMessage;
        End;

      If IsLive = 'N'  then
        NodeData.FontColor := clRed;

      // Check for Deceased Customer
      If Customerdeceased <> EmptyStr then
        Begin
          If NodeData.Index = 36 then
            NodeData.Index := 46
          Else
            NodeData.Index := 47;

          NodeData.Caption := NodeData.Caption + ' - (Customer Deceased ' + CustomerDeceased + ')';
        End;

      If IsProspect = 'Y' then
        Begin
         // SJ-BSL - 02/05/2021 - Replacing constant assignment by Global Variable.
         If (FCustIcon >= 306) and (FCustIcon <= 308) then // 83
           NodeData.Index := 131
         Else
           NodeData.Index := 36;
        End;

      If Mailing = EmptyStr then
        Mailing := 'Mailing Address NOT SPECIFIED';

      NodeData.Caption := NodeData.Caption + #10 + Mailing;
      MyNode1          := Treeview1.AddChild(XNode);
      NodeData         := Treeview1.GetNodeData(mynode1);
      NodeData.Caption := 'Click for full Details';
      NodeData.Index   := 2;
    End; // Else
End;

Procedure TFRM_Tree.ShowSuperCustomerOnly(XNode:PVirtualNode; In_Id_Is_Cust, Full, CustId, CustName, CustAddress, PremiseCount, CustomerDeceased, IsProspect: String;
                                                 Debt: String);
//////////////////////////////////////////////////////////////////////////////////////////
// Just Display Top Level Customer Details Only, from Search Query                      //
//////////////////////////////////////////////////////////////////////////////////////////
Var
 MyRecPtr       : PMyRec;
 DebtMessage,
 Fad,
 SuperCustomerId: String;
 f              : Integer;
begin
  DebtMessage    := EmptyStr;
  Inc(CustCount);

  NodeData       := Treeview1.GetNodeData(XNode);
  // BSL - 22/06/2021 - CRM-511 - Control over Super Customer.
  NodeData.Index := FSuperCustIcon; // FCustIcon;// aCustIcon1; // Customer
  MyNodeCustomer  := XNode;

 //mynodesuperCustomer.selectedindex:=mynodesupercustomer.imageindex;

  if In_Id_Is_Cust <> 'Y' then
    Begin
      // Find Super Customer Details, unless already a supercustomer
      with main_data_module.tempquery do
      Begin
       close;
       deletevariables;
       declarevariable('CUSTID',otlong);
       sql.clear;
       sql.add('select');
       sql.add('sc.customer_id,');
       sql.add('sc.legal_entity_name,');
       sql.add('sc.primary_mailing_address_id,');
       sql.add('scp.premise_line_1,');
       sql.add('scp.premise_line_2,');
       sql.add('scp.premise_line_3,');
       sql.add('scp.premise_line_4,');
       sql.add('scp.premise_line_5,');
       sql.add('scp.premise_line_6,');
       sql.add('scp.premise_line_7,');
       sql.add('scp.premise_line_8,');
       sql.add('scp.premise_line_9,');
       sql.add('scp.premise_postcode from ');
       sql.add('crm.customer sc,');
       sql.add('crm.premises scp,');
       sql.add('crm.customer_to_super_customer csc');
       sql.add('where');
       sql.add('sc.primary_mailing_address_id=scp.premise_id');
       sql.add('and');
       sql.add('sc.customer_id=csc.super_customer_id and csc.customer_id=:CUSTID') ;
       setvariable('CUSTID',custid);
       open;
       deletevariables;
      end;
    end
  else
    begin
      with main_data_module.tempquery do
      Begin
       close;
       deletevariables;
       declarevariable('CUSTID',otlong);
       sql.clear;
       sql.add('select');
       sql.add('sc.customer_id,');
       sql.add('sc.legal_entity_name,');
       sql.add('sc.primary_mailing_address_id,');
       sql.add('scp.premise_line_1,');
       sql.add('scp.premise_line_2,');
       sql.add('scp.premise_line_3,');
       sql.add('scp.premise_line_4,');
       sql.add('scp.premise_line_5,');
       sql.add('scp.premise_line_6,');
       sql.add('scp.premise_line_7,');
       sql.add('scp.premise_line_8,');
       sql.add('scp.premise_line_9,');
       sql.add('scp.premise_postcode from ');
       sql.add('crm.customer sc,');
       sql.add('crm.premises scp');
       sql.add('where');
       sql.add('sc.primary_mailing_address_id=scp.premise_id');
       sql.add('and');
       sql.add('sc.customer_id=:CUSTID') ;
       setvariable('CUSTID',custid);
       open;
       deletevariables
      end;
    end;

  if main_data_module.tempquery.recordcount <> 0 then
    Begin
      SuperCustomerId := main_data_module.tempquery.fields[0].text;
      CustName        := main_data_module.tempquery.fields[1].text;
      Mailing         := main_data_module.tempquery.fields[2].text;
      Fad             := EmptyStr;

      if main_data_module.tempquery.fields[3].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[3].text+', ';
      if main_data_module.tempquery.fields[4].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[4].text+', ';
      if main_data_module.tempquery.fields[5].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[5].text+', ';
      if main_data_module.tempquery.fields[6].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[6].text+', ';
      if main_data_module.tempquery.fields[7].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[7].text+', ';
      if main_data_module.tempquery.fields[8].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[8].text+', ';
      if main_data_module.tempquery.fields[9].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[9].text+', ';
      if main_data_module.tempquery.fields[10].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[10].text+', ';
      if main_data_module.tempquery.fields[11].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[11].text+', ';

      Fad     := Fad + main_data_module.tempquery.fields[12].text + EmptyStr;
      Mailing := Fad;
    end;

  Password := EmptyStr;

 //New(MyRecPtr);
 NodeData.D_Customer_ID   := SuperCustomerId;
 NodeData.D_LIB_CUST_ID   := EmptyStr;
 NodeData.D_Customer_Name := Custname;
 NodeData.D_PremiseCount  := PremiseCount;
 NodeData.D_CDEBT         := DebtMessage;
  //mynodesupercustomer.data:=MyRecPtr;

 NodeData.Caption := 'Super Customer ' + SuperCustomerId + ' - ' + CustName;

 if Mailing = EmptyStr then
   Mailing := 'Mailing Address NOT SPECIFIED';

 NodeData.Caption := NodeData.Caption + #10 + Mailing;

 if Full = 'N' then
   exit;

 with FRM_main_search.customercontacts do
 Begin
  close;
  setvariable('CUSTID',supercustomerid);
  open;
 end;

 mynodecustomer:=xnode;
 BuildCustomerNotes(supercustomerid);

 // Check for Statement Reviewr
 with main_data_module.generalquery do
 Begin
  close;
  DeleteVariables;
  DeclareVariable('CID', otstring);
  sql.clear;
  sql.add('Select * from crm.customer_statement_reviewer where customer_id=:CID');
  setvariable('CID',supercustomerid);
  open;
  deletevariables;
  if recordcount<>0 then
  Begin
   {mynode1:=Treeview1.AddChild(mynodeCustomer,'Statement Reviewer Requested by - '+fields[1].text+' on '+fields[2].text);
   mynode1.imageindex:=136;
   mynode1.selectedindex:=136;
   mynode1.font.color:=clpurple;
   mynode1.font.style:=[fsbold];}

   mynode1:=Treeview1.AddChild(mynodeCustomer);
   nodeData := Treeview1.GetNodeData(mynode1);
   NodeData.caption := 'Statement Reviewer Requested by - '+fields[1].text+' on '+fields[2].text;
   NodeData.index:=136;
   NodeData.fontcolor:=clpurple;
   NodeData.fontBold:=true;

   C1.caption:='Remove Statement Reviewer';
  End
  else C1.caption:='Add Statement Reviewer';
 End;

 //changed by maryam on 05/05/2016 for HH requested by Rosie & Martin
 with main_data_module.generalquery do
 Begin
  Close;
  DeleteVariables;
  DeclareVariable(':RESULTS', otcursor);
  sql.clear;
  sql.Add('begin');
  sql.Add('CRM.PK_UTILITIES.PR_RET_REFUSED_HH_DATA('''+supercustomerid+''',:RESULTS);');
  sql.Add('end;');
  open;
  deletevariables;
  if (recordcount<>0) then
  Begin
   mynode1:=Treeview1.AddChild(mynodeCustomer);
   nodeData := Treeview1.GetNodeData(mynode1);
   NodeData.caption := fields[0].Text;
   NodeData.index:=92;
   NodeData.fontcolor:=clRed;
   NodeData.fontBold:=true;
  End;
 End;

 if specialneeds='T' then
 Begin
  mynode1:=Treeview1.AddChild(mynodeCustomer);
  nodeData := Treeview1.GetNodeData(mynode1);
  NodeData.caption := 'Special Needs - '+SpecialNeeds;
  NodeData.index:=15;
  NodeData.fontcolor:=clpurple;
  NodeData.fontBold:=true;
 end;


 if Password<>'' then
 Begin
  mynode1:=Treeview1.AddChild(mynodeCustomer);
  nodeData := Treeview1.GetNodeData(mynode1);
  NodeData.caption := 'Password-      '+password+'      Effective From ('+PasswordDate+')';
  NodeData.index:=7;
  NodeData.fontcolor:=clpurple;
  NodeData.fontBold:=true;
 end;

 no_of_holders:=FRM_main_search.customercontacts.recordcount;

 prevah:='-1';

 mynodecustomer:=mynodesupercustomer;

 with FRM_main_search.customercontacts do
 Begin
  for f:=1 to no_of_holders do
   Begin
    AH_ID:=fields[16].text;
    AH_CONTACT_TITLE_ID:=fields[23].text;
    AH_INITIALS:=fields[25].text;
    AH_SURNAME:=fields[26].text;
    AH_FORENAME:=fields[27].text;
    AH_DISPLAY_NAME:=fields[28].text;
    AH_ADDITIONAL_INFORMATION:=fields[29].text;
    AH_SPECIAL_NEEDS_INFORMATION:=fields[30].text;
    AH_TELEPHONE_NO_DAY:=fields[31].text;
    AH_TELEPHONE_NO_EVE:=fields[32].text;
    AH_TELEPHONE_NO_MOBILE:=fields[33].text;
    AH_EMAIL:=fields[34].text;
    AH_FAX:=fields[35].text;
    AH_order:=fields[20].text;
    AH_DOB:=fields[48].text;
    ah_TYPE:='Contact';
    if fields[19].text='P' then AH_TYPE:='Primary';
    if fields[19].text='E' then AH_TYPE:='Emergency';

    if (ah_id<>'') and (ah_id<>prevah) then
    Begin
     if (f=3) and (no_of_holders>3) then
     Begin

      mynode1:=Treeview1.AddChild(mynodeCustomer);
      nodeData := Treeview1.GetNodeData(mynode1);
      NodeData.caption := 'Additional Account Holders ['+inttostr(no_of_holders-2)+']';
      NodeData.index:=74;
     end;

   Contdet:='';
   if AH_CONTACT_TITLE_ID<>'' then
   Begin
     Contdet:=ah_contact_title_id+' ';
   end;

   if AH_INITIALS<>'' then
   Begin
    contdet:=contdet+'('+ah_initials+') ';
   end;

   if AH_FORENAME<>'' then
   Begin
    contdet:=contdet+ah_Forename+' ';
   end;

   if AH_SURNAME<>'' then
   Begin
    contdet:=contdet+ah_surname+' ';
   end;


   if AH_DISPLAY_NAME<>'' then
   Begin
    Contdet:=contdet+'. Known as - '+ah_display_name;
    contdet:=ah_display_name;
   end;
   if contdet<>'' then
   Begin
    // First two account holders hang of customer
    if (f<3) then
    Begin
     premiseContactNode:=Treeview1.AddChild(mynodeCustomer);
     nodeData := Treeview1.GetNodeData(premiseContactNode);
     NodeData.caption := 'Account Holder - '+ah_type+' - '+contdet;
    end
    else
    // 3rd account holder will hang of tree if only 3 account holders
    if (f=3) and (no_of_holders=3) then
    Begin
     premiseContactNode:=Treeview1.AddChild(mynodeCustomer);
     nodeData := Treeview1.GetNodeData(premiseContactNode);
     NodeData.caption := 'Account Holder - '+ah_type+' - '+contdet;

    End
    else
    // more than 3 account holders get rolled up
    Begin
     premiseContactNode:=Treeview1.AddChild(mynode1);
     nodeData := Treeview1.GetNodeData(premiseContactNode);
     NodeData.caption := 'Account Holder - '+ah_type+' - '+contdet;
    end;

    // Create Account Holder Node
    // Check for deceased accountholder
    if fields[24].value <> null then
    nodedata.index:=fields[24].value
    else  nodedata.index := 0;

    if fields[46].text<>'1' then
    Begin
    end;
    if fields[46].text='3' then
    Begin
     nodedata.index:=43;
    end;
   // premiseContactNode.selectedindex:=premisecontactnode.imageindex;

    nodedata.D_customer_id :=FRM_main_search.customercontacts.fields[0].text;
    nodedata.D_Account_holder_id :=FRM_main_search.customercontacts.fields[16].text;
    nodedata.D_contact_id :=FRM_main_search.customercontacts.fields[18].text;
   // premiseContactNode.data:=MyRecPtr;
   end;

   if AH_ADDITIONAL_INFORMATION<>'' then
   Begin
    {premiseContactitemNode:=Treeview1.items.AddChild(premisecontactnode,ah_additional_information);
    premiseContactitemNode.imageindex:=73;
    premiseContactitemNode.selectedindex:=73;}

    premiseContactItemNode:=Treeview1.AddChild(premisecontactnode);
    nodeData := Treeview1.GetNodeData(premiseContactItemNode);
    NodeData.caption := ah_additional_information;
    NodeData.index:=73;

   end;

   if AH_SPECIAL_NEEDS_INFORMATION<>'' then
   Begin
   { premiseContactitemNode:=Treeview1.items.AddChild(premisecontactnode,'Special Needs - '+ah_special_needs_information);
    premiseContactitemNode.imageindex:=72;
    premiseContactitemNode.selectedindex:=72; }

    premiseContactItemNode:=Treeview1.AddChild(premisecontactnode);
    nodeData := Treeview1.GetNodeData(premiseContactItemNode);
    NodeData.caption := 'Special Needs - '+ah_special_needs_information;
    NodeData.index:=72;
   end;

   Telno:='';
   if AH_TELEPHONE_NO_DAY<>'' then
   Begin
    telno:=telno+'Tel No: Day - '+ah_Telephone_no_Day+'     ';
   end;

   if AH_TELEPHONE_NO_EVE<>'' then
   Begin
    telno:=telno+'Tel No: Eve - '+ah_Telephone_no_eve;
   end;

   if telno<>'' then
   Begin
    {premiseContactitemNode:=Treeview1.items.AddChild(premisecontactnode,telno);
    premiseContactitemNode.imageindex:=2;
    premiseContactitemNode.selectedindex:=2;}

    premiseContactItemNode:=Treeview1.AddChild(premisecontactnode);
    nodeData := Treeview1.GetNodeData(premiseContactItemNode);
    NodeData.caption := telno;
    NodeData.index:=2;
   end;

   if AH_TELEPHONE_NO_MOBILE<>'' then
   Begin
    {premiseContactitemNode:=Treeview1.items.AddChild(premisecontactnode,'Tel No: Mobile - '+ah_Telephone_no_Mobile);
    premiseContactitemNode.imageindex:=14;
    premiseContactitemNode.selectedindex:=14;}

    premiseContactItemNode:=Treeview1.AddChild(premisecontactnode);
    nodeData := Treeview1.GetNodeData(premiseContactItemNode);
    NodeData.caption := 'Tel No: Mobile - '+ah_Telephone_no_Mobile;
    Nodedata.d_tel:=ah_Telephone_no_Mobile;
    nodedata.D_customer_id :=FRM_main_search.customercontacts.fields[0].text;
    NodeData.index:=14;
   end;

   if AH_FAX<>'' then
   Begin
    {premiseContactitemNode:=Treeview1.items.AddChild(premisecontactnode,'Fax - '+ah_Fax);
    premiseContactitemNode.imageindex:=11;
    premiseContactitemNode.selectedindex:=11; }

    premiseContactItemNode:=Treeview1.AddChild(premisecontactnode);
    nodeData := Treeview1.GetNodeData(premiseContactItemNode);
    NodeData.caption := 'Fax - '+ah_Fax;
    NodeData.index:=11;

   end;

   if AH_EMAIL<>'' then
   Begin
    {premiseContactitemNode:=Treeview1.items.AddChild(premisecontactnode,'Email - '+ah_email);
    premiseContactitemNode.imageindex:=20;
    premiseContactitemNode.selectedindex:=20;
    premiseContactitemNode.font.color:=clblue;
    premiseContactitemNode.font.style:=[fsunderline]; }

    premiseContactItemNode:=Treeview1.AddChild(premisecontactnode);
    nodeData := Treeview1.GetNodeData(premiseContactItemNode);
    NodeData.caption := 'Email - '+ah_email;
    Nodedata.d_email:=ah_email;
    Nodedata.D_Customer_Id:=FRM_main_search.customercontacts.fields[0].text;
    NodeData.index:=20;
    NodeData.fontcolor:=clblue;
    NodeData.fontunderline:=true;

   end;

   if AH_DOB>'' then
   Begin
    premiseContactItemNode:=Treeview1.AddChild(premisecontactnode);
    nodeData := Treeview1.GetNodeData(premiseContactItemNode);
    NodeData.caption := 'Date Of Birth - '+ah_DOB;
    NodeData.index:=63;
   end;

   if AH_contact_method<>''  then
   begin
    premiseContactItemNode:=Treeview1.AddChild(premisecontactnode);
    nodeData := Treeview1.GetNodeData(premiseContactItemNode);
    NodeData.caption := 'Preferred Contact: - '+uppercase(AH_contact_method);
    NodeData.index:=271;
   end;

   end;
   prevah:=ah;
  next;
  end;
 end;

end;


Procedure TFRM_Tree.BuildpremiseNode(premisename,premiseid,premisetype:string);
//Var
//MyRecPtr: PMyRec;
Begin
 xnode:=treeview1.FocusedNode;
 nodeData := treeview1.GetNodeData(xnode);
 nodedata.D_premise_NAME := premisename;
 nodedata.D_premise_ID := premiseid;
 //mynodepremise.data:=MyRecPtr;
 inc(premisecount);
 // Format premise Address
 nodedata.index:=32; // premise
 if premisetype='01' then  Treedata.index:=49;
 if premisetype='02' then  Treedata.index:=32;
 if premisetype='03' then  Treedata.index:=50;
 if premisetype='04' then  Treedata.index:=35;
 if premisetype='05' then  Treedata.index:=35;
 //mynodepremise.selectedindex:=mynodepremise.imageindex;
// premiseContactNode:=Treeview1.items.AddChild(mynodepremise,'Click for Details');
 premiseContactNode:=Treeview1.Addchild(xnode);
 nodeData := Treeview1.GetNodeData(premiseContactNode);
 NodeData.caption := 'Click for Details'
end;

procedure TFRM_Tree.BuildMPANNode(MPAN,regstatus,energisation_status,new_connection:string);
Begin

 xnode:=treeview1.FocusedNode;
 nodeData := treeview1.GetNodeData(xnode);
 inc(mpancount);
 nodedata.index:=5;
 nodedata.fontBold:=true;
 // MPANNode.color:=rgb(251,248,203);
 if (regstatus='PENDING') then
 Begin
  nodedata.index:=4;
  nodedata.fontcolor:=clpurple;
 end;
 if (regstatus='IN PROGRESS') then
 Begin
  nodedata.index:=4;
  nodedata.fontcolor:=clolive;
 end;
 if (regstatus='DISCONNECTED') then
 Begin
  Treedata.index:=4;
  Treedata.fontcolor:=clblue;
 end;
 if (regstatus='REGISTERED') then Treedata.fontcolor:=clgreen;
 if (regstatus='LOSS PENDING') then Treedata.fontcolor:=clpurple;
 if (regstatus='FUTURE LOSS') or
    (regstatus='LOST') or
    (regstatus='REJECTED') or
    (regstatus='OBJECTED') then
 begin
  Treedata.index:=4;
  Treedata.fontcolor:=clred;
 end;
 if energisation_status='D' then
 Begin
  Treedata.index:=4;
 end
 else
 Begin
 end;
 if new_connection='T' then
 Begin
  Treedata.index:=19;
 end;
 mynode1:=Treeview1.AddChild(mpannode);
end;

procedure TFRM_Tree.BuildElectricMeterNode(Xnode:Pvirtualnode);
var
mpan,ENDDATE,TPR,desc,metertype,enstatus,ssc,sscdesc,daterem,nsr,ppmip,maketype:string;
//MyRecPtr: PMyRec;
Begin
   // Check for Meter Technical Details
 nodeData := treeview1.GetNodeData(xnode);
 mpan:=nodedata.D_Span;
 enddate:=nodedata.D_SpanEnd;
 if ENDDATE='' then ENDDATE:='10/10/2060';
 mpannode:=xnode;
 fPremiseDcc:= false;

  // Check For Last Known PPMIP
  ppmip:=frm_common.getlastppmip(mpan);
  if ppmip<>'' then
  Begin
   {MeterConfigNode:=Treeview1.items.AddChild(mpannode,'Last Known Legacy PPMIP '+ppmip);
   MeterConfigNode.imageindex:=217;
   MeterConfigNode.selectedindex:=217; }

   MeterConfigNode:=Treeview1.Addchild(xnode);
   nodeData := Treeview1.GetNodeData(MeterConfigNode);
   NodeData.caption := 'Last Known Legacy PPMIP '+ppmip;
   nodedata.index:=217;

  end;

  // Check if Customer Has Requested Single Rate Billing
  m_single.caption:='Default to Single Rate Billing';
  with main_data_module.generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('MPAN', otString);
   sql.clear;
   sql.add('select * from crm.mpans_single_rate_billing');
   sql.add('where mpancore=:mpan');
   sql.add('Order by effective_from desc');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
  End;
  if main_data_module.generalquery.recordcount<>0 then
  Begin
   m_single.caption:='Remove Single Rate Billing';
   efsdmsmtd:=main_data_module.generalquery.fields[1].text;
   {MeterConfigNode:=Treeview1.items.AddChild(mpannode,'Customer Requests Single Rate Billing from '+efsdmsmtd);
   MeterConfigNode.font.color:=clpurple;
   MeterConfigNode.font.style:=[fsbold];
   MeterConfigNode.imageindex:=140;
   MeterConfigNode.selectedindex:=140; }

   MeterConfigNode:=Treeview1.Addchild(xnode);
   nodeData := Treeview1.GetNodeData(MeterConfigNode);
   NodeData.caption := 'Customer Requests Single Rate Billing from '+efsdmsmtd;
   nodedata.index:=140;
   nodedata.fontcolor:=clpurple;
   nodedata.fontBold:=true;

  End;

    // Check if Span has SSC change
  with main_data_module.generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('MPAN', otString);
   sql.clear;
   sql.add('select s.*,t.description from crm.span_ssc_changes s,crm.span_type t');
   sql.add('where s.span=:MPAN');
   sql.add('and s.span_type=t.span_type_id');
   sql.add('Order by s.effective_from desc');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
  End;
  while not main_data_module.generalquery.eof do
  Begin
   efsdmsmtd:=main_data_module.generalquery.fields[2].text;
   {MeterConfigNode:=Treeview1.items.AddChild(mpannode,efsdmsmtd+' Billing changed to '+main_data_module.generalquery.fields[3].text);
   MeterConfigNode.font.color:=clpurple;
   MeterConfigNode.font.style:=[fsbold];
   MeterConfigNode.imageindex:=26;
   MeterConfigNode.selectedindex:=26; }

   MeterConfigNode:=Treeview1.Addchild(xnode);
   nodeData := Treeview1.GetNodeData(MeterConfigNode);
   NodeData.caption := efsdmsmtd+' Billing changed to '+main_data_module.generalquery.fields[3].text;
   nodedata.index:=26;
   nodedata.fontcolor:=clpurple;
   Nodedata.fontBold:=true;

   main_data_module.generalquery.next;
  End;

    // Check for Billing Profile Change
  with main_data_module.generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('MPAN', otString);
   sql.clear;
   sql.add('select * from billing.mpan_bill_profile');
   sql.add('where mpancore=:MPAN');
   sql.add('Order by effective_from desc');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
  End;
  while not main_data_module.generalquery.eof do
  Begin
   efsdmsmtd:=main_data_module.generalquery.fields[2].text;
   {pcConfigNode:=Treeview1.items.AddChild(mpannode,efsdmsmtd+' -  Billing Profile Set to 0'+main_data_module.generalquery.fields[1].text);
   pcConfigNode.font.color:=clpurple;
   pcConfigNode.font.style:=[fsbold];
   pcConfigNode.imageindex:=44;
   pcConfigNode.selectedindex:=44;}

   pcConfigNode:=Treeview1.Addchild(xnode);
   nodeData := Treeview1.GetNodeData(pcConfigNode);
   NodeData.caption := efsdmsmtd+' -  Billing Profile Set to 0'+main_data_module.generalquery.fields[1].text;
   nodedata.index:=217;
   Nodedata.fontcolor:=clpurple;
   Nodedata.fontBold:=true;

   main_data_module.generalquery.next;
  End;


  with mtds do
  begin
   close;
   setvariable('MPAN',MPAN);
   open;
  end;

  // Only Do This Block If Meter Records Exist
  if MTDs.recordcount<>0 then
  Begin
   msid:='LEEOK';
   oldefsdmsmtd:='lee';
   oldmeterid:='';
   oldregister:='';
   while not MTDs.eof do
   Begin
    // Build Tree Of MTDS
   // Create Subtree of Effective From Dates
    efsdmsmtd := FormatDateTime('dd/mm/yyyy',MTDs.fields[1].AsDateTime);
    MeterType:=MTDS.fields[23].text;
    EnStatus:=mtds.fields[2].text;
    SSC:=mtds.fields[5].text;
    SSCDesc:=mtds.fields[6].text;
    DateRem:=mtds.fields[37].text;
    NSR:=mtds.fields[36].text;
    Mregister:=mtds.fields[26].text;
    meterid:=MTDs.fields[10].text;
    maketype:='';
    if mtds.fields[14].text<>'' then maketype:=copy(MTDs.fields[14].text,1,3);
    if efsdmsmtd<>oldefsdmsmtd then
    Begin
     if oldefsdmsmtd='lee' then config:='Current Configuration'
     else config:='Previous Configuration';

     if ((config='Previous Configuration') and (historic1.checked=true)) or (config<>'Previous Configuration') then
     Begin
      if (efsdmsmtd<>'') and (MeterType='') and (SSC='') then
      Begin
      { MeterConfigNode:=Treeview1.items.AddChild(mpannode,config+' - '+efsdmsmtd+' - MOP Reports No Meters on this Supply. Energisation Status = '+EnStatus);
       MeterConfigNode.font.color:=clred;
       MeterConfigNode.font.style:=[fsbold];
       MeterConfigNode.imageindex:=24;
       MeterConfigNode.selectedindex:=24; }

       MeterConfigNode:=Treeview1.Addchild(xnode);
       nodeData := Treeview1.GetNodeData(MeterConfigNode);
       NodeData.caption := config+' - '+efsdmsmtd+' - MOP Reports No Meters on this Supply. Energisation Status = '+EnStatus;
       nodedata.index:=24;
       Nodedata.fontcolor:=clred;
       Nodedata.fontBold:=true;
      end
      else
      if MeterType='' then
      Begin
       {MeterConfigNode:=Treeview1.items.AddChild(mpannode,'Metering Configuration not Known (Missing / Incomplete meter technical details)');
       MeterConfigNode.font.color:=clred;
       MeterConfigNode.font.style:=[fsbold];
       MeterConfigNode.imageindex:=26;
       MeterConfigNode.selectedindex:=26; }

       MeterConfigNode:=Treeview1.Addchild(xnode);
       nodeData := Treeview1.GetNodeData(MeterConfigNode);
       NodeData.caption := 'Metering Configuration not Known (Missing / Incomplete meter technical details)';
       nodedata.index:=26;
       Nodedata.fontcolor:=clred;
       Nodedata.fontBold:=true;

      end;
      if (SSC<>'') then
      Begin
       desc:=config+' - '+efsdmsmtd+' - SSC ID ('+SSC+') - '+SSCDESC;
       if (MTDs.fields[4].text='') and (MTDs.fields[39].text='') then
       begin
        desc:=desc+#10+'(* WARNING: '+MTDs.fields[13].text+' *)';
       end;
       MeterConfigNode:=Treeview1.AddChild(xnode);
       nodeData := Treeview1.GetNodeData(MeterConfigNode);
       nodedata.caption:=desc;

       if config<>'Previous Configuration' then
       Begin
        nodedata.fontcolor:=clgreen;
        nodedata.fontBold:=true;
       end;
       Nodedata.index:=27;
       oldmeterid:='lee';
      end;
     end;
    end; // End Of Configuration Date
     // Do Meters

    if ((config='Previous Configuration') and (historic1.checked=true)) or (config<>'Previous Configuration') then
    Begin
     if meterid<>oldmeterid then
     Begin
      oldregister:='Lee';
      mtype:='';
      if MeterType='' then mtype:='*NO*';
      if (efsdmsmtd<>'') and (metertype='') then mtype:='*NO*';

      if mtype<>'*NO*' then
      Begin
       {if DateRem<>'' then Dateremoved:='    (Date Removed='+DateRem+')'
       else}
       dateremoved:='';
       if (NSR='Y') and (v_non.checked=false) then
       Begin
       //
       end
       else
       Begin
       { MeterNode:=Treeview1.items.AddChild(MeterConfigNode,'NHH Meter ID-'+Meterid+dateremoved);
        MeterNode.imageindex:=17; // NHH Credit Meter
        MeterNode.selectedindex:=17;}

        MeterNode:=Treeview1.Addchild(meterconfignode);
        nodeData := Treeview1.GetNodeData(MeterNode);
        nodedata.caption:='NHH Meter ID-'+Meterid+dateremoved;
        nodedata.index:=17;
        nodedata.D_SPAN :=mpan;
        nodedata.M_METERID :=Meterid;
        nodedata.M_SERVICE :='0';

        nodedata.Metertype := Metertype;
         // SMI check for DCC Meter
         nodedata.D_SpanDCC_E := false;
         nodedata.D_SpanE := '';

         with Generalquery do
         Begin
          close;
          DeleteVariables;
          DeclareVariable('MPXN', otstring);
          sql.clear;
          sql.add('Select ods.dcc_enrolled(:MPXN)from dual');
          setvariable('MPXN', mpan);
          open;
          deletevariables;

          if generalquery.Fields[0].Text = 'Y' then
          begin
            nodedata.D_SpanDCC_E := true;
            nodedata.D_SpanE := mpan;
            fPremiseDcc:= true;
          end;
         End;

        if MeterType='N' then
        Begin
         nodedata.caption:='NHH Credit Meter ID-'+MeterID+dateremoved;
        end;

        if MeterType='S' then
        Begin
         nodedata.caption:='NHH Smart Card Meter ID-'+MeterID+dateremoved;
         nodedata.index:=22; // NHH Smart Card meter
        end;

        if (MeterType='S') and (mtds.fields[20].text='R') then
        Begin
         Nodedata.caption:='Remote Read Smart Meter ID-'+MeterID+dateremoved;
         nodedata.index:=205; // NHH Smart Card meter
        end;
         if (copy(MeterType,1,4)='RCAM') or (MakeType='PRI') or (copy(MeterType,1,3)='NSS') then
        Begin
         nodedata.caption:='Smart Meter ID-'+MeterID+dateremoved;
         nodedata.index:=314;
        end;

        if MeterType = 'S1EA' then
        begin
          nodedata.caption:='SMETS1 E&A Smart Meter ID-' + MeterID + dateremoved;
          nodedata.index := 313;
        end
        else
        if (copy(MeterType,1,2)='S1') then
        Begin
          ShowSmetsMeterCommsSupplier(MeterNode,MPAN,METERID,'0',dateremoved,'X');
        end;

        if (copy(MeterType,1,2)='S2')  then
        Begin
         nodedata.caption:='SMETS 2 Meter ID-'+MeterID+dateremoved;
         nodedata.index:=205; // SMETS 2 ICON
        end;


        if MeterType='T' then
        Begin
         nodedata.caption:='NHH Token Meter ID-'+MeterID+dateremoved;
         nodedata.index:=23; // NHH token Meter
        end;
        if MeterType='K' then
        Begin
         nodedata.caption:='NHH Key Meter ID-'+MeterID+dateremoved;
         nodedata.index:=21; // NHH key Meter
        end;
        if MeterType='H' then
        Begin
         nodedata.caption:='HH Meter ID-'+MeterID+dateremoved;
         nodedata.index:=9; // HH Meter
        end;
        if dateremoved<>'' then
        Begin
         nodedata.index:=24; // Removed Meter
        end;

        if (MTDs.fields[4].text='') and (MTDs.fields[39].text='') then
        begin
         nodedata.caption:=nodedata.caption+#10+'(* WARNING: '+MTDs.fields[13].text+' *)';
        end;

        if MTDS.FieldByName('has_sc_tariff').AsString = 'Y' then
        begin
          nodedata.index:= 321;
        end;

       end;
      end
      else
      Begin
       // A Meter from D0149 but No D0150
       //MeterNode:=Treeview1.items.AddChild(MeterConfigNode,'NHH Meter ID-'+Meterid+' (Missing D0150)');
       MeterNode:=Treeview1.Addchild(MeterConfigNode);
       nodeData := Treeview1.GetNodeData(MeterNode);
       nodedata.caption:='NHH Meter ID-'+Meterid+' (Missing D0150)';
       nodedata.index:=26; // NHH Credit Meter
       nodedata.fontcolor:=clred;
       nodedata.D_SPAN :=mpan;
       nodedata.M_METERID :=Meterid;
       nodedata.M_SERVICE :='0';
      End;

     // Meternode.data:=MyRecPtr;
     end; // Change Of Meter
    end;  // End Of Configuration Block

        // Have any Meters been Removed?
    if (efsdmsmtd<>'') and (MeterType='') and (SSC='') and (daterem<>'')then
    Begin
     {MeterNode:=Treeview1.items.AddChild(MeterConfigNode,'Removed Meter -'+MTDS.fields[39].text+'. Date Removed ('+DateRem+')');
     nodedata.index:=24; // Removed Meter
     MeterNode.selectedindex:=24; }

     MeterNode:=Treeview1.Addchild(MeterConfigNode);
     nodeData := Treeview1.GetNodeData(MeterNode);
     NodeData.caption := 'Removed Meter -'+MTDS.fields[39].text+'. Date Removed ('+DateRem+')';
     nodedata.index:=24;
     nodedata.D_SPAN :=mpan;
     nodedata.M_METERID :=Meterid;
     nodedata.M_SERVICE :='0';
    end;


     if ((config='Previous Configuration') and (historic1.checked=true)) or (config<>'Previous Configuration') then
    Begin
     if mregister<>oldregister then
     Begin
      if (NSR='Y') and (v_non.checked=false) then
      begin
      //
      end
      else
      Begin
       TPR:=MTDs.fields[33].text;  // TPR ID e.g. 00001
       if tpr<>'' then
       Begin
        //MeterRegisterNode:=Treeview1.items.AddChild(MeterNode,mregister+' - TPR '+TPR+' ('+MTDs.fields[38].text+') - '+MTDs.fields[30].text);

        MeterRegisterNode:=Treeview1.Addchild(MeterNode);
        nodeData := Treeview1.GetNodeData(MeterRegisterNode);
        NodeData.caption := mregister+' - TPR '+TPR+' ('+MTDs.fields[38].text+') - '+MTDs.fields[30].text;
        nodedata.fontcolor:=clblack;
        if mtds.fields[30].text='' then nodedata.fontcolor:=clred;
       end
       else
       Begin
        if MeterID<>'' then
        Begin
        // MeterRegisterNode:=Treeview1.items.AddChild(MeterNode,mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * ) - '+MTDs.fields[30].text);
         MeterRegisterNode:=Treeview1.Addchild(MeterNode);
         nodeData := Treeview1.GetNodeData(MeterRegisterNode);
         NodeData.caption := mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * ) - '+MTDs.fields[30].text;
         nodedata.fontcolor:=clred;
        end;
       end;
       nodedata.Index:=Frm_Common.GetRegisterPic(MTDS.fields[33].text);

       nodedata.D_SPAN :=mpan;
       nodedata.M_EFSDMSMTD :=efsdmsmtd;
       nodedata.M_METERID :=Meterid;
       nodedata.M_REGISTERID :=mregister;
       if MeterType='H' then nodedata.M_HH_REGISTER :='H'
       else nodedata.M_HH_REGISTER :='N';
     //  MeterRegisterNodedata:=MyRecPtr;

       if MTDs.fields[29].text='RI' then
       Begin
        nodedata.index:=44;
       end;
       if NSR='Y' then
       Begin
        // non settlement register
        nodedata.Fontcolor:=clred;
        if v_non.checked then
        Begin
         nodedata.Index:=Frm_Common.GetNonRegisterPic(MTDS.fields[33].text);

        if MTDs.fields[29].text='RI' then
         Begin
          nodedata.index:=45;
         end;
        end;
       end;
      end;
     end; // End Of Add Register
    end;  // End Of Configuration Block
    oldefsdmsmtd:=efsdmsmtd;
    oldmeterid:=meterid;
    oldregister:=mregister;
    mtds.next;
   end;
  end; // End Of Meter Strucutre Tree

 ////////// Build Tree Of Orphaned Registers //////////////
 With generalquery do
  Begin
   Close;
   DeleteVariables;
   DeclareVariable('MPAN', otString);
   sql.clear;
   sql.add('select distinct R.mpancore,R.meterid,R.registerid,r.current_status');
   sql.add('from edmgr.readings R,edmgr.d0149a D, edmgr.d0150_293 M');
   sql.add('where');
   sql.add('r.MPANCORE=:MPAN');
   sql.add('and r.mpancore=d.mpancore (+)');
   sql.add('and');
   sql.add('r.meterid=d.meterid (+)');
   sql.add('and');
   sql.add('r.registerid=d.registerid (+)');
   sql.add('and');
   sql.add('r.mpancore=M.mpancore (+)');
   sql.add('and');
   sql.add('r.meterid=M.meterid (+)');
   sql.add('and');
   sql.add('r.registerid=M.meter_register_id (+)');
   sql.add('and d.mpancore is null');
   sql.add('and m.mpancore is null');
   sql.add('and r.current_status<>''D''');
  // sql.add('and r.rdngtype<>''W''');
   sql.add('order by R.meterid,R.registerid');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
  end;
   // Only Do This Block If Meter Records Exist
  if generalquery.recordcount<>0 then
  Begin
   oldmeterid:='OldMeter';
   msid:='LEEOK';
   oldefsdmsmtd:='lee';

   {MeterConfigNode:=Treeview1.items.AddChild(mpannode,'Register Readings (Orphans - No Mapping Details)');
   MeterConfigNode.font.color:=clred;
   MeterConfigNode.font.style:=[fsbold];
   MeterConfigNode.imageindex:=26;
   MeterConfigNode.selectedindex:=26; }

   MeterConfigNode:=Treeview1.Addchild(xnode);
   nodeData := Treeview1.GetNodeData(MeterConfigNode);
   NodeData.caption := 'Register Readings (Orphans - No Mapping Details)';
   nodedata.index:=26;
   nodedata.fontcolor:=clred;
   nodedata.fontBold:=true;

   while not generalquery.eof do
   Begin
     // Do Meters
    Begin
     meterid:=generalquery.fields[1].text;
     if meterid<>oldmeterid then
     Begin
      oldregister:='Lee';
      mtype:='';
       Begin
        {MeterNode:=Treeview1.items.AddChild(MeterConfigNode,'NHH Meter ID-'+Meterid+dateremoved);
        nodedata.index:=26; // NHH Meter
        MeterNode.selectedindex:=26; }

        MeterNode:=Treeview1.Addchild(meterconfignode);
        nodeData := Treeview1.GetNodeData(MeterNode);
        NodeData.caption := 'NHH Meter ID-'+Meterid+dateremoved;
        nodedata.index:=26;
        nodedata.D_SPAN :=mpan;
        nodedata.M_METERID :=Meterid;
        nodedata.M_SERVICE :='0';

       end;
      end; // Change Of Meter
    end;  // End Of Configuration Block
    Begin
     Mregister:=generalquery.fields[2].text;
     if mregister<>oldregister then
     Begin
      Begin
       Begin
        //MeterRegisterNode:=Treeview1.items.AddChild(MeterNode,mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * )');
        MeterRegisterNode:=Treeview1.Addchild(MeterNode);
        nodeData := Treeview1.GetNodeData(MeterRegisterNode);
        NodeData.caption := mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * )';
        nodedata.fontcolor:=clred;
       end;
        nodedata.index:=29;

       nodedata.D_SPAN :=mpan;
       nodedata.M_EFSDMSMTD :=efsdmsmtd;
       nodedata.M_METERID :=MeterID;
       nodedata.M_REGISTERID :=mregister;
       //MeterRegisterNode.data:=MyRecPtr;

      end;
     end; // End Of Add Register
    end;  // End Of Configuration Block
    oldmeterid:=meterid;
    oldregister:=mregister;
    generalquery.next;
   end;
  end; // End Of Meter Strucutre Tree
end;

procedure TFRM_Tree.Button1Click(Sender: TObject);
begin
 if pagecontrol1.ActivePage=TABSUPP then
 begin
  treeview1.BeginUpdate;
  Treeview1.FullExpand;
  treeview1.endUpdate;
 end;
 if pagecontrol1.ActivePage=TABMOP then
 begin
  moptree.BeginUpdate;
  MopTree.FullExpand;
  moptree.EndUpdate;
 end;
 treeupdating:=false;
end;

procedure TFRM_Tree.treeview1DblClick(Sender: TObject);
var
ag,qu,s,pid,ptype:string;
ts,BillType,agid,faultno :string;
mpancore,efsdmsmtd,meterid,registerid,hh_register,mprn:string;
refno,ps,  priority,pdesc: String;
// BSL - 15/12/2014 - Change JBS from Web to CRM.
Begin
 if fIsExpanding then
   exit;

 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 if assigned(Treedata)=false then exit;

 FRM_Main_Search.pUpdateCustomer(TreeData.D_Customer_Id);

 if Copy(TreeData.caption,1,9)='Agreement' then ViewAggreement1Click(sender)
 else

 if Copy(TreeData.caption,1,23)='Unactioned Pending Loss' then
 Begin
  custid:=Treedata.d_customer_id;
  refno:=Treedata.d_refno;
  priority:=Treedata.d_priority;
  pdesc:=TreeData.d_pdesc;
  if Treedata.d_order=1 then FRM_LOSS_RESULT.losses.first
  else FRM_LOSS_RESULT.losses.last;
  FRM_LOSS_RESULT.tag:=0;
  FRM_LOSS_RESULT.NOTE.lines.clear;
  FRM_LOSS_RESULT.outcomegroup.itemindex:=-1;
  FRM_LOSS_RESULT.showmodal;
  FRM_LOSS_RESULT.close;

  if (FRM_LOSS_RESULT.tag=1) and (treeview1.Selected[xnode]=true) then
  Begin
   treeview1.Expanded[xnode.parent]:=false;
   treeview1.Expanded[xnode.parent]:=true;
  end;
 End;

 if copy(TreeData.caption,1,8)='(FMS ID:' then
 begin
  FAULTNO:=TreeData.C_RECORD_ID;
  ts:=mastersource+'FMS.exe';
  ps:='"'+FRM_Login.edtUsername.Text+'" "'+FRM_Login.edtPassword.Text+'" "'+frm_login.MainSession.LogonDatabase+'" "'+faultno+'"';
  shellexecute(Handle,'open',pchar(ts),pchar(ps),nil,sw_shownormal);
 end
 else

 if (Copy(TreeData.caption,1,7)='Payment') and (TreeData.index <> 300) then
 Begin
  With FRM_Agreement Do
  Begin
   tag:=0;
   // Agreement tab

   //LK17
   custid.text:=frm_common.GetCustomerIdfromAgreementid(TreeData.D_agreement_id);
   custname.text:=frm_common.GetCustomerNameFromId(custid.text);


   Getfields(TreeData.D_Agreement_id);
   Statusbar.panels[0].text:=' Update';
   frm_agreement.pagecontrol1.activepage:=tabsheet3;
   showmodal;
  end;
  if treeview1.Selected[xnode]=true then
  Begin
   treeview1.Expanded[xnode]:=false;
   treeview1.Expanded[xnode]:=true;
  end;
 End
 else
 if (TreeData.index = 300) or (TreeData.index = 301) then
 begin
   HistoricPayPlanClick(sender);
 end
 else
 if (Copy(TreeData.caption,1,14)='Account Review') and Assigned(gFeatureAccessList) and gFeatureAccessList.IsEnabled(FEATURE_ACCESS__ACCOUNTREVIEW) then
 Begin
  With FRM_Agreement Do
  Begin
   AG:=TreeData.D_agreement_id;
   QU:=TreeData.D_Period_id;
   FRM_ACCOUNT_REVIEW.SHOWREVIEW(AG,QU);
   FRM_Account_Review.showmodal;

   if frm_account_review.tag<>0 then
   Begin
    { if treeview1.Selected[xnode.parent]=true then
     Begin
      treeview1.Selected[xnode.parent]:=false;
      treeview1.Selected[xnode.parent]:=true;
     end; }
    if frm_account_review.tag=2 then Messagedlg('DD Changes have now been applied to the account'+#13+
            'Check DD schedule to see changes.',mtinformation,[MBOK],0);
   end;

  end;
 End
 else
 // Check If Email
 if (TreeData.index=20) Then
 Begin
   ShellExecute(0, 'open', PChar('mailto:' + copy(TreeData.caption,9,100) + '?' + 'subject=' + ''), nil, nil, SW_SHOWNORMAL);
 end;

 // Final Bill or Statement
 if (Copy(TreeData.caption,1,14)='* FINAL BILL *')
 or (Copy(TreeData.caption,1,8)='* BILL *') then
 Begin
  if Copy(TreeData.caption,3,1)='B' then BillType:='B' else BillType:='Y';
  pid:=TreeData.D_Period_id;
  ptype:=TreeData.D_Period_type;
  if ptype='M' then frm_reports.ShowSelectedMonthBill(TreeData.D_agreement_id,pid,BillType,'V')
  else frm_reports.ShowSelectedQuarterBill(TreeData.D_agreement_id,pid,BillType,'V');
 End

 else
 if Copy(TreeData.caption,1,13)='Statement For' then
 Begin
   pid:=TreeData.D_Period_id;
   ptype:=TreeData.D_Period_type;
   if ptype='M' then frm_reports.ShowSelectedMonthBill(TreeData.D_agreement_id,pid,'','V')
   else frm_reports.ShowSelectedQuarterBill(TreeData.D_agreement_id,pid,'','V');
 End

 else
 if Copy(TreeData.caption,1,14)='Legacy Quantum' then
 Begin
  Application.CreateForm(TFrm_GAS_QUANTUM, Frm_GAS_QUANTUM);
  try
   MPRN:=TreeData.D_SPAN;
   agid:=TreeData.D_agreement_id;
   frm_gas_quantum.tag:=1;
   frm_gas_quantum.DoQueryMPRN(MPRN,AGID);
   frm_GAS_QUANTUM.showmodal;
  finally
   frm_GAS_QUANTUM.release;
  end;
 End
 else


  // Check if Meter Register
 if (TreeData.index=28) or
    (TreeData.index=29) or
    (TreeData.index=44) or
    (TreeData.index=45) then
 Begin
  MPANCore:=TreeData.D_SPAN;
  EFSDMSMTD:=TreeData.M_EFSDMSMTD;
  METERID:=TreeData.M_METERID;
  REGISTERID:=TreeData.M_REGISTERID;
  HH_Register:=TreeData.M_HH_REGISTER;
  if HH_register='H' then
  Begin
  { FRM_HH_DATA.show;
   FRM_HH_DATA.mpanlookup.keyvalue:=mpanCore;
   FRM_HH_DATA.quantitygroup.itemindex:=0;
   FRM_HH_DATA.showblank.checked:=false;
   application.processmessages;
   FRM_HH_DATA.doquery;}
  end
  else
  Begin
   {FRM_nhh_metering.show;
   FRM_nhh_metering.getmeterdetails(mpancore,efsdmsmtd,meterid,registerid);
  }
  end;
 end;

 // Check if Note or Enquiry
 if (TreeData.index=51) or
    (TreeData.index=52) or
    (TreeData.index=53) or

    (TreeData.index=257) or
    (TreeData.index=258) or
    (TreeData.index=259) or

    (TreeData.index=54) or
    (TreeData.index=55) or
    (TreeData.index=56) or
    (TreeData.index=57) or
    (TreeData.index=58) or
    (TreeData.index=59) or
    (TreeData.index=60) or
    (TreeData.index=61) or
    (TreeData.index=62) or
    (TreeData.index=208) or
    (TreeData.index=209) or
    (TreeData.index=210) or
    (TreeData.index=128) or
    (TreeData.index=63) or
    // Complaints
    (TreeData.index=263) or
    (TreeData.index=264) or
    (TreeData.index=265)
    then
 Begin
  S:=TreeData.C_record_id;
  ts:=TreeData.c_firstline;
  // What we Doing with doc
  if s='' then exit;

  screen.cursor:=crhourglass;
  Begin

   DisplayOrder:=5;
   TREEENQUIRYRESOLVED:=false;

   if (TreeData.index=63) or (TreeData.index=128) then
   Begin
    if Messagedlg('Select YES to open document, NO to open enquiry.',mtconfirmation,[MBYES,MBNO],0)=mryes
    then
    Begin
     // doc only
     // Need an alternative way to ge tthe document path, rather than lloking at enquiry table
     //ts:=FRM_ENQUIRY_SUMMARY.ENQUIRIES.fields[6].text;
     //shellexecute(Handle,'open',pchar(ts),nil,nil,sw_shownormal);
    FRM_COMMON.ShowImageDoc(ts);
    End
    else
    Begin
     // enquiry
     Show_Hot_Note(s); // LK 2016 SHOWNOTE
    End;
   end
   else
   Begin
    // Open Note or Enquiry
    Show_Hot_Note(s);   // LK 2016 SHOWNOTE
   End;
  end;

  //////////////////////////////////////////////////////////////////////////////
  // This no longer works as Enquiry is now show rather than showmodal;
 { if treeenquiryresolved=true then
  Begin
   treeview1.deletenode(xnode);
  end;}
 ///////////////////////////////////////////////////////////////////////////////
  screen.cursor:=crdefault;
 end;
end;

Procedure TFRM_Tree.BuildEnquiriesNode(MPANNODE:PvirtualNode);
Var
 MPAN,custid,premiseid:String;
Begin

 nodeData := treeview1.GetNodeData(mpannode);

 //mpannode:=treeview1.selected;
 mpan:=nodedata.D_SPAN;
 Custid:=nodedata.D_Customer_id;
 premiseid:=nodedata.D_premise_id;
 // Get Last Outstanding Service Order
 With GeneralQuery do
 Begin
  Close;
  DeleteVariables;
  DeclareVariable('MPAN', otString);
  sql.clear;
  sql.add('SELECT comments_1 FROM enquiry.enquiries');
  sql.add('Where contact_type=502 and');
  sql.add('MPANCORE=:MPAN');
  sql.add('and resolved=''N''');
  sql.add('and system_role=''X''');
  sql.add('order by date_raised desc');
  open;
  deletevariables;
  if recordcount<>0 then
  Begin
   Desc:='(S0001) - '+frm_common.getcomments(fields[0].text);
   {ServiceOrderNode:=Treeview1.items.AddChild(mpanNode,desc);
   ServiceOrderNode.imageindex:=6; // HH Meters
   ServiceOrderNode.selectedindex:=6;
   ServiceOrderNode.font.color:=clred; }

   ServiceOrderNode:=Treeview1.Addchild(mpannode);
   nodeData := Treeview1.GetNodeData(ServiceOrderNode);
   NodeData.caption := desc;
   nodedata.index:=6;
   nodedata.fontcolor:=clred;

  end;
 end;
end;

procedure TFRM_Tree.HideCheckClick(Sender: TObject);
begin
 try
  xNode := Treeview1.GetFirst();
  treeview1.expanded[xnode]:=false;
  treeview1.expanded[xnode]:=true;
 except
 end;
end;

procedure TFRM_Tree.Historic1Click(Sender: TObject);
begin
 historic1.checked:=not historic1.checked;
end;

procedure TFRM_Tree.HistoricPayPlanClick(Sender: TObject);
begin
  xnode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(xnode);

  FrmArgToPay := TFrmArgToPay.Create(self);
  try
    with FrmArgToPay Do
    begin
      tsCreatePlan.TabVisible := false;
      Agreement_Id := TreeData.D_Agreement_id;
      PageControl1.Activepage := tsHistory;
      FrmArgToPay.Caption := 'Arrangement To Pay: ' + Agreement_ID;
      ShowModal;
    end;
  finally
    FrmArgToPay.Free;
  end;
end;

procedure TFRM_Tree.S_DflowsClick(Sender: TObject);
begin
 if treeupdating=true then exit;
 mpannode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(mpannode);

// mpannode:=treeview1.selected;
 mpan:=TreeData.D_SPAN;
 //FRM_Dataflow_History.mpanlookup.visible:=false;
 FRM_Dataflow_History.DflowQuery(MPAN,'');
 FRM_Dataflow_History.Show_Mpan_Status(MPAN,'');
 if FRM_Dataflow_History.caption='' then
 Begin
  messagedlg('There is no Dataflow History for this MPAN',MTinformation,[MBOK],0);
  exit;
 end;
 FRM_Dataflow_history.show;
end;

procedure TFRM_Tree.Enquiries1Click(Sender: TObject);
var
  mpanID: String;
begin
  if treeupdating = True then
    exit;

  NodeData := Treeview1.GetNodeData(Treeview1.FocusedNode);
  FRM_ENQUIRY_SUMMARY.setoptions('X');
  mpanID := NodeData.D_Span;
  FRM_ENQUIRY_SUMMARY.Mpancore.Text := mpanID;
  MPAN := NodeData.D_Span;
  FRM_ENQUIRY_SUMMARY.findbtn.click();
  FRM_ENQUIRY_SUMMARY.show;
  FRM_ENQUIRY_SUMMARY.windowstate := wsnormal;
end;

procedure TFRM_Tree.BillpayCard1Click(Sender: TObject);
var
  agid,enddate:string;
  agreementId: Int64;
begin
  xnode:=treeview1.FocusedNode;
  TreeData:= treeview1.GetNodeData(xnode);
  Agid:=TreeData.D_Agreement_ID;

  if not TryStrToInt64(agid, agreementId) then
  begin
    MessageDlg('Invalid Agreement ID.', mtError, [mbOk], 0);
    Exit;
  end;

  // Need to Check if Non Prepay Agreement
  if TBillingUtil.UseDllRerating then
  begin
    if TRateAccountsWrapper.IsPrepayment(agreementId, false) then
    begin
      MessageDlg('This is not a CREDIT Agreement. BillPay is not available.', mtError, [mbok], 0);
      Exit;
    end;
  end
  else
  begin
    if Frm_Rate_Accounts.DoPrePayCheck(IntToStr(agreementId), '') then
    begin
      MessageDlg('This is not a CREDIT Agreement. BillPay is not available.', mtError, [mbok], 0);
      Exit;
    end;
  end;

  // Check if Agreement is Live?
  enddate:=TreeData.D_Agreement_end_date;
  if enddate<>'' then
  Begin
    if Messagedlg('Agreement is Terminated. Continue?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;
  End;

  Application.CreateForm(TFRM_BillPay,FRM_BillPay);
  try
    FRM_BillPay.Cardno.Caption:=frm_common.GET_BILLPAY_ID(Agid);
    //FRM_BillPay.Cardno.Caption:='98260135'+Agid;
    frm_billpay.ReasonCode.keyvalue:=0;
    FRM_BillPay.getDetails(Agid);
    FRM_BillPay.showmodal;
  finally
    FRM_BillPay.Release;
  end;
end;

procedure TFRM_Tree.BitBtn1Click(Sender: TObject);
begin
 {if pagecontrol1.ActivePage=TABSUPP then
 Begin
  Treeview1.printoptions.footer:=datetimetostr(now);
  Treeview1.Print(true);
 end;
 if pagecontrol1.ActivePage=TABMOP then
 Begin
  MopTree.printoptions.footer:=datetimetostr(now);
  MopTree.Print(true);
 End;}
end;

// BSL - 13/05/2015 - Fuel Direct Execute.
Procedure TFRM_Tree.BuildCustomerFuelDirect(C_Id: String);
Begin
  //  query table crm.agreements_fule_direct where status='A'
  //then Show line in tree
  // may need to externd this to look in payments recieved
  Try
    qrAgrFuelDir.Close;
    qrAgrFuelDir.SetVariable('Cust_Id', C_Id);
    qrAgrFuelDir.Open;

    If (Not qrAgrFuelDir.IsEmpty) and (qrAgrFuelDirSTATUS.AsString <> 'N') then
      Begin
        FuelDirectNode    := Treeview1.AddChild(MyNodeCustomer);
        NodeData          := Treeview1.GetNodeData(FuelDirectNode);
        NodeData.Caption  := 'Fuel Direct - Third Party Payments - ' + qrAgrFuelDirDESCRIPTION.AsString;
        NodeData.Index    := 254;
        NodeData.FontBold := True;
        FuelDirect        := qrAgrFuelDirSTATUS.AsString[1];

        If qrAgrFuelDirSTATUS.AsString = 'A' then
          Begin
            Nodedata.FontColor := $00358DAA;
          End // If
        Else
          Begin
            Nodedata.FontColor := clPurple;
          End; // Else
      End // If
    Else
      FuelDirect := 'N';

    qrAgrFuelDir.Close;
  Except
    On E: Exception do
      Begin
        Application.MessageBox(PChar('SQL Error qrAgrFuelDir --> ' + E.Message), Attn, MB_ICONERROR);
      End; // On
  End; // Try
End; // Proc


procedure TFRM_Tree.BuildCustomerNotifications(C_ID:string);
Var
 msg:string;
 ICONid:integer;
Begin
 With GeneralQuery do
  Begin
   Close;
   DeleteVariables;
   DeclareVariable(':RESULTS', otcursor);
   sql.clear;
   sql.Add('begin');
   sql.Add('CRM.PK_UTILITIES.PR_RET_NOTIFICATIONS('''+C_ID+''',:RESULTS);');
   sql.Add('end;');

   try
    open;
    deletevariables;
   except
    deletevariables;
    exit;
   end;

   if generalquery.RecordCount=0 then exit;

   NotificationsNode:=Treeview1.Addchild(mynodeCustomer);
   nodeData := Treeview1.GetNodeData(NotificationsNode);
   ICONID:=104;
   MSG:='Notifications';
   NodeData.caption := MSG;
   nodedata.index:=ICONID;

   while not generalquery.Eof do
   Begin
    NotificationNode:=Treeview1.Addchild(Notificationsnode);
    nodeData := Treeview1.GetNodeData(NotificationNode);
    ICONID:=strtoint(fields[0].Text);
    MSG:=fields[1].Text;
    NodeData.caption := MSG;
    nodedata.index:=ICONID;
  //  nodedata.fontcolor:=strtoint(fields[2].Text);
    next;
   end;
  end;
 ////////////////////////////////////////////////////////////////////////////////
end;

procedure TFRM_Tree.BuildCustomerNotes(C_ID:string);
Var
//MyRecPtr: PMyRec;
 x:integer;
 num,dt,wt:string;
 noteid:integer;
 NotificationsError: string;
 vCustBalance: Double;
 vCustHaveBalance: Boolean;
Begin
 if treelimit<1 then treelimit:=10;

 // Display Status of any IHD Cover Orders
 With GeneralQuery do
  Begin
   Close;
   DeleteVariables;
   DeclareVariable(':RESULTS', otcursor);
   sql.clear;
   sql.Add('begin');
   sql.Add('CRM.PK_UTILITIES.PR_RET_IHD_COVER_ORDER_STATUS('''+C_ID+''',:RESULTS);');
   sql.Add('end;');
   open;
   deletevariables;
   if generalquery.recordcount<>0 then
   Begin

    JbsNode:=Treeview1.Addchild(mynodeCustomer);
    nodeData := Treeview1.GetNodeData(JbsNode);
    noteid:=strtoint(fields[4].Text);
    status:=fields[5].Text;
    if fields[0].Text<>'' Then Nodedata.fontcolor:=clblack;
    if fields[1].Text<>'' Then Nodedata.fontcolor:=clpurple;
    if fields[2].Text<>'' Then Nodedata.fontcolor:=clgreen;

    desc:=status;
    NodeData.caption := desc;
    nodedata.index:=noteid;

   end;

  end;
 ////////////////////////////////////////////////////////////////////////////////


 ///////////////////////////////////////////////////////////////////////////////
 // Try and Get TreeDatafrom JBS Booking System. This work work for DUAL.        //
 ///////////////////////////////////////////////////////////////////////////////
 if frm_common.getvalue('CHECK_IN_JBS')='Y' then
 Begin
  try
   x:=1;
  With GeneralQuery do
  Begin
   Close;
   DeleteVariables;
   DeclareVariable(':RESULTS', otcursor);
   sql.clear;
   sql.Add('begin');
   sql.Add('SMIFF.PK_UTILITIES.pr_full_job_summary('''+C_ID+''',:RESULTS);');
   sql.Add('end;');
   open;
   deletevariables;
   if generalquery.recordcount<>0 then
   Begin
    //  add first note
    dt:='';
    NotificationsError := Trim(GeneralQuery.FieldByName('PREINSTALL_NOTIFICATIONS_ERROR').AsString);
    if generalquery.fields[9].text<>'' then
    Begin
     dt:=' - '+generalquery.fields[9].text;
     if generalquery.fields[9].text='25/12/2099' then dt:=' - Cold Call';
     if generalquery.fields[9].text='14/12/2099' then dt:=' - Faults';
     if generalquery.fields[9].text='01/01/2070' then dt:=' - Gas Only';
     if generalquery.fields[9].text='01/01/2069' then dt:=' - No Comms / Hub Checks';
     if generalquery.fields[9].text='31/12/2068' then dt:=' - No Vends';
    end;

    wt:=' - ';
    try
     wt:=' - '+generalquery.fields[26].Text+' - ';
    except
    end;
    desc:='(JBS ID:'+generalquery.fields[0].text+')'+wt+generalquery.fields[19].text+' - '+generalquery.fields[18].text;
    if fields[24].text='1' then desc:=desc+#10+'One Way Appointment'
    else
    if fields[24].text='2' then desc:=desc+#10+'Two Way Appointment';
    if fields[11].text='A' then desc:=desc+dt+' (AM)';
    if fields[11].text='P' then desc:=desc+dt+' (PM)';
    if Fields[27].Text<>'' then desc:=desc+' '+fields[27].Text;

    //desc:=desc+' - Last Updated: '+generalquery.fields[15].text;
    if v_fullnotes.checked then
    begin
     desc:=desc+#13+generalquery.fields[5].text+' - '+generalquery.fields[6].text;

    end;
    desc:=desc+#13+'Priority - '+ generalquery.fields[30].text;
    if NotificationsError <> EmptyStr then
      desc:=desc+#13+' - [Notifications Error - '+ NotificationsError + ']';

    noteid:=19;
    {JbsNode:=Treeview1.items.AddChild(mynodeCustomer,desc);
    JbsNode.imageindex:=noteid;
    JbsNode.selectedindex:=noteid;}

    JbsNode:=Treeview1.Addchild(mynodeCustomer);
    nodeData := Treeview1.GetNodeData(JbsNode);
    NodeData.caption := desc;
    nodedata.index:=noteid;

    nodedata.C_record_id :=generalquery.fields[0].text;
    //JbsNode.data:=MyRecPtr;
    Nodedata.fontcolor:=clblack;
    NodeData.C_record_status:=generalquery.fields[18].text;
    if generalquery.fields[18].text='Aborted' then Nodedata.fontcolor:=clred;
    if generalquery.fields[18].text='Booked' then Nodedata.fontcolor:=clmaroon;
    if generalquery.fields[18].text='Cancelled' then Nodedata.fontcolor:=clred;
    if generalquery.fields[18].text='Completed' then Nodedata.fontcolor:=clgreen;
    if generalquery.fields[18].text='On Route' then Nodedata.fontcolor:=clpurple;
    if generalquery.fields[18].text='On Site' then Nodedata.fontcolor:=clpurple;
    if generalquery.fields[18].text='No Access' then Nodedata.fontcolor:=clred;
    if generalquery.fields[18].text='Not Yet Booked' then Nodedata.fontcolor:=clblue;
    RefreshJBSPushBackNode(jbsnode);

    generalquery.next;
    while not generalquery.eof do   // add sub notes
    Begin
     inc(x);
     if x<10000 then num:=inttostr(x);
     if x<1000 then num:='0'+inttostr(x);
     if x<100 then num:='00'+inttostr(x);
     if x<10 then num:='000'+inttostr(x);
     // Desc:=getcomments(fields[1].text);
     dt:='';
     NotificationsError := Trim(GeneralQuery.FieldByName('PREINSTALL_NOTIFICATIONS_ERROR').AsString);
     if generalquery.fields[9].text<>'' then
     Begin
      dt:=' - '+generalquery.fields[9].text;
      if generalquery.fields[9].text='25/12/2099' then dt:=' - Cold Call';
      if generalquery.fields[9].text='14/12/2099' then dt:=' - Faults';
      if generalquery.fields[9].text='01/01/2070' then dt:=' - Gas Only';
      if generalquery.fields[9].text='01/01/2069' then dt:=' - No Comms / Hub Checks';
      if generalquery.fields[9].text='31/12/2068' then dt:=' - No Vends';
     end;
    wt:=' - ';
    try
     wt:=' - '+generalquery.fields[26].Text+' - ';
    except
    end;
    desc:='(JBS ID:'+generalquery.fields[0].text+')'+wt+generalquery.fields[19].text+' - '+generalquery.fields[18].text;

    if fields[24].text='1' then desc:=desc+#10+'One Way Appointment'
    else
    if fields[24].text='2' then desc:=desc+#10+'Two Way Appointment';
    if fields[11].text='A' then desc:=desc+dt+' (AM)';
    if fields[11].text='P' then desc:=desc+dt+' (PM)';
    if Fields[27].Text<>'' then desc:=desc+' '+fields[27].Text;
    //desc:=desc+' - Last Updated: '+generalquery.fields[15].text;
    if v_fullnotes.checked then
    begin
     desc:=desc+#13+generalquery.fields[5].text+' - '+generalquery.fields[6].text;
    end;
    desc:=desc+#13+'Priority - '+ generalquery.fields[30].text;

    if NotificationsError <> EmptyStr then
      desc:=desc+#13+' - [Notifications Error - '+ NotificationsError + ']';

    { desc:='(JBS ID:'+generalquery.fields[0].text+') - '+generalquery.fields[19].text+' - '+generalquery.fields[18].text+dt;
    // desc:=desc+' - Last Updated: '+generalquery.fields[15].text;
     if v_fullnotes.checked then desc:=desc+#10+generalquery.fields[5].text+' -
        '+generalquery.fields[6].text;  }
     noteid:=19;
     {jbssubnode:=Treeview1.items.AddChild(JbsNode,desc);
     jbssubnode.imageindex:=noteid;
     jbssubnode.selectedindex:=noteid; }

     jbssubnode:=Treeview1.Addchild(JbsNode);
     nodeData := Treeview1.GetNodeData(jbssubnode);
     NodeData.caption := desc;
     nodedata.index:=noteid;

     nodedata.C_record_id :=generalquery.fields[0].text;
     //jbssubnode.data:=MyRecPtr;
     nodedata.fontcolor:=clblack;
     if generalquery.fields[18].text='Aborted' then nodedata.fontcolor:=clred;
     if generalquery.fields[18].text='Booked' then nodedata.fontcolor:=clmaroon;
     if generalquery.fields[18].text='Cancelled' then nodedata.fontcolor:=clred;
     if generalquery.fields[18].text='Completed' then nodedata.fontcolor:=clgreen;
     if generalquery.fields[18].text='On Route' then nodedata.fontcolor:=clpurple;
     if generalquery.fields[18].text='On Site' then nodedata.fontcolor:=clpurple;
     if generalquery.fields[18].text='No Access' then nodedata.fontcolor:=clred;
     if generalquery.fields[18].text='Not Yet Booked' then nodedata.fontcolor:=clblue;
    RefreshJBSPushBackNode(jbssubnode);
      generalquery.next;
    end;

   end;
  end;
  except
  end;
 end; // end check JBS

 ///////////////////////////////////////////////////////////////////////////////
 // Try and Get TreeDatafrom FMS Booking System. This will not work for DUAL.        //
 ///////////////////////////////////////////////////////////////////////////////
 if frm_common.getvalue('CHECK_IN_FMS')='Y' then
 Begin
  x:=1;
  try
  With GeneralQuery do
  Begin
   Close;
     DeleteVariables;
   DeclareVariable(':RESULTS', otcursor);
   sql.clear;
   sql.Add('begin');
   sql.Add('CRM.PK_UTILITIES.PR_RET_FMS_FAULTS('''+C_ID+''',:RESULTS);');
   sql.Add('end;');
   open;
   deletevariables;
   if generalquery.recordcount<>0 then
   Begin
    desc:=fields[1].Text;
    noteid:=78;
    fmsnode:=Treeview1.Addchild(mynodecustomer);
    nodeData := Treeview1.GetNodeData(fmsnode);
    NodeData.caption := desc;
    nodedata.index:=noteid;
    nodedata.C_record_id :=generalquery.fields[0].text;
    nodedata.fontcolor:=clblack;
    generalquery.next;

    while not generalquery.eof do   // add sub notes
    Begin
     inc(x);
     if x<10000 then num:=inttostr(x);
     if x<1000 then num:='0'+inttostr(x);
     if x<100 then num:='00'+inttostr(x);
     if x<10 then num:='000'+inttostr(x);
     desc:=fields[1].Text;
     noteid:=78;
     fmssubnode:=Treeview1.Addchild(fmsnode);
     nodeData := Treeview1.GetNodeData(fmssubnode);
     NodeData.caption := desc;
     nodedata.c_firstline:=generalquery.fields[1].text;
     nodedata.index:=noteid;
     nodedata.C_record_id :=generalquery.fields[0].text;
     nodedata.fontcolor:=clblack;
     generalquery.next;
    end;
   end;
  end;
  except
  end;
 end; // end FMS

 ///////////////////////////////////////////////////////////////////////////////
 // Get Notes For Customer                                                    //
 ///////////////////////////////////////////////////////////////////////////////
 x:=1;
 With GeneralQuery do
 Begin
  Close;
  DeleteVariables;
  DeclareVariable(':RESULTS', otcursor);
  sql.clear;
  sql.Add('begin');
  sql.Add('CRM.PK_UTILITIES.PR_RET_NOTES('''+C_ID+''',:RESULTS);');
  sql.Add('end;');
  open;
  deletevariables;
  if generalquery.recordcount<>0 then
  Begin
  //  add first note
   Firstline:=frm_common.getFirstLine(generalquery.fields[1].text);
   desc:='(N0001) - '+generalquery.fields[2].text+' - '+generalquery.fields[0].text;
   if v_fullnotes.checked then desc:=desc+#10+generalquery.fields[1].text
   else desc:=desc+' '+firstline;
   noteid:=53;
   if generalquery.fields[5].text<>'' then noteid:=52;
   if generalquery.fields[4].text<>'' then noteid:=51;
   NoteNode:=Treeview1.Addchild(mynodeCustomer);
   nodeData := Treeview1.GetNodeData( NoteNode);
   NodeData.caption := desc;
   nodedata.c_firstline:=generalquery.fields[1].text;
   nodedata.index:=noteid;
   nodedata.C_record_id :=generalquery.fields[6].text;
   Nodedata.fontcolor:=clblack;
   if  generalquery.fields[7].Text='Y' then
   Begin
    NodeData.fontcolor:=clPurple;
    nodedata.index:=noteid+206;
   end;
   generalquery.next;
   while not generalquery.eof do   // add sub notes
   Begin
    inc(x);
    if x<10000 then num:=inttostr(x);
    if x<1000 then num:='0'+inttostr(x);
    if x<100 then num:='00'+inttostr(x);
    if x<10 then num:='000'+inttostr(x);
    Firstline:=frm_common.getFirstLine(generalquery.fields[1].text);
    desc:='(N'+num+') - '+generalquery.fields[2].text+' - '+generalquery.fields[0].text;
    if v_fullnotes.checked then desc:=desc+#10+generalquery.fields[1].text
    else desc:=desc+' '+firstline;
    noteid:=53;
    if generalquery.fields[5].text<>'' then noteid:=52;
    if generalquery.fields[4].text<>'' then noteid:=51;
    NotesubNode:=Treeview1.Addchild(NoteNOde);
    nodeData := Treeview1.GetNodeData(NotesubNode);
    NodeData.caption := desc;
    nodedata.c_firstline:=generalquery.fields[1].text;
    nodedata.index:=noteid;
    nodedata.C_record_id :=generalquery.fields[6].text;
    nodedata.fontcolor:=clblack;
    generalquery.next;
   end;
  end;
 end;

///////////////////////////////////////////////////////////////////////////////
 // Get Disput Note Notes For Customer
 //added by maryam on 15/08/2017 on wrike ticket  163254834                                       //
 ///////////////////////////////////////////////////////////////////////////////
  x:=1;
 With GeneralQuery do
 Begin
  Close;
  DeleteVariables;
  DeclareVariable(':RESULTS', otcursor);
  sql.clear;
  sql.Add('begin');
  sql.Add('CRM.PK_UTILITIES.PR_RET_DISPUTE_NOTES('''+C_ID+''',:RESULTS);');
  sql.Add('end;');
  open;
  deletevariables;
  if generalquery.recordcount<>0 then
  Begin
  //  add first note
   desc:='(N0001) - '+generalquery.FieldByName('COMMENTS').AsString;
   NoteNode:=Treeview1.Addchild(mynodeCustomer);
   nodeData := Treeview1.GetNodeData( NoteNode);
   NodeData.caption := desc;
   nodedata.index:=152;
   Nodedata.fontcolor:=clblack;

   generalquery.next;
   while not generalquery.eof do   // add sub notes
   Begin
    inc(x);
    if x<10000 then num:=inttostr(x);
    if x<1000 then num:='0'+inttostr(x);
    if x<100 then num:='00'+inttostr(x);
    if x<10 then num:='000'+inttostr(x);

    desc:='(N'+num+') - '+generalquery.FieldByName('COMMENTS').AsString;
    NotesubNode:=Treeview1.Addchild(NoteNOde);
    nodeData := Treeview1.GetNodeData(NotesubNode);
    NodeData.caption := desc;
    nodedata.index:=152;
    nodedata.fontcolor:=clblack;
    generalquery.next;
   end;
  end;
 end;
//-----------------------
//---------------------

   // Get Outstanding Flags for Customer
 x:=0;
 With GeneralQuery do
 Begin
  Close;
  DeleteVariables;
  DeclareVariable(':RESULTS', otcursor);
  sql.clear;
  sql.Add('begin');
  sql.Add('CRM.PK_UTILITIES.PR_RET_FLAGS('''+C_ID+''',:RESULTS);');
  sql.Add('end;');
  open;
  deletevariables;
  if generalquery.recordcount<>0 then
  Begin

   if generalquery.recordcount>TREELIMIT  then
   Begin
    desc:='(FLAGS Outstanding for Customer = '+inttostr(generalquery.recordcount)+')';
    NoteTopNode:=Treeview1.Addchild(mynodecustomer);
    nodeData := Treeview1.GetNodeData(NoteTopNode);
    NodeData.caption := desc;
    nodedata.index:=207;
   end;

   while not generalquery.eof do      // add sub notes
   Begin
    inc(x);
    if x<10000 then num:=inttostr(x);
    if x<1000 then num:='0'+inttostr(x);
    if x<100 then num:='00'+inttostr(x);
    if x<10 then num:='000'+inttostr(x);
    Firstline:=frm_common.getFirstLine(generalquery.fields[1].text);
    desc:='(F'+num+') - ';
    if generalquery.fields[4].text<>'' then desc:=desc+generalquery.fields[4].text+' - ';
    desc:=desc+generalquery.fields[2].text+' - ';
    desc:=desc+generalquery.fields[0].text;
    noteid:=208;
    if generalquery.fields[7].text<>'' then desc:=desc+' - (Owned by '+generalquery.fields[7].text+') ';
    if v_fullnotes.checked then desc:=desc+#10+generalquery.fields[1].text
    else desc:=desc+' '+firstline;
    if generalquery.recordcount<=TREELIMIT  then
    Begin
     NotesubNode:=Treeview1.Addchild(mynodecustomer);
     nodeData := Treeview1.GetNodeData(NotesubNode);
     NodeData.caption := desc;
     nodedata.c_firstline:=generalquery.fields[1].text;
    end
    else
    Begin
     NotesubNode:=Treeview1.Addchild(NoteTopNode);
     nodeData := Treeview1.GetNodeData(NotesubNode);
     NodeData.caption := desc;
     nodedata.c_firstline:=generalquery.fields[1].text;
    end;
    nodedata.index:=noteid;
    Nodedata.fontcolor:=clred;
    if generalquery.fields[3].text<>'' then
    Begin
     if strtodatetime(generalquery.fields[3].text)<(now) then
     Begin
      noteid:=210;
      nodedata.index:=noteid;
     end;
    end;

    nodedata.C_record_id :=generalquery.fields[6].text;
    nodedata.C_Owner :=generalquery.fields[7].text;
    nodedata.C_Raised_by :=generalquery.fields[8].text;
    nodedata.C_Date_Raised :=generalquery.fields[2].text;
    Nodedata.fontcolor:=clblack;
    generalquery.next;
   end;
  end;
 end;

   // Get Outstanding Eqnuiries for Customer
 x:=0;
 With GeneralQuery do
 Begin
  Close;
  DeleteVariables;
  DeclareVariable(':RESULTS', otcursor);
  sql.clear;
  sql.Add('begin');
  sql.Add('CRM.PK_UTILITIES.PR_RET_ENQUIRIES('''+C_ID+''',:RESULTS);');
  sql.Add('end;');
  open;
  deletevariables;
  if generalquery.recordcount<>0 then
  Begin

   if generalquery.recordcount>TREELIMIT  then
   Begin
    desc:='(ENQUIRIES Outstanding for Customer = '+inttostr(generalquery.recordcount)+')';
    NoteTopNode:=Treeview1.Addchild(mynodecustomer);
    nodeData := Treeview1.GetNodeData(NoteTopNode);
    NodeData.caption := desc;
    NodeData.index:=207;
   end;

   while not generalquery.eof do      // add sub notes
   Begin
    inc(x);
    if x<10000 then num:=inttostr(x);
    if x<1000 then num:='0'+inttostr(x);
    if x<100 then num:='00'+inttostr(x);
    if x<10 then num:='000'+inttostr(x);
    Firstline:=frm_common.getFirstLine(generalquery.fields[1].text);
    desc:='(E'+num+') - ';
    if generalquery.fields[4].text<>'' then desc:=desc+generalquery.fields[4].text+' - ';
    desc:=desc+generalquery.fields[2].text+' - ';
    desc:=desc+generalquery.fields[0].text;
    noteid:=56;
    if generalquery.fields[5].text<>'' then noteid:=55;
    if generalquery.fields[4].text<>'' then noteid:=54;
    if generalquery.fields[7].text<>'' then desc:=desc+' - (Owned by '+generalquery.fields[7].text+') ';
    if generalquery.fields[9].text<>'D' then
    Begin
     if v_fullnotes.checked then desc:=desc+#10+generalquery.fields[1].text
     else desc:=desc+' '+firstline;
    end;
    if generalquery.fields[9].text='D' then
    Begin
     //desc:=desc+generalquery.fields[2].text;
     if generalquery.fields[6].text<>'' then desc:=desc+'. Document ID -';
     desc:=desc+generalquery.fields[6].text;
     noteid:=128;
    End;
    if generalquery.recordcount<=TREELIMIT then
    Begin
     NotesubNode:=Treeview1.Addchild(mynodecustomer);
     nodeData := Treeview1.GetNodeData(NotesubNode);
     NodeData.caption := desc;
     nodedata.c_firstline:=generalquery.fields[1].text;
    end
    else
    Begin
     NotesubNode:=Treeview1.Addchild(NoteTopNode);
     nodeData := Treeview1.GetNodeData(NotesubNode);
     NodeData.caption := desc;
     nodedata.c_firstline:=generalquery.fields[1].text;
    end;
    nodedata.index:=noteid;
    Nodedata.fontcolor:=clred;
    if generalquery.fields[3].text<>'' then
    Begin
     if strtodatetime(generalquery.fields[3].text)<(now) then
     Begin
      if noteid<>128 then
      Begin
       noteid:=62;
       if generalquery.fields[5].text<>'' then noteid:=61;
       if generalquery.fields[4].text<>'' then noteid:=60;
      end
      else noteid:=63;
     NodeData.index:=noteid;
     end;
    end;

    nodedata.C_record_id :=generalquery.fields[6].text;
    nodedata.C_Owner :=generalquery.fields[7].text;
    nodedata.C_Raised_by :=generalquery.fields[8].text;
    nodedata.C_Date_Raised :=generalquery.fields[2].text;
    nodedata.fontcolor:=clblack;
    generalquery.next;
   end;

  end;
 end;

  // Get Outstanding Complaints for Customer
 x:=0;
 With GeneralQuery do
 Begin
  Close;
  DeleteVariables;
  DeclareVariable(':RESULTS', otcursor);
  sql.clear;
  sql.Add('begin');
  sql.Add('CRM.PK_UTILITIES.PR_RET_COMPLAINTS('''+C_ID+''',:RESULTS);');
  sql.Add('end;');
  open;
  deletevariables;
  if generalquery.recordcount<>0 then
  Begin

   if generalquery.recordcount>TREELIMIT  then
   Begin
    desc:='(COMPLAINTS Outstanding for Customer = '+inttostr(generalquery.recordcount)+')';
    NoteTopNode:=Treeview1.Addchild(mynodecustomer);
    nodeData := Treeview1.GetNodeData(NoteTopNode);
    NodeData.caption := desc;
    NodeData.index:=267; // Complaint ICON
   end;

   while not generalquery.eof do
   Begin
    inc(x);
    if x<10000 then num:=inttostr(x);
    if x<1000 then num:='0'+inttostr(x);
    if x<100 then num:='00'+inttostr(x);
    if x<10 then num:='000'+inttostr(x);
    Firstline:=frm_common.getFirstLine(generalquery.fields[1].text);
    //desc:='(Query ID: '+generalquery.fields[6].text+') - ';
    desc:='(Complaint Reference Number: '+generalquery.fields[6].text+') - ';

    if generalquery.fields[4].text<>'' then desc:=desc+generalquery.fields[4].text+' - ';
    desc:=desc+generalquery.fields[2].text+' - ';
    desc:=desc+generalquery.fields[0].text;
    if generalquery.fields[10].text<>'' then desc:=desc+' - ('+generalquery.fields[10].text+') ';
    if generalquery.fields[7].text<>'' then desc:=desc+' - (Owned by '+generalquery.fields[7].text+') ';

    if v_fullnotes.checked then desc:=desc+#10+generalquery.fields[1].text
    else desc:=desc+' '+firstline;

    if generalquery.recordcount<=TREELIMIT  then
    Begin
     NotesubNode:=Treeview1.Addchild(mynodecustomer);
     nodeData := Treeview1.GetNodeData(NotesubNode);
     NodeData.caption := desc;
     nodedata.c_firstline:=generalquery.fields[1].text;
    end
    else
    Begin
     NotesubNode:=Treeview1.Addchild(NoteTopNode);
     nodeData := Treeview1.GetNodeData(NotesubNode);
     NodeData.caption := desc;
     nodedata.c_firstline:=generalquery.fields[1].text;
    end;


    Nodedata.fontcolor:=clred;
    noteid:=263;
    IF generalquery.fields[11].text='G' THEN noteid:=263; // g
    IF generalquery.fields[11].text='A' THEN noteid:=264; // a
    IF generalquery.fields[11].text='R' THEN noteid:=265; // r

    nodedata.index:=noteid;

    nodedata.C_record_id :=generalquery.fields[6].text;
    nodedata.C_Owner :=generalquery.fields[7].text;
    nodedata.C_Raised_by :=generalquery.fields[8].text;
    nodedata.C_Date_Raised :=generalquery.fields[2].text;
    nodedata.fontcolor:=clblack;
    generalquery.next;
   end;

  end;
 end;

  vCustHaveBalance := TCrmUtil.CustHaveBalance(StrToInt64(C_ID), vCustBalance);

  if IsSmartPay(C_ID) then
  begin
    if not vCustHaveBalance then
      vCustBalance := 0;

    MyUtilitaWalletPowerPayNode := Treeview1.AddChild(MyNodeCustomer);
    NodeData := Treeview1.GetNodeData(MyUtilitaWalletPowerPayNode);
    NodeData.Caption := 'Power Pot: ' + Chr(163) + FormatCurr('#,##0.####', vCustBalance);
    NodeData.Index := FPowerPayIconIndex;
  end
  else
  begin
    if vCustHaveBalance then
    begin
      MyUtilitaWalletPowerPayNode := Treeview1.AddChild(MyNodeCustomer);
      NodeData := Treeview1.GetNodeData(MyUtilitaWalletPowerPayNode);
      NodeData.Caption := 'My Utilita Savings Balance: ' + Chr(163) + FormatCurr('#,##0.####', vCustBalance);
      NodeData.Index := FBalanceAccountIconIndex;
    end;
  end;

  TSavingsTransactions.EnableCreditDebitMenuOptions(StrToInt64(C_ID),
    mniTransferCreditFromSavingsToAgreement, mniTransferDebitFromSavingsToAgreement);
  TSavingsTransactions.EnableCreditDebitMenuOptions(StrToInt64(C_ID),
    mniTransferCreditFromSavingsToMeter, nil);
end;

procedure TFRM_Tree.BuildCustomerLosses(C_ID:string);
Var
 x:integer;
 priority,pdesc,odesc:string;
 noteid:integer;
Begin


 ///////////////////////////////////////////////////////////////////////////////
 // Get Pending Losses For Customer - Unactioned with Telesales               //
 ///////////////////////////////////////////////////////////////////////////////
 With Frm_loss_result.Losses do
 Begin
  close;
  setvariable('CID',C_ID);
  open;
  x:=1;
  while not Frm_loss_result.Losses.eof do   // add sub notes
  Begin
   if Frm_loss_result.Losses.fields[0].text='E' then
   Begin
    noteid:=5;  // Elec
    desc:='Unactioned Pending Loss on '+Frm_loss_result.Losses.fields[5].text+':  Electricity Supply No: '+Frm_loss_result.Losses.fields[2].text;
   end;
   if Frm_loss_result.Losses.fields[0].text='G' then
   Begin
    noteid:=66; // Gas
    desc:='Unactioned Pending Loss on '+Frm_loss_result.Losses.fields[5].text+':  Gas Supply No: '+Frm_loss_result.Losses.fields[2].text;
   end;
   priority:=Frm_loss_result.Losses.fields[6].text;
   pdesc:='';
   odesc:='';
   if Frm_loss_result.Losses.fields[13].Text='N' then odesc:=' - (Objection window Expired)';

   if priority='T' then Pdesc:='Priority 1 - Installed';
   if priority='U' then Pdesc:='Priority 2 - Installed';
   if priority='V' then Pdesc:='Priority 3 - Uninstalled';
   if priority='W' then Pdesc:='Priority 4 - Uninstalled';
   desc:=desc+#13+pdesc+odesc;
   NoteNode:=Treeview1.Addchild(mynodeCustomer);
   nodeData := Treeview1.GetNodeData(NoteNode);
   NodeData.caption := desc;
   nodedata.index:=noteid;
   nodedata.D_Customer_id :=Frm_loss_result.Losses.fields[1].text;
   nodedata.D_REFNO :=Frm_loss_result.Losses.fields[4].text;
   nodedata.D_SPAN :=Frm_loss_result.Losses.fields[2].text;
   nodedata.d_priority:=priority;
   nodedata.d_InObjPeriod:=Frm_loss_result.Losses.fields[13].text;
   nodedata.d_pdesc:=pdesc;
   nodedata.d_order:=x;
   Nodedata.fontcolor:=clPurple;
   Nodedata.fontBold:=true;
   inc(x);
   Frm_loss_result.Losses.next;
  end;
 end;
end;

procedure Tfrm_Tree.AddCustomerNote(pTreeView: TVirtualStringTree);
var
  selNode: PVirtualNode;
  data: PMyRec;
begin
  selNode := pTreeView.FocusedNode;
  if not Assigned(selNode) then
    exit;

  data := pTreeView.GetNodeData(selNode);
  if not Assigned(data) then
    exit;

  Raise_Hot_Note(3, data.D_Customer_Id);
end;

procedure TFRM_Tree.AddCustomerNote1Click(Sender: TObject);
begin
  AddCustomerNote(Treeview1);
end;

procedure TFRM_Tree.Raiseerroneoustransferrequest1Click(Sender: TObject);
Var
  regid,span,ssd:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 REGID:=TreeData.D_REGid;
 SPAN:=TreeData.D_SPAN;
 SSD:=TreeData.D_SSD;
 Application.CreateForm(TFRM_GAS_ET, FRM_GAS_ET);
 try
  FRM_GAS_ET.clearfields;
  FRM_GAS_ET.SETDefault(SPAN);
  FRM_GAS_ET.ShowModal;
 finally
  FRM_GAS_ET.release;
 end;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.Raise_Hot_Note(aNType: integer; aCustomerId: string);
var
  customerId : Int64;
begin

  try
    customerId := StrToInt64(aCustomerId);
  except
    MessageDlg('Invalid customer ID! ' + aCustomerId, mtError, [mbOk], 0);
    exit;
  end;

  gHotNoteList.AddHotNote(Self, aNType, customerId);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.Show_Hot_Note(aNID: string);
begin
  gHotNoteList.ShowHotNote(Self, aNID);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.AddComplaint1Click(Sender: TObject);
var
  bRaise: Boolean;
begin
 bRaise := False;
 // see if open complaints already exist.
 // and allow the user to decide if they wish to raise a new one.
  xnode:=treeview1.FocusedNode;
  TreeData:= treeview1.GetNodeData(xnode);
  with Main_Data_Module.TempQuery do
  begin
    Close;
    DeleteVariables;
    DeclareVariable('CUSTID',otString);
    DeclareVariable('COMPLAINT', otString);
    DeclareVariable('RESLV', otString);
    SQL.Clear;
    SQL.Add('SELECT COUNT(*) ');
    SQL.Add('FROM ENQUIRY.ENQUIRIES E, ENQUIRY.REQUEST_TYPE R ');
    SQL.Add('WHERE E.CUSTOMER_ID=:CUSTID ');
    SQL.Add('AND R.ID = E.REQUEST_TYPE ');
    SQL.Add('AND R.ENQIURY_OR_NOTE = :COMPLAINT ');
    SQL.Add('AND E.RESOLVED = :RESLV ');
    SetVariable('CUSTID', Treedata.D_Customer_ID);
    SetVariable('COMPLAINT', 'C');
    SetVariable('RESLV', 'N');
    Open;
    DeleteVariables;

    if (RecordCount <> 0) and (Fields[0].Value > 0) then
    begin
      if MessageDlg('Existing open complaints exist for this customer ' + #13#10 +
        'Are you sure you wish open a new complaint?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        bRaise := True;
      end;
    end
    else
    begin
      bRaise := True;
    end;
  end;

  if bRaise then
  begin
    xnode:=treeview1.FocusedNode;
    TreeData:= treeview1.GetNodeData(xnode);
    Raise_Hot_Note(5,Treedata.D_Customer_ID);
  end;
end;

procedure TFRM_Tree.AddCustomerEnquiry1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Raise_Hot_Note(2,Treedata.D_Customer_ID);
end;

procedure TFRM_Tree.V_FullNotesClick(Sender: TObject);
begin
 V_Fullnotes.checked:=not V_fullnotes.checked;
end;

procedure TFRM_Tree.V_NonClick(Sender: TObject);
begin
 v_non.checked:=not v_non.checked;
end;

procedure TFRM_Tree.ShowAllCustomerComplaints1Click(Sender: TObject);
Var
CustomerID,CustomerName:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;
 Begin
  CustomerID:=TreeData.D_Customer_ID;
  CustomerName:=TreeData.D_Customer_Name;
  FRM_enquiry_summary.custname_label.caption:=CUSTOMERID+' - '+CustomerName;

  frm_enquiry_summary.setoptions('C');
  FRM_ENQUIRY_SUMMARY.CustomerEnquiryNotes(CustomeriD);
  FRM_Enquiry_Summary.show;
  FRM_Enquiry_Summary.windowstate:=wsnormal;
 end;
end;

procedure TFRM_Tree.ShowAllEnquiriesNotes1Click(Sender: TObject);
Var
CustomerID,CustomerName:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;
 Begin
  CustomerID:=TreeData.D_Customer_ID;
  CustomerName:=TreeData.D_Customer_Name;
  FRM_enquiry_summary.custname_label.caption:=CUSTOMERID+' - '+CustomerName;

  frm_enquiry_summary.setoptions('X');
  FRM_ENQUIRY_SUMMARY.CustomerEnquiryNotes(CustomeriD);
  FRM_Enquiry_Summary.show;
  FRM_Enquiry_Summary.windowstate:=wsnormal;
 end;
end;

procedure TFRM_Tree.E_TakeOwnershipClick(Sender: TObject);
var
s,raisedby,dateraised:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 S:=TreeData.C_record_id;
 raisedby:=TreeData.C_raised_by;
 dateraised:=TreeData.C_date_raised;

  if MessageDlg('Are you sure you wish to take ownership of this Enquiry?',
    mtConfirmation, [mbYes, mbNo], 0) = mrno then exit;
 // Get Full Details Of Enquiry
 // Update Ownership
 // Add a log

 with main_data_module.updatequery do
 Begin
  Try
   close;
   sql.clear;
   sql.add('update enquiry.enquiries set owner='''+userid+''' where record_id='+s);
   execute;
   close;
   sql.clear;
   sql.add('Insert into enquiry.audit_log values('+
   S+','''+
   raisedby+''',to_date('''+
   dateraised+''',''DD/MM/YYYY hh24:mi:ss''),''Ownership Taken By User'','''+userid+''','''+
   userid+''',NULL,sysdate,trunc(sysdate))');
   execute;
  except
   Messagedlg('No Enquiries were Assigned',MTinformation,[MBOK],0);
   exit;
  end;
 end;
 FRM_Login.MainSession.commit;
 Messagedlg('Enquiry Assigned.',MTinformation,[MBOK],0);
 nodeData := treeview1.GetNodeData(xnode);
 nodedata.caption:=nodedata.caption+' - (Owned by '+userid+')';
 E_takeownership.enabled:=false;
end;

procedure TFRM_Tree.MeterDetailsReadings1Click(Sender: TObject);
begin
 if treeupdating=true then exit;
 mpannode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(mpannode);

 mpan:=TreeData.D_SPAN;
 with mtds do
 begin
  close;
  setvariable('MPAN',MPAN);
  open;
 end;
 if mtds.recordcount=0 then
 Begin
  with generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('MPAN', otString);
   sql.clear;
   sql.add('select count(*) from edmgr.reads_with_no_mtds');
   sql.add('where mpancore=:MPAN');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
  End;
  if generalquery.recordcount=0 then
  Begin
   Messagedlg('No Meter Technical Details Exist for this Supply Point',MTInformation,[MBOK],0);
   exit;
  End;
 end;

 FRM_nhh_metering.show;
 FRM_nhh_metering.getmeterdetails(mpan,'','','', custid);
end;

procedure TFRM_Tree.MaintainCustomerDetailsClick(Sender: TObject);
Var
Custid{,mpan}:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 custid:=TreeData.D_Customer_ID;

 Application.CreateForm(TFRM_Maintain_Customer, FRM_Maintain_Customer);
 try
 FRM_Maintain_Customer.GetCustomerDetails(Custid);
 FRM_Maintain_Customer.ShowModal;
 if FRM_maintain_Customer.tag<>0 then
 Begin
  // Customer Details Changed there for send any D0302s for Electric MPANS
  {with main_data_module.generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('CID', otstring);
   sql.clear;
   sql.add('select distinct span from crm.customer_to_registration');
   sql.add('where customer_id=:CID');
   sql.add('and span_type in (''O'',''E'',''D'',''S'',''1'',''G'',''B'')');
   setvariable('CID',custid);
   open;
   deletevariables;
  End;
  while not main_data_module.generalquery.eof do
  Begin
   mpan:=main_data_module.generalquery.fields[0].text;
   FRM_Customer.CreateD0302('R',mpan,'ALL'); // Distributor
   FRM_Customer.CreateD0302('C',mpan,'ALL'); // DC
   FRM_Customer.CreateD0302('D',mpan,'ALL'); // DC
   FRM_Customer.CreateD0302('M',mpan,'ALL'); // MO
   main_data_module.generalquery.next;
  end;}
 end;
 finally
  FRM_Maintain_Customer.release;
 end;
 FRM_Main.SearchForcust(custid);
 // Tree is refreshed so Node is always reset

 xNode := Treeview1.GetFirst();
 treeview1.expanded[xnode]:=false;
 treeview1.expanded[xnode]:=true;

end;


Procedure TFRM_Tree.RefreshCustomerNode(XNode: PVirtualNode);
Var
  productid,productname,effective_from,effective_to,status,
  paymentplan,collectionrate,desc,fad,debt,LIB_ID,msg:string;
  f,
  TelNum,EmailNum  : Integer;
  IsSuper, hasneeds: Boolean;
//////////////////////////////////////////////////////////////////////////////////////////////////
// Builds / Refreshes Customer Tree                                                             //
//////////////////////////////////////////////////////////////////////////////////////////////////
Begin
  NodeData       := Treeview1.GetNodeData(XNode);
  mynodeCustomer := XNode;
  //  SJ-BSL - 30/04/2021 - Change 214 Icon by 309..311  (Changed back 07/05/2021)SJ
  IsSuper := NodeData.Index = 214;
  fPremiseDcc:= false;

 inc(custcount);

 // DANBYT - 04/12/2024 - CRMX-164 - Ensure CustId persists even if nodedata.D_Customer_Id is empty
 PersistStringValue(CustID,nodedata.D_Customer_Id);

 LIB_ID:=nodedata.D_LIB_CUST_ID;
 Cdebt:=nodedata.D_CDEBT;

 issuper:=false;
 try
  treeview1.deletechildren(xnode);
 except
 end;

  ///////////////////////////////////////////////////////////////////////////////
 // Add Quick Shortcut to Customer MENU for SMETS meters
 SM_E.Tag:=0;
 SM_E.Visible:=false;
 SM_G.Tag:=0;
 SM_G.Visible:=false;
 SM_B.Visible:=false;
 nodedata.D_SPANE:='';
 nodedata.D_SPANG:='';
 nodedata.D_SpanDCC_E := false;
 nodedata.D_SpanDCC_G := false;

 with Generalquery do
 Begin
  close;
  DeleteVariables;
  DeclareVariable('CID', otstring);
  sql.clear;
  sql.add('Select crm.fn_has_smart_elec(:CID) from dual');
  setvariable('CID',custid);
  open;
  deletevariables;
 End;
 // Will return MPANCORE is customer has ONE live SMETS Elec Supply
 if generalquery.Fields[0].Text<>'' then
 Begin
  nodedata.D_SPANE:=generalquery.Fields[0].Text;
  sm_e.Visible:=true;
  sm_b.Visible:=true;
 End;

 with Generalquery do
 Begin
  close;
  DeleteVariables;
  DeclareVariable('CID', otstring);
  sql.clear;
  sql.add('Select crm.fn_has_smart_gas(:CID) from dual');
  setvariable('CID',custid);
  open;
  deletevariables;
 End;
  // Will return MPRN is customer has ONE live SMETS Gas Supply
 if generalquery.Fields[0].Text<>'' then
 Begin
  nodedata.D_SPANG:=generalquery.Fields[0].Text;
  sm_g.Visible:=true;
  sm_b.Visible:=true;
 End;

 // check DCC managed electric meter
 with Generalquery do
 Begin
  close;
  DeleteVariables;
  DeclareVariable('MPXN', otstring);
  sql.clear;
  sql.add('Select ods.dcc_enrolled(:MPXN)from dual');
  setvariable('MPXN', nodedata.D_SPANE);
  open;
  deletevariables;

  if generalquery.Fields[0].Text = 'Y' then
  begin
    nodedata.D_SpanDCC_E := true;
    fPremiseDcc:=true;
  end;
 End;

 // check DCC managed gas meter
 with Generalquery do
 Begin
  close;
  DeleteVariables;
  DeclareVariable('MPXN', otstring);
  sql.clear;
  sql.add('Select ods.dcc_enrolled(:MPXN)from dual');
  setvariable('MPXN', nodedata.D_SPANG);
  open;
  deletevariables;

  if generalquery.Fields[0].Text = 'Y' then
  begin
    nodedata.D_SpanDCC_G := true;
    fPremiseDcc:=true;
  end;
 End;


 ///////////////////////////////////////////////////////////////////////////////

 with FRM_main_search.customercontacts do
 Begin
  close;
  setvariable('CUSTID',custid);
  open;

  Custname:=fields[2].text;
  premiseCount:=0;
  password:=fields[12].text;
  passworddate:=fields[13].text;
  relationship:=fields[15].text;
  specialneeds:=fields[4].text;
  customertype:=fields[9].text;

  Customerdeceased:='';
  mailing:='';
  if fields[36].text<>'' then mailing:=mailing+fields[36].text+', ';
  if fields[37].text<>'' then mailing:=mailing+fields[37].text+', ';
  if fields[38].text<>'' then mailing:=mailing+fields[38].text+', ';
  if fields[39].text<>'' then mailing:=mailing+fields[39].text+', ';
  if fields[40].text<>'' then mailing:=mailing+fields[40].text+', ';
  if fields[41].text<>'' then mailing:=mailing+fields[41].text+', ';
  if fields[42].text<>'' then mailing:=mailing+fields[42].text+', ';
  if fields[43].text<>'' then mailing:=mailing+fields[43].text+', ';
  if fields[44].text<>'' then mailing:=mailing+fields[44].text+', ';
  mailing:=mailing+fields[45].text+'';
 // mynodeCustomer.imageindex:=36; // Customer
 // mynodeCustomer.selectedindex:=36;
  nodedata.caption:='Customer '+custid+' - '+custname + GetCustomerPronoun(EmptyStr, Cust);

  if lib_id<>'' then
  begin
   nodedata.caption:=nodedata.caption+' -(use '+lib_id+' in Liberty)';
  end;

  if Cdebt<>'' then
  Begin
   nodedata.fontcolor:=clmaroon;
   nodedata.fontBold:=true;
   nodedata.caption:=nodedata.caption+#10+CDEBT;
  End;

  if mailing='' then mailing:='Mailing Address NOT SPECIFIED';
  nodedata.caption:=nodedata.caption+#10+mailing;

  nodedata.D_Customer_ID := Custid;
  nodedata.D_Customer_Name := Custname;
  nodedata.D_LIB_CUST_ID:=LIB_ID;
  nodedata.D_CDEBT:=Cdebt;
  nodedata.D_premiseCount :=inttostr(premiseCount);
  //mynodecustomer.data:=MyRecPtr;

  {// Show Multi premise Customer
  if premiseCount>1 then
  Begin
   mynodecustomer.text:='Customer '+custid+' - '+custname+' - ('+inttostr(premiseCount)+' premises)';
   mynodeCustomer.selectedindex:=40;
   mynodeCustomer.Imageindex:=40;
  end; }

  // Is Prospect?
  If Fields[47].Text = 'Y' then
    Begin
      // SJ-BSL - 02/05/2021 - Replacing constant assignment by Global Variable.
      If ((NodeData.Index >= 306) and (NodeData.Index <= 308)) or (NodeData.Index = 131) then // 83
        Nodedata.Index := 131
      Else
        Nodedata.Index := 36;
    End;

  { // Check for Deceased Customer
  if Customerdeceased<>'' then
  Begin
   if mynodeCustomer.selectedindex=36 then
   Begin
    mynodeCustomer.selectedindex:=46;
    mynodeCustomer.Imageindex:=46;
   end
   else
   Begin
    mynodeCustomer.selectedindex:=47;
    mynodeCustomer.Imageindex:=47;
   end;
   mynodeCustomer.text:=mynodecustomer.text+' - (Customer Deceased '+CustomerDeceased+')';
   P_C_Rev.visible:=true;
   P_C_Dec.visible:=false;
  end;
  }

  // Check for Future Mailing Address e.g. Customer Move
  with main_data_module.generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('CID', otstring);
   sql.clear;
   sql.add('Select M.effective_from,');
   sql.add('p.premise_line_1,');
   sql.add('p.premise_line_2,');
   sql.add('p.premise_line_3,');
   sql.add('p.premise_line_4,');
   sql.add('p.premise_line_5,');
   sql.add('p.premise_line_6,');
   sql.add('p.premise_line_7,');
   sql.add('p.premise_line_8,');
   sql.add('p.premise_line_9,');
   sql.add('p.premise_Postcode');
   sql.add(' from crm.customer_mailing_history M, crm.premises p where M.customer_id=:CID');
   sql.add('and M.effective_from>sysdate');
   sql.add('and m.mailing_address_id=p.premise_id (+)');
   sql.add('order by M.effective_from');
   setvariable('CID',CUSTID);
   open;
   deletevariables;
   if recordcount<>0 then
   Begin
    Fad:='';
    if fields[1].text<>'' then Fad:=Fad+fields[1].text+', ';
    if fields[2].text<>'' then Fad:=Fad+fields[2].text+', ';
    if fields[3].text<>'' then Fad:=Fad+fields[3].text+', ';
    if fields[4].text<>'' then Fad:=Fad+fields[4].text+', ';
    if fields[5].text<>'' then Fad:=Fad+fields[5].text+', ';
    if fields[6].text<>'' then Fad:=Fad+fields[6].text+', ';
    if fields[7].text<>'' then Fad:=Fad+fields[7].text+', ';
    if fields[8].text<>'' then Fad:=Fad+fields[8].text+', ';
    if fields[9].text<>'' then Fad:=Fad+fields[9].text+', ';
    Fad:=Fad+fields[10].text+'';
    {mynode1:=Treeview1.items.AddChild(mynodeCustomer,'Forwarding Address - '+Fad+' as of  '+fields[0].text);
    mynode1.imageindex:=141;
    mynode1.selectedindex:=141;
    mynode1.font.color:=clmaroon;
    mynode1.Font.style:=[fsbold];}


    mynode1:=Treeview1.Addchild(mynodeCustomer);
    nodeData := Treeview1.GetNodeData(mynode1);
    NodeData.caption := 'Forwarding Address - '+Fad+' as of  '+fields[0].text;
    nodedata.index:=141;
    nodedata.fontcolor:=clmaroon;
    nodedata.fontBold:=true;

   end;
  End;
    // or 34 for an unhappy customer
  // Format Customer Address

  BuildCustomerNotifications(CustId);
  BuildCustomerFuelDirect(CustId); // BSL - 13/05/2015 - Fuel Direct Execute.
  BuildCustomerLosses(Custid);
  BuildCustomerNotes(Custid);



  {  if customertype<>'' then
  Begin
  mynode1:=Treeview1.items.AddChild(CustomerDetailsNode,'Customer Type - '+customertype);
  mynode1.imageindex:=49; // Customer Type
  mynode1.selectedindex:=49;
  end; }
{  if relationship<>'' then
  Begin
   mynode1:=Treeview1.items.AddChild(Mynodecustomer,'Relationship - '+relationship);
   mynode1.imageindex:=70; // RelationShip
   mynode1.selectedindex:=70;
  end;   }

  // Check for Statement Reviewr
  with main_data_module.generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('CID', otstring);
   sql.clear;
   sql.add('Select * from crm.customer_statement_reviewer where customer_id=:CID');
   setvariable('CID',custid);
   open;
   deletevariables;
   if recordcount<>0 then
   Begin
    {mynode1:=Treeview1.items.AddChild(mynodeCustomer,'Statement Reviewer Requested by - '+fields[1].text+' on '+fields[2].text);
    mynode1.imageindex:=136;
    mynode1.selectedindex:=136;
    mynode1.font.color:=clpurple;
    mynode1.font.style:=[fsbold];}

    mynode1:=Treeview1.Addchild(mynodeCustomer);
    nodeData := Treeview1.GetNodeData(mynode1);
    NodeData.caption := 'Statement Reviewer Requested by - '+fields[1].text+' on '+fields[2].text;
    nodedata.index:=136;
    nodedata.fontcolor:=clpurple;
    nodedata.fontBold:=true;

    C1.caption:='Remove Statement Reviewer';
    End
   else C1.caption:='Add Statement Reviewer';
  End;

  with main_data_module.generalquery do
  Begin
   Close;
   DeleteVariables;
   DeclareVariable(':RESULTS', otcursor);
   sql.clear;
   sql.Add('begin');
   sql.Add('CRM.PK_UTILITIES.PR_RET_REFUSED_HH_DATA('''+CUSTID+''',:RESULTS);');
   sql.Add('end;');
   open;
   deletevariables;
   if (recordcount<>0) then
   Begin
    mynode1:=Treeview1.AddChild(mynodeCustomer);
    nodeData := Treeview1.GetNodeData(mynode1);
    NodeData.caption := fields[0].Text;
    NodeData.index:=92;
    NodeData.fontcolor:=clRed;
    NodeData.fontBold:=true;
   End;
  end;

  if specialneeds='T' then
  Begin
   {mynode1:=Treeview1.items.AddChild(mynodeCustomer,'Special Needs - '+SpecialNeeds);
   mynode1.imageindex:=15; // Special Needs
   mynode1.selectedindex:=15;
   mynode1.font.color:=clpurple;
   mynode1.font.style:=[fsbold]; }

   mynode1:=Treeview1.Addchild(mynodeCustomer);
   nodeData := Treeview1.GetNodeData(mynode1);
   NodeData.caption := 'Special Needs - '+SpecialNeeds;
   nodedata.index:=15;
   nodedata.fontcolor:=clpurple;
   nodedata.fontBold:=true;


  end;


  if Password<>'' then
  Begin
   {mynode1:=Treeview1.items.AddChild(myNodecustomer,'Password-      '+password+'      Effective From ('+PasswordDate+')');
   mynode1.imageindex:=7;
   mynode1.selectedindex:=7;
   mynode1.font.color:=clpurple;
   mynode1.font.style:=[fsbold]; }

   mynode1:=Treeview1.Addchild(mynodeCustomer);
   nodeData := Treeview1.GetNodeData(mynode1);
   NodeData.caption := 'Password-      '+password+'      Effective From ('+PasswordDate+')';
   nodedata.index:=7;
   nodedata.fontcolor:=clpurple;
   nodedata.fontBold:=true;


  end;

  no_of_holders:=FRM_main_search.customercontacts.recordcount;

 end;
 prevah:='-1';
 TelNum := 0;
 EmailNum := 0;

 with FRM_main_search.customercontacts do
 Begin
  for f:=1 to no_of_holders do
   Begin
    AH_ID:=fields[16].text;
    AH_CONTACT_TITLE_ID:=fields[23].text;
    AH_INITIALS:=fields[25].text;
    AH_SURNAME:=fields[26].text;
    AH_FORENAME:=fields[27].text;
    AH_DISPLAY_NAME:=fields[28].text;
    AH_ADDITIONAL_INFORMATION:=fields[29].text;
    AH_SPECIAL_NEEDS_INFORMATION:=fields[30].text;
    AH_TELEPHONE_NO_DAY:=fields[31].text;
    AH_TELEPHONE_NO_EVE:=fields[32].text;
    AH_TELEPHONE_NO_MOBILE:=fields[33].text;
    AH_EMAIL:=fields[34].text;

    if (AH_TELEPHONE_NO_DAY <> '') or (AH_TELEPHONE_NO_EVE <> '') or (AH_TELEPHONE_NO_MOBILE <> '') or (Fields[46].Text = '3') then inc(TelNum);
    if (AH_EMAIL <> '') or (Fields[46].Text = '3') then Inc(EmailNum);

    AH_FAX:=fields[35].text;
    AH_order:=fields[20].text;
    AH_DOB:=fields[48].text;
    AH_contact_method:=fields[49].text;
    ah_TYPE:='Contact';
    if fields[19].text='P' then AH_TYPE:='Primary';
    if fields[19].text='E' then AH_TYPE:='Emergency';

    if (ah_id<>'') and (ah_id<>prevah) then
    Begin
     if (f=3) and (no_of_holders>3) then
     Begin
      {mynode1:=Treeview1.items.AddChild(mynodeCustomer,'Additional Account Holders ['+inttostr(no_of_holders-2)+']');
      mynode1.imageindex:=74;
      mynode1.selectedindex:=74; }

      mynode1:=Treeview1.Addchild(mynodeCustomer);
      nodeData := Treeview1.GetNodeData(mynode1);
      NodeData.caption := 'Additional Account Holders ['+inttostr(no_of_holders-2)+']';
      nodedata.index:=74;

     end;

   //   premiseContactNode:=Treeview1.items.AddChild(mynode1,'('+ah_order+') - '+ah_type+' - '+ah_display_name);
   //   premiseContactNode.imageindex:=38;
   //   premiseContactNode.selectedindex:=38;

   Contdet:='';
   if AH_CONTACT_TITLE_ID<>'' then
   Begin
     Contdet:=ah_contact_title_id+' ';
   end;

   if AH_INITIALS<>'' then
   Begin
    contdet:=contdet+'('+ah_initials+') ';
   end;

   if AH_FORENAME<>'' then
   Begin
    contdet:=contdet+ah_Forename+' ';
   end;

   if AH_SURNAME<>'' then
   Begin
    contdet:=contdet+ah_surname+' ';
   end;


   if AH_DISPLAY_NAME<>'' then
   Begin
    Contdet:=contdet+'. Known as - '+ah_display_name;
    contdet:=ah_display_name;
   end;
   if contdet<>'' then
   Begin
    // First two account holders hang of customer
    if ((fields[62].text<>'0') and (fields[62].text<>'')) or  (fields[56].text<>'') or (fields[64].text<>'') or (fields[65].text<>'') then hasneeds:=true
    else hasneeds:=false;

    if (f<3) then
    Begin
    // premiseContactNode:=Treeview1.items.AddChild(mynodecustomer,'Account Holder - '+ah_type+' - '+contdet);
     premiseContactNode:=Treeview1.Addchild(mynodeCustomer);
     nodeData := Treeview1.GetNodeData(premiseContactNode);
     NodeData.caption := 'Account Holder - '+ah_type+' - '+contdet + GetCustomerPronoun(FRM_main_search.CustomerContacts.FieldByName('CONTACT_ID').AsString, custid);
     nodedata.index:=74;
     if hasneeds=true then
     begin
      nodedata.fontcolor:=clblue;
      nodedata.fontBold:=true;
     end;


    end
    else
    // 3rd account holder will hang of tree if only 3 account holders
    if (f=3) and (no_of_holders=3) then
    Begin
     //premiseContactNode:=Treeview1.items.AddChild(mynodecustomer,'Account Holder - '+ah_type+' - '+contdet);
     premiseContactNode:=Treeview1.Addchild(mynodeCustomer);
     nodeData := Treeview1.GetNodeData(premiseContactNode);
     NodeData.caption := 'Account Holder - '+ah_type+' - '+contdet + GetCustomerPronoun(FRM_main_search.CustomerContacts.FieldByName('CONTACT_ID').AsString, custid);
      if hasneeds=true then
     begin
      nodedata.fontcolor:=clblue;
      nodedata.fontBold:=true;
     end;
    End
    else
    // more than 3 account holders get rolled up
    Begin
     //premiseContactNode:=Treeview1.items.AddChild(mynode1,'('+ah_order+') - '+ah_type+' - '+contdet);
     premiseContactNode:=Treeview1.Addchild(mynode1);
     nodeData := Treeview1.GetNodeData(premiseContactNode);
     NodeData.caption := 'Account Holder - '+ah_type+' - '+contdet + GetCustomerPronoun(FRM_main_search.CustomerContacts.FieldByName('CONTACT_ID').AsString, custid);
     if hasneeds=true then
     begin
      nodedata.fontcolor:=clblue;
      nodedata.fontBold:=true;
     end
    end;

    // Create Account Holder Node
    // Check for deceased accountholder
    if fields[24].value <> null then
    nodedata.index:=fields[24].value
    else  nodedata.index := 0;

    if fields[46].text<>'1' then
    Begin
    end;
    if fields[46].text='3' then
    Begin
     nodedata.index:=43;
    end;
    //premiseContactNode.selectedindex:=premisecontactnode.imageindex;

    nodedata.D_customer_id :=FRM_main_search.customercontacts.fields[0].text;
    nodedata.D_Account_holder_id :=FRM_main_search.customercontacts.fields[16].text;
    nodedata.D_contact_id :=FRM_main_search.customercontacts.fields[18].text;
    //premiseContactNode.data:=MyRecPtr;
   end;

      // Nominated Contact
   if fields[56].text<>'' then
   Begin
    desc:='Nominated Contact - ('+fields[56].text+') - ';
    if fields[52].Text<>'' then Desc:=desc+fields[52].Text+' ';
    if fields[53].Text<>'' then Desc:=desc+fields[53].Text;
    if fields[54].Text<>'' then Desc:=desc+'. Tel: '+fields[54].Text;
    if fields[59].Text='I' then Desc:=desc+#13+'Consent Type: INFORMATION ONLY';
    if fields[59].Text='F' then Desc:=desc+#13+'Consent Type: FULL';

    premiseContactitemNode:=Treeview1.Addchild(PremiseContactNode);
    nodeData := Treeview1.GetNodeData(premiseContactitemNode);
    NodeData.caption :=desc;
    NodeData.index:=87;
   end;

     // Vulnerability
   if fields[65].text<>'' then
   Begin
    msg:=fields[65].text;
    if length(MSG)>100 then msg:=copy(fields[65].Text,1, 100)+'...';
    premiseContactitemNode:=Treeview1.Addchild(PremiseContactNode);
    nodeData := Treeview1.GetNodeData(premiseContactitemNode);
    if fields[61].Text<>'' then  msg:=msg+#13+'Additional Information: '+fields[61].Text;
    NodeData.caption :='Vulnerabilities - '+MSG;


    NodeData.index:=72;
   end;

   // Suggested Support
   if fields[64].text<>'' then
   Begin
    msg:=fields[64].text;
    if length(MSG)>100 then msg:=copy(fields[64].Text,1, 100)+'...';
    premiseContactitemNode:=Treeview1.Addchild(PremiseContactNode);
    nodeData := Treeview1.GetNodeData(premiseContactitemNode);
    NodeData.caption :='Suggested Support - '+MSG;
    NodeData.index:=140;
   end;

     // Children Under 5
   if ((fields[62].text<>'0') and (fields[62].text<>'')) then
   Begin
    premiseContactitemNode:=Treeview1.Addchild(PremiseContactNode);
    nodeData := Treeview1.GetNodeData(premiseContactitemNode);
    if fields[62].text<>'1' then NodeData.caption :=fields[62].text+' Children Under 5, Youngest DOB '+fields[63].text
    else NodeData.caption :=fields[62].text+' Child Under 5. DOB '+fields[63].text;
    NodeData.index:=274;
   end;

     // Sharing with Newtwork Operators
   if (fields[57].text='N') and (fields[58].text<>'') then
   Begin
    premiseContactitemNode:=Treeview1.Addchild(PremiseContactNode);
    nodeData := Treeview1.GetNodeData(premiseContactitemNode);
    NodeData.caption :='DATA SHARING WITH NETWORK OPERATORS REFUSED';
    NodeData.fontcolor:=clred;
    NodeData.index:=78;
   end;

   // Additional Information
   if AH_ADDITIONAL_INFORMATION<>'' then
   Begin
    premiseContactitemNode:=Treeview1.Addchild(PremiseContactNode);
    nodeData := Treeview1.GetNodeData(premiseContactitemNode);
    NodeData.caption := ah_additional_information;
    NodeData.index:=73;
   end;

   Telno:='';
   if AH_TELEPHONE_NO_DAY<>'' then
   Begin
    telno:=telno+'Tel No: Day - '+ah_Telephone_no_Day+'     ';
   end;

   if AH_TELEPHONE_NO_EVE<>'' then
   Begin
    telno:=telno+'Tel No: Eve - '+ah_Telephone_no_eve;
   end;
   if telno<>'' then
   Begin
   { premiseContactitemNode:=Treeview1.items.AddChild(premisecontactnode,telno);
    premiseContactitemNode.imageindex:=2;
    premiseContactitemNode.selectedindex:=2;  }

     premiseContactitemNode:=Treeview1.Addchild(PremiseContactNode);
    nodeData := Treeview1.GetNodeData(premiseContactitemNode);
    NodeData.caption :=telno;
    NodeData.index:=2;

   end;

   if AH_TELEPHONE_NO_MOBILE<>'' then
   Begin
   { premiseContactitemNode:=Treeview1.items.AddChild(premisecontactnode,'Tel No: Mobile - '+ah_Telephone_no_Mobile);
    premiseContactitemNode.imageindex:=14;
    premiseContactitemNode.selectedindex:=14; }

      premiseContactitemNode:=Treeview1.Addchild(PremiseContactNode);
    nodeData := Treeview1.GetNodeData(premiseContactitemNode);
    NodeData.caption :='Tel No: Mobile - '+ah_Telephone_no_Mobile;
    Nodedata.d_tel:=ah_Telephone_no_Mobile;
    nodedata.D_customer_id :=FRM_main_search.customercontacts.fields[0].text;
    NodeData.index:=14;
   end;

   if AH_FAX<>'' then
   Begin
   { premiseContactitemNode:=Treeview1.items.AddChild(premisecontactnode,'Fax - '+ah_Fax);
    premiseContactitemNode.imageindex:=11;
    premiseContactitemNode.selectedindex:=11; }

        premiseContactitemNode:=Treeview1.Addchild(PremiseContactNode);
    nodeData := Treeview1.GetNodeData(premiseContactitemNode);
    NodeData.caption :='Fax - '+ah_Fax;
    NodeData.index:=11;
   end;

   if AH_EMAIL<>'' then
   Begin
    {premiseContactitemNode:=Treeview1.items.AddChild(premisecontactnode,'Email - '+ah_email);
    premiseContactitemNode.imageindex:=20;
    premiseContactitemNode.selectedindex:=20;
    premiseContactitemNode.font.color:=clblue;
    premiseContactitemNode.font.style:=[fsunderline];
    }

    premiseContactitemNode:=Treeview1.Addchild(PremiseContactNode);
    nodeData := Treeview1.GetNodeData(premiseContactitemNode);
    NodeData.caption :='Email - '+ah_email;
    Nodedata.d_email:=ah_email;
    Nodedata.D_Customer_Id:=FRM_main_search.customercontacts.fields[0].text;
    NodeData.index:=20;
    NodeData.fontcolor:=clblue;
    NodedAta.fontUnderline:=true;

   end;

   if AH_DOB>'' then
   Begin
    premiseContactItemNode:=Treeview1.AddChild(premisecontactnode);
    nodeData := Treeview1.GetNodeData(premiseContactItemNode);
    NodeData.caption := 'Date Of Birth - '+ah_DOB;
    NodeData.index:=63;
   end;

   if AH_contact_method<>''  then
   begin
    premiseContactItemNode:=Treeview1.AddChild(premisecontactnode);
    nodeData := Treeview1.GetNodeData(premiseContactItemNode);
    NodeData.caption := 'Preferred Contact: - '+uppercase(AH_contact_method);
    NodeData.index:=271;
   end;

   end;
   prevah:=ah;
  next;
  end;
 end;


 // Get Prospect Details For Customer
 with Generalquery do
 Begin
  close;
  DeleteVariables;
  DeclareVariable('CID', otstring);
  sql.clear;
  sql.add('Select * from crm.prospect_details where customer_id=:CID');
  sql.add('order by date_created desc');
  setvariable('CID',custid);
  open;
  deletevariables;
 End;
 if generalquery.recordcount<>0 then
 Begin
  Desc:='Prospect Details Exist - Last updated '+Generalquery.fields[8].text;
{  mynodeProsp:=Treeview1.items.AddChild(mynodeCustomer,desc);
  mynodeprosp.imageindex:=73;
  mynodeprosp.selectedindex:=mynodeprosp.imageindex; }

  mynodeProsp:=Treeview1.Addchild(mynodeCustomer);
  nodeData := Treeview1.GetNodeData(mynodeProsp);
  NodeData.caption :=desc;
  NodeData.index:=73;
  nodedata.D_customer_Id := custid;
  //mynodeprosp.data:=MyRecPtr;
 end;


 if FRM_Tree.tag<2 then
 Begin
  if FRM_main_search.customercontacts.fields[1].text<>'5' then   // Dont Show is a Sales Agent
  Begin
   // Get All Agreements For Customer
   generalquery := TCrmUtil.GetAllAgreementsForCustomer(StrToInt64(custid), ((hidecheck.Visible = true) and (hidecheck.checked = true)));

   while not generalquery.eof do
   Begin
    agreement_status_id:=generalquery.fields[6].text;
    agreement_status:=generalquery.fields[7].text;
    Agreement_Start_Date:=generalquery.fields[3].text;
    Agreement_End_Date:=generalquery.fields[4].text;
    Agreement_id:=generalquery.fields[0].text;

    desc:='';
   // mynodeAgreement:=Treeview1.items.AddChild(mynodeCustomer,desc);

    mynodeAgreement:=Treeview1.Addchild(mynodeCustomer);
    nodeData := Treeview1.GetNodeData(mynodeAgreement);
    Nodedata.caption:=desc;


    ShowAgreement(mynodeagreement,generalquery.fields[9].value,Agreement_id,custid,Agreement_status,agreement_start_Date,Agreement_end_date,generalquery.fields[8].text,
      false,generalquery.fields[10].text);
    Status:= TFinancialHistoryInfo.GetFinancialStatusText(StrToInt64(Agreement_id));  //FRM_Financial_History.GetStatus(Agreement_id);
    ShowProduct(mynodeagreement,Agreement_id,false,status);

    generalquery.next;
   end;
  end
  else
  Begin // Sales Agent Tab
   // Get All Agreements For Customer
   generalquery := TCrmUtil.GetAllAgreementsForCustomer(StrToInt64(custid));

   while not generalquery.eof do
   Begin
    Agreement_Start_Date:=generalquery.fields[3].text;
    Agreement_id:=generalquery.fields[0].text;
    Desc:='Agreement - None (Sales Agent) - Date Started '+agreement_start_date;
    {mynodeAgreement:=Treeview1.items.AddChild(mynodeCustomer,desc);
    mynodeagreement.imageindex:=111;
    mynodeagreement.selectedindex:=mynodeagreement.imageindex; }

    mynodeAgreement:=Treeview1.Addchild(mynodeCustomer);
    nodeData := Treeview1.GetNodeData(mynodeAgreement);
    Nodedata.caption:=desc;
    Nodedata.index:=111;
    nodedata.D_agreement_id := agreement_id;
    nodedata.D_Customer_id := custid;

    //Mynodeagreement.data:=MyRecPtr;
    //mynodeAgreementItem:=Treeview1.items.AddChild(mynodeagreement,'dummy');
    mynodeAgreementItem:=Treeview1.Addchild(mynodeagreement);
    nodeData := Treeview1.GetNodeData(mynodeAgreementItem);
    Nodedata.caption:='Dummy';

    generalquery.next;
   End;
  end;
 end;

 if ((TelNum <> no_of_holders) or (EmailNum <> no_of_holders)) and (dpanag = true) and (not (fCustomerAccountHolder = CustId)) then
  begin
    fCustomerAccountHolder := CustId;
    Messagedlg('Some contact details appear to be missing, please ensure we have TELEPHONE numbers and EMAIL addresses recorded for all account holders and contacts',mtWarning,[mbok],0);
  end;

 if FRM_Tree.tag=0 then
   exit;
  // Show Any Super Customer Details
  With main_data_module.tempquery do
  Begin
   Close;
   DeleteVariables;
   DeclareVariable('CID', otString);
   SetVariable('CID', CustId);
   // SJ-BSL - 30/04/2021 - Add Icon Variables
   SQL.Text :=
      'Select sc.customer_id, sc.legal_entity_name, sc.primary_mailing_address_id, scp.premise_line_1, scp.premise_line_2, scp.premise_line_3, ' +
              'scp.premise_line_4, scp.premise_line_5, scp.premise_line_6, scp.premise_line_7, scp.premise_line_8, scp.premise_line_9, scp.premise_postcode, ' +
              'csc.super_customer_id, COALESCE (cri.icon_index, ci.icon_index) icon_index1 ' +
      'From crm.customer sc, crm.premises scp, crm.customer_to_super_customer csc, crm.customer_type ci, crm.cust_type_relationship_icon cri ' +
      'Where sc.primary_mailing_address_id = scp.premise_id ' +
             'and sc.customer_id = csc.super_customer_id ' +
             'and sc.customer_type_id = ci.customer_type_id (+) ' +
             'and sc.customer_type_id = cri.customer_type_id (+) ' +
             'and sc.relationship_rating_id = cri.relationship_rating_id (+) ' +
             'and csc.customer_id = :CID ';

//   sql.clear;
//   sql.add('select');
//   sql.add('sc.customer_id,');
//   sql.add('sc.legal_entity_name,');
//   sql.add('sc.primary_mailing_address_id,');
//   sql.add('scp.premise_line_1,');
//   sql.add('scp.premise_line_2,');
//   sql.add('scp.premise_line_3,');
//   sql.add('scp.premise_line_4,');
//   sql.add('scp.premise_line_5,');
//   sql.add('scp.premise_line_6,');
//   sql.add('scp.premise_line_7,');
//   sql.add('scp.premise_line_8,');
//   sql.add('scp.premise_line_9,');
//   sql.add('scp.premise_postcode, ');
//   sql.add('csc.super_customer_id from');
//   sql.add('crm.customer sc,');
//   sql.add('crm.premises scp,');
//   sql.add('crm.customer_to_super_customer csc');
//   sql.add('where');
//   sql.add('sc.primary_mailing_address_id=scp.premise_id');
//   sql.add('and');
//   sql.add('sc.customer_id=csc.super_customer_id and csc.customer_id=:CID') ;
   Open;
   DeleteVariables;
  end;

  if main_data_module.tempquery.recordcount<>0 then
  Begin
   cust:=main_data_module.tempquery.fields[0].text;
   custname:=main_data_module.tempquery.fields[1].text;
   mailing:=main_data_module.tempquery.fields[2].text;
   Fad:='';
   if main_data_module.tempquery.fields[3].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[3].text+', ';
   if main_data_module.tempquery.fields[4].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[4].text+', ';
   if main_data_module.tempquery.fields[5].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[5].text+', ';
   if main_data_module.tempquery.fields[6].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[6].text+', ';
   if main_data_module.tempquery.fields[7].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[7].text+', ';
   if main_data_module.tempquery.fields[8].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[8].text+', ';
   if main_data_module.tempquery.fields[9].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[9].text+', ';
   if main_data_module.tempquery.fields[10].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[10].text+', ';
   if main_data_module.tempquery.fields[11].text<>'' then Fad:=Fad+main_data_module.tempquery.fields[11].text+', ';
   Fad:=Fad+main_data_module.tempquery.fields[12].text+'';
   mailing:=fad;
   premisecount:=0;
   customerdeceased:='0';
   FCustIcon := 214; //  SJ-BSL - 30/04/2021 - Change 214 Icon by 309..311 (Changed back 07/05/2021) SJ
   debt:='0';
   // Now add a sub node if theres a super customer
   if issuper=false then
   Begin
    //mynodeSuperCustomer:=treeview1.items.Addchild(mynodecustomer, 'Super Customer '+main_data_module.tempquery.fields[13].text+' - '+Custname);

    mynodeSuperCustomer:=Treeview1.Addchild(mynodeCustomer);
    nodeData := Treeview1.GetNodeData(mynodeSuperCustomer);
    Nodedata.caption:='Super Customer '+main_data_module.tempquery.fields[13].text+' - '+Custname;


    ShowSuperCustomerOnly(mynodeSuperCustomer,'Y','N',main_data_module.tempquery.fields[13].text,custname,mailing,inttostr(premisecount),customerdeceased,'0', debt);
   end;
  end;
end;
{
procedure TFRM_Tree.GetCustomerAgreements(const custId: string; agreeList: TStringList);
begin
  // Get All LIVE Agreements For Customer
   with generalquery do
  Begin
    close;
    DeleteVariables;
    DeclareVariable('CID', otstring);
    sql.clear;
    sql.Add('select');
    sql.add('A.AGREEMENT_ID ');
    sql.add('from CRM.Agreements A, crm.agreement_status S,crm.agreement_renewal_dates RD,');
    sql.add('crm.agreement_type T,CRM.initial_periods IP');
    sql.add('where A.customer_id=:CID');
    sql.add('and a.agreement_status_id=s.agreement_status_id (+)');
    sql.add('and a.agreement_type_id=T.agreement_type_id (+)');
    sql.add('and a.initial_period=ip.initial_period_id (+)');
    sql.add('and a.agreement_id=rd.agreement_id (+)');
    sql.Add(' and a.agreement_id in (select agreement_id from (select agreement_id,agreement_status_id,agreement_start_date');
    sql.Add(' from  crm.agreements');
    sql.Add(' where  customer_id=:CID');
    sql.Add(' and  agreement_status_id = 1');
    sql.Add(' union all');
    sql.Add(' select ca.agreement_id,ca.agreement_status_id,ca.agreement_start_date');
    sql.Add(' from   crm.agreements ca');
    sql.Add(' where  ca.customer_id=:CID');
    sql.Add(' and    not exists (select ''y'' from crm.agreements ca1 where ca1.CUSTOMER_ID = ca.customer_id and ca1.AGREEMENT_STATUS_id <> 3))) ');

    sql.add('order by A.agreement_id desc');
    setvariable('CID',custId);
    open;
    deletevariables;
  end;

  if generalQuery.RecordCount > 0 then
  begin
    generalQuery.First;
    while not(generalQuery.EOF) do
    begin
      agreeList.Add(generalQuery.FieldByName('AGREEMENT_ID').AsString);
      generalQuery.Next;
    end;
  end;
end;
}

{------------------------------------------------------------------------------}
Procedure TFRM_Tree.Refreshpremisenode(MyNodePremise: PVirtualNode);
Var
status,customerdeceased,btssd,debt,customer_id:string;
spanindex: Integer;
  Related,ISDNNode:boolean;

  s           : string;
  agreementId : variant;
  premiseId   : variant;
begin
  if not Assigned(MyNodePremise) then
    exit;

   // Anna: nodeData must be the global variable here as it is used by a called procedure
   // and changing them to a local one/parameter might be a larger task
  nodeData := TreeView1.GetNodeData(MyNodePremise);
  if not Assigned(nodeData) then
    exit;

  s := nodeData.D_Agreement_Id;
  if Trim(s) = '' then
    agreementId := null
  else
    agreementId := StrToInt64(s);

  s := nodeData.D_Premise_Id;
  if Trim(s) = '' then
    premiseId := null
  else
    premiseId := StrToInt64(s);

  customer_id := nodeData.D_Customer_Id;
  if Trim(customer_id) = '' then
  begin
    if VarIsNull(agreementId) then
      customer_id := ''
    else
      customer_id := Frm_Common.GetCustomerIdFromAgreementId(agreementId);
  end;

  if TBillingUtil.UseDllRerating then
    TRateAccountsWrapper.SetMoveOuts(agreementId, premiseId, TRANSACTION_YES)
  else
    TCustomerCommon.SetMoveOuts(agreementId, premiseId);

 try
  //treeview1.selected.deletechildren;
  treeview1.deletechildren(mynodepremise);
 except
 end;
// Get Customer Type
 with GeneralQuery Do
 begin
  close;
  DeleteVariables;
  DeclareVariable(':CUSTID',otLong);
  sql.clear;
  sql.add('SELECT');
  sql.add('C.CUSTOMER_TYPE_ID ');
  sql.add('FROM');
  sql.add('  CRM.CUSTOMER C ');
  sql.add('WHERE');
  sql.add('  C.CUSTOMER_ID =:CUSTID');
  SetVariable('CUSTID',StrToInt64(Customer_ID));
  open;
  deletevariables;
 end;
 if Generalquery.recordcount<>0 then
 begin
   cust_type := generalquery.fields[0].AsInteger;
 end;

 // Get Premise Details From Aggreement
 // Site Name & Contact Details
 with GeneralQuery Do
 Begin
  close;
  DeleteVariables;
  DeclareVariable('AID', otlong);
  DeclareVariable('PID', otlong);
  sql.clear;
  sql.add('SELECT');
  sql.add('AP.PREMISE_NAME, AP.SPECIAL_ACCESS, CO.DISPLAY_NAME, CO.ADDITIONAL_INFORMATION, CO.TELEPHONE_NO_DAY,');
  sql.add(' CO.TELEPHONE_NO_EVE, CO.TELEPHONE_NO_MOBILE, CO.EMAIL, CO.FAX,AP.DATE_MOVED_OUT');
  sql.add('FROM');
  sql.add('  CRM.AGREEMENT_PREMISES AP,');
  sql.add('  CRM.CONTACTS CO');
  sql.add('WHERE');
  sql.add('  AP.PREMISE_CONTACT_ID = CO.CONTACT_ID(+)');
  sql.add('and ap.agreement_id=:AID');
  sql.add('and ap.premise_id=:PID');
  setvariable('AID',agreementId);
  setvariable('PID', premiseId);
  open;
  deletevariables;
 End;
 if Generalquery.recordcount<>0 then
 Begin
  if generalquery.fields[1].text<>'' then
  Begin
   desc:='Special Access - '+generalquery.fields[1].text;
   {PremDetailsNode:=Treeview1.items.AddChild(mynodepremise,desc);
   PremDetailsNode.imageindex:=8;
   PremDetailsNode.selectedindex:=premDetailsNode.imageindex;
   }

   PremDetailsNode:=Treeview1.Addchild(mynodepremise);
   nodeData := Treeview1.GetNodeData(PremDetailsNode);
   NodeData.caption := desc;
   NodeData.D_Cust_Type := cust_type;
   nodedata.index:=8;
   end;

 End;

 if tag<>4 then
 Begin
  ISDNNode:=false;
  // Get Services for premise/agreement
  with Generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('PID', otlong);
   sql.clear;
   sql.add('select');
   sql.add('S.service_id,');
   sql.add('T.Description,');
   sql.add('T.icon_index,');
   sql.add('O.description,');
   sql.add('S.service_type_id');
   sql.add('from crm.service s, crm.service_type T,crm.order_status O');
   sql.add('where S.premise_id=:PID');
   if frm_tree.tag<>3 then
   Begin
    DeclareVariable('AID', otlong);
    setvariable('AID',agreementId);
    sql.add('and S.agreement_id=:AID');
   end;
   sql.add('and s.service_type_id=t.service_type_id');
   sql.add('and S.order_status_id=o.order_status_id');
   sql.add('order by s.start_date');
   setvariable('PID',premiseId);
   open;
   deletevariables;
   while not generalquery.eof do
   Begin
    // Now Get Spans for each service
    // if only one span create one node
    service_id:=generalquery.fields[0].text;
    with Spanquery do
    Begin
     deletevariables;
     DeclareVariable('SID', otlong);
     setvariable('SID',service_id);
     close;
     sql.clear;
     sql.add('select');
     sql.Add('S.SERVICE_ID,');
     sql.add('T.ICON_INDEX,');
     sql.add('T.DESCRIPTION,');
     sql.add('MAX(S.REGISTRATION_ID),');
     sql.add('S.SPAN_START_DATE,');
     sql.add('S.SPAN_END_DATE,');
     sql.add('S.SPAN_END_REASON,');
     sql.add('S.SPAN,');
     sql.add('S.SPAN_ADDRESS_1,');
     sql.add('S.SPAN_ADDRESS_2,');
     sql.add('S.SPAN_ADDRESS_3,');
     sql.add('S.SPAN_ADDRESS_4,');
     sql.add('S.SPAN_ADDRESS_5,');
     sql.add('S.SPAN_ADDRESS_6,');
     sql.add('S.SPAN_ADDRESS_7,');
     sql.add('S.SPAN_ADDRESS_8,');
     sql.add('S.SPAN_ADDRESS_9,');
     sql.add('S.SPAN_POSTCODE,');
     sql.add('S.RELATED,');
     sql.add('S.SERVICE_PRIOITY_NEEDS,');
     sql.add('S.CURRENT_SERVICE_STATUS,');
     sql.add('O.DESCRIPTION,');
     sql.add('S.T_BT_ACCOUNT_NO,');
     sql.add('S.Current_service_status,');
     sql.add('S.G_TRANSPORTER,');
     sql.add('S.SPAN_TYPE_ID,');
     sql.add('S.LOCK_STATUS');
     sql.add('from crm.spans s,crm.span_type T,crm.order_status O where s.service_id=:SID');
     sql.add('and s.span_type_id=t.span_type_id (+)');
     sql.add('and S.order_status_id=o.order_status_id (+)');

     sql.add('and (s.registration_id,S.service_id,S.span) in');
     sql.add('(select max(registration_id),service_id,span');
     sql.add('from crm.spans');
     sql.add('where service_id=:SID');
     sql.add('group by service_id,span)');

     sql.add('group by');
     sql.Add('S.SERVICE_ID,');
     sql.add('T.ICON_INDEX,');
     sql.add('T.DESCRIPTION,');
     sql.add('S.SPAN_START_DATE,');
     sql.add('S.SPAN_END_DATE,');
     sql.add('S.SPAN_END_REASON,');
     sql.add('S.SPAN,');
     sql.add('S.SPAN_ADDRESS_1,');
     sql.add('S.SPAN_ADDRESS_2,');
     sql.add('S.SPAN_ADDRESS_3,');
     sql.add('S.SPAN_ADDRESS_4,');
     sql.add('S.SPAN_ADDRESS_5,');
     sql.add('S.SPAN_ADDRESS_6,');
     sql.add('S.SPAN_ADDRESS_7,');
     sql.add('S.SPAN_ADDRESS_8,');
     sql.add('S.SPAN_ADDRESS_9,');
     sql.add('S.SPAN_POSTCODE,');
     sql.add('S.RELATED,');
     sql.add('S.SERVICE_PRIOITY_NEEDS,');
     sql.add('S.CURRENT_SERVICE_STATUS,');
     sql.add('O.DESCRIPTION,');
     sql.add('S.T_BT_ACCOUNT_NO,');
     sql.add('S.Current_service_status,');
     sql.add('S.G_TRANSPORTER,');
     sql.add('S.SPAN_TYPE_ID,S.LOCK_STATUS');
     sql.add('order by s.related,s.span_start_date desc,s.span');
     open;
     deletevariables;
     oldspan:='lee';
     related:=false;

     // If More Than 1 Span on Service Then it must be a Related Service
     // Create a Realted MPAN Node
     if spanquery.recordcount>1 then
     Begin

      IF (SPANQUERY.FIELDS[25].TEXT='C') or
      (SPANQUERY.FIELDS[25].TEXT='6') or
      (SPANQUERY.FIELDS[25].TEXT='S') or
      (SPANQUERY.FIELDS[25].TEXT='9') or
      (SPANQUERY.FIELDS[25].TEXT='D') or
      (SPANQUERY.FIELDS[25].TEXT='E') or
      (SPANQUERY.FIELDS[25].TEXT='F') or
      (SPANQUERY.FIELDS[25].TEXT='G') or
      (SPANQUERY.FIELDS[25].TEXT='O') or
      (SPANQUERY.FIELDS[25].TEXT='P') or
      (SPANQUERY.FIELDS[25].TEXT='B') or
      (SPANQUERY.FIELDS[25].TEXT='3') or
      (SPANQUERY.FIELDS[25].TEXT='1') or
      (SPANQUERY.FIELDS[25].TEXT='') then
      Begin
       {Servicenode:=Treeview1.items.AddChild(mynodepremise,'Related MPAN Service');
       servicenode.imageindex:=130;  }

       Servicenode:=Treeview1.Addchild(mynodepremise);
       nodeData := Treeview1.GetNodeData(Servicenode);
       NodeData.caption := 'Related MPAN Service';
       nodedata.index:=130;


      end
      else
      IF (SPANQUERY.FIELDS[25].TEXT='A') then
      Begin
       {Servicenode:=Treeview1.items.AddChild(mynodepremise,'Related MPRN Service');
       servicenode.imageindex:=203;  }

       Servicenode:=Treeview1.Addchild(mynodepremise);
       nodeData := Treeview1.GetNodeData(Servicenode);
       NodeData.caption := 'Related MPAN Service';
       NodeData.D_Cust_Type := cust_type;
       nodedata.index:=203;

      end
      else
      Begin
       {Servicenode:=Treeview1.items.AddChild(mynodepremise,'Telecoms Number Change');
       servicenode.imageindex:=183; }

       Servicenode:=Treeview1.Addchild(mynodepremise);
       nodeData := Treeview1.GetNodeData(Servicenode);
       NodeData.caption := 'Telecoms Number Change';
       NodeData.D_Cust_Type := cust_type;
       nodedata.index:=183;

      End;
      //servicenode.selectedindex:=servicenode.imageindex;
      related:=true;
     End;

     // Check if an ISDN Service
     if (spanquery.fields[25].text='W') or (spanquery.fields[25].text='H') then
     Begin
      if (ISDNNode=false) then
      Begin
       {Servicenode:=Treeview1.items.AddChild(mynodepremise,'ISDN Service');
       servicenode.imageindex:=134;
       servicenode.SelectedIndex:=134; }

       Servicenode:=Treeview1.Addchild(mynodepremise);
       nodeData := Treeview1.GetNodeData(Servicenode);
       NodeData.caption := 'ISDN Service';
       NodeData.D_Cust_Type := cust_type;
       nodedata.index:=134;


       ISDNNode:=true;
      End;
     end;

     while not spanquery.eof do
     Begin
      if spanquery.fields[7].text<>oldspan then
      Begin
       spantype:='SPAN ';
       //fields[2].text;
        if Generalquery.fields[2].text<>'' then spanindex:=Generalquery.fields[2].value
        else spanindex:=1;

       if (copy(spanquery.fields[23].text,11, 10)<>'Y') and (copy(spanquery.fields[23].text,11, 10)<>'') then
       Begin
       try
        strtodate(copy(spanquery.fields[23].text,11, 10));
        BTSSD:=copy(spanquery.fields[23].text,11, 10)
       except
        BTSSD:='';
       end;
       end
       else BTSSD:='';



       // ISDN Node
       if (spanquery.fields[25].text='W') or (spanquery.fields[25].text='H') then
       Begin
        if related=false then
        Begin
         //spannode:=Treeview1.items.AddChild(servicenode,'test');
         spannode:=Treeview1.Addchild(servicenode);
         nodeData := Treeview1.GetNodeData(spannode);
         NodeData.caption := 'Test';
         NodeData.D_Cust_Type := cust_type;
         NodeData.D_Service_ID := service_id;
        end;
        if spanquery.fields[25].text='W' then nodedata.caption:=spanquery.fields[2].text;
       End
       else
       // Node ISDN Node
       Begin
        if related=false then
        Begin
         //spannode:=Treeview1.items.AddChild(mynodepremise,'test')
         spannode:=Treeview1.Addchild(mynodepremise);
         nodeData := Treeview1.GetNodeData(spannode);
         NodeData.caption := 'Test';
         NodeData.D_Cust_Type := cust_type;
        end
        else
        Begin
         //spannode:=Treeview1.items.AddChild(servicenode,'test');
         spannode:=Treeview1.Addchild(servicenode);
         nodeData := Treeview1.GetNodeData(spannode);
         NodeData.caption := 'Test';
         NodeData.D_Cust_Type := cust_type;
        end;
       end;

       //ShowSpan(status,span,spantype,spanindex,regid,servicetype,btacno,btssd);
       ShowSpan(spannode,spanquery.fields[2].text,spanquery.fields[21].text,spanquery.fields[4].text,spanquery.fields[7].text,generalquery.fields[4].text,spanindex,spanquery.fields[3].text,generalquery.fields[1].text,spanquery.fields[22].text,btssd,spanquery.fields[24].text,spanquery.fields[5].text,spanquery.fields[6].text,agreementId,spanquery.fields[26].text,premiseId,'N/A','','','','','','','','', cust_type );
      end;

      oldspan:=spanquery.fields[7].text;
      spanquery.next;
     end;
    End;
    generalquery.next;
   end;
  end;
 end;

 if (tag=3) or (tag=4) then
 Begin
  with generalquery do
  Begin
   close;
   deletevariables;
   DeclareVariable('PID', otlong);
   sql.clear;
   sql.Add('select distinct');
   sql.add('A.AGREEMENT_ID,');
   sql.Add('T.DESCRIPTION,');
   sql.Add('A.CUSTOMER_ID,');
   sql.Add('A.AGREEMENT_START_DATE,');
   sql.Add('A.AGREEMENT_END_DATE,');
   sql.Add('A.ADDITIONAL_INFORMATION,');
   sql.add('A.AGREEMENT_STATUS_ID,');
   sql.add('S.DESCRIPTION S,');
   sql.add('IP.description,');
   sql.add('S.ICONINDEX,');
   sql.add('RD.RENEWAL_DATE');
   sql.add('from CRM.Agreements A, crm.agreement_status S,crm.service SE,crm.agreement_renewal_dates RD,');
   sql.add('crm.agreement_type T,crm.initial_periods IP');
   sql.add('where a.agreement_id=se.agreement_id');
   sql.add('and se.premise_id=:PID');
   sql.add('and a.agreement_status_id=s.agreement_status_id (+)');
   sql.add('and a.agreement_type_id=T.agreement_type_id (+)');
   sql.add('and a.initial_period=ip.initial_period_id (+)');
   sql.add('and a.agreement_id=rd.agreement_id (+)');
   sql.add('order by A.agreement_id desc');
   setvariable('PID',premiseId);
   open;
   deletevariables;
  end;

  while not generalquery.eof do
  Begin
   agreement_status_id:=generalquery.fields[6].text;
   agreement_status:=generalquery.fields[7].text;
   Agreement_Start_Date:=generalquery.fields[3].text;
   Agreement_End_Date:=generalquery.fields[4].text;
   AgreementId:=generalquery.fields[0].text;
   desc:='';
   //mynodeAgreement:=Treeview1.items.AddChild(mynodepremise,desc);

   mynodeAgreement:=Treeview1.Addchild(mynodepremise);
   nodeData := Treeview1.GetNodeData(mynodeAgreement);
   NodeData.caption := desc;
   NodeData.D_Cust_Type := cust_type;

   ShowAgreement(mynodeagreement,generalquery.fields[9].value,AgreementId,generalquery.fields[2].text,Agreement_status,agreement_start_Date,Agreement_end_date,generalquery.fields[8].text,false,generalquery.fields[10].text);
   Status := TFinancialHistoryInfo.GetFinancialStatusText(StrToInt64(AgreementId)); //FRM_Financial_History.GetStatus(Agreement_id);
   showproduct(mynodeagreement,AgreementId,false,status);
   generalquery.next;
  end;
 end;
 // Add Customer Tree
  // If sort order <> 1 then display customer Details for agreement
 if tag>1 then
 Begin
  with generalquery do
  Begin
   close;
   deletevariables;
   DeclareVariable('AID', otlong);
   sql.clear;
   sql.add('select C.customer_id,C.legal_entity_name,T.icon_index,');
   sql.add('P.PREMISE_LINE_1,');
   sql.add('P.PREMISE_LINE_2,');
   sql.add('P.PREMISE_LINE_3,');
   sql.add('P.PREMISE_LINE_4,');
   sql.add('P.PREMISE_LINE_5,');
   sql.add('P.PREMISE_LINE_6,');
   sql.add('P.PREMISE_LINE_7,');
   sql.add('P.PREMISE_LINE_8,');
   sql.add('P.PREMISE_LINE_9,');
   sql.add('P.PREMISE_POSTCODE,');
   sql.add('C.IS_PROSPECT,');
   sql.add('DT.DATE_SENT,');
   sql.Add('case when lv.customer_id is not null then ''Y'' else ''N'' end IS_LIVE_CUST');
   sql.add('from crm.customer c, crm.customer_type T, crm.agreements a,crm.premises P,');
   sql.add('crm.agreements_sent_to_debt_col DT,  (select customer_id,count(*) from crm.agreements where agreement_status_id=1');
   sql.Add('group by customer_id) lv');
   sql.add('where c.customer_id=a.customer_id');
   sql.add('and c.customer_type_id=t.customer_type_id');
   sql.add('and c.customer_id=lv.customer_id (+)');
   sql.add('and c.primary_mailing_address_id=p.premise_id');
   sql.add('and a.agreement_id=:AID');
   sql.add('and a.agreement_id=dt.agreement_id (+)');
   setvariable('AID',agreementId);
   open;
   deletevariables;
  end;
  debt:=generalquery.fields[14].text;
  cust:=GeneralQuery.fields[0].text;
  custname:=GeneralQuery.fields[1].text;

  If GeneralQuery.fields[2].Value <> Null then
    fCustIcon := GeneralQuery.fields[2].Value;

  mailing:='';
  if GeneralQuery.fields[3].text<>'' then mailing:=mailing+GeneralQuery.fields[3].text+', ';
  if GeneralQuery.fields[4].text<>'' then mailing:=mailing+GeneralQuery.fields[4].text+', ';
  if GeneralQuery.fields[5].text<>'' then mailing:=mailing+GeneralQuery.fields[5].text+', ';
  if GeneralQuery.fields[6].text<>'' then mailing:=mailing+GeneralQuery.fields[6].text+', ';
  if GeneralQuery.fields[7].text<>'' then mailing:=mailing+GeneralQuery.fields[7].text+', ';
  if GeneralQuery.fields[8].text<>'' then mailing:=mailing+GeneralQuery.fields[8].text+', ';
  if GeneralQuery.fields[9].text<>'' then mailing:=mailing+GeneralQuery.fields[9].text+', ';
  if GeneralQuery.fields[10].text<>'' then mailing:=mailing+GeneralQuery.fields[10].text+', ';
  if GeneralQuery.fields[11].text<>'' then mailing:=mailing+GeneralQuery.fields[11].text+', ';
  mailing:=mailing+GeneralQuery.fields[12].text+'';
  if cust='' then cust:='0';
  customerdeceased:='';
  premiseCount:=0;  // premise Count
  //mynodeCustomer:=treeview1.items.Addchild(mynodepremise, 'Customer '+cust+' - '+Custname);

  mynodeCustomer:=Treeview1.Addchild(mynodepremise);
  nodeData := Treeview1.GetNodeData(mynodeCustomer);
  NodeData.caption := 'Customer '+cust+' - '+Custname + GetCustomerPronoun(EmptyStr, Cust);
  NodeData.D_Cust_Type := cust_type;

  ShowCustomerOnly(mynodeCustomer,cust,custname,mailing,inttostr(premisecount),customerdeceased,generalquery.fields[13].text,  debt,'',generalquery.Fields[15].text);
 end;
 CrmCommon.InsertCustomerAccessedAudit(UserID, customer_id, agreementId, premiseId, null);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.ExpandCustomerNode(aSender: TBaseVirtualTree; aCustomerNode: PVirtualNode; var oDataProtectionOK: boolean);
var
  accountData : TAccountData;
begin
  oDataProtectionOK := true;

  // Now Show TreeDataProtection Warning
  if DPANAG then
  begin
    if not (fCustomerNote = CustId) then
    begin
      if (StrToFloat(FRM_Common.GETVALUE('CUSTOMER_DATA_CONFIRMATION_TIMEFRAME_MONTHS')) >= 0) then
      begin
        // Conversion will throw an exception if CustId is not a valid Int64
        {oDataProtectionOK := }TFrm_Dpa_Check.StartModal(Self, StrToInt64(CustId), accountData);
        if oDataProtectionOK then
          if accountData.ConfirmTel or accountData.ConfirmEmail then
            {oDataProtectionOK := }TFrm_Data_Capture.StartModal(Self, accountData);

        // Anna (ISC-848) settings oDataProtectionOk have been removed to allow expanding all nodes
        // regardless the outcome of either DPA check or Data Capture dialogue windows.
      end;

      fCustomerNote := CustId;
      Raise_Hot_Note(3, CustId);
    end;
  end;

  if oDataProtectionOK then
  begin
    if main.PriorityNotificationPopUp and (not (fCustomerPriorityNotification = CustId)) then
    begin
      fCustomerPriorityNotification := CustId;
      mnuViewPriorityNotificationClick(aSender);
      ShowInvoluntaryModeChangePopUp(CustId.ToInt64);
    end;
    ShowCursor(showfeedback);
    RefreshCustomerNode(aCustomerNode);

    // Automatically add a log to say that customer record access
    CrmCommon.InsertCustomerAccessedAudit(UserID, CustId, null, null, null);
  end;
end;

procedure TFRM_Tree.ExpandSpanNode(aSpanNode: PVirtualNode);
var
  customerId, agreementId, premiseId, spanId: string;
begin
  customerId := TreeData.D_Customer_Id;
  agreementId := TreeData.D_Agreement_ID;
  premiseId := TreeData.D_Premise_Id;
  spanId := TreeData.D_Span;
  if customerId = EmptyStr then
  begin
    if agreementId = EmptyStr then
      customerId := FRM_Common.GetCustomerId(spanId)
    else
      customerId := FRM_Common.GetCustomerIdfromAgreementid(agreementId);
  end;

  Treeview1.DeleteChildren(aSpanNode);

  if (spantype = 'G') or (spantype = 'C') then
  begin
    BuildS1EnrolledNode(spanId, aSpanNode);
    BuildGasMeterNode(aSpanNode);
  end;

  if (spantype = 'E') or (spantype = 'F') then
  begin
    BuildS1EnrolledNode(spanId, aSpanNode);
    BuildEnquiriesNode(aSpanNode);
    BuildElectricMeterNode(aSpanNode);
  end;

  if (spantype = 'T') or (spantype = 'J') then
    BuildTelecomMeterNode(aSpanNode);

  if tag = 4 then
    ShowSites(aSpanNode, spanId);

  CrmCommon.InsertCustomerAccessedAudit(userid, customerId, agreementId, premiseId, spanId);
end;

procedure TFRM_Tree.RefundActionClick(Sender: TObject);
begin
  Xnode    := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(Xnode);
  FormRefunds := TFormRefunds.Create(Self);
  FormRefunds.ClearAll;
  FormRefunds.EditCustomerNumber.Text := TreeData.D_Customer_Id;
  FormRefunds.Showmodal;
  If Assigned(FormRefunds) Then
   Begin
     FormRefunds.Free;
   End;
end;

Procedure TFRM_Tree.RefreshagreementNode(mynodeagreement:Pvirtualnode);
Var
salesagent:boolean;
//MyRecPtr: PMyRec;
Agreement_ID,status,Customer_ID:String;
Begin
 TreeData:= treeview1.GetNodeData(mynodeagreement);

 Agreement_id:=TreeData.D_agreement_id;
 Customer_id:=TreeData.D_Customer_id;
 try
  treeview1.deletechildren(mynodeagreement);
 except
 end;

 // Get Details about selected Agreement
 SalesAgent:=true;
 if copy(TreeData.caption,1,16)<>'Agreement - None' then
 Begin
  salesagent:=false;
  with generalquery do
  Begin
   close;
   deletevariables;
   DeclareVariable('AID', otlong);
   sql.clear;
   sql.Add('select');
   sql.add('A.AGREEMENT_ID,');
   sql.Add('T.DESCRIPTION,');
   sql.Add('A.CUSTOMER_ID,');
   sql.Add('A.AGREEMENT_START_DATE,');
   sql.Add('A.AGREEMENT_END_DATE,');
   sql.Add('A.ADDITIONAL_INFORMATION,');
   sql.add('A.AGREEMENT_STATUS_ID,');
   sql.add('S.DESCRIPTION S,');
   sql.add('IP.description,');
   sql.add('S.ICONINDEX,');
   sql.add('RD.RENEWAL_DATE');
   sql.add('from CRM.Agreements A, crm.agreement_status S,crm.agreement_renewal_dates RD,');
   sql.add('crm.agreement_type T,crm.initial_periods IP');
   sql.add('where A.agreement_id=:AID');
   sql.add('and a.agreement_status_id=s.agreement_status_id (+)');
   sql.add('and a.agreement_type_id=T.agreement_type_id (+)');
   sql.add('and a.initial_period=ip.initial_period_id (+)');
   sql.add('and a.agreement_id=rd.agreement_id (+)');
   sql.add('order by A.agreement_id desc');
   setvariable('AID',agreement_id);
   open;
   deletevariables;
  end;
  agreement_status_id:=generalquery.fields[6].text;
  agreement_status:=generalquery.fields[7].text;
  Agreement_Start_Date:=generalquery.fields[3].text;
  Agreement_End_Date:=generalquery.fields[4].text;
  Agreement_id:=generalquery.fields[0].text;

  IF SHOWREASSIGN = FALSE THEN
  BEGIN
   ShowAgreement(mynodeagreement,generalquery.fields[9].value,Agreement_id,customer_id,Agreement_status,agreement_start_Date,Agreement_end_date,generalquery.fields[8].text,false,generalquery.fields[10].text);
   Status := TFinancialHistoryInfo.GetFinancialStatusText(StrToInt64(Agreement_id)); //FRM_Financial_History.GetStatus(Agreement_id);
   Showproduct(mynodeagreement,Agreement_id,true,status);
   ShowAnyDisputes(mynodeagreement,Agreement_id);
   ShowLatestAccountReview(MyNodeAgreement,Agreement_id);
   ShowLatestRatedUsage(mynodeagreement,agreement_id);
  END;

  CrmCommon.InsertCustomerAccessedAudit(UserID, Customer_id, Agreement_id, null, null);
 end;

 if tag<3 then
 Begin
  // Get Premise in Agreement
  with generalquery do
  Begin
   close;
   deletevariables;
   DeclareVariable('AID', otlong);
   sql.clear;
   sql.add('select Distinct');
   sql.add('APS.PREMISE_NAME,');
   sql.add('APS.SPECIAL_ACCESS,');
   sql.add('APS.PREMISE_CONTACT_ID,');
   sql.add('P.PREMISE_ID,');
   sql.add('T.ICON_INDEX,');
   sql.add('P.PREMISE_LINE_1,');
   sql.add('P.PREMISE_LINE_2,');
   sql.add('P.PREMISE_LINE_3,');
   sql.add('P.PREMISE_LINE_4,');
   sql.add('P.PREMISE_LINE_5,');
   sql.add('P.PREMISE_LINE_6,');
   sql.add('P.PREMISE_LINE_7,');
   sql.add('P.PREMISE_LINE_8,');
   sql.add('P.PREMISE_LINE_9,');
   sql.add('P.PREMISE_POSTCODE,');
   sql.add('APS.DATE_MOVED_OUT');
   sql.add('from crm.premises P,');
   sql.add('crm.agreement_premises APS,');
   sql.add('crm.premise_type t');
   if salesagent=false then
   Begin
    sql.add(', crm.service S');
    sql.add('where S.agreement_id=:AID');
    sql.add('and S.premise_id=P.premise_id (+)');
    sql.add('and P.premise_type_id=t.premise_type_id (+)');
    sql.add('and s.agreement_id=aps.agreement_id (+)');
    sql.add('and s.premise_id=aps.premise_id (+)');
   end
   else
   Begin
    sql.add('where APS.agreement_id=:AID');
    sql.add('and P.premise_id=APS.premise_id (+)');
    sql.add('and P.premise_type_id=t.premise_type_id (+)');
   End;
   setvariable('AID',agreement_id);
   open;
   deletevariables;
  end;
  with generalquery do
  Begin
   while not eof do
   Begin
    premaddr:='';
    if fields[5].text<>'' then premaddr:=premaddr+fields[5].text+',';
    if fields[6].text<>'' then premaddr:=premaddr+fields[6].text+',';
    if fields[7].text<>'' then premaddr:=premaddr+fields[7].text+',';
    if fields[8].text<>'' then premaddr:=premaddr+fields[8].text+',';
    if fields[9].text<>'' then premaddr:=premaddr+fields[9].text+',';
    if fields[10].text<>'' then premaddr:=premaddr+fields[10].text+',';
    if fields[11].text<>'' then premaddr:=premaddr+fields[11].text+',';
    if fields[12].text<>'' then premaddr:=premaddr+fields[12].text+',';
    if fields[13].text<>'' then premaddr:=premaddr+fields[13].text+',';
    premaddr:=premaddr+fields[14].text;
    premaddr:=premaddr+' - ['+fields[3].text+']';
   { mynodepremise:=Treeview1.items.AddChild(mynodeagreement,'Premises - '+premaddr);
    mynodepremise.imageindex:=fields[4].value;
    mynodepremise.selectedindex:=mynodepremise.imageindex; }

    mynodepremise:=Treeview1.Addchild(mynodeagreement);
    nodeData := Treeview1.GetNodeData(mynodepremise);
    NodeData.caption := 'Premises - '+premaddr;
    nodedata.index:=fields[4].value;


    cot1.visible:=true;
    cot2.visible:=false;
    if fields[15].text<>'' then
    Begin
      cot1.visible:=false; // Dont Show COT option if already vavacted
      cot2.visible:=true;  // Show COT Tools if Vacated
      if strtodate(fields[15].text)<date then
      Begin
       nodedata.caption:=nodedata.caption+' - Vacated on '+fields[15].text;
       nodedata.index:=142;
       //mynodepremise.selectedindex:=mynodepremise.imageindex;
       nodedata.Fontcolor:=clmaroon;
      end
      else
      Begin
       nodedata.caption:=nodedata.caption+' - Vacating on '+fields[15].text;
       nodedata.index:=139;
       //mynodepremise.selectedindex:=mynodepremise.imageindex
      End;
    end;


    nodedata.D_premise_Id :=   fields[3].text;
    nodedata.D_agreement_Id := agreement_id;
    nodedata.D_customer_Id := customer_id;
    //mynodepremise.data:=MyRecPtr;
    //mynode1:=Treeview1.items.AddChild(mynodepremise,'test');
    mynode1:=Treeview1.Addchild(mynodepremise);
    nodeData := Treeview1.GetNodeData(mynode1);
    NodeData.caption := 'Test';

    next;
   end;
  end;
 end;
 // Refresh statement TreeDatainto users schema.
 // Nut inly if rated TreeDataexists
 frm_main.get_statement_data(agreement_id);
end;

Procedure TFRM_Tree.RefreshJBSPushBackNode(mynodeJOB:Pvirtualnode);
Var
JBS_ID:String;
Begin

 TreeData:= treeview1.GetNodeData(mynodeJOB);

 JBS_id:=TreeData.C_Record_id;

 try
//  treeview1.deletechildren(mynodeJOB);
 except
 end;

 with main_data_module.GeneralQuery do
 begin
   close;
   deletevariables;
   declarevariable('JOBID',otstring);
   setvariable('JOBID',JBS_ID);
   sql.clear;
   sql.add('select * from SMIFF.WMOL_VW_PUSHBACK_HISTORY where job_id=:JOBID order by pushback_counter desc');
   Open;
   deletevariables;
 end;
 while not  main_data_module.GeneralQuery.eof do
 begin
   desc:=main_data_module.generalquery.Fields[1].Text;
   jbsPushnode:=Treeview1.Addchild(mynodeJOB);
   nodeData := Treeview1.GetNodeData(jbsPushnode);
   NodeData.caption := desc;
   nodedata.index:=219;
   nodedata.fontcolor:=clpurple;
   main_data_module.GeneralQuery.next;
  end;
end;


procedure TFRM_Tree.P1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 FRM_Main.SearchForpremise(TreeData.D_premise_id);
end;

procedure TFRM_Tree.MenuItem1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 FRM_Main.SearchForAgreement(TreeData.D_Agreement_id);
end;

procedure TFRM_Tree.MaintainAccountHolder1Click(Sender: TObject);
begin
FRM_Maintain_accountholders.showmodal;
end;

procedure TFRM_Tree.MaintainAccountHoldersClick(Sender: TObject);
var
  node     : PVirtualNode;
  treeData : PMyRec;
  custId   : Int64;
  contId   : Int64;
begin
  node := treeview1.FocusedNode;

  if not Assigned(node) then
  begin
    MessageDlg('Select a valid tree node.', mtError, [mbOk], 0);
    Exit;
  end;

  treeData := TreeView1.GetNodeData(node);

  if not Assigned(treeData) then
  begin
    MessageDlg('No data avaliable.', mtError, [mbOk], 0);
    Exit;
  end;

  if not TryStrToInt64(TreeData.D_Customer_Id,custId) then
  begin
    MessageDlg('Invalid customer Id: ' + treeData.D_Customer_Id, mtError, [mbOk], 0);
    Exit;
  end;

  contId := 0;
  if not frm_add_account_holder.StartAddAccountHolder(custId,contId,smInsert) then
    exit;

  // refresh customer node;
  treeview1.Expanded[xnode]:=false;
  treeview1.Expanded[xnode]:=true;
end;

procedure TFRM_Tree.AddAgreement1Click(Sender: TObject);
begin
  xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

  With FRM_Agreement Do
 Begin
  tag:=0;
  // Agreement tab
  custid.text:=TreeData.D_Customer_id;
  custname.text:=TreeData.D_Customer_name;
  clearfields;
  showmodal;
 end;
 if frm_agreement.tag=0 then exit;

  // refresh customer node;
 if treeview1.Selected[xnode]=true then
 Begin
  treeview1.Expanded[xnode]:=false;
  treeview1.Expanded[xnode]:=true;
 end;
end;

procedure TFRM_Tree.P3Click(Sender: TObject);
Var
Customer_id,agreement_id:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Agreement_id:=TreeData.D_agreement_id;
 Customer_id:=TreeData.D_Customer_id;

 if (TreeData.index=142) or (TreeData.index=139) then
 Begin
  Messagedlg('This action cannot be performed against a VACATED Property',mtwarning,[mbok],0);
  exit;
 End;

 // Get a List Of Premises For Agreement
 // Default Site Address to Customer Mailing Address
 FRM_Premise_Details.clearfields;
 with FRM_Premise_Details.premisequery do
 Begin
  close;
  FRM_Premise_Details.agreementid.text:=agreement_id;
  setvariable('Customerid',customer_id);
  setvariable('aggreementid',agreement_id);
  open;
  FRM_Premise_Details.sitelookup.keyvalue:=FRM_Premise_Details.premisequery.fields[14].text;
 End;
 FRM_Premise_Details.premisecontrol.ActivePageIndex:=2;
 FRM_Premise_Details.showmodal;

 With FRM_Agreement Do
 Begin
  tag:=0;
  // Agreement tab
  custid.text:=customer_id;
  custname.text:=customer_id;
  Getfields(agreement_id);
  Statusbar.panels[0].text:=' Update';
  Frm_agreement.pagecontrol1.activepageindex:=2;
  frm_agreement.productquery.last;
  frm_agreement.openproduct;
  frm_agreement.maintain_schedule;
 end;

 // refresh customer node;
 treeview1.Expanded[xnode]:=false;
 treeview1.Expanded[xnode]:=true;
end;


procedure TFRM_Tree.ViewAggreement1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 With FRM_Agreement Do
 Begin
  tag:=0;
  // Agreement tab
  custid.text:=TreeData.D_customer_id;
  Getfields(TreeData.D_Agreement_id);
  Statusbar.panels[0].text:=' Update';
  frm_agreement.pagecontrol1.activepage:=tabsheet1;
  showmodal;
 end;
  if treeview1.Selected[xnode]=true then
  Begin
   treeview1.Expanded[xnode]:=false;
   treeview1.Expanded[xnode]:=true;
  end;
end;

procedure TFRM_Tree.A_ProdWizardClick(Sender: TObject);
Var
agreementid,custid:string;
begin
 xnode:=treeview1.FocusedNode;
 nodeData := treeview1.GetNodeData(xnode);

 Agreementid:=nodedata.D_agreement_id;
 custid:=nodedata.D_customer_id;
 frm_add_customer_wizard.Step3aWhichPremise(agreementid,custid);
  // refresh customer node;
  if treeview1.Selected[xnode]=true then
  Begin
   treeview1.Expanded[xnode.parent]:=false;
   treeview1.Expanded[xnode.parent]:=true;
  end;
end;

procedure TFRM_Tree.RegisterItemClick(Sender: TObject);
var
agreement_id:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 agreement_id:=TreeData.D_Agreement_id;
 FRM_Add_Customer_Wizard.Step8RegisterAgreement(agreement_id);
 // refresh customer node;
 treeview1.Expanded[xnode.parent]:=false;
 treeview1.Expanded[xnode.parent]:=true;
end;

procedure TFRM_Tree.EditAccountHoldersClick(Sender: TObject);
var
  node     : PVirtualNode;
  treeData : PMyRec;
  custId   : Int64;
  contId   : Int64;
begin
  node := TreeView1.FocusedNode;

  if not Assigned(node) then
  begin
    MessageDlg('Select a valid tree node.', mtError, [mbOk], 0);
    Exit;
  end;

  treeData := TreeView1.GetNodeData(node);

  if not Assigned(treeData) then
  begin
    MessageDlg('No data avaliable.', mtError, [mbOk], 0);
    Exit;
  end;

  if not TryStrToInt64(TreeData.D_Customer_Id,custId) then
  begin
    MessageDlg('Invalid customer Id: ' + treeData.D_Customer_Id, mtError, [mbOk], 0);
    Exit;
  end;

  if not TryStrToInt64(treeData.D_Contact_Id,contId) then
  begin
    MessageDlg('Invalid contact Id: ' + treeData.D_Contact_Id, mtError, [mbOk], 0);
    Exit;
  end;

  if not Frm_Add_Account_Holder.StartAddAccountHolder(custId, contId, smEdit) then
    exit;

  RefreshAccountOrder(treeData.D_Customer_Id);
  // Refresh AccountHolder node
  TreeView1.Expanded[node.Parent] := false;
  TreeView1.Expanded[node.Parent] := true;
end;

procedure TFRM_Tree.CallDataRecords1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Application.CreateForm(TFRM_TELEPHONE, FRM_TELEPHONE);
 try
  frm_telephone.CDR_Telephone.close;
  frm_telephone.CDR_Telephone.open;
  frm_telephone.tellookup.enabled:=true;
  frm_telephone.tellookup.keyvalue:=stringreplace(TreeData.D_SPAN,' ','',[rfreplaceall]);
  frm_telephone.tellookup.enabled:=false;
  frm_telephone.showmodal;
 finally
  frm_telephone.release;
 end;
end;

procedure TFRM_Tree.A_decClick(Sender: TObject);
var
ahid,custid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 custid:=TreeData.D_customer_id;
 ahid:=TreeData.D_Account_Holder_Id;

 if MessageDlg('Are you sure you wish to make this Account holder Deceased?'+#13+Treedata.caption,
  mtConfirmation, [mbYes, mbNo], 0) = mrYes then
 begin


  with main_data_module.updatequery do
  Begin
   close;
   DeleteVariables;
   sql.clear;
   sql.add('Insert into crm.account_holders_validity values('+ahid+','+custid+',');
   sql.add('sysdate,'''+userid+''',3,''Account Holder Deceased'')');
   execute;
  End;
  with main_data_module.updatequery do
  Begin
   close;
   DeleteVariables;
   sql.clear;
   sql.add('update crm.account_holders');
   sql.add('set account_holder_status_id=3');
   sql.add('where account_holder_id='+ahid);
   sql.add('and customer_id='+custid);
   execute;
  End;

  FRM_Login.MainSession.commit;
  RefreshAccountOrder(custid);
  treeview1.expanded[xnode.parent]:=false;
  treeview1.expanded[xnode.parent]:=true;
 end;

end;

procedure TFRM_Tree.A_revClick(Sender: TObject);
var
ahid,custid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 custid:=TreeData.D_customer_id;
 ahid:=TreeData.D_Account_Holder_Id;

 if MessageDlg('Are you sure you wish to Revive this Account Holder?'+#13+Treedata.caption,
  mtConfirmation, [mbYes, mbNo], 0) = mrYes then
 begin

  with main_data_module.updatequery do
  Begin
   close;
   DeleteVariables;
   sql.clear;
   sql.add('Insert into crm.account_holders_validity values('+ahid+','+custid+',');
   sql.add('sysdate,'''+userid+''',4,''Account Holder Revived'')');
   execute;
  End;
  with main_data_module.updatequery do
  Begin
   close;
   DeleteVariables;
   sql.clear;
   sql.add('update crm.account_holders');
   sql.add('set account_holder_status_id=1');
   sql.add('where account_holder_id='+ahid);
   sql.add('and customer_id='+custid);
   execute;
  End;
  FRM_Login.MainSession.commit;
  refreshaccountorder(custid);
  treeview1.expanded[xnode.parent]:=false;
  treeview1.expanded[xnode.parent]:=true;
 end;
end;

procedure TFRM_Tree.RemoveAccountHolder1Click(Sender: TObject);
var
ahid,custid:string;
ahname:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 custid:=TreeData.D_customer_id;
 ahid:=TreeData.D_Account_Holder_Id;
 AHNAME:=TreeData.caption;

 with main_data_module.generalquery do
 Begin
  close;
  deletevariables;
  declarevariable('CUSTID',otlong);
  declarevariable('AHID',otlong);
  sql.clear;
  sql.add('select * from crm.account_holders where account_holder_status_id=1');
  sql.add('and account_holder_id<>:AHID and customer_id=:CUSTID');
  setvariable('AHID',ahid);
  setvariable('CUSTID',custid);
  open;
  deletevariables;
 End;
 if main_data_module.generalquery.recordcount=0 then
 Begin
  Messagedlg('You cannot delete this account holder. Customer must'+#13+
             'have at least One account holder assigned.',MTinformation,[MBOK],0);
  exit;
 End;
 // What if only One Active Account Holder?
 // Can They be Deleted

 if MessageDlg('Are you sure you wish to Remove this Account Holder?'+#13+#13+AHNAME+#13+#13+
               'This Account Holder Cannot be re-instated!',
  mtConfirmation, [mbYes, mbNo], 0) = mrYes then
 begin

  with main_data_module.updatequery do
  Begin
   close;
   DeleteVariables;
   sql.clear;
   sql.add('Insert into crm.account_holders_validity values('+ahid+','+custid+',');
   sql.add('sysdate,'''+userid+''',2,''Account Holder Removed'')');
   execute;
  End;
  with main_data_module.updatequery do
  Begin
   close;
   DeleteVariables;
   sql.clear;
   sql.add('update crm.account_holders');
   sql.add('set account_holder_status_id=2');
   sql.add('where account_holder_id='+ahid);
   sql.add('and customer_id='+custid);
   execute;
  End;
  refreshaccountorder(custid);
  FRM_Login.MainSession.commit;
  treeview1.DeleteNode(xnode);
  treeview1.expanded[xnode.parent]:=false;
  treeview1.expanded[xnode.parent]:=true;
 end;
end;

procedure TFRM_Tree.MopHistoryClick(Sender: TObject);
begin
  if treeupdating=true then exit;
 mpannode:=moptree.FocusedNode;
 TreeData:= moptree.GetNodeData(mpannode);

 //mpannode:=MopTree.selected;
 mpan:=TreeData.D_SPAN;
 if not Assigned(FRM_DFLOW_HISTORY_MOP) then Application.CreateForm(TFRM_DFLOW_HISTORY_MOP, FRM_DFLOW_HISTORY_MOP);
 FRM_DFLOW_History_MOP.mpanedit.text:=mpan;
 FRM_DFLOW_History_MOP.DflowQuery(MPAN,'');
 if FRM_DFLOW_History_MOP.caption='' then
 Begin
  messagedlg('There is no Dataflow History for this MPAN',MTinformation,[MBOK],0);
  exit;
 end;
 FRM_DFLOW_History_MOP.show;
end;

procedure TFRM_Tree.NEWBTNClick(Sender: TObject);
begin
 if MessageDlg('Do you wish to create a new customer?',
 mtConfirmation, [mbYes, mbNo], 0) = mryes then
 Begin
  Frm_Add_Customer_Wizard.showmodal;
 End;
end;

procedure TFRM_Tree.BuildGasMeterNode(MPANNODE:Pvirtualnode);
var
mpan,MeterText,m,enstatus,agid:string;
showmeter:boolean;
Begin
 nodeData := treeview1.GetNodeData(mpannode);
 fPremiseDcc:= false;

  // Check for Meter Technical Details
  //mpannode:=treeview1.selected;
  mpan:=nodedata.D_SPAN;
  agid:=nodedata.D_agreement_id;
  // Check For Last Known PPMIP
  with main_data_module.tempquery do
  Begin
   close;
   deletevariables;
   declarevariable('MPAN',otstring);
   sql.clear;
   sql.add('select * from gdmgr.quantum_data where meter_point_reference=:MPAN');
   sql.add('and (last_trans is not null or last_supplier_file is not null or last_talir011_file is not null)');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
  end;
  if main_data_module.tempquery.recordcount<>0 then
  Begin
   desc:='Legacy Quantum Gas Meter Data.';
   if main_data_module.tempquery.fields[1].text<>'' then desc:=desc+#10+'Last Transaction: '+main_data_module.tempquery.fields[1].text;
   if main_data_module.tempquery.fields[2].text<>'' then desc:=desc+#10+'LIVE IN QUANTUM: '+main_data_module.tempquery.fields[2].text;
   if main_data_module.tempquery.fields[3].text<>'' then desc:=desc+#10+'Data Exists in CASH to CLOSED: '+main_data_module.tempquery.fields[3].text;
   {MeterConfigNode:=Treeview1.items.AddChild(mpannode,desc);
   MeterConfigNode.imageindex:=217;
   MeterConfigNode.selectedindex:=217;}

   MeterConfigNode:=Treeview1.Addchild(mpannode);
   nodeData := Treeview1.GetNodeData(MeterConfigNode);
   NodeData.caption := desc;
   nodedata.index:=217;

   nodedata.D_SPAN := MPAN;
   nodedata.D_Agreement_id := agid;

   //meterconfignode.data:=MyRecPtr;
  end;


  With GasMeter do
  Begin
   Close;
   setvariable('SPAN',mpan);
   open;
  end;
   // Only Do This Block If Meter Records Exist
  if GasMeter.recordcount<>0 then
  Begin
   msid:='LEEOK';
   oldefsdmsmtd:='lee';
   oldmeterid:='lee';

   // what if two meters on same date

   while not GasMeter.eof do
   begin
    efsdmsmtd := FormatDateTime('dd/mm/yyyy',gasmeter.fields[2].asDateTime);

    if oldefsdmsmtd='lee' then config:='Current Configuration'
      else config:='Previous Configuration';

    if GasMeter.fields[3].text <> '' then
      dateremoved := '    (Date Removed=' + GasMeter.fields[3].text + ')'
    else
      dateremoved := '';

    if GasMeter.FieldByName('METERMECHANISM').AsString = 'S2' then
    begin
      MeterText := 'SMETS 2 ' + GasMeter.fields[0].text + dateremoved;
    end
    else if GasMeter.FieldByName('METERMECHANISM').AsString = 'S1EA' then
    begin
      MeterText := 'SMETS 1 E&A ' + GasMeter.fields[0].text + dateremoved;
    end
    else if GasMeter.FieldByName('METERTYPE').AsString = 'F' then
    begin
      MeterText := 'Imperial Meter ID - ' + GasMeter.fields[0].text + dateremoved;
    end
    else
    begin
      MeterText := 'Metric Meter ID - ' + GasMeter.fields[0].text + dateremoved;
    end;

    if GasMeter.fields[4].text = 'N' then
      enstatus := 'De-Energised'
    else
      enstatus := 'Energised';

    showmeter:=false;

    // SHow Multiple Meters for Same Date?
    if efsdmsmtd=oldefsdmsmtd then showmeter:=true;

    if efsdmsmtd<>oldefsdmsmtd then
    begin
     if ((historic1.checked=true) and(config='Previous Configuration')) then
     Begin
      {meterconfignode:=Treeview1.items.AddChild(mpannode,Config+' -'+efsdmsmtd+' - '+enstatus);
      MeterConfigNode.font.color:=clred;
      MeterConfigNode.font.style:=[fsbold];
      MeterConfigNode.imageindex:=27;
      MeterConfigNode.selectedindex:=27;}

      MeterConfigNode:=Treeview1.Addchild(mpannode);
      nodeData := Treeview1.GetNodeData(MeterConfigNode);
      desc:= Config+' -'+efsdmsmtd+' - '+enstatus;
      if gasmeter.fields[7].text<>'GAS' then desc:=desc+#10+'(* WARNING: '+gasmeter.fields[7].text+' *)';
      NodeData.caption:=desc;

      nodedata.index:=27;
      nodedata.fontcolor:=clred;
      nodedata.fontBold:=true;

      showmeter:=true;
      oldmeterid:='lee';
     end
     else
     Begin
      if config='Current Configuration' then
      Begin
       {meterconfignode:=Treeview1.items.AddChild(mpannode,Config+' -'+efsdmsmtd+' - '+enstatus);
       MeterConfigNode.font.color:=clgreen;
       MeterConfigNode.font.style:=[fsbold];
       MeterConfigNode.imageindex:=27;
       MeterConfigNode.selectedindex:=27; }

       MeterConfigNode:=Treeview1.Addchild(mpannode);
       nodeData := Treeview1.GetNodeData(MeterConfigNode);
       desc:= Config+' -'+efsdmsmtd+' - '+enstatus;
       if gasmeter.fields[7].text<>'GAS' then desc:=desc+#10+'(* WARNING: '+gasmeter.fields[7].text+' *)';
       NodeData.caption:=desc;
       nodedata.index:=27;
       nodedata.fontcolor:=clgreen;
       nodedata.fontBold:=true;

       showmeter:=true;
       oldmeterid:='lee';
      end;
     end
    end;

    if showmeter=true then
    Begin
     //MeterNode:=Treeview1.items.AddChild(meterconfignode,MeterText);
     if gasmeter.fields[7].text<>'GAS' then metertext:=metertext+#10+'(* WARNING: '+gasmeter.fields[7].text+' *)';
     MeterNode:=Treeview1.Addchild(meterconfignode);
     nodeData := Treeview1.GetNodeData(MeterNode);
     NodeData.caption := MeterText;
     nodedata.fontcolor:=clblack;
     nodedata.index:=17; // NHH Credit Meter
     nodedata.D_SPAN :=mpan;
     nodedata.M_METERID :=gasmeter.fields[0].text;
     nodedata.M_SERVICE :='1';

     nodedata.Metertype := gasmeter.FieldByName('METERTYPE').text;

  // SMI check for DCC Meter
     nodedata.D_SpanDCC_G := false;
     nodedata.D_SpanG := '';

     with Generalquery do
     Begin
      close;
      DeleteVariables;
      DeclareVariable('MPXN', otstring);
      sql.clear;
      sql.add('Select ods.dcc_enrolled(:MPXN)from dual');
      setvariable('MPXN', mpan);
      open;
      deletevariables;

      if generalquery.Fields[0].Text = 'Y' then
      begin
        nodedata.D_SpanDCC_G := true;
        nodedata.D_SpanG := mpan;
        fPremiseDcc:= true;
      end;
     End;

     // Try and Swow correct Icons
     if (GasMeter.FieldByName('METERMECHANISM').AsString='CM') or
       (GasMeter.FieldByName('METERMECHANISM').AsString='ET') or
       (GasMeter.FieldByName('METERMECHANISM').AsString='MT') or
       (GasMeter.FieldByName('METERMECHANISM').AsString='PP') or
       (GasMeter.FieldByName('METERMECHANISM').AsString='TH') then nodedata.index:=23;

     if GasMeter.FieldByName('MANUFACTURECODE').AsString='PRI' then nodedata.index:=314;
     if (GasMeter.FieldByName('MANUFACTURECODE').AsString='PRI') and (GasMeter.FieldByName('METERMECHANISM').AsString='NS') then nodedata.index:=314;
     if (GasMeter.FieldByName('MANUFACTURECODE').AsString='SCM') and (GasMeter.FieldByName('METERMECHANISM').AsString='NS') then nodedata.index:=314;
     if (GasMeter.FieldByName('MANUFACTURECODE').AsString='PRI') and (GasMeter.FieldByName('METERMECHANISM').AsString='S1') then nodedata.index:=239;
     if (GasMeter.FieldByName('MANUFACTURECODE').AsString='SCM') and (GasMeter.FieldByName('METERMECHANISM').AsString='S1') then nodedata.index:=239;
     if (GasMeter.FieldByName('METERMECHANISM').AsString='S2') then nodedata.index:=205;

     if dateremoved<>'' then
     Begin
      nodedata.index:=24; // Removed Meter
     end;
     //MeterNode.selectedindex:=nodedata.index;
     if (GasMeter.FieldByName('ENDDATE').AsString <> '') or (GasMeter.FieldByName('ACTIVE').AsString='N') then nodedata.fontcolor:=clred;
     if GasMeter.FieldByName('METERTYPE').AsString='F' then M:='Feet' else m:='Metres';

     if (GasMeter.FieldByName('METERMECHANISM').AsString = 'S1EA') then
     begin
       nodedata.index := 313;
     end
     else
     begin
       if (GasMeter.FieldByName('METERMECHANISM').AsString = 'S1') and
        ((GasMeter.FieldByName('MANUFACTURECODE').AsString ='PRI') or (GasMeter.FieldByName('MANUFACTURECODE').AsString ='SCM'))  then
       begin
         ShowSmetsMeterCommsSupplier(MeterNode,MPAN,GasMeter.FieldByName('SERIALNUM').AsString,'1',dateremoved,'X');

         if GasMeter.FieldByName('DATA_SOURCE').AsString <>'GAS' then
         begin
           nodedata.caption:=nodedata.caption+#10+'(* WARNING: '+ GasMeter.FieldByName('DATA_SOURCE').AsString+' *)';
         end;
       end;
     end;

     //Meternode.data:=MyRecPtr;

     {MeterRegisterNode:=Treeview1.items.AddChild(MeterNode,'(All Periods) Active Cubic '+m+' Import');
     MeterRegisterNode.font.color:=clblack;
     MeterRegisternode.imageindex:=28;}

     MeterRegisterNode:=Treeview1.Addchild(MeterNode);
     nodeData := Treeview1.GetNodeData(MeterRegisterNode);
     NodeData.caption := '(All Periods) Active Cubic '+m+' Import';
     Nodedata.index:=28;

     if (Gasmeter.fields[3].text<>'') or (Gasmeter.fields[4].text='N') then nodedata.index:=29;
     //Meterregisternode.selectedindex:=Meterregisternode.imageindex;

      if GasMeter.FieldByName('hot_shoe').AsString <> '' then
      begin
        if Assigned(MeterNode) then
        begin
          MeterRegisterNode := TreeView1.AddChild(MeterNode);

          if Assigned(MeterRegisterNode) then
          begin
            nodeData := TreeView1.GetNodeData(MeterRegisterNode);
            if Assigned(nodeData) then
            begin
              nodeData.Caption := 'Hot Shoe - ' + GasMeter.FieldByName('hot_shoe').AsString;
              nodeData.Index   := iiHotShoe;
            end;
          end;
        end;
      end;

    end;
    oldmeterid:=meterid;
    oldefsdmsmtd:=efsdmsmtd;
    GasMeter.next;
   end;
  end; // End Of Meter Strucutre Tree
end;

procedure TFRM_Tree.BuildTelecomMeterNode(MPANNODE:PVirtualnode);
var
CLID,efd:string;
Begin
 // mpannode:=treeview1.selected;
 nodeData := treeview1.GetNodeData(mpannode);

  CLID:=nodedata.D_SPAN;
  With Generalquery do
  Begin
   close;
   deletevariables;
   DeclareVariable('CLID', otstring);
   sql.clear;
   sql.add('select min(C.START_DATE_TIME) FirstCall,');
   sql.add('max(C.START_DATE_TIME) LastCall');
   sql.add('FROM BILLING.RATED_T C');
   sql.add('Where C.CALLING_LINE_ID =:CLID');
   sql.add('and c.item_charge_type=''CA''');
   setvariable('CLID',stringreplace(CLID,' ','',[rfreplaceall]));
   open;
   deletevariables;
  end;
  if GeneralQuery.fields[0].text<>'' then
  Begin
   {MeterNode:=Treeview1.items.AddChild(mpannode,CLID+' Calls Recorded from '+generalquery.fields[0].text+' - '+generalquery.fields[1].text);
   nodedata.fontcolor:=clblack;
   nodedata.index:=101;
   MeterNode.selectedindex:=101;}


   MeterNode:=Treeview1.Addchild(mpannode);
   nodeData := Treeview1.GetNodeData(MeterNode);
   NodeData.caption := CLID+' Calls Recorded from '+generalquery.fields[0].text+' - '+generalquery.fields[1].text;
   nodedata.index:=101;
   nodedata.fontcolor:=clblack;
   nodedata.D_SPAN :=mpan;
   nodedata.M_METERID :=Meterid;
   nodedata.M_SERVICE :='0';

  end;

  // Now Get Current List Of Phone Feature Features
  With Generalquery do
  Begin
   Close;
   deletevariables;
   declarevariable('CLID',otstring);
   sql.clear;
   sql.add('select * from telecoms.services_and_features_SPAN');
   sql.add('Where span like :CLID');
   sql.add('order by effective_from desc,last_updated desc');
   setvariable('CLID',CLID+'%');
   open;
   deletevariables;
  end;
  if GeneralQuery.recordcount<>0 then
  Begin
   {FeatureNode:=Treeview1.items.AddChild(mpannode,'Services & Network Features');
   FeatureNode.imageindex:=93;
   FeatureNode.selectedindex:=FeatureNode.imageindex; }

   FeatureNode:=Treeview1.Addchild(mpannode);
   nodeData := Treeview1.GetNodeData(FeatureNode);
   NodeData.caption := 'Services & Network Features';
   nodedata.index:=93;


   if Generalquery.fields[2].text='Y' then addFeature('Z1',clid);
   if Generalquery.fields[3].text='Y' then addFeature('Z2',clid);
   if Generalquery.fields[4].text='Y' then addFeature('Z3',clid);
   if Generalquery.fields[5].text='Y' then addFeature('Z4',clid);
   if Generalquery.fields[6].text='Y' then addFeature('Z5',clid);
   if Generalquery.fields[7].text='Y' then addFeature('Z6',clid);
   if Generalquery.fields[8].text='Y' then addFeature('Z7',clid);
   if Generalquery.fields[9].text='Y' then addFeature('Z8',clid);
   if Generalquery.fields[10].text='Y' then addFeature('Z9',clid);
   if Generalquery.fields[11].text='Y' then addFeature('ZA',clid);
   if Generalquery.fields[12].text='Y' then addFeature('ZB',clid);
   if Generalquery.fields[13].text='Y' then addFeature('ZC',clid);
   if Generalquery.fields[14].text='Y' then addFeature('ZD',clid);
   if Generalquery.fields[15].text='Y' then addFeature('ZE',clid);
   if Generalquery.fields[16].text='Y' then addFeature('ZF',clid);
   if Generalquery.fields[17].text='Y' then addFeature('ZG',clid);
   if Generalquery.fields[18].text='Y' then addFeature('ZH',clid);
   if Generalquery.fields[19].text='Y' then addFeature('ZI',clid);
   if Generalquery.fields[20].text='Y' then addFeature('ZJ',clid);
   if Generalquery.fields[21].text='Y' then addFeature('ZK',clid);
   if Generalquery.fields[22].text='Y' then addFeature('ZL',clid);
   if Generalquery.fields[23].text='Y' then addFeature('ZM',clid);
   if Generalquery.fields[24].text='Y' then addFeature('ZN',clid);
   if Generalquery.fields[25].text='Y' then addFeature('ZO',clid);
   if Generalquery.fields[26].text='Y' then addFeature('ZP',clid);
   if Generalquery.fields[27].text='Y' then addFeature('ZQ',clid);
   if Generalquery.fields[28].text='Y' then addFeature('ZR',clid);
   if Generalquery.fields[29].text='Y' then addFeature('ZS',clid);
   if Generalquery.fields[30].text='Y' then addFeature('ZT',clid);
   if Generalquery.fields[31].text='Y' then addFeature('ZU',clid);
  End
  else
  begin
   {FeatureNode:=Treeview1.items.AddChild(MPANnode,'Services && Network Features - None');
   FeatureNode.imageindex:=93;
   FeatureNode.selectedindex:=93; }

   FeatureNode:=Treeview1.Addchild(mpannode);
   nodeData := Treeview1.GetNodeData(FeatureNode);
   NodeData.caption := 'Services && Network Features - None';
   nodedata.index:=93;

  end;
  // Get Friends & Family
  with Generalquery do
  Begin
   close;
   deletevariables;
   DeclareVariable('CLID', otstring);
   sql.clear;
   sql.add('select * from telecoms.friends_and_family_span');
   sql.add('Where span = :CLID');
   sql.add('order by effective_from desc,last_updated desc');
   setvariable('CLID',CLID);
   open;
   deletevariables;
  end;
  if GeneralQuery.recordcount=0 then
  Begin
   {FeatureItemNode:=Treeview1.items.AddChild(Featurenode,'Friends && Family - None');
   FeatureItemNode.imageindex:=112;
   FeatureItemNode.selectedindex:=112; }

   FeatureItemNode:=Treeview1.Addchild(FeatureNode);
   nodeData := Treeview1.GetNodeData(FeatureItemNode);
   NodeData.caption := 'Friends && Family - None';
   nodedata.index:=112;

   exit;
  End
  else
  Begin
   EFD:=generalquery.fields[1].text;
   {FeatureItemNode:=Treeview1.items.AddChild(Featurenode,'Friends & Family - Effective From '+efd);
   FeatureItemNode.imageindex:=112;
   FeatureItemNode.selectedindex:=112;}

   FeatureItemNode:=Treeview1.Addchild(FeatureNode);
   nodeData := Treeview1.GetNodeData(FeatureItemNode);
   NodeData.caption := 'Friends & Family - Effective From '+efd;
   nodedata.index:=112;

   Showfriend(generalquery.fields[2].text,Generalquery.fields[3].text,generalquery.fields[4].text);
   Showfriend(generalquery.fields[5].text,Generalquery.fields[6].text,generalquery.fields[7].text);
   Showfriend(generalquery.fields[8].text,Generalquery.fields[9].text,generalquery.fields[10].text);
   Showfriend(generalquery.fields[11].text,Generalquery.fields[12].text,generalquery.fields[13].text);
   Showfriend(generalquery.fields[14].text,Generalquery.fields[15].text,generalquery.fields[16].text);
   Showfriend(generalquery.fields[17].text,Generalquery.fields[18].text,generalquery.fields[19].text);
   Showfriend(generalquery.fields[20].text,Generalquery.fields[21].text,generalquery.fields[22].text);
   Showfriend(generalquery.fields[23].text,Generalquery.fields[24].text,generalquery.fields[25].text);
   Showfriend(generalquery.fields[26].text,Generalquery.fields[27].text,generalquery.fields[28].text);
   Showfriend(generalquery.fields[29].text,Generalquery.fields[30].text,generalquery.fields[31].text);
   Showfriend(generalquery.fields[32].text,Generalquery.fields[33].text,generalquery.fields[34].text);
   Showfriend(generalquery.fields[35].text,Generalquery.fields[36].text,generalquery.fields[37].text);
   Showfriend(generalquery.fields[38].text,Generalquery.fields[39].text,generalquery.fields[40].text);
   Showfriend(generalquery.fields[41].text,Generalquery.fields[42].text,generalquery.fields[43].text);
   Showfriend(generalquery.fields[44].text,Generalquery.fields[45].text,generalquery.fields[46].text);
  end;
end;


procedure TFRM_Tree.MaintainPremiseDetails1Click(Sender: TObject);
Var
Agreement_id,Customer_Id,premise_id:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Agreement_id:=TreeData.D_agreement_id;
 Customer_id:=TreeData.D_Customer_id;
 Premise_id:=TreeData.D_Premise_id;

 FRM_Premise_Details.clearfields;
 FRM_Premise_Details.agreementid.text:=Agreement_id;
 with FRM_Premise_Details.premisequery do
 Begin
  close;
  setvariable('Customerid',Customer_id);
  setvariable('aggreementid',agreement_id);
  open;
  FRM_Premise_Details.sitelookup.enabled:=true;
  FRM_Premise_Details.BtnAddSite.enabled:=true;
  if FRM_Premise_Details.premisequery.recordcount>1 then
  Begin
   FRM_Premise_Details.premisequery.first;
   Repeat
    if FRM_Premise_Details.premiseid.text=premise_id then break;
    FRM_Premise_Details.premisequery.next;
   until FRM_Premise_Details.premiseid.text=premise_id;
  end; // End Multiple Premise
 End;
 FRM_Premise_Details.sitelookup.keyvalue:=FRM_Premise_Details.premisequery.fields[14].text;
 with FRM_Premise_Details.premiseTypequery do
 Begin
  close;
  open;
 End;
 with FRM_Premise_Details.Regionquery do
 Begin
  close;
  open;
 End;
 // Now Select The Premise
 FRM_Premise_Details.btnaddsite.enabled:=false;
 FRM_Premise_Details.btnaddsite.visible:=false;
 FRM_Premise_Details.premisecontrol.activepageindex:=0;
 FRM_Premise_Details.showmodal;
   // refresh customer node;
  if treeview1.Selected[xnode]=true then
  Begin
   treeview1.Expanded[xnode]:=false;
   treeview1.Expanded[xnode]:=true;
  end;
end;

procedure TFRM_Tree.MenuItem2Click(Sender: TObject);
begin
 if treeupdating=true then exit;
 //mpannode:=treeview1.selected;
 mpannode:=treeview1.FocusedNode;
 nodeData := treeview1.GetNodeData(mpannode);

 mpan:=nodedata.D_SPAN;
 FRM_Gas_History.mprnlookup.visible:=false;
 Frm_gas_history.MPRNEdit.Visible:=true;
 FRM_Gas_History.mprnedit.text:=mpan;
 if FRM_Gas_History.caption='' then
 Begin
  messagedlg('There is no Transaction History for this MPRN',MTinformation,[MBOK],0);
  exit;
 end;
 FRM_Gas_History.show;
end;

procedure TFRM_Tree.G_ReOrderClick(Sender: TObject);
begin
ReOrder('4');
end;

procedure TFRM_Tree.Reorder(OStatus:String);
Var
SPAN,Status,Regid,NewSSD,NewRegiD,GSTATUS:String;
result:Integer;
begin
 if treeupdating=true then exit;
 //mpannode:=treeview1.selected;

  mpannode:=treeview1.FocusedNode;
  nodeData := treeview1.GetNodeData(mpannode);

  STATUS:=nodedata.D_STATUS;
  SPAN:=nodedata.D_SPAN;
  REGID:=nodedata.D_REGID;

  if (STATUS<>'Flunked') and (status<>'Cancelled') then
  Begin
   Messagedlg('Services Can only be Re-Ordered if existing Status is Flunked, or Cancelled',MTinformation,[MBOK],0);
   exit;
  End;
  // Check Span Markers
  if frm_common.is_dap(regid,'')=true then exit;

  // Now do a check to see if Gas large Site
  GStatus:='X';
  with main_data_module.tempquery do
  begin
   Close;
   sql.clear;
   sql.add('select g_large_site from crm.spans where registration_id='+regid);
   open;
  end;
  if Main_Data_Module.tempquery.fields[0].text='Y' then
  Begin
   repeat
   result:=0;
   with CreateMessageDialog('You are Re-Ordering a Gas LARGE Supply Point. Please choose process.', mtConfirmation,[mbyes,mbno,mbok]) do
    try
    TButton(FindComponent('Yes')).Caption :='Enquiry';
    TButton(FindComponent('No')).Caption := 'Nomination';
    TButton(FindComponent('Ok')).Caption := 'Register';
    Position := poScreenCenter;
    Result := ShowModal;
    if result=6 then GStatus:='E';
    if result=7 then GStatus:='N';
    if result=1 then GStatus:='';
    finally
    Free;
    end;
    until (result= 1) or (result= 6) or( result=7);
  end;

  if GSTATUS='X' then if Messagedlg('Are you sure you wish to re order this service?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
  if GSTATUS='E' then if Messagedlg('Are you sure you wish to re order this Large Gas Supply point as a ENQUIRY request?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
  if GSTATUS='N' then if Messagedlg('Are you sure you wish to re order this Large Gas Supply point as a NOMINATION request?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
  if GSTATUS=''  then if Messagedlg('Are you sure you wish to re order this Large Gas Supply point as a REGISTRATION request?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;

 if frm_common.authoritycheck=false then exit;

 // Get A New Registration date // not Past

  frm_batch_calendar.batchruns.close;
  frm_batch_calendar.batchruns.open;
  frm_batch_calendar.a_ssd.visible:=true;
  frm_batch_calendar.a_ssd.text:='  /  /    ';
  frm_batch_calendar.tag:=0;
  frm_batch_calendar.showmodal;
  frm_batch_calendar.a_ssd.visible:=true;
  if frm_batch_calendar.tag=1 then
  Begin
   Newssd:=frm_batch_calendar.batchruns.fields[4].text;
   if frm_batch_calendar.a_ssd.text<>'  /  /    ' then
   newssd:=frm_batch_calendar.a_ssd.text;;
  end
  else Exit;

  NEWREGID:=frm_common.nextregistrationid ;

  // REORDER SAME REGISTRATION
  if (GSTATUS='')  or (GSTATUS ='N') then
  begin
  // Update status back to Order ready
  with main_data_module.UpdateQuery do
  begin
   Close;

   sql.clear;
   sql.add('update crm.spans set span_start_date=to_date('''+newssd+''',''DD/MM/YYYY''),order_status_id=4 where registration_id='+regid);
   execute;
   Close;
   sql.clear;
   sql.add('update crm.spans_gas_extra set status='''+GSTATUS+''' where registration_id='+RegiD);
   execute;;
  end;
  end
  else
   // REORDER AS NEW REGISTRATION
  begin
   with main_data_module.updatequery do
   Begin
    close;
    sql.clear;
    sql.add('Insert into crm.spans');
    sql.add('select ');
    sql.add('SERVICE_ID,');
    sql.add(NEWREGID+',');
    sql.add('SPAN,');
    sql.add('SPAN_TYPE_ID,');
    sql.add('to_date('''+newssd+''',''DD/MM/YYYY''),');
    sql.add('null,');
    sql.add('null,');
    sql.add('SPAN_ADDRESS_1,');
    sql.add('SPAN_ADDRESS_2,');
    sql.add('SPAN_ADDRESS_3,');
    sql.add('SPAN_ADDRESS_4,');
    sql.add('SPAN_ADDRESS_5,');
    sql.add('SPAN_ADDRESS_6,');
    sql.add('SPAN_ADDRESS_7,');
    sql.add('SPAN_ADDRESS_8,');
    sql.add('SPAN_ADDRESS_9,');
    sql.add('SPAN_POSTCODE,');
    sql.add('RELATED,');
    sql.add('SERVICE_PRIOITY_NEEDS,');
    sql.add('''ORDER READY'',');
    sql.add('E_PROFILE_CLASS,');
    sql.add('E_MTC,');
    sql.add('E_LLF,');
    sql.add('E_SSC,');
    sql.add('E_ENERGISATION_STATUS,');
    sql.add('E_COT,');
    sql.add('E_MEASUREMENT_CLASS,');
    sql.add('E_NEW_CONNECTION,');
    sql.add('E_GSP_GROUP_ID,');
    sql.add('E_COMMS_METHOD,');
    sql.add('E_REGULAR_READING_CYCLE,');
    sql.add('E_MAX_POWER_REQUIREMENT,');
    sql.add('E_DA_ID,');
    sql.add('E_DA_EFD,');
    sql.add('E_DC_ID,');
    sql.add('E_DC_EFD,');
    sql.add('E_MO_ID,');
    sql.add('E_MO_EFD,');
    sql.add('E_EAC,');
    sql.add('T_BT_ACCOUNT_NO,');
    sql.add('T_OUR_REF,');
    sql.add('T_CHANGE_NUMBER,');
    sql.add('T_THREE_WAY_CALLING,');
    sql.add('T_ADDITIONAL_INFO,');
    sql.add(ostatus+',');
    sql.add('G_POST_CODE_OUT,');
    sql.add('G_POST_CODE_IN,');
    sql.add('G_METERID,');
    sql.add('G_LARGE_SITE,');
    sql.add('G_ZONE,null,G_TRANSPORTER,SALES_REFERENCE, SALES_REFERENCE_STATUS,null,null,null');
    sql.add('from crm.spans where registration_id='+regid);
    execute;
   End;
   //Call this function  that will add extra span TreeDataif Gas Registration. NEXUS
   FRM_COMMON.execute_Oracle_Procedure('CRM.PR_ADD_GAS_SPANS_EXTRA('+NEWREGID+')');
   with main_data_module.updatequery do
   begin
    Close;
    sql.clear;
    sql.add('update crm.spans_gas_extra set status='''+GSTATUS+''' where registration_id='+NewRegiD);
    execute;
   end;
  end;


  FRM_Login.MainSession.commit;
  Messagedlg('Refresh Tree to view changes',Mtinformation,[MBOK],0);
  if treeview1.Selected[mpannode]=true then
  Begin
   treeview1.Selected[mpannode]:=false;
   treeview1.Selected[mpannode]:=true;
  end;
end;

procedure TFRM_Tree.E_Re_OrderClick(Sender: TObject);
begin
if frm_common.authoritycheck=false then exit;
ReOrder('4');
end;

procedure TFRM_Tree.T_ReOrderClick(Sender: TObject);
begin
ReOrder('5');
end;

procedure TFRM_Tree.MaintainServicesNetworkFeatures1Click(Sender: TObject);
Var
Clid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Clid:=trim(TreeData.D_SPAN);
 Application.CreateForm(TFRM_T_Services, FRM_T_Services);
 try
  FRM_T_Services.CLID.text:=clid;
  FRM_T_Services.GetCurrentStatus(clid);
  FRM_T_Services.showmodal;
 finally
  FRM_T_Services.release;
 end;
end;

procedure TFRM_Tree.AddFeature(code,span:string);
Var
ActiveDate,Feature:string;
iconindex:integer;
Begin
 With Featurequery do
 Begin
  close;
  sql.clear;
  sql.add('select description,icon_index from telecoms.services_and_features where item_id='''+code+'''');
  open;
 End;
 Feature:=featurequery.fields[0].text;
 try
  iconindex:=featurequery.fields[1].value;
 except
  iconindex:=95;
 end;
  // Get First Activation Date of Feature
 With Featurequery do
 Begin
  close;
  sql.clear;
  sql.add('select min(effective_from),max(last_updated) from telecoms.services_and_featureD_SPAN');
  sql.add('where '+code+'=''Y''');
  sql.add('and span='''+span+'''');
  open;
 End;
 if featurequery.fields[0].text<>'' then ActiveDate:=' - '+FeatureQuery.fields[0].text
 else activedate:='';
 {FeatureitemNode:=Treeview1.items.AddChild(featureNode,feature+ActiveDate);
 Featureitemnode.imageindex:=iconindex;
 Featureitemnode.SelectedIndex:=Featureitemnode.imageindex; }

 FeatureitemNode:=Treeview1.Addchild(Featurenode);
 nodeData := Treeview1.GetNodeData(FeatureitemNode);
 NodeData.caption := feature+ActiveDate;
 nodedata.index:=iconindex;

end;

procedure TFRM_Tree.RefreshAgentPremiseNode;
Var
Customer_id:string;
Begin
 //mynodeagreement:=treeview1.selected;
 mynodeagreement:=treeview1.FocusedNode;
 nodeData := treeview1.GetNodeData(mynodeagreement);

 Customer_id:=nodedata.D_Customer_id;
 try
  //treeview1.selected.deletechildren;
  treeview1.deletechildren(mynodeagreement);
 except
 end;

 // Get Premises for customer
  with generalquery do
  Begin
   close;
   deletevariables;
   DeclareVariable('CID', otstring);
   sql.clear;
   sql.add('select Distinct');
   sql.add('APS.PREMISE_NAME,');
   sql.add('APS.SPECIAL_ACCESS,');
   sql.add('APS.PREMISE_CONTACT_ID,');
   sql.add('P.PREMISE_ID,');
   sql.add('T.ICON_INDEX,');
   sql.add('P.PREMISE_LINE_1,');
   sql.add('P.PREMISE_LINE_2,');
   sql.add('P.PREMISE_LINE_3,');
   sql.add('P.PREMISE_LINE_4,');
   sql.add('P.PREMISE_LINE_5,');
   sql.add('P.PREMISE_LINE_6,');
   sql.add('P.PREMISE_LINE_7,');
   sql.add('P.PREMISE_LINE_8,');
   sql.add('P.PREMISE_LINE_9,');
   sql.add('P.PREMISE_POSTCODE');
   sql.add('from crm.premises P,');
   sql.add('crm.agreements A,');
   sql.add('crm.agreement_premises APS,');
   sql.add('crm.premise_type t');
   sql.add('where A.Customer_id=:CID');
   sql.add('and ApS.agreement_id=A.Agreement_id');
   sql.add('and P.premise_id=APS.premise_id (+)');
   sql.add('and P.premise_type_id=t.premise_type_id (+)');
   setvariable('CID',customer_id);
   open;
   deletevariables;
  end;
  with generalquery do
  Begin
   while not eof do
   Begin
    premaddr:='';
    if fields[5].text<>'' then premaddr:=premaddr+fields[5].text+',';
    if fields[6].text<>'' then premaddr:=premaddr+fields[6].text+',';
    if fields[7].text<>'' then premaddr:=premaddr+fields[7].text+',';
    if fields[8].text<>'' then premaddr:=premaddr+fields[8].text+',';
    if fields[9].text<>'' then premaddr:=premaddr+fields[9].text+',';
    if fields[10].text<>'' then premaddr:=premaddr+fields[10].text+',';
    if fields[11].text<>'' then premaddr:=premaddr+fields[11].text+',';
    if fields[12].text<>'' then premaddr:=premaddr+fields[12].text+',';
    if fields[13].text<>'' then premaddr:=premaddr+fields[13].text+',';
    premaddr:=premaddr+fields[14].text;
    premaddr:=premaddr+' - ['+fields[3].text+']';
   { mynodepremise:=Treeview1.items.AddChild(mynodeagreement,'Premises - '+premaddr);
    mynodepremise.imageindex:=fields[4].value;
    mynodepremise.selectedindex:=mynodepremise.imageindex; }

    mynodepremise:=Treeview1.Addchild(mynodeagreement);
    nodeData := Treeview1.GetNodeData( mynodepremise);
    NodeData.caption := 'Premises - '+premaddr;
    nodedata.index:=fields[4].value;

    nodedata.D_premise_Id :=   fields[3].text;
    nodedata.D_agreement_Id := agreement_id;
    nodedata.D_customer_Id := customer_id;
    //mynodepremise.data:=MyRecPtr;



    //mynode1:=Treeview1.items.AddChild(mynodepremise,'test');
    mynode1:=Treeview1.Addchild(mynodepremise);
    nodeData := Treeview1.GetNodeData( mynode1);
    NodeData.caption := 'test';



    next;
   end;
  end;
end;

procedure TFRM_Tree.FriendsFamily1Click(Sender: TObject);
Var
Clid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Clid:=trim(TreeData.D_SPAN);
 Application.CreateForm(TFRM_FAF, FRM_FAF);
 try
  FRM_FAF.CLID.text:=clid;
  FRM_FAF.GetRecomended;
  FRM_FAF.showmodal;
 finally
  FRM_FAF.release;
 end;
end;

procedure TFRm_Tree.ShowFriend(No,IType,Bf:string);
Var
Desc:string;
Begin
 if no='' then exit;
 Desc:=No;
 desc:=desc+' - '+frm_common.getstdarea(no);
 //FeatureitemSubNode:=Treeview1.items.AddChild(featureItemNode,desc);

 FeatureitemSubNode:=Treeview1.Addchild(featureItemNode);
 nodeData := Treeview1.GetNodeData( FeatureitemSubNode);
 NodeData.caption := desc;


 if Itype='U' then nodedata.index:=122; // Friend
 if Itype='I' then
 Begin
  try
   nodedata.index:=main_data_module.intlareaquery.fields[1].value;
   if main_data_module.intlareaquery.fields[2].text<>'' then  nodedata.caption:=nodedata.caption+' - ('+main_data_module.intlareaquery.fields[2].text+')';
  except
   nodedata.index:=115; // Intl
  end;
 end;
 if Itype='M' then NodedAta.index:=116; // Mobile
 if BF='Y' then
 Begin
  NOdedata.index:=114; // Best Friend
  nodedata.caption:=nodedata.caption+' - (Best Friend)';
 end;
// FeatureitemSubnode.SelectedIndex:=FeatureitemSubnode.imageindex;
End;

function TFRM_Tree.ShowInvoluntaryModeChangePopUp(const aCustomerId: Int64): Integer;
const
  INVOLUNTARY_MESSAGE = 'THIS IS AN INVOLUNTARY METER MODE SWITCH (IMMS) ACCOUNT. ENSURE THAT A FULL AND THOROUGH SAFE AND REASONABLY PRACTICABLE (S&&RP) ASSESSMENT IS COMPLETED AND IF NOT SAFE THEN FOLLOW GUIDANCE ON HELPJUICE';
begin
  if CustomerHasInvoluntaryModeChangeFlag(aCustomerid) then
    Result := TCrmUtil.ShowMessageDialog(INVOLUNTARY_MESSAGE, mtInformation, [mbOK], clRed, [fsBold]);
end;

procedure TFRM_Tree.QuarterlyStatement1Click(Sender: TObject);
Var
DCLID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 DCLID:=trim(TreeData.D_SPAN);
 FRM_Reports.PrintThisReport('Telecoms\TELECOM_STATEMENT_SUMMARY.rpt','Telecom Statement','{TELECOMS_SUMMARY.SPAN}='''+DCLID+'''','','','','');
end;

procedure TFRM_Tree.QuarterlyStatementItemised1Click(Sender: TObject);
Var
DCLID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 DCLID:=trim(TreeData.D_SPAN);
 FRM_Reports.PrintThisReport('Telecoms\TELECOM_STATEMENT_SUMMARY_ITEMISED.rpt','Telecom Statement Itemised','{TELECOMS_SUMMARY.SPAN}='''+DCLID+'''','','','','');
end;


procedure TFRM_Tree.ErroneousTransfer1Click(Sender: TObject);
Var
Custid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Custid:=TreeData.D_Customer_ID;
 FRM_Reports.PrintThisReport('CRM\ET Reports\ET_Electric_Letter.rpt','Electricity - Erroneous Transfer','{ACCOUNT_HOLDERS.CONTACT_ORDER}=1 and {CUSTOMER.CUSTOMER_ID}='+custid+'','','',custid,'');
end;

procedure Tfrm_tree.ShowAgreement(mynodeagreement:Pvirtualnode;Iconindex:integer;AID,CID,Astatus,Astart,Aend,Aperiod:String;sc:boolean;Renewaldate:string);
var
dys:real;
rd:string;
begin
 nodeData := treeview1.GetNodeData(mynodeagreement);

 desc:='Agreement '+Aid;
 desc:=desc+', '+astatus;
 desc:=desc+', Start '+aStart;
 if Aperiod='0' then desc:=desc+', (Open Ended)' else
 desc:=desc+', (Initial Period '+Aperiod+')';
 if agreement_status_id='3' then desc:='Agreement '+aid+', Start '+aStart+' - '+astatus+' on '+Aend;

 // If No end Date on Agreement then show renewal Date
 if aend='' then
 Begin
  if renewaldate<>'' then
  Begin
   desc:=desc+' - Renewal Date ('+renewaldate+')';
   // DANBYT - 31/10/2024 - CRMX-123 - Changed StrToDate to StrToDateTime
   dys:=StrToDateTime(renewaldate)-date;
   if dys<=0 then rd:=' OVERDUE';
   if (dys>0) and (dys<=28) then rd:=' Renewal Due';
   if dys>14 then rd:='';
   desc:=desc+rd;
  end;
 end;

  if (astart='31/12/2099') and (agreement_status_id='0') then
 Begin
  desc:='Agreement '+Aid;
  desc:=desc+', PENDING - Installation Date Not Yet Booked';
  astatus:='Pending';
 End;

 desc:=desc+TCrmUtil.CheckContractAttributes(StrToInt64(Aid));

 nodedata.fontcolor:=frm_common.Orderstatuscolor(astatus);
 Nodedata.caption:=desc;

 if (ICONINDEX=71) AND (COPY(AID,1,10)<>cid)  then ICONINDEX:=275;
 if (ICONINDEX=88) AND (COPY(AID,1,10)<>cid)  then ICONINDEX:=276;

 nodedata.index:=iconindex;
 //mynodeagreement.selectedindex:=mynodeagreement.imageindex;


 // Pending Agreement - Change



 nodedata.D_agreement_id := aid;
 nodedata.D_agreement_end_date := aEnd;
 nodedata.D_agreement_Start_date := astart;
 nodedata.D_Customer_id := cid;
 //nodedata.CheckType:=virtualtrees.ctCheckBox;
 //Mynodeagreement.data:=MyRecPtr;


 if sc=true then
 Begin
 // mynodeAgreementItem:=Treeview1.items.AddChild(mynodeagreement,'dummy');

  mynodeAgreementItem:=Treeview1.Addchild(mynodeagreement);
  nodeData := Treeview1.GetNodeData(mynodeAgreementItem);
  NodeData.caption := 'Dummy'

 end;


end;

procedure TFRM_TREE.ShowProduct(mynodeagreement:Pvirtualnode;Agreement_id:string;ShowBD:Boolean;Fstatus:string);
Var
//MyRecPtr: PMyRec;
 status,effective_from,effective_to,paymentplan,collectionrate,
 productid,getbname,nextdddate,bundlecode:string;
 atpid: integer;
 brokenDate: Tdate;
Begin

 nodeData := treeview1.GetNodeData(mynodeagreement);

 // Get Products for Agreement/Premise
 with productquery do
 Begin
  deletevariables;
  DeclareVariable('AID', otlong);
  setvariable('AID',agreement_id);
  close;
  sql.clear;
  sql.Add('select');
  sql.add('AP.AGREEMENT_ID,');
  sql.add('AP.GETB_TYPE,');
  sql.add('PR.Description,');
  sql.Add('AP.EFFECTIVE_FROM,');
  sql.Add('AP.EFFECTIVE_TO,');
  sql.Add('OS.Description,');
  sql.add('AP.ADDITIONAL_INFORMATION,');
  sql.Add('PP.DESCRIPTION,');
  sql.add('PP.ICON_INDEX,');
  sql.Add('AP.DIRECT_DEBIT_FIRST_DATE,');
  sql.Add('AP.COLLECTION_RATE,');
  sql.Add('AP.CEF_EST_COST_GAS,');
  sql.Add('AP.CEF_EST_COST_ELECTRIC,');
  sql.Add('AP.CEF_EST_COST_TELECOMS,');
  sql.Add('AP.CEF_EST_COST_BROADBAND,');
  sql.Add('AP.CEF_TOTAL,');
  sql.add('AP.CREDITS_TYPE,');
  sql.add('AP.DATE_SETUP,');
  sql.add('PPL.ALIAS,');
  sql.add('A.SALES_REFERENCE,');

  sql.add('AP.CASHBACk_TYPE,');
  sql.add('AP.CREDITS_TYPE,');
  sql.add('AP.PROMOTIONS_TYPE,');
  sql.add('AP.PRICE_PLAN_ID,');
  sql.add('AP.PAYMENT_PLAN_ID,');

  sql.add('c.agreement_id,c.comments,c.added_by,c.date_added');

  sql.add('from CRM.AGREEMENT_PRODUCTS AP,');
  sql.add('crm.dd_suppress_catchups c,');
  sql.add('CRM.order_status OS,');
  sql.add('BILLING.TARIFF_CODE_1 PR,');
  sql.add('BILLING.TARIFF_CODE_3 PP,');
  sql.add('BILLING.TARIFF_CODE_5 PPL,');
  //sql.add('CRM.BUNDLE_CODE_4 CB,');
  //sql.add('CRM.BUNDLE_CODE_2 CGL,');
  //sql.add('CRM.BUNDLE_CODE_6 PM,');
  sql.add('CRM.AGREEMENTS A');
  sql.add('where AP.agreement_id=:AID');
  sql.add('and AP.ORDER_STATUS_id=OS.ORDER_STATUS_id (+)');
  sql.add('and AP.GETB_TYPE=PR.CODE (+)');
  sql.add('and AP.PAYMENT_PLAN_ID=PP.CODE (+)');
  sql.add('and AP.PRICE_PLAN_ID=PPL.CODE (+)');
  sql.add('and AP.AGREEMENT_ID=A.AGREEMENT_ID (+)');
  sql.add('and ap.agreement_id=c.agreement_id (+)');
  sql.add('order by AP.DATE_SETUP desc');
  open;
  deletevariables;
  first;
 end;

 // Get Next DD amunt to be collected and Date of Collection for Agreement
 with schedulequery do
 Begin
  deletevariables;
  DeclareVariable('AID', otlong);
  setvariable('AID',agreement_id);
  close;
  sql.clear;
  sql.add('select collection_date,collection_amount,order_status_id from crm.dd_request_schedule');
  sql.add('where agreement_id=:AID');
  sql.add('and (order_status_id=4 or ');
  sql.add('order_status_id=5 or ');
  sql.add('order_status_id=6 or ');
  sql.add('order_status_id=9) ');
  sql.add('and collection_date>=sysdate');
  sql.add('order by collection_date asc');
  open;
  deletevariables;
 End;

 nextdddate:=schedulequery.fields[0].text;

 // Display Product Details
 if productquery.recordcount=0 then
 Begin

  desc:='No Products Defined';
  nodedata.caption:=nodedata.caption+#10+desc;

  try
   paymentplan:=productquery.fields[7].text;
   if productquery.fields[10].text='' then collectionrate:=''
   else CollectionRate:=chr(163) + ' ' +Formatfloat('0.00', productquery.fields[10].value)+'p';
   desc:='Payment Plan - '+paymentplan;
   if (productquery.fields[24].text='D') or (productquery.fields[24].text='Q') or (productquery.fields[24].text='S') then desc:=desc+'     Collection Rate  ['+collectionrate+']';
   //mynodepaymentplan:=Treeview1.items.AddChild(mynodeagreement,desc);

   mynodepaymentplan:=Treeview1.Addchild(mynodeagreement);
   nodeData := Treeview1.GetNodeData(mynodepaymentplan);
   nodedata.caption:=desc;



   Paymenticon:=productquery.fields[8].value;
   // if DD and HelpCo then Change to show Help CO DD ICON
   if (paymenticon=89) and (copy(productquery.fields[19].text,9,3)='SPK') then paymenticon:=202
   else
   if (paymenticon=89) and (copy(productquery.fields[19].text,9,3)='HLP') then paymenticon:=179
   else
   nodedata.index:=PaymentIcon;
   //mynodepaymentplan.selectedindex:=mynodepaymentplan.imageindex;
  except
   nodedata.caption:='No Payment Plan Defined';
   nodedata.index:=73;
   //mynodepaymentplan.selectedindex:=mynodepaymentplan.imageindex;
  end;

  a_prodwizard.visible:=true;
 end

 else
 Begin
  a_prodwizard.visible:=false;
  productid:=productquery.fields[1].text;
  GETBname:=productquery.fields[2].text;
  effective_from:=productquery.fields[3].text;
  effective_to:=productquery.fields[4].text;
  Status:=productquery.fields[5].text;

  bundlecode:='';
  bundlecode:=bundlecode+productquery.fields[1].text;  // GetB Type;
  bundlecode:=bundlecode+productquery.fields[21].text; // Credits Type;
  bundlecode:=bundlecode+productquery.fields[24].text; // Payment Plan;
  bundlecode:=bundlecode+productquery.fields[20].text; // Cashback;
  bundlecode:=bundlecode+'N';                          // Promotion;
  bundlecode:=bundlecode+productquery.fields[23].text; // Tariff;

  desc:='Product Bundle - '+frm_common.decodebundlecode(BUNDLECODE);

  // Smart Pay
  if IsSmartPay(agreement_id) then
    desc := desc + ' - Power Pay';

  if effective_from<>'' then desc:=desc+#10+'Product Effective From '+effective_from;
  desc:=desc+'. Status ['+status+']';

  nodedata.caption:=nodedata.caption+#10+desc;


  paymentplan:=productquery.fields[7].text;
  if productquery.fields[10].text='' then collectionrate:=''
  else CollectionRate:=chr(163) + ' ' +Formatfloat('0.00', productquery.fields[10].value)+'p';

  if paymentplan<>'Non Specified' then
  Begin
   desc:='Payment Plan - '+paymentplan;
   if (productquery.fields[24].text='D') or (productquery.fields[24].text='Q') or (productquery.fields[24].text='S') then desc:=desc+'     Collection Rate  ['+collectionrate+']';
   //mynodepaymentplan:=Treeview1.items.AddChild(mynodeagreement,desc);

   mynodepaymentplan:=Treeview1.Addchild(mynodeagreement);
   nodeData := Treeview1.GetNodeData(mynodepaymentplan);
   NodeData.caption := desc;



   try
   Paymenticon:=productquery.fields[8].value;
   except
    paymenticon:=8;
   end;
   // if DD and HelpCo then Change to show Help CO DD ICON
   if (paymenticon=89) and (copy(productquery.fields[19].text,9,3)='HLP') then paymenticon:=179
   else if (paymenticon=89) and (copy(productquery.fields[19].text,9,3)='SPK') then paymenticon:=202;
   nodedata.index:=paymenticon;
   //mynodepaymentplan.selectedindex:=mynodepaymentplan.imageindex;
  end
  else
  Begin
   {mynodepaymentplan:=Treeview1.items.AddChild(mynodeagreement,desc);
   mynodepaymentplan.text:='No Payment Plan Defined';
   mynodepaymentplan.imageindex:=73;
   mynodepaymentplan.selectedindex:=mynodepaymentplan.imageindex; }

   mynodepaymentplan:=Treeview1.Addchild(mynodeagreement);
   nodeData := Treeview1.GetNodeData(mynodepaymentplan);
   NodeData.caption := desc;
   nodedata.index:=73;

  End;

  // Show latest Bank Detail DD request under payment plan
  // Get bank Details
  if (ShowBD=true) then
  Begin
   with generalquery do
   Begin
    close;
    deletevariables;
    DeclareVariable('AID', otlong);
    sql.clear;
    sql.Add('select');
    sql.Add('O.DESCRIPTION,o.order_status_id,B.direct_debit_signed');
    sql.add('from CRM.AGREEMENT_BANK_DETAILS B,');
    sql.add('CRM.order_status O');
    sql.add('where B.agreement_id=:AID');
    sql.add('and B.bank_details_status_id=O.order_status_id (+)');
    sql.add('order by B.effective_from desc,b.effective_to desc');
    setvariable('AID',agreement_id);
    open;
    deletevariables;
   end;
   // Format bank Details
   if generalquery.fields[0].text<>'' then
   Begin
    desc:='Direct Debit Instruction Status = '+generalquery.fields[0].text;
    if generalquery.fields[2].text='Y' then desc:=desc+' (Signed)'
    else desc:=desc+' (NOT SIGNED)';
   end
   else
   Begin
    if Paymentplan='Direct Debit' then desc:='No Bank Details exist'
    else desc:='';
   end;

   if nextdddate<>'' then
   Begin
    desc:=desc+' - Next DD Collection = ' + chr(163) + ' ' +Formatfloat('0.00', schedulequery.fields[1].value)+'p';
    desc:=desc+' - on '+nextdddate;
   End;
    if desc<>'' then nodedata.caption:=nodedata.caption+#10+desc;
    nodedata.fontcolor:=frm_common.Orderstatuscolor(generalquery.fields[0].text);
  end;


  nodedata.D_Agreement_id:=agreement_id;
  //mynodepaymentplan.data:=MyRecPtr;

  if productquery.fields[25].text<>'' then
  Begin
   {mynodesuppress:=Treeview1.items.AddChild(mynodeagreement,desc);
   mynodeSuppress.text:='DD Catch Ups Suppressed - Marker added by '+productquery.fields[27].text+' on '+productquery.fields[28].text;
   //+#10+productquery.fields[26].text;
   mynodeSuppress.font.color:=clpurple;
   mynodeSuppress.font.style:=[fsbold];
   mynodeSuppress.imageindex:=168;
   mynodeSuppress.selectedindex:=mynodeSuppress.imageindex;}


  // Mynodesuppress.data:=MyRecPtr;

   mynodesuppress:=Treeview1.Addchild(mynodeagreement);
   nodeData := Treeview1.GetNodeData(mynodesuppress);
   NodeData.caption := 'DD Catch Ups Suppressed - Marker added by '+productquery.fields[27].text+' on '+productquery.fields[28].text;
   nodedata.index:=168;
   NodeData.fontcolor:=clpurple;
   NodeData.fontBold:=true;
   NodeData.D_agreement_id := agreement_id;
  end;

 end;
 if Fstatus<>'' then
 Begin
  nodeData := Treeview1.GetNodeData(mynodeagreement);
  nodedata.caption:=nodedata.caption+#10+fStatus;
 end;

 with ATPQuery do                           //checks if cust has payment plan active
 begin
   Close;
   DeleteVariables;
   DeclareVariable('AGID',otLong);
   Sql.Clear;
   Sql.Add('SELECT');
   Sql.Add('ATP.AGREEMENT_ID,');
   Sql.Add('PAYMETHOD.DESCRIPTION AS PAYMENT,');
   Sql.Add('FREQ.DESCRIPTION AS FREQUENCY,');
   Sql.Add('SCHED.AMOUNT_DUE,');
   Sql.Add('ATP.NO_OF_PAYMENTS,');
   Sql.Add('ATP.STATUS,');
   Sql.Add('ATP.ID,');
   Sql.Add('(SELECT COUNT(*) FROM CRM.ARRANGEMENT_TO_PAY_SCHEDULE WHERE STATUS LIKE ''%Uncollected%'' AND ID_ARRANGEMENT_TO_PAY =ATP.ID AND CATCH_UP IS NULL)NUM_LEFT');
   Sql.Add('FROM CRM.ARRANGEMENT_TO_PAY ATP');
   Sql.Add('JOIN CRM.ARRANGEMENT_TO_PAY_METHOD PAYMETHOD ON ATP.PAYMENT_METHOD = PAYMETHOD.ID');
   Sql.Add('JOIN CRM.ARRANGEMENT_TO_PAY_FREQ FREQ ON ATP.FREQUENCY = FREQ.ID');
   Sql.Add('JOIN CRM.ARRANGEMENT_TO_PAY_SCHEDULE SCHED ON ATP.ID = SCHED.ID_ARRANGEMENT_TO_PAY');
   Sql.Add('WHERE ATP.AGREEMENT_ID =:AGID');
   Sql.Add('AND ATP.STATUS = 0');
   Sql.Add('AND SCHED.PAYMENT_ORDER = 1');
   Sql.Add('AND ROWNUM <= 1');
   SetVariable('AGID',StrToInt64(Agreement_ID));
   Open;
   DeleteVariables;
 end;

 if ATPQuery.RecordCount <> 0 then
 begin                                            //adds payment plan details to the tree using the original plan data- not taking into account breaches
   atpnode := Treeview1.AddChild(mynodeagreement);
   NodeData := Treeview1.GetNodeData(atpnode);
   nodedata.D_Agreement_id:=agreement_id;
   NodeData.caption := 'Payment Plan Agreed - ' + ATPQuery.FieldByName('NUM_LEFT').Text + ' x ' + chr(163) + ATPQuery.FieldByName('AMOUNT_DUE').text
   + ' ' + ATPQuery.FieldByName('FREQUENCY').Text + ' - ' + ATPQuery.FieldByName('PAYMENT').Text;
   NodeData.index := 300;
 end

 else
 begin
  with ATPQuery do                           //checks if cust has a broken payment plan
  begin
    Close;
    DeleteVariables;
    DeclareVariable('AGID',otLong);
    Sql.Clear;
    Sql.Add('SELECT');
    Sql.Add('ATP.AGREEMENT_ID,');
    Sql.Add('ATP.ID,');
    Sql.Add('ATP.STATUS,');
    Sql.Add('SCHED.DUE_DATE');
    Sql.Add('FROM CRM.ARRANGEMENT_TO_PAY ATP');
    Sql.Add('JOIN (SELECT * FROM CRM.ARRANGEMENT_TO_PAY_SCHEDULE ORDER BY DUE_DATE DESC) SCHED ON ATP.ID = SCHED.ID_ARRANGEMENT_TO_PAY');
    Sql.Add('WHERE SCHED.STATUS LIKE ''%Breached%''');
    Sql.Add('AND ATP.AGREEMENT_ID =:AGID');
    Sql.Add('AND ATP.STATUS = 7');
    Sql.Add('AND ATP.CANCELLATION_REASON IS NULL');
    Sql.Add('AND ROWNUM <= 1');
    SetVariable('AGID',StrToInt64(Agreement_ID));
    Open;
    DeleteVariables;
  end;
  if ATPQuery.RecordCount <> 0 then
  begin
   AtpId := StrToInt(ATPQuery.FieldByName('ID').Text);
   BrokenDate:= IncDay(StrToDate(ATPQuery.FieldByName('DUE_DATE').Text),14);
   if Date < BrokenDate then    //if last due date 14 in past then don't show icon on tree
   begin
     with ATPQuery do
     begin
       Close;
       DeleteVariables;
       DeclareVariable('ATPID',otInteger);
       Sql.Clear;
       Sql.Add('SELECT * FROM (SELECT *');
       Sql.Add('FROM CRM.ARRANGEMENT_TO_PAY_SCHEDULE WHERE ID_ARRANGEMENT_TO_PAY =:ATPID AND STATUS LIKE ''%Collected%''');
       Sql.Add('ORDER BY DUE_DATE DESC)');
       Sql.Add('WHERE ROWNUM <=1');
       SetVariable('ATPID',AtpId);
       Open;
       DeleteVariables;
     end;
     atpnode := Treeview1.AddChild(mynodeagreement);                //adds broken plan to tree
     NodeData := Treeview1.GetNodeData(atpnode);
     nodedata.D_Agreement_id:=agreement_id;
     if ATPQuery.RecordCount <> 0 then NodeData.caption := 'Broken Payment Plan - Last Payment Received: ' + ATPQuery.FieldByName('DUE_DATE').Text
     else NodeData.caption := 'Broken Payment Plan - Agreed Payment Not Received';
     NodeData.index := 301;
   end;
  end;
 end;
end;

procedure TFRM_Tree.showspan(spannode:Pvirtualnode;spandesc,status,ssd,Span,spantype:string;spanindex:integer;regid,servicetype,btacno,btssd,IGT,SPANEND,SPANEndReason,AGID,locked,premid,fl1,fl2,fl3,fl4,fl5,fl6,fl7,fl8,fl9:string; custtype: Integer);
Var
Liveicon:boolean;
Spandisplay,ssddisp,enstatus,ob,rel,DAP,THREEPHASE:string;
et,dnb,NC:boolean;
begin
 TreeData:= treeview1.GetNodeData(spannode);
 // Check Span Markers
 if FL1='N/A' then
 begin
 with main_data_module.generalquery do
 Begin
  close;
  deletevariables;
  DeclareVariable('RID', otlong);
  sql.clear;
  sql.add('select * from CRM.SPAN_REGISTRATION_FLAGS');
  sql.add('where regid=:RID');
  setvariable('RID',regid);
  open;
  deletevariables;
 End;
 if main_data_module.generalquery.fields[2].text='T' then NC:=true else NC:=false ;
 fl2:=main_data_module.generalquery.fields[3].text;
 if main_data_module.generalquery.fields[4].text='' then OB:='' else ob:=' ('+main_data_module.generalquery.fields[4].text+' OBJ)';
 if main_data_module.generalquery.fields[5].text='' then REL:='' else REL:=' (** Release DO NOT OBJECT **)';
 if main_data_module.generalquery.fields[6].text='' then ET:=false else ET:=true;
 if main_data_module.generalquery.fields[7].text='' then DNB:=false else DNB:=true;
 if main_data_module.generalquery.fields[8].text='' then DAP:='' else DAP:=' (** DEBT ASSIGNMENT PROCESS **)';
 try
 if main_data_module.generalquery.fields[9].text='' then THREEPHASE:='' else THREEPHASE:='Y';
 except
  ThreePhase:='';
 end;
 end
 else
 begin
  if FL1='T' then NC:=true else NC:=false ;
  if FL3='' then OB:='' else ob:=' ('+fl3+' OBJ)';
  if FL4='' then REL:='' else REL:=' (** Release DO NOT OBJECT **)';
  if FL5='' then ET:=false else ET:=true;
  if FL6='' then DNB:=false else DNB:=true;
  if FL7='' then DAP:='' else DAP:=' (** DEBT ASSIGNMENT PROCESS **)';
  if FL8='' then THREEPHASE:='' else THREEPHASE:='Y';

 end;


 if THREEPHASE='Y' then servicetype:=stringreplace(servicetype,'tricity',' 3PHASE',[rfreplaceall]);


 // Check If Elect De-Energised
 enstatus:='                    ';
 if (spantype='E') or (spantype='F') then if fl2='D' then enstatus:='DE-ENERGISED        ';

 if spantype='Y' then spandisplay:=copy(span,2,30)
 else spandisplay:=span;
 if ssd='' then ssddisp:=' WAITING  '
 else ssddisp:=ssd;

 desc:=copy('SPAN '+ENSTATUS,1,17);
 desc:=desc+'     '+copy(Spandisplay+'               ',1,15);            // SPAN  No
 desc:=desc+'        SSD = '+SSDdisp;                                    // SSD
 if SpanEnd<>'' then desc:=desc+' - ' +SpanEnd+' ';
 if spanend='' then desc:=desc+'         ';


 nodedata.fontcolor:=clblack;
 liveicon:=false;
 if (status='Live') then Liveicon:=true;
 nodedata.fontcolor:=frm_common.Orderstatuscolor(Status);

 if (spanend<>'')  and (SpanEndReason='13') then
 Begin
  Status:=Status+' (Vacating)';
 End;
 if (spanend<>'') and (SpanEndReason='14') then
 Begin
  Status:='Vacated';
  liveicon:=false;
  nodedata.fontcolor:=frm_common.OrderStatusColor('Vacated');
 End;

 if (spanend<>'') and (SpanEndReason='15') then
 Begin
  Status:='Supply LOST';
  liveicon:=false;
  nodedata.fontcolor:=frm_common.OrderStatusColor('Lost');
  //////////////////////////////////////////////////////////////////////////////
  // WRIKE 163823640: Spans already showing lost with future dates
  //////////////////////////////////////////////////////////////////////////////
  if strtodate(spanend)>now then SpanEndReason:='18';

 End;

 if (spanend<>'') and (SpanEndReason='18') then
 Begin
  Status:='Future Loss';
  liveicon:=false;
  nodedata.fontcolor:=frm_common.OrderStatusColor('Future Loss');
 End;


 if (spanend<>'') and (SpanEndReason='16') then
 Begin
  Status:='Disconnected';
  liveicon:=false;
  nodedata.fontcolor:=frm_common.OrderStatusColor('Disconnected');
 End;


 desc:=desc+Status;
 desc:=desc+ob;
 if locked='Y' then desc:=desc+' ('+chr(163)+')';

 nodedata.caption:=desc;
 nodedata.FontName:='Lucida Console';

 if liveicon=true then nodedata.index:=spanindex;

 if liveicon=false then
 Begin
  if (spantype='G') or (spantype='C') then nodedata.index:=67; // Dead Gas
  if (spantype='E') or (spantype='F') then nodedata.index:=4;  // Dead Electric
  if ((spantype='E') or (spantype='F')) and nc=true then nodedata.index:=222;  // Dead Electric NC
  if (spantype='T') or (spantype='J') then nodedata.index:=68; // Dead Tel
  if (spantype='Y') or (spantype='Y') then nodedata.index:=127;// Dead Broadband Tel
 End;

 if ((spantype='E') or (spantype='F')) then
 Begin
  if (enstatus[1]='D') then nodedata.index:=218;
  if (nc=true) and (liveicon=true) then nodedata.index:=221;  // Live Electric NC
  if rel='' then e_relo.visible:=false else e_relo.visible:=true;
  e_rel.visible:=not e_relo.visible;
  if dap='' then e_debt.visible:=true else e_debt.visible:=false;
  e_debtr.Visible:=not e_debt.visible;
 end;

 if (spantype='U') or (spantype='V') then nodedata.index:=225; // Fresh Water
 if (spantype='W') or (spantype='X') then nodedata.index:=224; // Grey Water
 if (spantype='A') or (spantype='B') then nodedata.index:=226; // Heat
 if (spantype='H') or (spantype='I') then nodedata.index:=227;  // Sub Gas
 if (spantype='K') or (spantype='L') then nodedata.index:=228;   // Sub Elec
 if (spantype='R') or (spantype='S') then nodedata.index:=231;   // Sub Elec
 if (spantype='M') then nodedata.index:=230;   // Sub Elec

 if ((spantype='G') or (spantype='C')) then
 begin
  if rel='' then g_relo.visible:=false else g_relo.visible:=true;
  g_rel.visible:=not g_relo.visible;
  if dap='' then g_debt.visible:=true else g_debt.visible:=false;
  g_debtr.Visible:=not g_debt.visible;
 end;

 //spannode.selectedindex:=nodedata.index;

 nodedata.D_SPAN :=   span;
 //data.D_SPAN_NC :=   span_NC;
 nodedata.D_REGID :=  regid ;
 nodedata.D_STATUS := Status;
 nodedata.D_SSD := ssd;
 nodedata.D_SPANEND := SPANEND;
 nodedata.D_SPANDESC := SPANDESC;
 nodedata.D_DESC :=   'SUPPLY';
 nodedata.D_SPANTYPE := spantype;
 nodedata.D_Agreement_id:= agid;
 nodedata.D_Premise_id:= premid;
 nodedata.D_ET:= ET;
 nodedata.D_Cust_Type := custtype;

 //Spannode.data:=MyRecPtr;
 nodedata.caption:=nodedata.caption+#10+copy(servicetype+'                    ',1,20);

 //if (spantype<>'T') and (spantype<>'J') then
 Begin
  nodedata.caption:=nodedata.caption+'  '+copy(spandesc+'                      ',1,28);
 end;

 if (spantype='T') or (spantype='J') then
 Begin
    if BTSSD<>ssd then nodedata.caption:=nodedata.caption+' '+BTSSD
 //   spannode.text:=spannode.text+'  A/C - '+btacno;
 end;

 if igt<>'' then nodedata.caption:=nodedata.caption+' '+IGT;

 if et=true then
 Begin
  nodedata.caption:=nodedata.caption+' ** ET **';
  nodedata.fontcolor:=clred;
 end;

 if DNB=true then
 Begin
  nodedata.caption:=nodedata.caption+' ** Do Not Bill Disconnected/De-Energised **';
  nodedata.fontcolor:=clred;
 end;

 if dap<>'' then nodedata.caption:=nodedata.caption+#10+dap;
 if rel<>'' then nodedata.caption:=nodedata.caption+#10+rel;

 // Enable or Disable IGT Registration FOrm
 if (spantype='G') or (spantype='C') then
 Begin
  if igt<>'' then
  Begin
   IGTMENU.enabled:=true;
   IGTMENU.visible:=true;
  end
  else
  Begin
   IGTMENU.enabled:=false;
   IGTMENU.visible:=False;
  End;
 end;

 //mynode1:=Treeview1.items.AddChild(spanNode,'Temp');
 mynode1:=Treeview1.Addchild(spanNode);
 nodeData := Treeview1.GetNodeData(mynode1);
 NodeData.caption := 'Temp';


end;


procedure TFRm_TREE.AddScannedDoc(custid,Filename,role:string);
Var
SrceName:String;
DestDir:String;
DestName,fext,nid,dnow,fmon,resolved:String;
begin
 if Select_File_To_Attach.InitialDir='' then Select_File_To_Attach.InitialDir:=DIR_SCANNED_SOURCE;
 if filename='' then
 Begin
  if Select_File_To_Attach.execute=true then
  Begin
   SrceName:=Select_File_To_Attach.filename;
  end
  else exit;
 end
 else
 srcename:=filename;

 Application.CreateForm(TFRM_FILE_ATTATCH, FRM_FILE_ATTATCH);
 try
 FRM_FILE_ATTATCH.DocTypes.close;
 FRM_FILE_ATTATCH.DocTypes.open;
 FRM_FILE_ATTATCH.Doclookup.keyvalue:='';
 FRM_FILE_ATTATCH.lfilename.Caption:=Select_File_To_Attach.filename;
 FRM_FILE_ATTATCH.tag:=0;
 FRM_FILE_ATTATCH.pdate.Date:=now;
 FRM_FILE_ATTATCH.followupdate.Date:=now;
 FRM_FILE_ATTATCH.ShowModal;
 if FRM_FILE_ATTATCH.tag=0 then exit;
 Destdir:=DIR_SCANNED_DOCS;
 dnow:=datetimetostr(now);
 Fmon:=copy(FRM_FILE_ATTATCH.pdate.text,9,2)+'-'+copy(FRM_FILE_ATTATCH.pdate.text,4,2);
 try
  createdir(destdir+fmon);
 except
 end;
 Fext:=ExtractFileExt(SrceName);
 NID:=FRM_Common.nextnoteid;
 DestName:=Custid+'_'+NID+fext;
 if FRM_FILE_ATTATCH.movecheck.checked then
 Begin
  if renamefile(srcename,destdir+fmon+'\'+destname)=false then
  Begin
   Messagedlg('Failed to move file. Document could not be attached.',MTInformation,[MBOK],0);
   exit;
  End;
 End
 else
 begin
  Copyfile(srcename,destdir+fmon+'\'+destname);
 End;

 resolved:='N';
 if messagedlg('Is this document enquiry outstanding?',mtconfirmation,[mbyes,mbno],0)<>mrno then resolved:='N'
 else resolved:='Y';

 // Enquiry
 With main_data_module.updatequery Do
 Begin
  close;
  sql.clear;
  if length(custid)=13 then
  Begin
   if resolved='Y' then sql.add('Insert into enquiry.enquiries values ('''+custid+''','''+uppercase(userid)+''',to_date('''+FRM_FILE_ATTATCH.pdate.text+' '+timetostr(now)+''',''DD/MM/YYYY hh24:mi:ss''),''16'','+FRM_FILE_ATTATCH.doctypes.fields[0].text+',to_date('''+FRM_FILE_ATTATCH.followupdate.text+''',''DD/MM/YYYY''),'''+destdir+fmon+'\'+destname+''',null,''Y'',null,'''+uppercase(userid)+''',null,NULL,to_date('''+FRM_FILE_ATTATCH.pdate.text+''',''DD/MM/YYYY''),null,null,null,'+NID+','''+ROLE+''',null)');
   if resolved='N' then sql.add('Insert into enquiry.enquiries values ('''+custid+''','''+uppercase(userid)+''',to_date('''+FRM_FILE_ATTATCH.pdate.text+' '+timetostr(now)+''',''DD/MM/YYYY hh24:mi:ss''),''16'','+FRM_FILE_ATTATCH.doctypes.fields[0].text+',to_date('''+FRM_FILE_ATTATCH.followupdate.text+''',''DD/MM/YYYY''),'''+destdir+fmon+'\'+destname+''',null,''N'',null,null,null,NULL,to_date('''+FRM_FILE_ATTATCH.pdate.text+''',''DD/MM/YYYY''),null,null,null,'+NID+','''+ROLE+''',null)');
  end
  else
  Begin
   IF RESOLVED='Y' THEN sql.add('Insert into enquiry.enquiries values (NULL,'''+uppercase(userid)+''',to_date('''+FRM_FILE_ATTATCH.pdate.text+' '+timetostr(now)+''',''DD/MM/YYYY hh24:mi:ss''),''16'','+FRM_FILE_ATTATCH.doctypes.fields[0].text+',to_date('''+FRM_FILE_ATTATCH.followupdate.text+''',''DD/MM/YYYY''),'''+destdir+fmon+'\'+destname+''',null,''Y'',null,'''+uppercase(userid)+''',null,NULL,to_date('''+FRM_FILE_ATTATCH.pdate.text+''',''DD/MM/YYYY''),null,'+custid+',null,'+NID+','''+ROLE+''',null)');
   IF RESOLVED='N' THEN sql.add('Insert into enquiry.enquiries values (NULL,'''+uppercase(userid)+''',to_date('''+FRM_FILE_ATTATCH.pdate.text+' '+timetostr(now)+''',''DD/MM/YYYY hh24:mi:ss''),''16'','+FRM_FILE_ATTATCH.doctypes.fields[0].text+',to_date('''+FRM_FILE_ATTATCH.followupdate.text+''',''DD/MM/YYYY''),'''+destdir+fmon+'\'+destname+''',null,''N'',null,null,null,NULL,to_date('''+FRM_FILE_ATTATCH.pdate.text+''',''DD/MM/YYYY''),null,'+custid+',null,'+NID+','''+ROLE+''',null)');
  end;
  execute;
  Frm_login.mainsession.commit;
 end;
 Messagedlg('Document Attached Successfully. Doc Store ID is '+NID,MTINFORMATION,[MBOK],0);
 finally
  FRM_FILE_ATTATCH.RELEASE;
 end;

 treeview1.Expanded[xnode]:=false;
 treeview1.Expanded[xnode]:=true;
end;

procedure TFRM_Tree.ShowLatestRatedUsage(xnode:Pvirtualnode;Agreement_id:string);
Var
CorD,isfinal,isfinaltext,istry2,oldorder,ms,qend,per,perdesc:String;
fcolor:tcolor;
Begin
 // Show Status from Rating.rated_agreements
 mynodeagreement:=xnode;
 istry2:='1';
 with ratedusage do
 Begin
  deletevariables;
  DeclareVariable('AID', otlong);
  setvariable('AID',agreement_id);
  close;
  sql.clear;
  sql.add('Select A.quarter,A.Gas_inc_vat,A.Electric_inc_vat,');
  sql.add('A.Telecoms_inc_vat,A.Broadband_inc_vat,A.Other_inc_vat,A.heat_inc_vat,A.water_inc_vat, ');
  sql.add('A.Total_charges_inc_vat,A.effective_from,A.effective_to,b.effective_to, ');
  sql.add('S.RATED_STATUS,S.ADDITIONAL_INFO,S.SENT_STATUS,C.CREDIT_OR_DEBIT_STATUS, C.CREDIT_DEBIT_AMOUNT, ');
  sql.add('S.IS_FINAL,D.SERVICE,D.REASON,S.PERIOD_TYPE ');
  sql.add('from  ');
  sql.add('salesledger.agreement_summary_q A, ');
  sql.add('(select distinct quarter_id,q_to effective_to from billing.billing_periods_months) B,');
  sql.add('SALESLEDGER.Q_ACCOUNT_SUMMARY_AGREEMENT C,');
   sql.add(' (select * from Rating.rated_agreements where period_type=''Q'') S, ');
  sql.add(' (select * from  ');
  sql.add('Rating.missing_rated_reads ');
  sql.add('where override_error is null) D ');
  sql.add('where a.quarter=b.quarter_id ');
  sql.add('and  A.agreement_id=:AID');
  sql.add('and  A.agreement_id=s.agreement_id(+)');
  sql.add('and  A.quarter=s.period ');
  sql.add('and  A.agreement_id=C.agreement_id(+) ');
  sql.add('and  A.quarter=C.quarter_id(+)  ');
  sql.add('and  A.Agreement_ID=D.Agreement_ID (+) ');
  sql.add('and  A.EFFECTIVE_TO=D.QUARTER_END (+) ');
  sql.add('union  ');
  sql.add('(  ');
  {sql.add(' Select A.month,A.Gas_inc_vat,A.Electric_inc_vat, ');
  sql.add(' A.Telecoms_inc_vat,A.Broadband_inc_vat,A.Other_inc_vat, ');
  sql.add(' A.Total_charges_inc_vat,A.effective_from,A.effective_to,b.effective_to, ');
  sql.add('S.RATED_STATUS,S.ADDITIONAL_INFO,S.SENT_STATUS,C.CREDIT_OR_DEBIT_STATUS, C.CREDIT_DEBIT_AMOUNT,');
  sql.add('S.IS_FINAL,D.SERVICE,D.REASON,S.PERIOD_TYPE ');
  sql.add(' from  ');
  sql.add(' salesledger.agreement_summary_m A, ');
  sql.add(' billing.billing_periods_months B,');
  sql.add(' SALESLEDGER.M_ACCOUNT_SUMMARY_AGREEMENT C,');
  sql.add(' (select * from Rating.rated_agreements where period_type=''M'') S, ');
  sql.add(' (select * from ');
  sql.add(' Rating.missing_rated_reads');
  sql.add(' where override_error is null) D ');
  sql.add(' where a.month=b.month');
  sql.add(' and  A.agreement_id=:AID');
  sql.add(' and  A.agreement_id=s.agreement_id(+)');
  sql.add(' and  A.month=s.period  ');
  sql.add(' and  A.agreement_id=C.agreement_id(+)');
  sql.add(' and  A.month=C.month_id(+) ');
  sql.add(' and  A.Agreement_ID=D.Agreement_ID (+)');
  sql.add(' and  A.EFFECTIVE_TO=D.QUARTER_END (+) '); }

  sql.add(' Select s.period,A.Gas_inc_vat,A.Electric_inc_vat,');
  sql.add(' A.Telecoms_inc_vat,A.Broadband_inc_vat,A.Other_inc_vat,A.heat_inc_vat,A.water_inc_vat,');
  sql.add(' A.Total_charges_inc_vat,s.period_from,s.period_to,b.effective_to,');
  sql.add(' S.RATED_STATUS,S.ADDITIONAL_INFO,S.SENT_STATUS,C.CREDIT_OR_DEBIT_STATUS, C.CREDIT_DEBIT_AMOUNT,');
  sql.add(' S.IS_FINAL,D.SERVICE,D.REASON,S.PERIOD_TYPE');
  sql.add(' from ');
  sql.add(' salesledger.agreement_summary_m A,');
  sql.add(' billing.billing_periods_months B,');
  sql.add(' SALESLEDGER.M_ACCOUNT_SUMMARY_AGREEMENT C,');
  sql.add(' (select * from Rating.rated_agreements where period_type=''M'') S,');
  sql.add(' (select * from');
  sql.add(' Rating.missing_rated_reads');
  sql.add(' where override_error is null) D');
  sql.add(' where s.period=b.month (+)');
  sql.add(' and  s.agreement_id=:AID');
  sql.add(' and  s.agreement_id=a.agreement_id(+)');
  sql.add(' and  s.period=A.month (+)');
  sql.add(' and  s.agreement_id=C.agreement_id(+)');
  sql.add(' and  s.period=C.month_id(+)');
  sql.add(' and  s.Agreement_ID=D.Agreement_ID (+)');
  sql.add(' and  s.period_TO=D.QUARTER_END (+)');
  sql.add(') ');
  sql.add(' order by 12 desc');

  open;
  deletevariables;
 End;
 // If no errors, then check if NO Rated TreeDataexists
 if ratedusage.recordcount=0 then
 Begin
  istry2:='2';
  with ratedusage do
  Begin
   deletevariables;
   DeclareVariable('AID', otlong);
   setvariable('AID',agreement_id);
   close;
   sql.clear;
   sql.add('Select S.period,0,0,');
   sql.add('0,0,0,');
   sql.add('0,0,0,A.effective_from,A.effective_to,b.effective_to,');
   sql.add('S.RATED_STATUS,S.ADDITIONAL_INFO,S.SENT_STATUS,C.CREDIT_OR_DEBIT_STATUS,0,');
   sql.add('S.IS_FINAL,D.SERVICE,D.REASON,S.PERIOD_TYPE');
   sql.add('from');
   sql.add('salesledger.agreement_summary_q A,');
   sql.add('(select distinct quarter_id,q_to effective_to from billing.billing_periods_months) B,');
   sql.add('SALESLEDGER.Q_ACCOUNT_SUMMARY_AGREEMENT C,');
   sql.add('RATING.MISSING_RATED_READS M,');
    sql.add(' (select * from Rating.rated_agreements where period_type=''Q'') S, ');
   sql.add(' (select * from');
   sql.add('Rating.missing_rated_reads');
   sql.add('where override_error is null) D');
   sql.add('where S.period=b.quarter_id');
   sql.add('and  S.agreement_id=:AID');
   sql.add('and  S.agreement_id=A.agreement_id(+)');
   sql.add('and  S.period=A.quarter(+)');
   sql.add('and  A.agreement_id=C.agreement_id(+)');
   sql.add('and  A.quarter=C.quarter_id(+)');
   sql.add('and  A.Agreement_ID=D.Agreement_ID (+)');
   sql.add('and  A.EFFECTIVE_TO=D.QUARTER_END (+)');
   sql.add('and S.agreement_id=M.Agreement_id and M.Agreement_id is not null');
   sql.add('union');

   sql.add('(Select S.period,0,0,');
   sql.add('0,0,0,0,0,');
   sql.add('0,A.effective_from,A.effective_to,b.effective_to,');
   sql.add('S.RATED_STATUS,S.ADDITIONAL_INFO,S.SENT_STATUS,C.CREDIT_OR_DEBIT_STATUS,0,');
   sql.add('S.IS_FINAL,D.SERVICE,D.REASON,S.PERIOD_TYPE');
   sql.add('from');
   sql.add('salesledger.agreement_summary_m A,');
   sql.add('billing.billing_periods_months B,');
   sql.add('SALESLEDGER.M_ACCOUNT_SUMMARY_AGREEMENT C,');
   sql.add('RATING.MISSING_RATED_READS M,');
    sql.add(' (select * from Rating.rated_agreements where period_type=''M'') S, ');
   sql.add(' (select * from');
   sql.add('Rating.missing_rated_reads');
   sql.add('where override_error is null) D');
   sql.add('where S.period=b.month');
   sql.add('and  S.agreement_id=:AID');
   sql.add('and  S.agreement_id=A.agreement_id(+)');
   sql.add('and  S.period=A.month(+)');
   sql.add('and  A.agreement_id=C.agreement_id(+)');
   sql.add('and  A.month=C.month_id(+)');
   sql.add('and  A.Agreement_ID=D.Agreement_ID (+)');
   sql.add('and  A.EFFECTIVE_TO=D.QUARTER_END (+)');
   sql.add('and S.agreement_id=M.Agreement_id and M.Agreement_id is not null)');
   sql.add('order by 10 desc');
   open;
   deletevariables;
  end;
 end; // End of Quarterly


 ratedissues.enabled:=true;
 if ratedusage.recordcount<>0 then
 Begin
  isfinal:=ratedusage.fields[17].text;
  per:=ratedusage.fields[20].text;
  if Per='M' then perdesc:='Month'
  else perdesc:='Quarter';
  if isfinal='Y' then isfinaltext:='* FINAL BILL * '
  else if isfinal='B' then isfinaltext:='* BILL * '
  else isfinaltext:='Statement For '+perdesc+' ';
  cord:=copy(ratedusage.fields[15].text,12,6);
  if cord='hing' then cord:='';
  try
   desc:=isfinaltext+ratedusage.fields[0].text+' Ending - '+ratedusage.fields[10].text+' -   Total = '+frm_common.moneyformat(ratedusage.fields[8].value)+' - ('+frm_common.moneyformat(ratedusage.fields[16].value)+' '+cord+')';
  except
   istry2:='2';
   desc:=isfinaltext+ratedusage.fields[0].text+' Ending - '+ratedusage.fields[10].text+' (NO BILLED DATA)';
  end;
  fcolor:=clblack;
  if isfinal='Y' then fcolor:=clblue;
  if isfinal='B' then fcolor:=clblue;
  if ratedusage.Fields[12].text='N' then
  Begin
   ms:=ratedusage.Fields[18].text+' - '+ratedusage.Fields[19].text;
   if ms=' - ' then ms:='Fix errors on Previous Statement(s)';
   if istry2='2' then ms:='No Rated Usage. Check Errors';
  // if istry2='2' then ms:='There are billing warnings/issues. Please check.';
   desc:=desc+#10+'** '+ms+' **';
   fcolor:=clred;
  end;

  if per='Q' then qend:=frm_common.quarterend(ratedusage.fields[0].text)
  else
  qend:=ratedusage.fields[11].text;

  if (ratedusage.Fields[12].text='Y') and (strtodate(qend)>now) then
  Begin
   desc:=desc+#10+'** Incomplete, Check for Errors. Print as Off Cycle Statement Only **';
   fcolor:=clblue;
  end;


  ratednode:=Treeview1.Addchild(mynodeagreement);
  nodeData := Treeview1.GetNodeData(ratednode);
  NodeData.caption := desc;
  NodeData.FontName:='Lucida Console';
  NodeData.fontcolor:=fcolor;
  NodeData.index:=129;
  nodedata.D_AGREEMENT_ID := Agreement_ID;
  nodedata.D_PERIOD_ID :=   ratedusage.fields[0].text;
  nodedata.D_PERIOD_TYPE :=  per;

 { ratednode:=treeview1.items.addchild(mynodeagreement,desc);
  ratednode.Font.Name:='Lucida Console';
  ratednode.font.color:=fcolor;
  ratednode.imageindex:=129;  }
  if isfinal='Y' then nodedata.fontBold:=true;
  if isfinal='B' then nodedata.fontBold:=true;
  if ratedusage.Fields[12].text='N' then nodedata.index:=180;
  if ratedusage.Fields[14].text='Y' then nodedata.index:=181; // Printed
  if ratedusage.Fields[14].text='P' then nodedata.index:=181; // PDFed
  //ratednode.selectedindex:=ratednode.imageindex;

  nodedata.D_AGREEMENT_ID := Agreement_ID;
  nodedata.D_PERIOD_ID :=   ratedusage.fields[0].text;
  nodedata.D_PERIOD_TYPE :=  per;
  //ratednode.data:=MyRecPtr;



  oldorder:=ratedusage.fields[11].text;
  if ratedusage.recordcount>1 then
  Begin
   ratedusage.Next;
   repeat
    if ratedusage.fields[11].text<>oldorder then
    Begin
     isfinal:=ratedusage.fields[17].text;
      per:=ratedusage.fields[20].text;
     if Per='M' then perdesc:='Month'
     else perdesc:='Quarter';
     if isfinal='Y' then isfinaltext:='* FINAL BILL * '
     else if isfinal='B' then isfinaltext:='* BILL * '
     else isfinaltext:='Statement For '+perdesc+' ';
     cord:=copy(ratedusage.fields[15].text,12,6);
     if cord='hing' then cord:='';
     try
      desc:=isfinaltext+ratedusage.fields[0].text+' Ending - '+ratedusage.fields[10].text+' -   Total = '+frm_common.moneyformat(ratedusage.fields[8].value)+' - ('+frm_common.moneyformat(ratedusage.fields[16].value)+' '+cord+')';
     except
      istry2:='2';
      desc:=isfinaltext+ratedusage.fields[0].text+' Ending - '+ratedusage.fields[10].text+' (NO BILLED DATA)';
     end;

     fcolor:=clblack;
     if (isfinal='Y') or (isfinal='B') then
     Begin
      fcolor:=clblue;
     End;
     if ratedusage.Fields[12].text='N' then
     Begin
      ms:=ratedusage.Fields[18].text+' - '+ratedusage.Fields[19].text;
      if ms=' - ' then ms:='Fix errors on Previous Statement(s)';
      if istry2='2' then ms:='No Rated Usage. Check Errors';
      desc:=desc+#10+'** '+ms+' **';
      fcolor:=clred;
      ratedissues.enabled:=true;
     end;
     if ratedusage.Fields[12].text='' then
     Begin
      desc:=desc+#10+'** Incomplete, Check for Errors. Print as Off Cycle Statement Only **';
      fcolor:=clblue;
     end;

     ratedsubnode:=Treeview1.Addchild(ratednode);
     nodeData := Treeview1.GetNodeData(ratedsubnode);
     NodeData.caption :=desc;
     //ratedsubnode:=treeview1.items.addchild(ratednode,desc);
     nodedata.FontName:='Lucida Console';
     if (isfinal='Y') or (isfinal='B') then
     Begin
      NodeData.fontBold:=true;
     End;
     NodeData.fontcolor:=fcolor;
     NodeData.index:=129;
     if ratedusage.Fields[12].text='N' then NodeData.index:=180;
     if ratedusage.Fields[14].text='Y' then NodeData.index:=181; // Printed
     if ratedusage.Fields[14].text='P' then NodeData.index:=181; // PDFed
    // ratedsubnode.selectedindex:=ratedsubnode.imageindex;

     nodedata.D_AGREEMENT_ID := Agreement_ID;
     nodedata.D_PERIOD_ID :=   ratedusage.fields[0].text;
     nodedata.D_PERIOD_TYPE :=  per;
    // ratedsubnode.data:=MyRecPtr;
    end;
    oldorder:=ratedusage.fields[11].text;
    ratedusage.Next;
   until ratedusage.eof;
  End;
 End
 else
 Begin
  // No Rated TreeDataExists.
  // Check if Any Exists in POSTED INVOICE TABLE. THIS MAY NEED TO BE REVERSED
  with ratedusage do
  Begin
   deletevariables;
   DeclareVariable('AID', otlong);
   setvariable('AID',agreement_id);
   close;
   sql.clear;
   sql.add('Select agreement_id,period from salesledger.posted_invoices where agreement_id=:AID');
   open;
   deletevariables;
  end;
  ratedissues.enabled:=true;
  if ratedusage.recordcount<>0 then
  Begin
   desc:='Posted Invoices Exist - NO current BILLED Data';

   ratednode:=Treeview1.Addchild(mynodeagreement);
   nodeData := Treeview1.GetNodeData(ratednode);
   NodeData.caption :=desc;
   NodeData.FontName:='Lucida Console';
   NodeData.fontcolor:=clred;
   NodeData.index:=129;

  { ratednode:=treeview1.items.addchild(mynodeagreement,desc);
   ratednode.Font.Name:='Lucida Console';
   ratednode.font.color:=clred;
   ratednode.imageindex:=129;
   ratednode.selectedindex:=ratednode.imageindex;
   }
   nodedata.D_AGREEMENT_ID := Agreement_ID;
   nodedata.D_PERIOD_ID :=   ratedusage.fields[1].text;
   nodedata.D_PERIOD_TYPE :=  'NONE';
 //;  noderatednode.data:=MyRecPtr;
  End;
 end;
End;

procedure TFRM_Tree.ShowLatestAccountReview(xnode:Pvirtualnode;Agreement_id:string);
Var
act:String;
fcolor:tcolor;
olddd,newdd:Real;
Begin
 // Check if Account Reviewer on Hold (Manual Review)
 with ratedusage do
 begin
  deletevariables;
  DeclareVariable('AID', otlong);
  setvariable('AID',agreement_id);
  close;
  sql.clear;
  sql.add('select * from crm.account_reviews_on_hold');
  sql.add('where agreement_id=:AID');
  open;
  deletevariables;
 end;
 if ratedusage.recordcount<>0 then
 Begin
  desc:='System Reviewer Placed on Manual Hold by '+ratedusage.Fields[1].text;
  {reviewnode:=treeview1.items.addchild(mynodeagreement,desc);
  reviewnode.font.color:=clred;
  reviewnode.imageindex:=146;
  reviewnode.selectedindex:=reviewnode.imageindex;
  }
  //reviewnode.data:=MyRecPtr;

  reviewnode:=Treeview1.Addchild(xnode);
  nodeData := Treeview1.GetNodeData(reviewnode);
  NodeData.caption := desc;
  nodedata.index:=146;
  nodedata.D_AGREEMENT_ID := Agreement_ID;
  nodedata.D_PERIOD_ID :=   ratedusage.fields[1].text;
  nodedata.D_ACTIONED := ACT;

  exit;
 End;

 // Show Latest Review
 with ratedusage do
 Begin
  deletevariables;
  DeclareVariable('AID', otlong);
  setvariable('AID',agreement_id);
  close;
  sql.clear;
  sql.add('Select * from crm.account_reviews');
  sql.add('where agreement_id=:AID');
  sql.add('order by date_of_review desc');
  open;
  deletevariables;
  if recordcount=0 then exit;
 End;
 fcolor:=clblack;

 // Line 1
 desc:='Account Reviewed on '+ratedusage.fields[15].text+' (Based on '+ratedusage.fields[1].text+' Statement) ';

 if ratedusage.fields[17].text='' then desc:=desc+', Not Printed '
 else
 Begin
  fcolor:=clblue;
  desc:=desc+', Printed ';
 end;
 if ratedusage.fields[18].text='' then
 Begin
  ACT:='N';
  desc:=desc+', Not Actioned.'
 end
 else
 Begin
  ACT:='Y';
  desc:=desc+', DD Adjustments actioned.';
  fcolor:=clpurple;
 end;
 // Show Changes to DD (Line 2)
 olddd:=ratedusage.fields[3].value;
 newdd:=ratedusage.fields[7].value;

 if act='Y' then
 Begin
  if olddd<newdd then desc:=desc+#10+'DD Increased from '+frm_common.moneyformat(olddd)+' to '+frm_common.moneyformat(newdd)+' as of '+ratedusage.fields[8].text+'.'
  else
  if olddd>newdd then desc:=desc+#10+'DD Decreased from '+frm_common.moneyformat(olddd)+' to '+frm_common.moneyformat(newdd)+' as of '+ratedusage.fields[8].text+'.'
  else
  if ratedusage.fields[24].value<0 then desc:=desc+#10+'Suggest lowering DD' else desc:=desc+#10+'No Change in DD amount';
 End
 else
 Begin
  if olddd<newdd then desc:=desc+#10+'Suggest DD Increase from '+frm_common.moneyformat(olddd)+' to '+frm_common.moneyformat(newdd)+' as of '+ratedusage.fields[8].text+'.'
  else
  if olddd>newdd then desc:=desc+#10+'Suggest DD Decrease from '+frm_common.moneyformat(olddd)+' to '+frm_common.moneyformat(newdd)+' as of '+ratedusage.fields[8].text+'.'
  else
 if ratedusage.fields[24].value<0 then desc:=desc+#10+'Suggest lowering DD' else desc:=desc+#10+'No Change in DD amount';
 end;

 // Show Additional Catch up payments (Line 3)
 if ratedusage.fields[24].value>0 then
 Begin
 desc:=desc+#10'Catch up amount of '+frm_common.moneyformat(ratedusage.fields[24].value)+
 ' due. Payable by '+ratedusage.fields[13].text+' DDs of '+
 frm_common.moneyformat(ratedusage.fields[10].value)+' starting on '+ratedusage.fields[11].text+'.';
 End;

 if ratedusage.fields[24].value<0 then
 Begin
 desc:=desc+' and suggest refund to customer for the amount of '+frm_common.moneyformat(ratedusage.fields[24].value*-1)+'.';
 End;

 if ratedusage.fields[31].text<>'' then
 Begin
  desc:=desc+#10+ratedusage.fields[31].text;
 end;

  reviewnode:=Treeview1.Addchild(mynodeagreement);
  nodeData := Treeview1.GetNodeData(reviewnode);
  NodeData.caption := desc;
  nodedata.index:=146;
  nodedata.D_AGREEMENT_ID := Agreement_ID;
  nodedata.D_PERIOD_ID :=   ratedusage.fields[1].text;
  nodedata.D_ACTIONED := ACT;

 //reviewnode:=treeview1.items.addchild(mynodeagreement,desc);
 //reviewnode.imageindex:=146;
 // Account on Hold
 if ratedusage.fields[32].text='Y' then
 Begin
  fcolor:=clred;
  nodedata.Index:=186;
  desc:='Account Reviewed on '+ratedusage.fields[15].text+' (Based on '+ratedusage.fields[1].text+' Statement) ';
  desc:=desc+#10+'** Review set to DO NOT Action **';
  nodedata.caption:=desc;
 end;

 nodedata.fontcolor:=fcolor;
 //reviewnode.selectedindex:=reviewnode.imageindex;

{ Treedata.D_AGREEMENT_ID := Agreement_ID;
 Treedata.D_PERIOD_ID :=   ratedusage.fields[1].text;
 Treedata.D_ACTIONED := ACT;
 reviewnode.data:=MyRecPtr;   }




End;

procedure TFRM_Tree.ShowAnyDisputes(xnode:pvirtualnode;Agreement_id:string);
Var
z:integer;
Begin
 // Check if Account In Dispute
 with Disputes do
 begin
  deletevariables;
  DeclareVariable('AID', otlong);
  setvariable('AID',agreement_id);
  close;
  sql.clear;
  sql.add('select R.description,D.dispute_reason_code from crm.agreements_in_dispute D, crm.dispute_reason_codes R');
  sql.add('where d.agreement_id=:AID');
  sql.add('and d.dispute_reason_code=r.dispute_reason_code');
  open;
  deletevariables;
 end;
 for z:=1 to Disputes.recordcount do
 Begin
  desc:='** Account In Dispute. '+Disputes.Fields[0].text+' **';

  Disputesnode:=Treeview1.Addchild(xnode);
  nodeData := Treeview1.GetNodeData(Disputesnode);
  NodeData.caption := desc;
  nodedata.index:=217;
  nodedata.fontcolor:=clred;
  nodedata.fontBold:=true;
  nodedata.index:=193;

  {Disputesnode:=treeview1.items.addchild(mynodeagreement,desc);
  Disputesnode.font.color:=clred;
  Disputesnode.font.style:=[fsbold];
  Disputesnode.imageindex:=193;
  Disputesnode.selectedindex:=Disputesnode.imageindex;
  }
  nodedata.D_AGREEMENT_ID := Agreement_ID;
  nodedata.D_REASON := disputes.fields[1].text;
  //Disputesnode.data:=MyRecPtr;
  disputes.next;
 End;
End;


procedure TFRM_Tree.ShowRatedUsage(Agreement_id:string);
begin
 	Application.CreateForm(Tfrm_ratedusage, frm_ratedusage);
 	try
  	frm_Main.InitialiseUnfocusedSelectionColour(frm_ratedusage.Treeview1);

 		with frm_ratedusage do
 		begin
  		showtabs(agreement_id);
  		showmodal;
 		end;
 	finally
 		frm_ratedusage.release;
 	end;
end;

procedure TFRM_Tree.Showsites(xnode:pvirtualnode;SPAN:string);
begin
 // Get List Of Premises For SPan
 Begin
  // Get Premise in Agreement
  with generalquery do
  Begin
   close;
   deletevariables;
   DeclareVariable('SPAN', otstring);
   sql.clear;
   sql.add('select Distinct');
   sql.add('APS.PREMISE_NAME,');
   sql.add('APS.SPECIAL_ACCESS,');
   sql.add('APS.PREMISE_CONTACT_ID,');
   sql.add('P.PREMISE_ID,');
   sql.add('T.ICON_INDEX,');
   sql.add('P.PREMISE_LINE_1,');
   sql.add('P.PREMISE_LINE_2,');
   sql.add('P.PREMISE_LINE_3,');
   sql.add('P.PREMISE_LINE_4,');
   sql.add('P.PREMISE_LINE_5,');
   sql.add('P.PREMISE_LINE_6,');
   sql.add('P.PREMISE_LINE_7,');
   sql.add('P.PREMISE_LINE_8,');
   sql.add('P.PREMISE_LINE_9,');
   sql.add('P.PREMISE_POSTCODE,');
   sql.add('S.AGREEMENT_ID,');
   sql.add('A.CUSTOMER_ID, APS.DATE_MOVED_OUT,a.agreement_status_id,a.agreement_end_Date,a.agreement_start_date,sp.span_start_Date');
   sql.add('from crm.premises P,');
   sql.add('crm.agreement_premises APS,');
   sql.add('crm.premise_type t,');
   sql.add('crm.service S,');
   sql.add('crm.spans SP,');
   sql.add('crm.AGREEMENTS A');
   sql.add('where SP.SPAN=:SPAN');
   sql.add('and SP.Service_id=S.Service_id');
   sql.add('and P.premise_id=S.premise_id (+)');
   sql.add('and P.premise_type_id=t.premise_type_id (+)');
   sql.add('and s.agreement_id=aps.agreement_id (+)');
   sql.add('and s.premise_id=aps.premise_id (+)');
   sql.add('and S.agreement_id=A.Agreement_id');
   sql.Add('order by a.agreement_start_date desc,a.agreement_end_Date desc nulls first,sp.span_start_date desc');
   setvariable('SPAN',SPAN);
   open;
   deletevariables;
  end;
  with generalquery do
  Begin
   while not eof do
   Begin
    premaddr:='';
    if fields[5].text<>'' then premaddr:=premaddr+fields[5].text+',';
    if fields[6].text<>'' then premaddr:=premaddr+fields[6].text+',';
    if fields[7].text<>'' then premaddr:=premaddr+fields[7].text+',';
    if fields[8].text<>'' then premaddr:=premaddr+fields[8].text+',';
    if fields[9].text<>'' then premaddr:=premaddr+fields[9].text+',';
    if fields[10].text<>'' then premaddr:=premaddr+fields[10].text+',';
    if fields[11].text<>'' then premaddr:=premaddr+fields[11].text+',';
    if fields[12].text<>'' then premaddr:=premaddr+fields[12].text+',';
    if fields[13].text<>'' then premaddr:=premaddr+fields[13].text+',';
    premaddr:=premaddr+fields[14].text;
    premaddr:=premaddr+' - ['+fields[3].text+']';
    {mynodepremise:=Treeview1.items.AddChild(treeview1.selected,'Premises - '+premaddr);
    mynodepremise.imageindex:=fields[4].value;
    mynodepremise.selectedindex:=mynodepremise.imageindex; }

    mynodepremise:=Treeview1.Addchild(xnode);
    nodeData := Treeview1.GetNodeData(mynodepremise);
    NodeData.caption := 'Premises - '+premaddr;
    nodedata.index:=fields[4].value;


    cot1.Visible:=true;
    cot2.visible:=false;

    // check if agreement is terminated
    if (fields[18].text='3') and (fields[17].text='') then
    begin
      nodedata.caption:=nodedata.caption+' - Terminated on '+fields[19].text;
     nodedata.Fontcolor:=clred;
    end;

    // Check if Moving Out
    if fields[17].text<>'' then
    Begin
     cot1.visible:=false; // Dont Show COT option if already vavacted
     cot2.visible:=true;  // Show COT Tools if Vacated
     if strtodate(fields[17].text)<date then
     Begin
      nodedata.caption:=nodedata.caption+' - Vacated on '+fields[17].text;
      nodedata.index:=142;
      //mynodepremise.selectedindex:=mynodepremise.imageindex;
      nodedata.Fontcolor:=clmaroon;
     end
     else
     Begin
      nodedata.caption:=nodedata.caption+' - Vacating on '+fields[17].text;
      nodedata.index:=139;
      //mynodepremise.selectedindex:=mynodepremise.imageindex;
     End;
    end;


    nodedata.D_premise_Id :=   fields[3].text;
    nodedata.D_agreement_Id := fields[15].text;
    nodedata.D_customer_Id :=  fields[16].text;
    //mynodepremise.data:=MyRecPtr;
    //mynode1:=Treeview1.items.AddChild(mynodepremise,'test');
    mynode1:=Treeview1.Addchild(mynodepremise);
    nodeData := Treeview1.GetNodeData(mynode1);
    NodeData.caption := 'Test';

    next;
   end;
  end;
 end;
end;


procedure TFRM_Tree.MeterDetailsReadings2Click(Sender: TObject);
Var
regid:string;
begin
 if treeupdating=true then exit;
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 regid:=TreeData.D_REGID;
 //LK17
 custid:=frm_common.GetCustomerIdfromAgreementid(TreeData.D_agreement_id);

 FRM_Gasmetering.show;
 FRM_Gasmetering.getmeterdetails(mpan,'','','',regid,custid);
end;

procedure TFRM_Tree.ViewRefreshMPRN1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 FRM_Main.SearchForSpan(TreeData.D_SPAN,0);
end;

procedure TFRM_Tree.ViewRefreshMPAN1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 FRM_Main.SearchForSpan(TreeData.D_SPAN,0);
end;

procedure TFRM_Tree.ViewRefreshCallerLineID1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
FRM_Main.SearchForSpan(TreeData.D_SPAN,0);
end;

procedure TFRM_Tree.MenuItem15Click(Sender: TObject);
Var
agreement_id:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 agreement_id:=TreeData.D_agreement_id;
 ShowRatedUsage(agreement_id);
end;

procedure TFRM_Tree.SavingsTransMenuItemClick(Sender: TObject);
begin
  ShowSmartPayScreen(StrToInt64(CustId));
end;

procedure TFRM_Tree.AddProspectDetails1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 custid:=TreeData.D_Customer_id;

 Application.CreateForm(TFrm_Initial_Prospect, Frm_Initial_Prospect);
 Frm_Initial_Prospect.custid.enabled:=true;
 Frm_Initial_Prospect.custid.text:=custid;
 Frm_Initial_Prospect.custid.enabled:=False;
 Frm_Initial_Prospect.tag:=0;
 Frm_Initial_Prospect.clearfields;
 try
  Frm_Initial_Prospect.showmodal;
 finally
 Frm_Initial_Prospect.RELEASE;
 end;
 if frm_add_account_holder.tag=0 then exit;
 // refresh customer node;
  if treeview1.Selected[xnode]=true then
  Begin
   treeview1.Expanded[xnode]:=false;
   treeview1.Expanded[xnode]:=true;
  end;
end;

procedure TFRM_Tree.MenuItem32Click(Sender: TObject);
Var
  email, Ecard, Gcard, ts, ps, sURL, ccolor: String;
begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);
  CustId := TreeData.D_Customer_Id;
  email := TreeData.D_Email;

  with main_data_module.tempquery do
  begin
    DeleteVariables;
    DeclareVariable('CID', otstring);
    Close;
    Sql.Clear;
    Sql.Add('select');
    Sql.Add('a.card_no,a.service,');
    Sql.Add('lower(i.card_colour) color');
    Sql.Add('from ');
    Sql.Add(' crm.smartcard_spans a, ');
    Sql.Add('smiff.wmol_assets_iin_types i');
    Sql.Add(' where');
    Sql.Add('substr(a.card_no,1,9)=i.iin_start_range');
    Sql.Add('and');
    Sql.Add('a.customer_id=:CID');
    SetVariable('CID', CustId);
    Open;
    DeleteVariables;
  end;

  if main_data_module.tempquery.RecordCount = 0 then
  begin
    Messagedlg('Cannot find any payment Card numbers for this customer', mtwarning, [MBOK], 0);
    exit;
  end;
  Ecard := 'null';
  Gcard := 'null';
  ccolor := '';

  while not main_data_module.tempquery.eof do
  begin
    if main_data_module.tempquery.Fields[1].Text = 'E' then
      Ecard := main_data_module.tempquery.Fields[0].Text;
    if main_data_module.tempquery.Fields[1].Text = 'G' then
      Gcard := main_data_module.tempquery.Fields[0].Text;
    ccolor := main_data_module.tempquery.Fields[2].Text;
    main_data_module.tempquery.Next;
  end;

  ps := ' --window-size=400,800 --window-position=0,0';
  sURL := 'https://utilita.co.uk/contact-us/barcode/gen/' + email + '/' + Ecard + '/' + Gcard;
  ShellExecute(Handle, 'open', pchar(sURL), pchar(ps), nil, sw_shownormal);
end;

procedure TFRM_Tree.MenuItem3Click(Sender: TObject);
var
custid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 custid:=TreeData.D_Customer_id;
 Application.CreateForm(TFrm_Initial_Prospect, Frm_Initial_Prospect);
 Frm_Initial_Prospect.custid.enabled:=true;
 Frm_Initial_Prospect.custid.text:=custid;
 Frm_Initial_Prospect.custid.enabled:=False;
 Frm_Initial_Prospect.tag:=0;
 Frm_Initial_Prospect.clearfields;
 Frm_Initial_Prospect.GetDetails;
 Frm_Initial_Prospect.showmodal;
 if frm_Initial_Prospect.tag=0 then exit;
 Frm_Initial_Prospect.release;

 // refresh customer node;
  treeview1.Expanded[xnode]:=false;
  treeview1.Expanded[xnode]:=true;
end;

procedure TFRM_Tree.BroadbandReorderClick(Sender: TObject);
begin
ReOrder('5');
end;

procedure TFRM_Tree.MenuItem4Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 FRM_Main.SearchForSpan(TreeData.D_SPAN,0);
end;

procedure TFRM_Tree.WANMatrix1Click(Sender: TObject);
var
bFlag:boolean;
begin
xnode:=treeview1.FocusedNode;
TreeData:= treeview1.GetNodeData(xnode);
premiseid:= TreeData.D_premise_id;

  Application.CreateForm(TWANMatrix, WANMatrix);
  try
      With WANQuery Do
      Begin
        Close;
        Sql.Clear;
        DeleteVariables;
        DeclareVariable('PREMID', OtFloat);
        Sql.Add('select PR.PREMISE_ID, PR.PREMISE_POSTCODE, PR.PREMISE_LINE_1, PR.PREMISE_LINE_2, PR.PREMISE_LINE_3, PR.PREMISE_LINE_4, PR.PREMISE_LINE_5, PR.PREMISE_LINE_6, PR.PREMISE_LINE_7, PR.PREMISE_LINE_8, PR.PREMISE_LINE_9');
        Sql.Add('from CRM.PREMISES PR where PR.PREMISE_ID = :PREMID');
        SetVariable('PREMID', premiseid);
        Try
          Open;
          gCustPremCd := FieldByName('PREMISE_POSTCODE').AsString;
          gCustPremId := FieldByName('PREMISE_ID').AsInteger;
          gCustPremLn1 := FieldByName('PREMISE_LINE_1').AsString;
          gCustPremLn2 := FieldByName('PREMISE_LINE_2').AsString;
          gCustPremLn3 := FieldByName('PREMISE_LINE_3').AsString;
          gCustPremLn4 := FieldByName('PREMISE_LINE_4').AsString;
          gCustPremLn5 := FieldByName('PREMISE_LINE_5').AsString;
          gCustPremLn6 := FieldByName('PREMISE_LINE_6').AsString;
          gCustPremLn7 := FieldByName('PREMISE_LINE_7').AsString;
          gCustPremLn8 := FieldByName('PREMISE_LINE_8').AsString;
          gCustPremLn9 := FieldByName('PREMISE_LINE_9').AsString;
          Close;
          Sql.Clear;
        Except
          Messagedlg('An error has occured please contact the system administrator.',mtwarning,[mbok],0);
        End;
      End;
  finally
    WANMatrix.show;
  end;
end;

procedure TFRM_Tree.DoWarrentJobLock(Sender: TObject);
var
  bCaption : String;
Begin
  If Assigned(Treedata) then
  begin
    bCaption := Treedata.Caption;
    If (Pos('Warrant', bCaption) > 0) then
      TAction(Sender).Enabled := USER_FEATURE__AMEND_WARRANT_JOBS
    Else
      TAction(Sender).Enabled := True;
  end;
end;

procedure TFRM_Tree.WebSignupLetter1Click(Sender: TObject);
Var
AgId,salesref,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AgID:=TreeData.D_Agreement_id;
 with main_data_module.generalquery do
 Begin
  close;
  deletevariables;
  DeclareVariable('AID', otlong);
  sql.clear;
  sql.add('Select sales_reference,customer_id from crm.agreements where agreement_id=:AID');
  setvariable('AID',agid);
  open;
  deletevariables;
 End;
 salesref:=main_data_module.generalquery.Fields[0].text;
 CID:=main_data_module.generalquery.fields[1].text;
 if copy(salesref,9,3)='HLP' then
 Begin
  FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Signup_Letter_HELPCO.rpt','Signup Letter HelpCo','{ACCOUNT_HOLDERS.CONTACT_ORDER} = 1 and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
  exit;
 end
 else FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Signup_Letter.rpt','Signup Letter','{ACCOUNT_HOLDERS.CONTACT_ORDER} = 1 and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
 // PPM Yes
 // NOR Yes
 // PAP No
 if (Salesref='PCW WEB ENX') or
    (salesref='PCW WEB UNR') or
    (salesref='PCW WEB SPS') or
    (salesref='PCW WEB EHL') or
    (salesref='PCW WEB ESH') or
    (salesref='PCW WEB SYB') or
    (salesref='DIR WEB BIL') then
 Begin
  // if DDsigned then exit;
  with main_data_module.generalquery do
  Begin
   close;
   deletevariables;
   DeclareVariable('AID', otlong);
   sql.clear;
   sql.add('Select direct_debit_signed,effective_from from crm.agreement_bank_details where agreement_id=:AID');
   sql.add('order by effective_from desc');
   setvariable('AID',agid);
   open;
   deletevariables;
  End;
  if main_data_module.generalquery.recordcount<>0 then
  Begin
   if main_data_module.generalquery.Fields[0].text='Y' then exit;
  end
 end;

  {// We Dont Want one of these if a PPM Customer
  with main_data_module.generalquery do
  Begin
   close;
   sql.clear;
   sql.add('Select payment_plan_id from crm.agreement_products where agreement_id='+agid);
   sql.add('order by effective_from desc');
   open;
  End;
  if main_data_module.generalquery.recordcount<>0 then
  Begin
   if main_data_module.generalquery.Fields[0].text='P' then exit;
  end;}

 //FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Application_Form.rpt','Application Form','{ACCOUNT_HOLDERS.CONTACT_ORDER} = 1 and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID);
end;

procedure TFRM_Tree.InformationRequest11Click(Sender: TObject);
Var
Custid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Custid:=TreeData.D_Customer_ID;
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Information_Request.rpt','Information Request','{ACCOUNT_HOLDERS.CONTACT_ORDER}=1 and {CUSTOMER.CUSTOMER_ID}='+custid+'','','PRINTER',custid,'');
 // Print Default Application Form. Note this Goes to DEFAULT Printer.
 FRM_Reports.PrintAFile(FRM_Common.GETVALUE('APPLICATION_FORM'));
end;

procedure TFRM_Tree.CustomerCredits1Click(Sender: TObject);
Var
CustomerID:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;
 CustomerID:=TreeData.D_Customer_ID;

 Application.CreateForm(TFrm_Credit_Summary, Frm_Credit_Summary);
 try
  Frm_Credit_summary.showcustcredits(Customerid);
  frm_credit_summary.tabsheet1.tabvisible:=false;
  frm_credit_summary.menu:=nil;
  Frm_Credit_Summary.showmodal;
 finally
  Frm_Credit_Summary.release;
 end;

end;

procedure TFRM_Tree.SACClick(Sender: TObject);
Var
CustomerID,agreementid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;
 Begin
  AgreementID:=TreeData.D_Agreement_ID;
  CustomerID:=TreeData.D_Customer_ID;

  Application.CreateForm(TFrm_Credit_summary, Frm_Credit_summary);
  try
   Frm_Credit_summary.showAgreementcredits(AgreementID,Customerid);
   frm_credit_summary.tabsheet1.tabvisible:=true;
   frm_credit_summary.Menu:=frm_credit_summary.mainmenu1;
   frm_credit_summary.pagecontrol1.activepageindex:=0;
   Frm_Credit_Summary.showmodal;
  finally
   Frm_Credit_summary.release;
  end;

 end;
end;

function TFRM_Tree.IsValidFinancial: Boolean;
begin
  Result := (main.isFinance) or (Custid <> SuspenseCustID);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.ShowFinancialHistory1Click(Sender: TObject);
var
  nodeData    : PMyRec;
  agreementId : Int64;
begin
  if TreeUpdating then
    exit;

  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  agreementID := StrToInt64(nodeData.D_Agreement_ID);
  TFrm_Financial_History.StartModal(Self, ctShowFinancialHistory, agreementId, IsValidFinancial);
end;

procedure TFRM_Tree.ViewSPANSetupDetails1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;
 Application.CreateForm(TFRM_SPAN_DETAILS, FRM_SPAN_DETAILS);
 try
 WIth FRM_SPAN_DETAILS.spanquery do
 Begin
  close;
  setvariable('Registrationid',Treedata.D_REGID);
  open;
 end;
 FRM_SPAN_DETAILS.G.tabvisible:=true;
 FRM_SPAN_DETAILS.E.tabvisible:=false;
 FRM_SPAN_DETAILS.T.tabvisible:=false;
 FRM_SPAN_DETAILS.B.tabvisible:=false;
 FRM_SPAN_DETAILS.SHOWMODAL;
 finally
  FRM_SPAN_DETAILS.RELEASE;
 end;
end;

procedure TFRM_Tree.ViewSPANSetupDetails2Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;
 Application.CreateForm(TFRM_SPAN_DETAILS, FRM_SPAN_DETAILS);
 try
 WIth FRM_SPAN_DETAILS.spanquery do
 Begin
  close;
  setvariable('Registrationid',Treedata.D_REGID);
  open;
 end;
 FRM_SPAN_DETAILS.G.tabvisible:=false;
 FRM_SPAN_DETAILS.E.tabvisible:=true;
 FRM_SPAN_DETAILS.T.tabvisible:=false;
 FRM_SPAN_DETAILS.B.tabvisible:=false;
 FRM_SPAN_DETAILS.SHOWMODAL;
 finally
 FRM_SPAN_DETAILS.RELEASE;
 end;
end;

procedure TFRM_Tree.View1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;
 Application.CreateForm(TFRM_SPAN_DETAILS, FRM_SPAN_DETAILS);
 try
 WIth FRM_SPAN_DETAILS.spanquery do
 Begin
  close;
  setvariable('Registrationid',Treedata.D_REGID);
  open;
 end;
 FRM_SPAN_DETAILS.G.tabvisible:=false;
 FRM_SPAN_DETAILS.E.tabvisible:=false;
 FRM_SPAN_DETAILS.T.tabvisible:=True;
 FRM_SPAN_DETAILS.B.tabvisible:=false;
 FRM_SPAN_DETAILS.SHOWMODAL;
 finally
 FRM_SPAN_DETAILS.RELEASE;
 end;
end;

procedure TFRM_Tree.ViewSPANSetupDetails3Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;
 Application.CreateForm(TFRM_SPAN_DETAILS, FRM_SPAN_DETAILS);
 try
 WIth FRM_SPAN_DETAILS.spanquery do
 Begin
  close;
  setvariable('Registrationid',Treedata.D_REGID);
  open;
 end;
 FRM_SPAN_DETAILS.G.tabvisible:=false;
 FRM_SPAN_DETAILS.E.tabvisible:=false;
 FRM_SPAN_DETAILS.T.tabvisible:=false;
 FRM_SPAN_DETAILS.B.tabvisible:=true;
 FRM_SPAN_DETAILS.SHOWMODAL;
 finally
  FRM_SPAN_DETAILS.RELEASE;
 end;
end;

procedure TFRM_Tree.WelcomeLetter1Click(Sender: TObject);
Var
AgId,ddsigned,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AgID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 with main_data_module.generalquery do
 Begin
  close;
  deletevariables;
  DeclareVariable('AID', otlong);
  sql.clear;
  sql.add('Select direct_debit_signed from crm.agreement_bank_details');
  sql.add('where agreement_id=:AID');
  sql.add('order by effective_from desc');
  setvariable('AID',agid);
  open;
  deletevariables;
 End;
 if main_data_module.generalquery.recordcount<>0 then ddsigned:=main_data_module.generalquery.fields[0].text
 else ddsigned:='';
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Welcome_Letter.rpt','Welcome Letter','{ACCOUNT_HOLDERS.CONTACT_ORDER} = 1 and {AGREEMENTS.AGREEMENT_ID} = '+AgID,'','PRINTER',CID,'');
 if DDSIGNED='Y' then FRM_Reports.PrintThisReport('CRM\Direct Debit Reports\DD_Confirmation.rpt','DD Confirmation','isnull({AGREEMENT_BANK_DETAILS.DIRECT_DEBIT_SIGNED})=false and {AGREEMENTS.AGREEMENT_ID} = '+AgID,'','PRINTER',CID,'')
 else FRM_Reports.PrintThisReport('CRM\Direct Debit Reports\DD_Mandate.rpt','DD Mandate','isnull({AGREEMENT_BANK_DETAILS.DIRECT_DEBIT_SIGNED})=true and {AGREEMENTS.AGREEMENT_ID} = '+AgID,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.CancelAllSubOrders1Click(Sender: TObject);
Var
AgId,enddate,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AgID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 enddate:=TreeData.D_Agreement_end_date;

 if enddate='' then
 Begin
  Messagedlg('Agreement must be Terminated before you can cancel any sub orders',MTinformation,[MBOK],0);
  exit;
 End;
 if Messagedlg('Are you sure you wish to cancel all Sub orders?'+#13+
               'This will cancel any DDI''s, Products, Services and SPANS'+#13+
               'that are Order Placed or Order ready and all DDI/DDR requests.',Mtconfirmation,[MBYES,MBNO],0)<>MRyes then exit;
 // Now Cancel Sub Orders
 FRM_Agreement.CANCEL_Agreement_sub_orders(Agid,enddate);
 // Check for Statement Reviewer
 // If Not Exist then Add one, detailing Cancellation
 // Check for Statement Reviewr
 frm_common.addstatementreviewer(cid); //LK17

 with main_data_module.updatequery do
 Begin
  // Now Add A Note
  close;
  sql.clear;
  sql.add('Insert into enquiry.enquiries values(NULL,'''+uppercase(USERID)+''',sysdate,'+inttostr(14)+','+inttostr(232)+',null,''Customer Cancellation.'',null,''N'',null,null,null,NULL,trunc(sysdate),null,'+CID+',null,'+frm_Common.NextNoteId+',''X'',null)');
  execute;
  // Now Add An Enquiry For Switching Agent Cancellation
  close;
  sql.clear;
  sql.add('Insert into enquiry.enquiries values(NULL,'''+uppercase(USERID)+''',sysdate,'+inttostr(14)+','+inttostr(232)+',null,''Customer Cancellation.'',null,''N'',null,null,null,NULL,trunc(sysdate),null,'+CID+',null,'+frm_Common.NextNoteId+',''X'',null)');
 //execute;
 End;
 frm_login.mainsession.commit;

 treeview1.Expanded[xnode]:=false;
 treeview1.Expanded[xnode]:=true;
end;

Procedure TFRM_Tree.C1Click(Sender: TObject);
Var CustID: String;

  // SJ-BSL 01/04/2021- BRM-601 Statement Reviewer - Reduce code for optimisation
  Procedure pDeleteRecord;
  Begin
    With Main_Data_Module.UpdateQuery do
      Try
        DeleteVariables;
        DeclareVariable('CID', otString);
        SetVariable    ('CID', CustID);
        Close;
        SQL.Clear;
        SQL.Add('Delete From CRM.CUSTOMER_STATEMENT_REVIEWER ');
        SQL.Add('Where Customer_id = :CID');
        Execute;
        DeleteVariables;
        FRM_Login.MainSession.Commit;
      Except
        On E: Exception do
          Begin
            Application.MessageBox(PChar('SQL Error deleting Statement Reviewer --> ' + E.Message), Attn, MB_ICONERROR);
            FRM_Login.MainSession.Rollback;
          End;
      End;
  End; // SubProc

Begin
  XNode    := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(xnode);

  FRM_Enquiry_Note.Tag := 3;
  CustID               := TreeData.D_Customer_ID;

  // SJ-BSL 26/03/2021- BRM-601 Statement Reviewer
  If Copy(C1.Caption, 3, 1) = 'd' then
    Begin
      FRM_Enquiry_Note.StatReviewer := True;
      //FRM_Enquiry_Summary.adddata(CustID,'Internal Request','Statement Reviewer');
      // SJ-BSL 01/04/2021- BRM-601 Changed 295 to 2805 because 2805 is currently Live
      FRM_Enquiry_Note.AddData('', CustID, 'Internal Request', '2805', 'X');

      Try
        FRM_Enquiry_Note.Close;
      Except
      End;

      With Main_Data_Module.UpdateQuery do
        Try
          Close;
          SQL.Clear;
          SQL.Add('Insert into CRM.CUSTOMER_STATEMENT_REVIEWER VALUES(');
          SQL.Add(CustID + ',''' + UpperCase(UserID) + ''', Trunc(SysDate))');
          Execute;
          FRM_Login.MainSession.Commit;
        Except
          On E: Exception do
            Begin
              Application.MessageBox(PChar('SQL Error Inserting Statement Reviewer --> ' + E.Message), Attn, MB_ICONERROR);
              FRM_Login.MainSession.Rollback;
            End;
        End;

       If FRM_Enquiry_Note.ShowModal = mrCancel then
        Begin
          pDeleteRecord;
        End; // IF
    End
  Else
    Begin
      If Messagedlg('Are you sure you wish to remove Statement Reviewer for customer id ' + CustID + ' ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        pDeleteRecord;
    End;

  Treeview1.Expanded[XNode] := False;
  Treeview1.Expanded[XNode] := True;
End;

procedure TFRM_Tree.COT1Click(Sender: TObject);
var
  Customerid, Premise_ID: string;
begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  Agreement_id := TreeData.D_Agreement_ID;
  Customerid := TreeData.D_Customer_Id;
  Premise_ID := TreeData.D_Premise_Id;

  Application.CreateForm(TFRM_PREMISE_MOVE_OUT, FRM_PREMISE_MOVE_OUT);
  try
    FRM_PREMISE_MOVE_OUT.premiseid.Text := Premise_ID;
    FRM_PREMISE_MOVE_OUT.agreementId.Text := Agreement_id;
    FRM_PREMISE_MOVE_OUT.Tag := 0;
    FRM_PREMISE_MOVE_OUT.ShowModal;
    if FRM_PREMISE_MOVE_OUT.Tag = 2 then
    Begin
      MessageDlg('COT Actions complete.', mtinformation, [mbOk], 0);
      Treeview1.Expanded[XNode.parent] := False;
      Treeview1.Expanded[XNode.parent] := True;
    end;
  finally
    FRM_PREMISE_MOVE_OUT.release;
  end;
end;

procedure TFRM_Tree.M_SingleClick(Sender: TObject);
Var
MPAN,SSD:String;
begin
 mpannode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(mpannode);

 mpan:=TreeData.D_SPAN;
 SSD:=TreeData.D_SSD;

 if m_single.caption[1]='R' then
 Begin
  If Messagedlg('Are you sure you wish to remove Single Rate Billing on this MPAN?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('delete from crm.mpans_single_rate_billing where mpancore='''+mpan+'''');
   execute;
   frm_login.mainsession.commit;
  End;
 End
 else
 Begin
  If Messagedlg('Are you sure you wish to set this SPAN to Single Rate Billing?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
  // Check if Not Already got a default Single Rate billing Added
  with main_data_module.generalquery do
  Begin
   close;
   deletevariables;
   declarevariable('MPAN',otstring);
   sql.clear;
   sql.add('select * from crm.mpans_single_rate_billing');
   sql.add('where mpancore=:MPAN');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
   if recordcount<>0 then
   Begin
    Messagedlg('This Span has already been set to Single Rate Billing',mtinformation,[MBOK],0);
    exit;
   End;
  End;
  // Get User to Enter a Date for Single Rate billing
  //
  frm_date.efd.date:=strtodate(ssd);
  frm_date.showmodal;
  if frm_Date.tag=0 then exit;
  ssd:=frm_date.efd.text;
  //
  if Messagedlg('Are you sure you wish to set this SPAN to Single Rate Billing from '+SSD+'?',MTCONFIRMATION,[MBYES,MBNO],0)<>mryes then exit;
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('Insert into crm.mpans_single_rate_billing values('''+mpan+''',to_date('''+ssd+''',''DD/MM/YYYY''),null,'''+uppercase(userid)+''',trunc(sysdate))');
   execute;
   frm_login.mainsession.commit;
  End;
 end;

 treeview1.Expanded[mpannode]:=false;
 treeview1.Expanded[mpannode]:=true;

end;

procedure TFRM_Tree.A_C_1Click(Sender: TObject);
Var
CustID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 CUSTID:=TreeData.D_Customer_ID;
 frm_common.deletecustomer(Custid,'1');
end;

procedure TFRM_Tree.MakeLIVeCOTmoveIN1Click(Sender: TObject);
Var
AgId,startdate:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AgID:=TreeData.D_Agreement_id;
 Startdate:=TreeData.D_Agreement_Start_date;

 if Messagedlg('Are you sure you wish to make LIVE all Sub orders?'+#13+
               'This will make LIVE any DDI''s, Products, Services and SPANS',Mtconfirmation,[MBYES,MBNO],0)<>MRyes then exit;
 // Now MAKE LIVE Sub Orders
 with main_data_module.updatequery do
 Begin
  Begin
    DeleteVariables;
    DeclareVariable('AID', otLong);
    SetVariable    ('AID', AgId);
    // BSL - SJ - 29/06/2021 - CRM-520-ORA-01036 Illegal Variable name/number
                                 // Declared variable must be after previous Query,
                                 // because it doesn't required as a parameter in that Query.
    // DeclareVariable('SSD', otString);
    //SetVariable    ('SSD', StartDate);

    Close;
    Sql.Clear;
    Sql.Add('Update crm.agreement_bank_details');
    Sql.Add('set bank_details_status_id = 8');
    Sql.Add('where (bank_details_status_id = 3 or bank_details_status_id = 4)');
    Sql.Add('and agreement_id = :AID');
    Execute;

    // BSL - SJ - 29/06/2021 - CRM-520-ORA-01036 Illegal Variable name/number
                                 // Declared variable must be after previous Query,
                                 // because it doesn't required as a parameter in that Query.
    DeclareVariable('SSD', otString);
    SetVariable    ('SSD', StartDate);

    Close;
    Sql.Clear;
    Sql.Add('Update crm.agreement_products');
    Sql.Add('set order_status_id=8,effective_from=to_date(:SSD,''DD/MM/YYYY'')');
    Sql.Add('where agreement_id=:AID');
    Execute;

    Close;
    Sql.Clear;
    Sql.Add('Update crm.service');
    Sql.Add('set order_status_id=8,start_date=to_date(:SSD,''DD/MM/YYYY'')');
    Sql.Add('where agreement_id=:AID');
    Execute;

    Close;
    Sql.Clear;
    Sql.Add('Update crm.spans');
    Sql.Add('set order_status_id=8,span_start_date=to_date(:SSD,''DD/MM/YYYY'')');
    Sql.Add('where (service_id) in');
    Sql.Add('(select service_id from crm.service where order_status_id=8 and');
    Sql.Add('agreement_id=:AID)');
    Execute;

    DeleteVariables;
  End;

  FRM_Login.MainSession.Commit;

  Try
    Treeview1.Expanded[XNode.Parent] := False;
    Treeview1.Expanded[XNode.Parent] := True;
  Except
  End;
 End; // If
End;

procedure TFRM_Tree.UndofromOrderReadytoOrderPlaced1Click(Sender: TObject);
Var
AgId:string;
begin

 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AgID:=TreeData.D_Agreement_id;

 if Messagedlg('Are you sure you wish to set all Sub orders to Order PLACED?'+#13+
               'This will update any DDI''s, Products, Services and SPANS'+#13+
               'that are ORDER READY, or CANCELLED to ORDER PLACED',Mtconfirmation,[MBYES,MBNO],0)<>MRyes then exit;
 // Now Cancel Sub Orders
 with main_data_module.updatequery do
 Begin
  deletevariables;
  DeclareVariable('AID', otlong);
  setvariable('AID',agid);
  close;
  sql.clear;
  sql.add('Update crm.agreement_bank_details');
  sql.add('set bank_details_status_id=3,effective_to=null');
  sql.add('where (bank_details_status_id=4 or bank_details_status_id=10 or bank_details_status_id=11)');
  sql.add('and agreement_id=:AID');
  execute;
  close;
  sql.clear;
  sql.add('Update crm.agreement_products');
  sql.add('set order_status_id=3,effective_from=null,effective_to=null,additional_information=null');
  sql.add('where (order_status_id=4 or order_status_id=10) ');
  sql.add('and agreement_id=:AID');
  execute;
  close;
  sql.clear;
  sql.add('Update crm.service');
  sql.add('set order_status_id=3,start_date = null');
  sql.add('where (order_status_id=4 or order_status_id=10) ');
  sql.add('and agreement_id=:AID');
  execute;
  close;
  sql.clear;
  sql.add('Update crm.spans');
  sql.add('set order_status_id=3,span_start_date=null,batch_no=null');
  sql.add('where (order_status_id=4 or order_status_id=10)');
  sql.add('and (service_id) in');
  sql.add('(select service_id from crm.service where');
  sql.add('agreement_id=:AID)');
  execute;
  deletevariables;
 End;
 frm_login.mainsession.commit;

 treeview1.Expanded[xnode]:=false;
 treeview1.Expanded[xnode]:=true;

end;

procedure TFRM_Tree.UsageCalculatorAutomatic1Click(Sender: TObject);
Var
  AgId:string;
begin
  xnode:=treeview1.FocusedNode;
  TreeData:= treeview1.GetNodeData(xnode);
  AgID:=TreeData.D_Agreement_id;

  Application.CreateForm(TFRM_UsageCalculatorAuto, FRM_UsageCalculatorAuto);
  try
    Frm_UsageCalculatorAuto.edtAgreementId.Text := AgID;
    Frm_UsageCalculatorAuto.edtAgreementId.Enabled := False;
    Frm_UsageCalculatorAuto.showmodal;
  finally
    Frm_UsageCalculatorAuto.release;
  end;
end;

Procedure TFRM_Tree.ViewCustomerTree1Click(Sender: TObject);
Begin
  XNode    := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  If (Copy(TreeData.Caption, 1, 14) = 'Super Customer') then
    Begin
      FRM_Main.SearchForSupercust(TreeData.D_Customer_ID);
      FCustIcon := 214;
    End // If
  Else
    FRM_Main.SearchForcust(TreeData.D_Customer_ID);
End; // Proc

procedure TFRM_Tree.OnTesting(Sender: TObject);
begin
treeview1.fullcollapse;
//Frm_Tree.WindowState:=wsnormal;

end;

procedure TFRM_Tree.ChangeSuplyStartDate1Click(Sender: TObject);
Var
SSD,regid,ddate:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 SSD:=TreeData.D_SSD;
 regid:=TreeData.D_RegID;
 if ssd='' then
 Begin
  Messagedlg('You cannot change the Supply Start Date when it is Blank',MTInformation,[MBOK],0);
  exit;
 End;

 Application.CreateForm(Tfrm_datepicker, frm_datepicker);
 try
 frm_datepicker.cal1.date   :=strtodate(ssd);
 frm_datepicker.DatePan.caption:='Please Select Supply Start Date';
 Frm_datepicker.showmodal;
 ddate:=datetostr(frm_datepicker.cal1.date);
 finally
  frm_datepicker.release;
 end;
 if Messagedlg('Are you sure you wish to change the Supply Start Date to '+ddate+' ?',MTConfirmation,[MBYES,MBNO],0)<>MRYES then exit;

 with main_data_module.updatequery do
 Begin
  Close;
  sql.clear;
  sql.add('Update crm.spans');
  sql.add('Set Span_start_date=to_date('''+ddate+''',''DD/MM/YYYY'')');
  sql.add('where registration_id='+regid);
  execute;
 End;
  if treeview1.Selected[xnode.parent]=true then
  Begin
   treeview1.Selected[xnode.parent]:=false;
   treeview1.Selected[xnode.parent]:=true;
  end;

end;

procedure TFRM_Tree.AquireAttachDocument1Click(Sender: TObject);
{Var
Custid,custname:string;}
begin
 {xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Custid:=TreeData.D_Customer_ID;
 CustName:=TreeData.D_Customer_Name;
 FRM_Twain.custlabel.caption:=custid;
 FRM_Twain.custname.caption:=custname;
 FRM_Twain.RoleLabel.caption:='X';
 Frm_Twain.showmodal; }
end;

procedure TFRM_Tree.AddScannedDocument1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Addscanneddoc(TreeData.D_Customer_ID,'','X');
end;

procedure TFRM_Tree.Cancellation1Click(Sender: TObject);
Var
Custid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Custid:=TreeData.D_Customer_ID;
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Cancellation.rpt','Cancellation','{ACCOUNT_HOLDERS.CONTACT_ORDER}=1 and {CUSTOMER.CUSTOMER_ID}='+custid+'','','PRINTER',custid,'');
end;


procedure TFRM_Tree.CancelServices1Click(Sender: TObject);
var
  AGRE: String;
begin
  if treeupdating=true then exit;

  mpannode:=treeview1.FocusedNode;
  nodeData := treeview1.GetNodeData(mpannode);

  AGRE  := NOdedata.D_Agreement_ID;

  CancellationFrm := TCancellationFrm.Create(self, AGRE);
  CancellationFrm.ShowModal;
  CancellationFrm.Free;
end;

function TFRM_Tree.CanSendFurtherNOI: Boolean;
var
  storedProcedure, AgreementID, CustomerID : String;
  res : variant;
begin
  Xnode           := Treeview1.FocusedNode;
  TreeData        := Treeview1.GetNodeData(Xnode);

  AgreementID := Trim(TreeData.D_Agreement_ID);

  storedProcedure :=
  ' CBT.pk_letters_cbt_ni2.pr_check_eligibility('+
  ':p_agreement_id , '+
  ':p_response)';
  Try
    gSqlUtil.ExecProc(storedProcedure, TRANSACTION_YES,
    [
      'p_agreement_id',  otString, pdInput, AgreementID,
      'p_response',      otString, pdOutput, @res]);

    Result := (UpperCase(res) = 'TRUE')
  except on E: Exception do
    Result := False;
  end;
end;

procedure TFRM_Tree.elecomsApplicationForm1Click(Sender: TObject);
Var
Agid,PremiseID,cid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 PREMISEID:=TreeData.D_Premise_id;
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Telecoms.rpt','Telecoms Application','{PREMISES.PREMISE_ID}='+premiseid+' and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.RequestMissingMPRN1Click(Sender: TObject);
Var
Agid,PremiseID,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 PREMISEID:=TreeData.D_Premise_id;
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Missing_MPRNS.rpt','Missing MPRNS','{SITE_ADDRESS.PREMISE_ID}='+premiseid+' and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.RequestMissingMPAN1Click(Sender: TObject);
Var
Agid,PremiseID,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 PREMISEID:=TreeData.D_Premise_id;
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Missing_MPANS.rpt','Missing MPANS','{SITE_ADDRESS.PREMISE_ID}='+premiseid+' and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.RequestMissingMPRNMPAN1Click(Sender: TObject);
Var
Agid,PremiseID,cid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 PREMISEID:=TreeData.D_Premise_id;
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Missing_SPANS.rpt','Missing SPANS','{SITE_ADDRESS.PREMISE_ID}='+premiseid+' and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.ObjectionReceived1Click(Sender: TObject);
Var
RegID,agid,cid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Objection_received.rpt','Objection Received','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.ObjectionReveived1Click(Sender: TObject);
Var
RegID,agid,cid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Objection_received.rpt','Objection Received','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;


procedure TFRM_Tree.IGTMENUClick(Sender: TObject);
Var
RegID,agid,cid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('GAS\IGT CONFIRMATION SHEET.rpt','IGT Confirmation Form','{SPANS.REGISTRATION_ID}='+REGid,'','',CID,'');
end;


procedure TFRM_Tree.ETUtilitaFault1Click(Sender: TObject);
Var
RegID,agid,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('CRM\ET Reports\Signup_ET_U_Gas_Letter.rpt','Gas ET','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.ElectricETLetter1Click(Sender: TObject);
Var
RegID,agid,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('CRM\ET Reports\Signup_ET_U_Electric_Letter.rpt','Electric ET','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.DUALETUtilitaFault1Click(Sender: TObject);
Var
Agid,PremiseID,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 PREMISEID:=TreeData.D_Premise_id;
 FRM_Reports.PrintThisReport('CRM\ET Reports\Signup_ET_U_Dual_Letter.rpt','Dual ET','{AGREEMENTS.AGREEMENT_ID}='+agid+' and {AGREEMENT_PREMISES.PREMISE_ID}='+premiseid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.ManuallySetSPANStatusSSD1Click(Sender: TObject);
begin
 DoSpanOverride;
end;

procedure TFRM_Tree.MarketingConsent1Click(Sender: TObject);
Var
CustID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Custid:=TreeData.D_Customer_ID;
 FRM_MARKETING_AND_CONSENT.Consent_Query.SetVariable('CID',Custid);
 FRM_MARKETING_AND_CONSENT.doquery;
 FRM_MARKETING_AND_CONSENT.Showmodal;
end;

procedure TFRM_Tree.DoSpanOverride;
var
 SPAN:string;
 Regid,AgreementId: Int64;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 REGID:=strtoint64(TreeData.D_REGid);
 SPAN:=TreeData.D_SPAN;
 AgreementId := strtoint64(treedata.D_Agreement_ID);
 frm_span_override.GetDetails(REGID, AgreementId);
 FRM_Span_override.showmodal;


 if Frm_span_override.tag=0 then exit;

  // Refresh Parent Node
 if treeview1.Selected[xnode]=true then
 Begin
  treeview1.Expanded[xnode.parent]:=false;
  treeview1.Expanded[xnode.parent]:=true;
  //if node is first item, then no paretn so refresh span
  if frm_tree.tag = 4 then FRM_Main.SearchForSpan(SPAN,0);
 end;
end;

procedure TFRM_Tree.RatedIssuesClick(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  if TFrm_Rating_Errors.StartModal(Self, StrToInt64Def(nodeData.D_Agreement_Id, 0)) then
    MessageDlg('Agreement tree must be refreshed to show the changes!', mtWarning, [mbOk], 0);
(*
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Application.CreateForm(TFRM_Rating_Errors, FRM_Rating_Errors);
 try
 with frm_rating_errors.ratingerrors do
 Begin
  close;
  setvariable('Agreement_id',(TreeData.D_agreement_id));
  open;
 End;
 frm_rating_errors.tag:=0;
 Frm_Rating_errors.showmodal;
 if frm_rating_errors.tag=1 then
 Begin
  Messagedlg('You will need to refresh agreement tree to see any changes.',Mtinformation,[MBOK],0);
 End;
 finally
  frm_rating_errors.release;
 end;
*)
end;

procedure TFRM_Tree.ReRateMPAN1Click(Sender: TObject);
begin;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.miRe_RateNowClick(Sender: TObject);
  {----------------------------------------------------------------------------}
  procedure Rerate_OldVersion;
  var
    sAgreementId  : string;
    sConfig0      : string;
    sConfig1      : string;
    sConfig2      : string;
    sConfig3      : string;
    sResult       : string;
  begin
    if Messagedlg('Are you sure you wish to Re-Rate this Agreement?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then Exit;

    Screen.Cursor := crHourGlass;
    XNode := Treeview1.FocusedNode;
    TreeData := Treeview1.GetNodeData(XNode);
    sAgreementId := TreeData.D_Agreement_id;
    try
      // Get List Of Spans for Agreement
      SKIPRATINGERRORS := False;

      try
        sConfig1 := FRM_Main.Config[Ord(ciGas)];
        sConfig2 := FRM_Main.Config[Ord(ciElectricity)];
      except
        on e:Exception do
        begin
          MessageDlg(Format('Invalid configuration: %s - %s!', [FRM_Main.Config, e.Message]), mtError, [mbOk], 0);
          Exit;
        end;
      end;

      if sConfig1 = 'g' then sConfig1 := EmptyStr;
      if sConfig2 = 'e' then sConfig2 := EmptyStr;

      sConfig0 := IfThen(sConfig1 = 'G', 'C', EmptyStr);
      sConfig3 := IfThen(Billing_Do_Credits = 'YES', 'C', EmptyStr);

      sResult := FRM_RATE_ACCOUNTS.ReRateAgreement(sAgreementId, sConfig0, sConfig1,
                           sConfig2, 'P', 'L', 'R', sConfig3, 'X', 'SHOW WARNINGS');

      if sResult <> EmptyStr then
         MessageDlg(sResult, mtError, [mbOk], 0);

      Actioning(EmptyStr);

      Treeview1.Expanded[XNode] := False;
      Treeview1.Expanded[XNode] := True;
    finally
      Screen.Cursor := crDefault;
    end;
  end;
  {----------------------------------------------------------------------------}
  procedure Rerate;
  var
    node          : PVirtualNode;
    noteData      : PMyRec;
    agreementId   : Int64;
    oldCursor     : TCursor;
    c             : char;
    rerateOptions : TRerateOptionSet;
    errorMsg      : string;
  begin
    if MessageDlg('Do you want to rerate the selected agreement?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
      exit;


    node := TreeView1.FocusedNode;
    if not Assigned(node) then
      exit;

    nodeData := TreeView1.GetNodeData(node);
    if not Assigned(nodeData) then
      exit;

    agreementId := StrToInt64Def(nodeData.D_Agreement_Id, 0);
    if agreementId = 0 then
      exit;

    oldCursor     := Screen.Cursor;
    Screen.Cursor := crHourGlass;
    try
      rerateOptions := [];

      // Is Gas should be rated?
      c := Frm_Main.Config[Ord(ciGas)];
      if (c <> '') and (c <> 'g') then
        Include(rerateOptions, roGas);
      if c = 'G' then
        Include(rerateOptions, roCustomBilling);

      // Is Electric should be rated?
      c := Frm_Main.Config[Ord(ciElectricity)];
      if (c <> '') and (c <> 'e') then
        Include(rerateOptions, roElectric);

      if Billing_Do_Credits = RES_YES then
        Include(rerateOptions, roDoCredits);

      Include(rerateOptions, roPayments);
      Include(rerateOptions, roLedger);
      Include(rerateOptions, roRefreshAccountSummary);
      Include(rerateOptions, roCheckForErrors);

      if not TRateAccountsWrapper.RateAgreement(
        agreementId,
        '',
        true,
        true,
        X_MPID,
        StrToIntDef(GlobalBilling, 0),
        false,
        TBillingUtil.GetRerateOptions(rerateOptions),
        errorMsg) then
      begin
        MessageDlg(errorMsg, mtError, [mbOk], 0);
      end;

      // Refresh node
      Treeview1.Expanded[node] := false;
      Treeview1.Expanded[node] := true;
    finally
      Screen.Cursor := oldCursor;
    end;
  end;
  {----------------------------------------------------------------------------}
var
  msg : string;
begin
  msg := 'Do you want to execute the DLL version?' + CRLF +
         'Selecting the "No" button will execute the original version of rerating.';

  if TBillingUtil.UseDllRerating then
  begin
    Rerate;
  end
  else
  begin
    Rerate_OldVersion;
  end;
end;

procedure TFRM_Tree.miRe_RateOvernightClick(Sender: TObject);
begin
  with Main_Data_Module.UpdateQuery do
  begin
    Close;
    DeleteVariables;
    SQL.Clear;
    SQL.Add('Begin');
    SQL.Add(' SALESLEDGER.PK_ACCOUNTS_TO_REQUEST.PR_RERATING_REQUEST(:p_Agreement_Id, :p_PriorityId);');
    SQL.Add('End;');
    DeclareAndSet('p_Agreement_Id', otFloat, StrToFloat(TreeData.D_Agreement_ID));
    DeclareAndSet('p_PriorityId', otInteger, 1);
    try
      Execute;
      DeleteVariables;
      FRM_Login.MainSession.Commit;
      MessageDlg('Overnight ReRate Request has been' + #13 +'added successfully to the queue.', mtInformation, [mbOk], 0);
    except
      on E: Exception do
      begin
        Application.MessageBox(PChar('Error in Overnight ReRate Request -> ' + E.Message), Attn, MB_ICONERROR);
        FRM_Login.MainSession.Rollback;
      end;
    end;
  end;
end;

procedure TFRM_Tree.ReRateHistory1Click(Sender: TObject);
begin
  FRM_ReRate_History := TFRM_Rerate_History.Create(self);

  try
    FRM_ReRate_History.Agreement_id := TreeData.D_Agreement_id;
    FRM_ReRate_History.ShowModal;
  finally
    FreeAndNil(FRM_ReRate_History);
  end;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.ShowSiteVisitInformation1Click(Sender: TObject);
Var
Mpan:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 with Generalquery do
 begin
  close;
  deletevariables;
  declarevariable('MPAN',otstring);
  sql.clear;
  sql.add('Select 1 from edmgr.site_visit_info');
  sql.add('Where mpancore=:MPAN');
  setvariable('MPAN',mpan);
  open;
  deletevariables;
 end;
 if Generalquery.recordcount=0 then
 Begin
  messagedlg('There is no site visit information available for this MPAN.',mtinformation,[MBOK],0);
  exit;
 end;
 FRM_Reports.PrintThisReport('ELECTRIC\Dataflow Reports\MRPT_Site_Visit_Info.rpt','Site Visit Information','{SITE_VISIT_INFO.MPANCORE}='''+mpan+'''','P','','','');
end;

// BSL - 16/12/2014 - Two Way Appointment
Function TFRM_Tree.fOpenTwoWayAppointment(Const aNoCancel: Boolean): Boolean;
Var bMsg,
    bQueryName: String;
Begin
  Try
    Result     := DM_JBS.fOpenTwoWayAppDetails(bMsg);
    bQueryName := 'qrTwoWayAppDetails';

    If Result then
      Begin
        DM_JBS.CustomerId := DM_JBS.qrTwoWayAppDetailsCUSTOMER_ID.AsString;

        If aNoCancel then
          Begin
            bQueryName := 'qrCustContactTel';
            Result     := DM_JBS.fOpenContactTel(bMsg);
          End; // If
      End; // If
  Finally
    If Result then
      Begin
        If DM_JBS.qrTwoWayAppDetails.IsEmpty then
          Begin
            Application.MessageBox(PChar('Job No. --> ' + IntToStr(DM_JBS.JobId) + ', is NOT available for Two Way Appointment.' + bMsg), Attn, MB_ICONINFORMATION);
            Result := False;
          End; // Else
      End // If
    Else
      Application.MessageBox(PChar('SQL Error ' + bQueryName + ' --> ' + bMsg), Attn, MB_ICONERROR);
  End; // Try
End; // Proc

Function TFRM_Tree.fCloseTwoWayAppointment(Const aNoCancel: Boolean): Boolean;
Begin
  If aNoCancel then
    DM_JBS.qrCustContactTel.Close;

  DM_JBS.qrTwoWayAppDetails.Close;
End; // Proc

Procedure TFRM_Tree.acAmendJobExecute(Sender: TObject);
Begin
  If fOpenTwoWayAppointment(True) then
    Begin
      TFRM_TwoWayAppAmend.Launch;
      fCloseTwoWayAppointment(True);
    End; // Proc

  FRM_Main.SearchforCust(DM_JBS.CustomerId);
End; // Proc

Procedure TFRM_Tree.acArrangeJobDtlsExecute(Sender: TObject);
Var bSQL,
    bMsg: String;
    bOk : Boolean;
Begin
  If DM_JBS.fOpenOneWayAppDetails(bMsg) then
    Begin
      bOk := True;

      if ((DM_JBS.qrOneWayAppDetailsJOB_TYPE.AsInteger = 1) or (DM_JBS.qrOneWayAppDetailsJOB_TYPE.AsInteger = 2)) and
           not string.IsNullOrEmpty(DM_JBS.qrOneWayAppDetailsMPAN.AsString) then
      begin
          with main_data_module.TempQuery do
          begin
            DeleteVariables;
            DeclareVariable('MPANID', otString);
            Close;
            SQL.Clear;
            SQL.Add('SELECT CRM.FN_IS_E7(:MPANID) FROM DUAL');
            // 1900028316032
            SetVariable('MPANID', DM_JBS.qrOneWayAppDetailsMPAN.AsString);
            Open;
            DeleteVariables;
          end;

          if UpperCase(main_data_module.TempQuery.fields[0].text) = 'TRUE' then
          begin
            if (Messagedlg('This is an E7 Meter, do you wish to continue booking this job?', mtwarning, [MbYes, MbNo], 0) = MrYes) then
            begin
              if frm_common.Superauthoritycheck=False then
              begin
                Exit;
              end;
            end
            else
            begin
              Exit;
            end;
          end;
      end;

      If Not DM_JBS.fCheckPushBacks then
        If Application.MessageBox('This Job has been pushed back. Would you like to expire the Push Back now, in order to book this Job?', Attn, MB_ICONQUESTION + MB_YESNO) = IDYES then
          Begin
            bSQL := Format('Update SMIFF.WMOL_JOB_PUSHBACK_COUNTER ' +
                           'Set PUSHBACK_DATE     = SYSDATE - INTERVAL ''1'' MINUTE, ' +
                               'LAST_UPDATED_BY   = USER,          ' +
                               'LAST_UPDATED_DATE = SYSDATE        ' +
                           'Where PUSHBACK_DATE > SYSDATE and JOB_ID = %d', [DM_JBS.JobId]);

            bOk := DM_JBS.fExecQuery(bSQL, bMsg);

            FRM_Login.MainSession.Commit;

            If bOk then
              DM_JBS.fOpenOneWayAppDetails(bMsg)
            Else
              Application.MessageBox(PChar('SQL Error Updating WMOL_JOB_PUSHBACK_COUNTER --> ' + bMsg), Attn, MB_ICONERROR);
          End // If
        Else
          bOk := False;

      If bOk then
        If DM_JBS.fLockAJob(bMsg) then
          Begin
            If TFRM_OneWayAppointment.Launch(DM_JBS.JobId,trunc((DM_JBS.qrOneWayAppDetailsWORK_TYPE).Value), DM_JBS.qrOneWayAppDetails) then
              Begin
                FRM_Main.SearchforCust(DM_JBS.CustomerId);

                If Not DM_JBS.fDeleteJobLocked(bMsg) then
                  Application.MessageBox(PChar(bMsg), Attn, MB_ICONERROR);
              End // If
            Else
              If Not DM_JBS.fDeleteJobLocked(bMsg) then
                Application.MessageBox(PChar(bMsg), Attn, MB_ICONERROR);
          End // If
        Else
          Application.MessageBox(PChar(bMsg), Attn, MB_ICONERROR);
    End // If
  Else
    Application.MessageBox(PChar(bMsg), Attn, MB_ICONERROR);
End; // Proc


Procedure TFRM_Tree.acCancelExecute(Sender: TObject);
Var bReason: Byte;
Begin
  If fOpenTwoWayAppointment(False) then
    Begin
      TFRM_CancelJob.Launch(bReason);
      fCloseTwoWayAppointment(False);
    End; // If

  FRM_Main.SearchforCust(DM_JBS.CustomerId);
end;
// BSL - 20/04/2015 - Add items to the pup Menu.
Procedure TFRM_Tree.acChangeJobTypeExecute(Sender: TObject);
Begin
  // Do not delete.  It is necessary to show the action in the PopMenu.
End; // Proc

// BSL - 20/04/2015 - Add items to the pup Menu.
Procedure TFRM_Tree.acChangeJobTypeUpdate(Sender: TObject);
Var i, j    : Integer;
    bCaption: String;
Begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  If Assigned(Treedata) then
    Begin
       Val(TreeData.C_RECORD_ID, i, j);
       bCaption     := Treedata.Caption;
       DM_JBS.JobId := i;

       TAction(Sender).Visible := (j = 0) and (i > 0);

       TAction(Sender).Visible := TAction(Sender).Visible and
                                  Not ((Pos('Completed', bCaption) > 0) or
                                       (Pos('Aborted',   bCaption) > 0) or
                                       (Pos('Cancelled', bCaption) > 0));
       DoWarrentJobLock(Sender);
     End // If
   Else
     TAction(Sender).Visible := False;
End; // Proc

// BSL - 20/04/2015 - Add items to the pup Menu.
Procedure TFRM_Tree.acDualFuelInstallExecute(Sender: TObject);
Var i, j    : Integer;
    bCaption,
    bSQL,
    bMsg    : String;
Begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  If Assigned(Treedata) then
    Begin
       Val(TreeData.C_RECORD_ID, i, j);
       bCaption     := TAction(Sender).Caption;
       DM_JBS.JobId := i;

       If DM_JBS.fOpenAccounts(bMsg) then
         Begin
           If Trim(DM_JBS.qrAccountsCUSTOMER_LIVE.AsString) = EmptyStr then
             Application.MessageBox('You cannot change the Job Type.  Job is not live.', Attn, MB_ICONERROR)
           Else
             // BSL - 25/08/2021 - BRM-1162 - CRM - New work type request for JBS (Replace Battery).
             If (DM_JBS.qrAccountsJOB_TYPE.AsInteger = 4) and (DM_JBS.qrAccountsWORK_TYPE.AsInteger = 33) then // Replace Battery
               Application.MessageBox('You cannot change this Job Type.  Replace Battery belongs only to Check Comms.', Attn, MB_ICONERROR)
             Else
               If (Application.MessageBox(PChar('Are you sure you wish to change the Job Type to: ' + bCaption), Attn, MB_ICONQUESTION + MB_YESNO) =
                IDYES) then
                 Begin
                   bSQL := Format('Update SMIFF.WMOL_ACCOUNTS ' +
                                  'Set JOB_TYPE        = (Select JOB_TYPE_ID ' +
                                                         'From SMIFF.WMOL_JOB_TYPES JT ' +
                                                         'Where JT.DESCRIPTION = %s), ' +
                                      'LAST_UPDATED_BY = USER, ' +
                                      'LAST_UPDATED    = SYSDATE ' +
                                  'Where RECORD_ID = %d ',
                                  [QuotedStr(bCaption), DM_JBS.JobId]);

                    If DM_JBS.fExecQuery(bSQL, bMsg) then
                      Begin
                        FRM_Login.MainSession.Commit;
                        FRM_Main.SearchforCust(DM_JBS.qrAccountsCUSTOMER_ID.AsString);
                      End // If
                    Else
                      Begin
                        FRM_Login.MainSession.Rollback;
                        Application.MessageBox(PChar('SQL Error Update SMIFF.WMOL_ACCOUNTS --> ' + bMsg), Attn, MB_ICONERROR);
                      End; // Else
                 End // If
              Else
                Application.MessageBox('Job Type change, Aborted!.', Attn, MB_ICONWARNING);

           DM_JBS.qrAccounts.Close;
         End // If
       Else
         Application.MessageBox(PChar('SQL Error qrAccount --> ' + bMsg), Attn, MB_ICONERROR);
    End; // If
End; // Proc

// BSL - 20/04/2015 - Add items to the pup Menu.
Procedure TFRM_Tree.acDualFuelInstallUpdate(Sender: TObject);
Var i, j    : Integer;
    bCaption: String;
Begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  If Assigned(Treedata) then
    Begin
       Val(TreeData.C_RECORD_ID, i, j);
       bCaption     := Treedata.Caption;
       DM_JBS.JobId := i;

       TAction(Sender).Visible := (j = 0) and (i > 0);
       // BSL - 17/04/2015 - Be carefull with the Caption of the Action, because it must match with the Caption of the Leyend.
       TAction(Sender).Visible := TAction(Sender).Visible and
                                  (Pos(TAction(Sender).Caption, bCaption) = 0);
       DoWarrentJobLock(Sender);
     End // If
   Else
     TAction(Sender).Visible := False;
End; // Proc

Procedure TFRM_Tree.acReschedulingExecute(Sender: TObject);
var
  compdue: Boolean;
  bMsg, answer: String;
Begin
  compdue := false;
  If (fOpenTwoWayAppointment(True)) and
     (Application.MessageBox(PChar(DM_JBS.qrTwoWayAppDetailsFORMATTED_DATE.AsString + #13 + DM_JBS.qrTwoWayAppDetailsFORMATTED_TIME_SLOT.AsString +
      #13#13 + 'Are you sure you wish to reschedule job date?'), Attn, MB_ICONQUESTION + MB_YESNO) = IDYES) then
    Begin
      bMsg := 'Is the supplier rescheduling? If yes then compensation may be due.';

      if (Application.MessageBox(PChar(bMsg), Attn, MB_ICONQUESTION + MB_YESNO) = IDYES) then
      Begin
        with main_data_module.GeneralQuery do
        Begin
          close;
          sql.clear;
          sql.add('Select * From SMIFF.VW_LIVE_JOBS_IN_NEXT_24HRS ');
          sql.add('Where Record_ID = '+IntToStr(DM_JBS.JobId));
          Open;
          if RecordCount <> 0 then
          begin
            close;
            sql.clear;
            sql.add('Select * From ENQUIRY.GS_JOBS_WITH_COMPENSATION ');
            sql.add('Where JBS_ID = '+IntToStr(DM_JBS.JobId));
            Open;
            compdue := RecordCount = 0;
          end;
          close;
        end;
        if compdue then
        begin
         if (Application.MessageBox(PChar(chr(163)+'30 GSOS Compensation is due.'+#13+'Does the Customer consent to this Reschedule?'), Attn, MB_ICONQUESTION + MB_YESNO)= IDYES) then
         begin
           answer := '''Y''';
           Compdue:= false;
         end
         else
         begin
           answer := '''N''';
           Compdue:= true;
         end;
         with main_data_module.UpdateQuery do
         begin
           Close;
           Sql.Clear;
           DeleteVariables;
           Sql.Add('begin');
           SQL.Add(' SMIFF.PK_ACCOUNTS_ATTRIBUTES.PR_RECORD_ATTRIBUTE(' +IntToStr(DM_JBS.JobId) + ',' + '2' + ',' + answer + ',' +  '''' + UpperCase(UserId) + '''' + ');');
           SQL.Add('end;');
           Execute;
         end;
        end;
      End; // If supplier

      TFRM_OneWayAppointment.Launch(DM_JBS.JobId,trunc((
         DM_JBS.qrOneWayAppDetailsWORK_TYPE).Value), DM_JBS.qrTwoWayAppDetails,
         compdue);
      fCloseTwoWayAppointment(True);
    End; // If contine with reshedule

  FRM_Main.SearchforCust(DM_JBS.CustomerId);
End; // Proc

Procedure TFRM_Tree.acReschedulingUpdate(Sender: TObject);
Var i, j: Integer;
    bCaption: String;
Begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  If Assigned(Treedata) then
    Begin
       Val(TreeData.C_RECORD_ID, i, j);
       bCaption     := Treedata.Caption;
       DM_JBS.JobId := i;

       TAction(Sender).Visible := (j = 0) and (i > 0);

       If TAction(Sender) = acArrangeJobDtls then
         TAction(Sender).Visible := TAction(Sender).Visible and (Pos('Not Yet Booked', bCaption) > 0)
       Else
         Begin
           TAction(Sender).Visible := TAction(Sender).Visible and (Pos('Booked', bCaption) > 0);

           If TAction(Sender) = acRescheduling then
             TAction(Sender).Visible := TAction(Sender).Visible and (Pos('Not Yet Booked', bCaption) = 0)
         End; // Else
       DoWarrentJobLock(Sender);
     End // If
   Else
     TAction(Sender).Visible := False;

End; // Proc
// BSL - 16/12/2014 - End - Two Way Appointment
// BSL - 13/05/2015 - Fuel Direct Execute.
Procedure TFRM_Tree.acAddFuelDirectUpdate(Sender: TObject);
Begin
  TAction(Sender).Visible := FFuelDirect <> 'A';
End; // Proc

Procedure TFRM_Tree.acRemoveFuelDirectUpdate(Sender: TObject);
Begin
  TAction(Sender).Visible := FFuelDirect = 'A';
End; // Proc

Procedure TFRM_Tree.acAddFuelDirectExecute(Sender: TObject);
var
  bSQL, bMsg, bCustId: String;
begin
  bCustId := Treedata.D_Customer_Id;
  If Application.MessageBox('Please, confirm that this customer is enrolled in Fuel Direct', Attn, MB_ICONINFORMATION + MB_YESNO) = IDYES then
    Begin
      // add fuel Direct Mrker to account
      // INSERT OR UPDATE CRM.AGREEMENT_FULE_DIRECT_STATUS
      // STATUS='A' (ADDED) AND LAST UPDATED_DATE=SYSDATE
      // THEN Refresh TREE TO SEE CHANGES.Self
      Try
        If FFuelDirect = 'N' then
          Begin // ADD Record
            bSQL := Format('Insert Into CRM.AGREEMENT_FUEL_DIRECT_STATUS ' +
                           'Values (%s, %s, USER, SYSDATE)',
                            [bCustId, QuotedStr('A')]);

            If DM_JBS.fExecQuery(bSQL, bMsg) then
              Begin
                FRM_Login.MainSession.Commit;
                Application.MessageBox('Customer is enrolled in Fuel Direct.', Attn, MB_ICONINFORMATION);
              End // If
            Else
              Application.MessageBox(PChar('SQL Error adding qrAgrFuelDir --> ' + bMsg), Attn, MB_ICONERROR);
          End // If
        Else // Update the Record - FFuelDirect = 'R'
          Begin
            bSQL := Format('Update CRM.AGREEMENT_FUEL_DIRECT_STATUS ' +
                           'Set STATUS            = %s, ' +
                               'LAST_UPDATED_BY   = USER, ' +
                               'LAST_UPDATED_DATE = SYSDATE ' +
                               'Where agreement_id = %s',
                          [QuotedStr('A'), bCustId]);

            If DM_JBS.fExecQuery(bSQL, bMsg) then
              Begin
                FRM_Login.MainSession.Commit;
                Application.MessageBox('Customer is updated in Fuel Direct.', Attn, MB_ICONINFORMATION);
              End // If
            Else
              Application.MessageBox(PChar('SQL Error updating qrAgrFuelDir --> ' + bMsg), Attn, MB_ICONERROR);
          End; // Else

        qrAgrFuelDir.Close;
        FRM_Main.SearchforCust(bCustId);
      Except
        On E: Exception do
          Begin
            Application.MessageBox(PChar('SQL Error qrAgrFuelDir --> ' + E.Message), Attn, MB_ICONERROR);
          End; // On
      End; // Try
    End; // If
End; // Proc

Procedure TFRM_Tree.acFuelDirectExecute(Sender: TObject);
Begin
  // Do not delete.  It is necessary to show the action in the PopMenu.
End;

Procedure TFRM_Tree.acHistoricCustDtlExecute(Sender: TObject);
Var i, j: Integer;
Begin
  XNode    := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  Val(TreeData.D_Contact_ID, i, j);

  If j = 0 then
    Begin
      TFRM_HistoricCustDtls.Launch(i);
    End; // If
End; // Proc

procedure TFRM_Tree.AcPriorityChangeExecute(Sender: TObject);
Begin
 with frm_job_priority do
 begin
  clearfields;

  ShowModal;
 end;
end;

procedure TFRM_Tree.AcPriorityChangeUpdate(Sender: TObject);
Var i, j: Integer;
    bCaption: String;
Begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  If Assigned(Treedata) then
    Begin
       Val(TreeData.C_RECORD_ID, i, j);
       bCaption     := Treedata.Caption;
       DM_JBS.JobId := i;

       TAction(Sender).Visible := (j = 0) and (i > 0);
       TAction(Sender).Visible := False;

       If TAction(Sender) = acPriorityChange then
       Begin
         TAction(Sender).Visible :=  (Pos('Not Yet Booked', bCaption) > 0);
         TAction(Sender).Visible :=  (Pos('Booked', bCaption) > 0);
         DoWarrentJobLock(Sender);
       End;

     End // If

end;

// Proc

Procedure TFRM_Tree.acRemoveFuelDirectExecute(Sender: TObject);
var
  bSQL, bMsg, bCustId: String;
Begin
  bCustId := Treedata.D_customer_Id;
  If Application.MessageBox('Please, confirm you wish to remove Fuel Direct Marker', Attn, MB_ICONINFORMATION + MB_YESNO) = IDYES then
    Begin
      // Add Fuel Direct Mrker to account
      // INSERT OR UPDATE CRM.AGREEMENT_FULE_DIRECT_STATUS
      // STATUS='R' (ADDED) AND LAST UPDATED_DATE=SYSDATE
      // THEN Refresh TREE TO SEE CHANGES.Self
      Try
        // Update the Record - FFuelDirect = 'A'
        Begin
          bSQL := Format('Update CRM.AGREEMENT_FUEL_DIRECT_STATUS ' +
                         'Set STATUS            = %s, ' +
                             'LAST_UPDATED_BY   = USER, ' +
                             'LAST_UPDATED_DATE = SYSDATE ' +
                         'Where agreement_id = %s',
                          [QuotedStr('R'), bCustId]);

          If DM_JBS.fExecQuery(bSQL, bMsg) then
            Begin
              FRM_Login.MainSession.commit;
              Application.MessageBox('Customer removed from Fuel Direct.', Attn, MB_ICONINFORMATION);
            End // If
          Else
            Application.MessageBox(PChar('SQL Error removing qrAgrFuelDir --> ' + bMsg), Attn, MB_ICONERROR);
        End; // Else

        qrAgrFuelDir.Close;
           FRM_Main.SearchforCust(bCustId);
      Except
        On E: Exception do
          Begin
            Application.MessageBox(PChar('SQL Error qrAgrFuelDir --> ' + E.Message), Attn, MB_ICONERROR);
          End; // On
      End; // Try
    End; // If
End; // Proc
// BSL - 13/05/201

procedure TFRM_Tree.acVATExemptionExecute(Sender: TObject);
var
  bcustid,bSPAN, bSPANType: string;
begin
   xnode:=treeview1.FocusedNode;
   TreeData:= treeview1.GetNodeData(xnode);
   Agreement_id:=TreeData.D_agreement_id;
   bCUSTID := frm_common.getcustomeridfromagreementid(Agreement_id);
   bSPAN := '';
   bSPANType := '';
   bSPAN:=TreeData.D_SPAN;
   bSPANType := TreeData.D_spantype;

   TFRM_VAT_Exemptions.Launch(bcustid, bSPAN, bSPANType);
end;

procedure TFRM_Tree.acVATExemptionUpdate(Sender: TObject);
begin

   // Commentted out for now 26/03/2018,
   // Not Required for Go Live
   TAction(Sender).Visible := False;
   xnode:=treeview1.FocusedNode;
   TreeData:= treeview1.GetNodeData(xnode);
   SPAN:=TreeData.D_SPAN;
   Agreement_id:=TreeData.D_AGREEMENT_ID;

   // only allow VAT updates for commercial customer type from both customer tree and span search tree
   if (TreeData.D_Cust_Type = 2) or (TreeData.D_Cust_Type =  4) then
   begin
      TAction(Sender).Visible := True;
   end
   else
   begin
     TAction(Sender).Visible := False;
   end;
end;

procedure TFRM_Tree.Actioning(ActionText:string);
Begin
 if showfeedback=false then exit;

 FRM_PROCESSING.SETMSG(actiontext);
 Statusbar.Panels[0].text:=actiontext;
 application.processmessages;
End;

procedure TFRM_Tree.AcViewExecute(Sender: TObject);
var
  jbsid, ps, ts, sURL: string;
begin
  // Open Job Rad only in Chrome Browser
  if copy(TreeData.Caption, 1, 8) = '(JBS ID:' then
  begin
    jbsid := TreeData.C_Record_id;
    ps := ' --window-position=0,0 --window-size=600,1000';
    sURL := FRM_Common.getvalue('JBS_WEBPAGE') + 'job.php?job=' + jbsid + '&bypass=crm';
    ShellExecute(Handle, 'open', pchar(sURL), pchar(ps), nil, sw_shownormal);
  end;
end;

procedure TFRM_Tree.ACWhereEngineerExecute(Sender: TObject);
var
  jbsid, ps, ts, eid, sURL: string;
begin
  // Open Job Rad only in Chrome Browser
  if copy(TreeData.Caption, 1, 8) = '(JBS ID:' then
  begin
    jbsid := TreeData.C_Record_id;
    with main_data_module.tempquery do
    begin
      Close;
      Sql.Clear;
      DeleteVariables;
      DeclareVariable(':JOBID', otstring);
      Sql.Add('select engineer_id from SMIFF.VW_TODAYS_PENDING_JOBS where job_id=:JOBID');
      SetVariable('JOBID', jbsid);
      Open;
      DeleteVariables;
    end;
    if main_data_module.tempquery.RecordCount = 0 then
    begin
      Messagedlg
        ('You can only view Engineer Location when job is scheduled for today and status is Booked,On Route or On Site',
        mtinformation, [MBOK], 0);
      exit;
    end;
    eid := main_data_module.tempquery.Fields[0].Text;
    ps := ' --window-position=0,0 --window-size=1180,680';
    sURL := FRM_Common.getvalue('JBS_WEBPAGE') + 'map.php?e=' + eid + '&bypass=crm';
    ShellExecute(Handle, 'open', pchar(sURL), pchar(ps), nil, sw_shownormal);
  end;
end;

procedure TFRM_Tree.ACWhereEngineerUpdate(Sender: TObject);
var
  status: String;
Begin
 // Is Option Enable in Database
 if frm_common.GETVALUE('JBS_ENGINEER_LOCATION') <>'Y' then
 begin
   TAction(Sender).Visible := False;
   exit;
 end;

  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  If Assigned(Treedata) then
    Begin
      Status:=TreeData.C_RECORD_status;
      if (status='Booked') or
         (status='On Route') or
         (status='On Site')  then
           TAction(Sender).Visible := true
           else TAction(Sender).Visible := false;
     End // If
   Else
     TAction(Sender).Visible := False;
End;

procedure TFRM_Tree.G_SETClick(Sender: TObject);
var
regid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Regid:=(TreeData.D_REGid);
 if frm_common.authoritycheck=false then exit;
 setasET(REGID);;
end;

procedure TFRM_Tree.E_SETClick(Sender: TObject);
var
regid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Regid:=(TreeData.D_REGid);
 if frm_common.authoritycheck=false then exit;
 setasET(REGID);
end;

procedure TFRM_TREE.SetasET(Regid:string);
var
agid,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 nodeData := treeview1.GetNodeData(xnode);
 agid:=nodedata.D_Agreement_id;

 If Messagedlg('Are you sure you wish to set this Registration as an Erroneous Transfer?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('Insert into crm.ET_registrations values('+regid+','''+uppercase(userid)+''',sysdate)');
  try
   execute;
  except
  end;
 End;

 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 frm_common.addstatementreviewer(CID);
 frm_login.mainsession.commit;

 // Refresh Parent Node
 if treeview1.Selected[xnode]=true then
 Begin
  treeview1.Expanded[xnode.parent]:=false;
  treeview1.Expanded[xnode.parent]:=true;
  //if node is first item, then no paretn so refresh span
  if treeview1.AbsoluteIndex(XNode)=0 then FRM_Main.SearchForSpan(nodedata.D_SPAN,0);
 end;
end;

procedure TFRM_TREE.SetasDNB(Regid:string);
begin
 xnode:=treeview1.FocusedNode;
 nodeData := treeview1.GetNodeData(xnode);

 If Messagedlg('Are you sure you wish to set this Registration as Do Not Bill?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('Insert into crm.DONOTBILL_registrations values('+regid+')');
  try
   execute;
  except
  end;
 End;
 frm_login.mainsession.commit;

 // Refresh Parent Node
 if treeview1.Selected[xnode]=true then
 Begin
  treeview1.Expanded[xnode.parent]:=false;
  treeview1.Expanded[xnode.parent]:=true;
  //if node is first item, then no paretn so refresh span
  if treeview1.AbsoluteIndex(XNode)=0 then FRM_Main.SearchForSpan(nodedata.D_SPAN,0);
 end;

end;

procedure TFRM_Tree.SendD03061Click(Sender: TObject);
var
MPANTemp : String;
Begin
  Xnode    := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(Xnode);
  FDAP306Requests := TFDAP306Requests.Create(Self);
  FDAP306Requests.ClearAll;
  FDAP306Requests.MPAN.Text     := TreeData.D_SPAN;
  FDAP306Requests.HeaderText.Caption := 'Manual Send D0306';
  MPANTemp                         := TreeData.D_SPAN;
  FDAP306Requests.EditT.Enabled    := True;
  FDAP306Requests.ID.Enabled       := True;
  FDAP306Requests.Address1.Enabled := True;
  FDAP306Requests.Address2.Enabled := True;
  FDAP306Requests.Address3.Enabled := True;
  FDAP306Requests.Address4.Enabled := True;
  FDAP306Requests.Address5.Enabled := True;
  FDAP306Requests.Address6.Enabled := True;
  FDAP306Requests.Address7.Enabled := True;
  FDAP306Requests.Address8.Enabled := True;
  FDAP306Requests.Address9.Enabled := True;
  FDAP306Requests.Postcode.Enabled := True;
  FDAP306Requests.Forename.Enabled := True;
  FDAP306Requests.Surname.Enabled  := True;
  Screen.Cursor                    := Crhourglass;
  With Main_data_module.GeneralQuery Do
  /// get address details. Get Info from DAP.DAP_Generic
    Begin
           Close;
      Sql.Clear;
      DeleteVariables;
      DeclareVariable('SPAN', OtString);
      Sql.Add('select SPAN, CUSTOMER_TYPE, CUSTOMER_ID, TITLE, FORENAME, SURNAME, ADDRESS_LINE_1, ADDRESS_LINE_2, ADDRESS_LINE_3, ADDRESS_LINE_4, ADDRESS_LINE_5, ADDRESS_LINE_6, ADDRESS_LINE_7,');
      Sql.Add(' ADDRESS_LINE_8, ADDRESS_LINE_9, ADDRESS_LINE_10, contact_order from DAP.VW_ALL_CUST_DATA where SPAN = :SPAN and contact_order = 1');
      SetVariable('SPAN', MPANTemp);
      Try
        Open;
        FDAP306Requests.Address1.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_1').Text;
        FDAP306Requests.Address2.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_2').Text;
        FDAP306Requests.Address3.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_3').Text;
        FDAP306Requests.Address4.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_4').Text;
        FDAP306Requests.Address5.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_5').Text;
        FDAP306Requests.Address6.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_6').Text;
        FDAP306Requests.Address7.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_7').Text;
        FDAP306Requests.Address8.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_8').Text;
        FDAP306Requests.Address9.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_9').Text;
        FDAP306Requests.Postcode.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_10').Text;
        FDAP306Requests.Forename.Text := Main_data_module.Generalquery.FieldByName('FORENAME').Text;
        FDAP306Requests.Surname.Text  := Main_data_module.Generalquery.FieldByName('SURNAME').Text;
        FDAP306Requests.EditT.Text  := Main_data_module.Generalquery.FieldByName('TITLE').Text;
        FDAP306Requests.HoldCustID    :=  Main_data_module.Generalquery.FieldByName('CUSTOMER_ID').Text;
        // FDAP306Requests.ID.Text := main_data_module.generalquery.FieldByName('OSID').text;
        If Main_data_module.Generalquery.FieldByName('SPAN').Text = '' Then
          Begin
            If (MessageDlg('No D0306 DAP data on customer ' + MPANTemp + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
              FDAP306Requests.Showmodal;
          End
        Else If Main_data_module.Generalquery.FieldByName('CUSTOMER_TYPE').Text <> 'Domestic' Then
          Begin
            If (MessageDlg('None Domestic Customer Type ' + MPANTemp + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
              FDAP306Requests.Showmodal;
          End
        Else
          FDAP306Requests.Showmodal;
        Close;
      Except
        Screen.Cursor := CrDefault;
        Application.MessageBox('Failed to find D0306 DAP data on customer', 'Warning', MB_OK);
        FDAP306Requests.Showmodal;
      End;
      Close;
      Sql.Clear;
    End;
  If Assigned(FDAP306Requests) Then
    Begin
      FDAP306Requests.Free;
    End;
End;

Procedure TFRM_Tree.SendD03071Click(Sender: TObject);
Var
  DAPPackage: TPkDebtInfo;
  DAPDATA   : PkDebtInfoREstData;
  DAPRes    : String;
Begin
  Xnode           := Treeview1.FocusedNode;
  TreeData        := Treeview1.GetNodeData(Xnode);
  FDAP307Requests := TFDAP307Requests.Create(Self);
  FDAP307Requests.ClearAll;
  FDAP307Requests.MPAN.Text          := TreeData.D_SPAN;
  FDAP307Requests.HeaderText.Caption := 'Manual Send D0307';
  FDAP307Requests.ID.Enabled         := True;
  Screen.Cursor                      := Crhourglass;
  If (TreeData.D_Customer_Id <> EmptyStr) Then
    FDAP307Requests.HoldCustID := TreeData.D_Customer_Id;
  Try
    DAPPackage         := TPkDebtInfo.Create(Self);
    DAPPackage.Session := FRM_Login.MainSession;
    DAPDATA            := PkDebtInfoREstData.Create(FRM_Login.MainSession);
    DAPPackage.PrGetEstData(TreeData.D_SPAN, DAPDATA, DAPRes);

    FDAP307Requests.AddInfo.Text         := DAPDATA.AddInfo;
    FDAP307Requests.DebtRate.Text        := Floattostr(DAPDATA.RecoveryRate);
    FDAP307Requests.DebtOutstanding.Text := Floattostr(DAPDATA.Debt);

    If DAPDATA.Complex = 'T' Then
      FDAP307Requests.ComplexDebt.Text := 'True';
    If DAPDATA.Complex = 'F' Then
      FDAP307Requests.ComplexDebt.Text := 'False';
    If (DAPDATA.Complex <> 'F') And (DAPDATA.Complex <> 'T') Then
      FDAP307Requests.ComplexDebt.Text := 'Unknown';
  Except
    On E: EOracleError Do
      Showmessage(E.Message);
  End;
  If Assigned(DAPPackage) Then
    DAPPackage.Free;
  If Assigned(DAPDATA) Then
    DAPDATA.Free;
  Screen.Cursor := CrDefault;
  If DAPRes = '' Then
    FDAP307Requests.Showmodal;
  If DAPRes <> '' Then
    Begin
      If DAPRes = 'No meter found' Then
        Begin
          If (MessageDlg('No Meter found for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP307Requests.Showmodal;
        End
      Else If DAPRes = 'Multiple meters found' Then
        Begin
          If (MessageDlg('Multiple meters found for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP307Requests.Showmodal;
        End
      Else
        Begin
          If (MessageDlg('Unknown Error for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP307Requests.Showmodal;
        End;
    End;
  FDAP307Requests.Free;
End;

Procedure TFRM_Tree.SendD0308Click(Sender: TObject);
Var
  MPANTemp: String;
  DAPPackage: TPkDebtInfo;
Begin
  Xnode           := Treeview1.FocusedNode;
  TreeData        := Treeview1.GetNodeData(Xnode);
  FDAP308Requests := TFDAP308Requests.Create(Self);
  FDAP308Requests.ClearAll;
  FDAP308Requests.MPAN.Text          := TreeData.D_SPAN;
  FDAP308Requests.HeaderText.Caption := 'Manual Send D0308';
  FDAP308Requests.D0308TimePick.Date := Now + 14;
  MPANTemp                           := TreeData.D_SPAN;
  Screen.Cursor                      := Crhourglass;
  FDAP308Requests.ID.Enabled         := True;
  FDAP308Requests.EditT.Enabled      := True;
  FDAP308Requests.Address1.Enabled   := True;
  FDAP308Requests.Address2.Enabled   := True;
  FDAP308Requests.Address3.Enabled   := True;
  FDAP308Requests.Address4.Enabled   := True;
  FDAP308Requests.Address5.Enabled   := True;
  FDAP308Requests.Address6.Enabled   := True;
  FDAP308Requests.Address7.Enabled   := True;
  FDAP308Requests.Address8.Enabled   := True;
  FDAP308Requests.Address9.Enabled   := True;
  FDAP308Requests.Postcode.Enabled   := True;
  FDAP308Requests.Forename.Enabled   := True;
  FDAP308Requests.Surname.Enabled    := True;
  With Main_data_module.GeneralQuery Do
    Begin
      Close;
      Sql.Clear;
      DeleteVariables;
      DeclareVariable('SPAN', OtString);
      Sql.Add('select SPAN, CUSTOMER_TYPE, CUSTOMER_ID, TITLE, FORENAME, SURNAME, ADDRESS_LINE_1, ADDRESS_LINE_2, ADDRESS_LINE_3, ADDRESS_LINE_4, ADDRESS_LINE_5, ADDRESS_LINE_6, ADDRESS_LINE_7,');
      Sql.Add(' ADDRESS_LINE_8, ADDRESS_LINE_9, ADDRESS_LINE_10, contact_order from DAP.VW_ALL_CUST_DATA where SPAN = :SPAN and contact_order = 1');
      SetVariable('SPAN', MPANTemp);
       DAPPackage                         := TPkDebtInfo.Create(Self);
      Try
        Try
          DAPPackage.Session                 := FRM_Login.MainSession;
          FDAP308Requests.D0308TimePick.Date := DAPPackage.FnResubDate(Now);
        Finally
          If Assigned(DAPPackage) Then
            DAPPackage.Free;
        End;
        Open;
        FDAP308Requests.Address1.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_1').Text;
        FDAP308Requests.Address2.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_2').Text;
        FDAP308Requests.Address3.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_3').Text;
        FDAP308Requests.Address4.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_4').Text;
        FDAP308Requests.Address5.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_5').Text;
        FDAP308Requests.Address6.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_6').Text;
        FDAP308Requests.Address7.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_7').Text;
        FDAP308Requests.Address8.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_8').Text;
        FDAP308Requests.Address9.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_9').Text;
        FDAP308Requests.Postcode.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_10').Text;
        FDAP308Requests.Forename.Text := Main_data_module.Generalquery.FieldByName('FORENAME').Text;
        FDAP308Requests.Surname.Text  := Main_data_module.Generalquery.FieldByName('SURNAME').Text;
        FDAP308Requests.EditT.Text    := Main_data_module.Generalquery.FieldByName('TITLE').Text;
        FDAP308Requests.HoldCustID    := Main_data_module.Generalquery.FieldByName('CUSTOMER_ID').Text;
        Screen.Cursor                 := CrDefault;
        If Main_data_module.Generalquery.FieldByName('SPAN').Text = '' Then
          Begin
            If (MessageDlg('No D0308 DAP data on customer ' + MPANTemp + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
              FDAP308Requests.Showmodal;
          End
        Else If Main_data_module.Generalquery.FieldByName('CUSTOMER_TYPE').Text <> 'Domestic' Then
          Begin
            If (MessageDlg('None Domestic Customer Type ' + MPANTemp + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
              FDAP308Requests.Showmodal;
          End
        Else
          FDAP308Requests.Showmodal;
        Close;
      Except
        Screen.Cursor := CrDefault;
        Application.MessageBox('Failed to find D0308 DAP data on customer', 'Warning', MB_OK);
        FDAP308Requests.Showmodal;
      End;
      Close;
      Sql.Clear;
    End;
  If Assigned(FDAP308Requests) Then
    Begin
      FDAP308Requests.Free;
    End;
End;

Procedure TFRM_Tree.SendD0309Click(Sender: TObject);
Var
  MPANTemp, DAPRes: String;
  DAPPackage: TPkDebtInfo;
  DAPDATA   : PkDebtInfoRActData;
Begin
  Xnode           := Treeview1.FocusedNode;
  TreeData        := Treeview1.GetNodeData(Xnode);
  FDAP309Requests := TFDAP309Requests.Create(Self);
  FDAP309Requests.ClearAll;
  FDAP309Requests.MPAN.Text          := TreeData.D_SPAN;
  FDAP309Requests.HeaderText.Caption := 'Manual Send D0309';
  if (TreeData.D_Customer_Id <> EmptyStr) then
    FDAP309Requests.HoldCustID := TreeData.D_Customer_Id;
  MPANTemp                           := TreeData.D_SPAN;
  FDAP309Requests.ID.Enabled         := True;
  screen.Cursor := crhourglass;
  Try
    DAPPackage         := TPkDebtInfo.Create(Self);
    DAPPackage.Session := FRM_Login.MainSession;
    DAPDATA            := PkDebtInfoRActData.Create(FRM_Login.MainSession);
    DAPPackage.PrGetActData(MPANTemp, DAPDATA, DAPRes);
    FDAP309Requests.RecoveryRate.Text       := Floattostr(DAPDATA.RecoveryRate);
    FDAP309Requests.EstimatedTotalDebt.Text := Floattostr(DAPDATA.EstDebt);
    FDAP309Requests.TotalDebt.Text          := Floattostr(DAPDATA.TtlDebt);
    FDAP309Requests.VAT.Text                := Floattostr(DAPDATA.Vat);
    FDAP309Requests.TotalPayments.Text      := Floattostr(DAPDATA.FactoredPayment);
  Except
    On E: EOracleError Do
      Showmessage(E.Message);
  End;
  If Assigned(DAPPackage) Then
    DAPPackage.Free;
  If Assigned(DAPDATA) Then
    DAPDATA.Free;
  Screen.Cursor := CrDefault;
  If DAPRes = '' Then
    FDAP309Requests.Showmodal;
  If DAPRes <> '' Then
    Begin
      If DAPRes = 'No meter found' Then
        Begin
          If (MessageDlg('No Meter found for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP309Requests.Showmodal;
        End
      Else If DAPRes = 'Multiple meters found' Then
        Begin
          If (MessageDlg('Multiple meters found for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP309Requests.Showmodal;
        End
      Else
        Begin
          If (MessageDlg('Unknown Error for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP309Requests.Showmodal;
        End;
    End;
  If Assigned(FDAP309Requests) Then
    FDAP309Requests.Free;
End;

procedure TFRM_Tree.SendExtracareletter1Click(Sender: TObject);
var
  storedProcedure, AgreementID, CustomerID : String;
begin
  Xnode           := Treeview1.FocusedNode;
  TreeData        := Treeview1.GetNodeData(Xnode);

  AgreementID := Trim(TreeData.D_Agreement_ID);
  CustomerID := Trim(TreeData.D_Customer_Id);

  if not ((CustomerID = EmptyStr) or (AgreementID = EmptyStr)) then
  begin
    if MessageDlg('Are you sure you wish to send extra care letter for this customer?', MtConfirmation, [MbYes, MbNo], 0) = MrYes Then
    begin
      storedProcedure :=
      'crm.pr_add_to_letter_queue_d_me_3('+
      ':p_customer_id, '+
      ':p_agreement_id, '+
      ':p_letter_ref)';
      Try
        gSqlUtil.ExecProc(storedProcedure, TRANSACTION_YES,
        [
          'p_customer_id',   otString, pdInput, CustomerID,
          'p_agreement_id',  otString, pdInput, AgreementID,
          'p_letter_ref',    otString, pdInput, 'D-ME-3']);
        MessageDlg('The extra care letter has been sent to the print queue.', mtInformation, [mbOK], 0);
      except on E: Exception do
           Messagedlg('An error occurred when sending extra care letter: ' + E.Message, TMsgDlgType.mtError,[mbok],0);
      end;
    end;
  end
  else
    MessageDlg('An error occurred, please try again.', mtInformation, [mbOK], 0);
end;

procedure TFRM_Tree.SendfurtherNOI1Click(Sender: TObject);
var
  storedProcedure, AgreementID, CustomerID : String;
begin
  Xnode           := Treeview1.FocusedNode;
  TreeData        := Treeview1.GetNodeData(Xnode);

  AgreementID := Trim(TreeData.D_Agreement_ID);
  CustomerID := Trim(TreeData.D_Customer_Id);

  if not ((CustomerID = EmptyStr) or (AgreementID = EmptyStr)) then
  begin
    if MessageDlg('Are you sure you wish to send further NOI for this customer?', MtConfirmation, [MbYes, MbNo], 0) = MrYes Then
    begin
      storedProcedure :=
      'cbt.pk_letters_cbt_ni2.pr_add_letter_to_print_queue('+
      ':p_customer_id, '+
      ':p_agreement_id)';
      Try
        gSqlUtil.ExecProc(storedProcedure, TRANSACTION_YES,
        [
          'p_customer_id',   otString, pdInput, CustomerID,
          'p_agreement_id',  otString, pdInput, AgreementID]);
        MessageDlg('The further NOI has been sent to the print queue.', mtInformation, [mbOK], 0);
      except on E: Exception do
           Messagedlg('An error occurred when sending further NOI: ' + E.Message, TMsgDlgType.mtError,[mbok],0);
      end;
    end;
  end;
end;

procedure TFRM_Tree.SendG08061Click(Sender: TObject);
Var
  MPRNTemp: String;
Begin
  Xnode           := Treeview1.FocusedNode;
  TreeData        := Treeview1.GetNodeData(Xnode);
  FDAP306Requests := TFDAP306Requests.Create(Self);
  FDAP306Requests.ClearAll;
  FDAP306Requests.MPAN.Text          := TreeData.D_SPAN;
  FDAP306Requests.HeaderText.Caption := 'Manual Send G0806';
  MPRNTemp                           := TreeData.D_SPAN;
  FDAP306Requests.EditT.Enabled      := True;
  FDAP306Requests.ID.Enabled         := True;
  FDAP306Requests.Address1.Enabled   := True;
  FDAP306Requests.Address2.Enabled   := True;
  FDAP306Requests.Address3.Enabled   := True;
  FDAP306Requests.Address4.Enabled   := True;
  FDAP306Requests.Address5.Enabled   := True;
  FDAP306Requests.Address6.Enabled   := True;
  FDAP306Requests.Address7.Enabled   := True;
  FDAP306Requests.Address8.Enabled   := True;
  FDAP306Requests.Address9.Visible := False;
  FDAP306Requests.Postcode.Visible := False;
  FDAP306Requests.Label1.Caption   := 'Building No:';
  FDAP306Requests.Label2.Caption   := 'Sub Building:';
  FDAP306Requests.Label3.Caption   := 'Building:';
  FDAP306Requests.Label4.Caption   := 'Street:';
  FDAP306Requests.Label5.Caption   := 'Locality:';
  FDAP306Requests.Label6.Caption   := 'Town:';
  FDAP306Requests.Label7.Caption   := 'Postcode 1:';
  FDAP306Requests.Label8.Caption   := 'Postcode 2:';
  FDAP306Requests.ID.MaxLength     := 3;
  FDAP306Requests.Label9.Visible   := False;
  FDAP306Requests.Label10.Visible  := False;
  FDAP306Requests.Height           := 310;
  Screen.Cursor                      := Crhourglass;
  FDAP306Requests.Forename.Enabled := True;
  FDAP306Requests.Surname.Enabled  := True;
  With Main_data_module.GeneralQuery Do
    Begin
      Close;
      Sql.Clear;
      DeleteVariables;
      DeclareVariable('SPAN', OtString);
      Sql.Add('select SPAN, CUSTOMER_TYPE, CUSTOMER_ID, TITLE, FORENAME, SURNAME, ADDRESS_LINE_1, ADDRESS_LINE_2, ADDRESS_LINE_3, ADDRESS_LINE_4, ADDRESS_LINE_5, ADDRESS_LINE_6, ADDRESS_LINE_7,');
      Sql.Add(' ADDRESS_LINE_8, ADDRESS_LINE_9, ADDRESS_LINE_10, contact_order from DAP.VW_ALL_CUST_DATA where SPAN = :SPAN and contact_order = 1');
      SetVariable('SPAN', MPRNTemp);
      Try
        Open;
        FDAP306Requests.Address1.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_1').Text;
        FDAP306Requests.Address2.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_2').Text;
        FDAP306Requests.Address3.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_3').Text;
        FDAP306Requests.Address4.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_4').Text;
        FDAP306Requests.Address5.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_5').Text;
        FDAP306Requests.Address6.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_6').Text;
        FDAP306Requests.Address7.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_7').Text;
        FDAP306Requests.Address8.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_8').Text;
        FDAP306Requests.Forename.Text := Main_data_module.Generalquery.FieldByName('FORENAME').Text;
        FDAP306Requests.Surname.Text  := Main_data_module.Generalquery.FieldByName('SURNAME').Text;
        FDAP306Requests.EditT.Text    := Main_data_module.Generalquery.FieldByName('TITLE').Text;
        FDAP306Requests.HoldCustID    := Main_data_module.Generalquery.FieldByName('CUSTOMER_ID').Text;
        Screen.Cursor                 := CrDefault;
        If Main_data_module.Generalquery.FieldByName('SPAN').Text = '' Then
          Begin
            If (MessageDlg('No G0806 DAP data on customer ' + MPRNTemp + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
              FDAP306Requests.Showmodal;
          End
        Else If Main_data_module.Generalquery.FieldByName('CUSTOMER_TYPE').Text <> 'Domestic' Then
          Begin
            If (MessageDlg('None Domestic Customer Type ' + MPRNTemp + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
              FDAP306Requests.Showmodal;
          End
        Else
          FDAP306Requests.Showmodal;
        Close;
      Except
        Screen.Cursor := CrDefault;
        Application.MessageBox('Failed to find G0806 DAP data on customer', 'Warning', MB_OK);
        FDAP306Requests.Showmodal;
      End;
      Close;
      DeleteVariables;
      Sql.Clear;
    End;
  If Assigned(FDAP306Requests) Then
    Begin
      FDAP306Requests.Free;
    End;
End;

Procedure TFRM_Tree.SendG08071Click(Sender: TObject);
Var
  DAPPackage: TPkDebtInfo;
  DAPDATA   : PkDebtInfoREstData;
  DAPRes : String;
Begin
  Xnode           := Treeview1.FocusedNode;
  TreeData        := Treeview1.GetNodeData(Xnode);
  FDAP307Requests := TFDAP307Requests.Create(Self);
  FDAP307Requests.ClearAll;
  FDAP307Requests.MPAN.Text          := TreeData.D_SPAN;
  FDAP307Requests.HeaderText.Caption := 'Manual Send G0807';
  FDAP307Requests.ID.Enabled         := True;
  FDAP307Requests.ID.MaxLength       := 3;
  screen.Cursor := crhourglass;
    if (TreeData.D_Customer_Id <> EmptyStr) then
    FDAP307Requests.HoldCustID := TreeData.D_Customer_Id;
  Try
    DAPPackage         := TPkDebtInfo.Create(Self);
    DAPPackage.Session := FRM_Login.MainSession;
    DAPDATA            := PkDebtInfoREstData.Create(FRM_Login.MainSession);
    DAPPackage.PrGetEstData(TreeData.D_SPAN, DAPDATA, DAPRes);
    FDAP307Requests.AddInfo.Text         := DAPDATA.AddInfo;
    FDAP307Requests.DebtRate.Text        := Floattostr(DAPDATA.RecoveryRate);
    FDAP307Requests.DebtOutstanding.Text := Floattostr(DAPDATA.Debt);
    If DAPDATA.Complex = 'T' Then
      FDAP307Requests.ComplexDebt.Text := 'True';
    If DAPDATA.Complex = 'F' Then
      FDAP307Requests.ComplexDebt.Text := 'False';
    If (DAPDATA.Complex <> 'F') And (DAPDATA.Complex <> 'T') Then
      FDAP307Requests.ComplexDebt.Text := 'Unknown';
  Except
    On E: EOracleError Do
      Showmessage(E.Message);
  End;
  If Assigned(DAPPackage) Then
    DAPPackage.Free;
  If Assigned(DAPDATA) Then
    DAPDATA.Free;
  screen.Cursor := crDefault;
  If DAPRes = '' Then
    FDAP307Requests.Showmodal;
    If DAPRes <> '' Then
    Begin
      If DAPRes = 'No meter found' Then
        Begin
          If (MessageDlg('No Meter found for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP307Requests.Showmodal;
        End
      Else If DAPRes = 'Multiple meters found' Then
        Begin
          If (MessageDlg('Multiple meters found for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP307Requests.Showmodal;
        End
      Else
        Begin
          If (MessageDlg('Unknown Error for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP307Requests.Showmodal;
        End;
    End;
  FDAP307Requests.Free;
End;

Procedure TFRM_Tree.SendG08081Click(Sender: TObject);
Var
  MPRNTemp: String;
  DAPPackage: TPkDebtInfo;
Begin
  Xnode           := Treeview1.FocusedNode;
  TreeData        := Treeview1.GetNodeData(Xnode);
  FDAP308Requests := TFDAP308Requests.Create(Self);
  FDAP308Requests.ClearAll;
  FDAP308Requests.MPAN.Text          := TreeData.D_SPAN;
  FDAP308Requests.HeaderText.Caption := 'Manual Send G0808';
  MPRNTemp                           := TreeData.D_SPAN;
  FDAP308Requests.EditT.Enabled      := True;
  FDAP308Requests.ID.Enabled         := True;
  FDAP308Requests.Address1.Enabled   := True;
  FDAP308Requests.Address2.Enabled   := True;
  FDAP308Requests.Address3.Enabled   := True;
  FDAP308Requests.Address4.Enabled   := True;
  FDAP308Requests.Address5.Enabled   := True;
  FDAP308Requests.Address6.Enabled   := True;
  FDAP308Requests.Address7.Enabled   := True;
  FDAP308Requests.Address8.Enabled   := True;
  FDAP308Requests.Address9.Visible   := False;
  FDAP308Requests.Postcode.Visible   := False;
  FDAP308Requests.Label1.Caption     := 'Building No:';
  FDAP308Requests.Label2.Caption     := 'Sub Building:';
  FDAP308Requests.Label3.Caption     := 'Building:';
  FDAP308Requests.Label4.Caption     := 'Street:';
  FDAP308Requests.Label5.Caption     := 'Locality:';
  FDAP308Requests.Label6.Caption     := 'Town:';
  FDAP308Requests.Label7.Caption     := 'Postcode 1:';
  FDAP308Requests.Label8.Caption     := 'Postcode 2:';
  FDAP308Requests.ID.MaxLength       := 3;
  FDAP308Requests.Label9.Visible     := False;
  FDAP308Requests.Label10.Visible    := False;
  FDAP308Requests.Height             := 390;
  Screen.Cursor                      := Crhourglass;
  FDAP308Requests.Forename.Enabled   := True;
  FDAP308Requests.Surname.Enabled    := True;
  With Main_data_module.GeneralQuery Do
    Begin
      Close;
      Sql.Clear;
      DeleteVariables;
      DeclareVariable('SPAN', OtString);
      Sql.Add('select SPAN, CUSTOMER_TYPE, CUSTOMER_ID, TITLE, FORENAME, SURNAME, ADDRESS_LINE_1, ADDRESS_LINE_2, ADDRESS_LINE_3, ADDRESS_LINE_4, ADDRESS_LINE_5, ADDRESS_LINE_6, ADDRESS_LINE_7,');
      Sql.Add(' ADDRESS_LINE_8, ADDRESS_LINE_9, ADDRESS_LINE_10, contact_order from DAP.VW_ALL_CUST_DATA where SPAN = :SPAN and contact_order = 1');
      SetVariable('SPAN', MPRNTemp);
      DAPPackage                         := TPkDebtInfo.Create(Self);
      Try
        Try
          DAPPackage.Session                 := FRM_Login.MainSession;
          FDAP308Requests.D0308TimePick.Date := DAPPackage.FnResubDate(Now);
        Finally
          If Assigned(DAPPackage) Then
            DAPPackage.Free;
        End;
        Open;
        FDAP308Requests.Address1.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_1').Text;
        FDAP308Requests.Address2.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_2').Text;
        FDAP308Requests.Address3.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_3').Text;
        FDAP308Requests.Address4.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_4').Text;
        FDAP308Requests.Address5.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_5').Text;
        FDAP308Requests.Address6.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_6').Text;
        FDAP308Requests.Address7.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_7').Text;
        FDAP308Requests.Address8.Text := Main_data_module.Generalquery.FieldByName('ADDRESS_LINE_8').Text;
        FDAP308Requests.Forename.Text := Main_data_module.Generalquery.FieldByName('FORENAME').Text;
        FDAP308Requests.Surname.Text  := Main_data_module.Generalquery.FieldByName('SURNAME').Text;
        FDAP308Requests.EditT.Text    := Main_data_module.Generalquery.FieldByName('TITLE').Text;
        FDAP308Requests.HoldCustID    := Main_data_module.Generalquery.FieldByName('CUSTOMER_ID').Text;
        Screen.Cursor                 := CrDefault;
        If Main_data_module.Generalquery.FieldByName('SPAN').Text = '' Then
          Begin
            If (MessageDlg('No G0808 DAP data on customer ' + MPRNTemp + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
              FDAP308Requests.Showmodal;
          End
        Else If Main_data_module.Generalquery.FieldByName('CUSTOMER_TYPE').Text <> 'Domestic' Then
          Begin
            If (MessageDlg('None Domestic Customer Type ' + MPRNTemp + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
              FDAP308Requests.Showmodal;
          End
        Else
          FDAP308Requests.Showmodal;
        Close;
      Except
        Screen.Cursor := CrDefault;
        Application.MessageBox('Failed to find G0808 DAP data on customer', 'Warning', MB_OK);
        FDAP308Requests.Showmodal;
      End;
      If Assigned(FDAP308Requests) Then
        Begin
          FDAP308Requests.Free;
        End;
    End;
End;

  Procedure TFRM_Tree.SendG08091Click(Sender: TObject);
Var
  MPRNTemp, DAPRes: String;
    DAPPackage: TPkDebtInfo;
  DAPDATA   : PkDebtInfoRActData;
Begin
  Xnode           := Treeview1.FocusedNode;
  TreeData        := Treeview1.GetNodeData(Xnode);
  FDAP309Requests := TFDAP309Requests.Create(Self);
  FDAP309Requests.ClearAll;
  FDAP309Requests.MPAN.Text          := TreeData.D_SPAN;
  FDAP309Requests.HeaderText.Caption := 'Manual Send G0809';
  FDAP309Requests.ID.MaxLength       := 3;
  MPRNTemp                           := TreeData.D_SPAN;
  FDAP309Requests.ID.Enabled         := True;
  screen.Cursor := crhourglass;
    if (TreeData.D_Customer_Id <> EmptyStr) then
    FDAP309Requests.HoldCustID := TreeData.D_Customer_Id;
  Try
    DAPPackage         := TPkDebtInfo.Create(Self);
    DAPPackage.Session := FRM_Login.MainSession;
    DAPDATA            := PkDebtInfoRActData.Create(FRM_Login.MainSession);
    DAPPackage.PrGetActData(MPRNTemp, DAPDATA, DAPRes);
    FDAP309Requests.RecoveryRate.Text       := Floattostr(DAPDATA.RecoveryRate);
    FDAP309Requests.EstimatedTotalDebt.Text := Floattostr(DAPDATA.EstDebt);
    FDAP309Requests.TotalDebt.Text          := Floattostr(DAPDATA.TtlDebt);
    FDAP309Requests.VAT.Text                := Floattostr(DAPDATA.Vat);
    FDAP309Requests.TotalPayments.Text      := Floattostr(DAPDATA.FactoredPayment);

  Except
   On E: EOracleError Do
      Showmessage(E.Message);
  End;

  If Assigned(DAPPackage) Then
      DAPPackage.Free;
  If Assigned(DAPDATA) Then
      DAPDATA.Free;

  Screen.Cursor := CrDefault;
  If DAPRes = '' Then
    FDAP309Requests.Showmodal;
   If DAPRes <> '' Then
    Begin
      If DAPRes = 'No meter found' Then
        Begin
          If (MessageDlg('No Meter found for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP309Requests.Showmodal;
        End
      Else If DAPRes = 'Multiple meters found' Then
        Begin
          If (MessageDlg('Multiple meters found for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP309Requests.Showmodal;
        End
      Else
        Begin
          If (MessageDlg('Unknown Error for ' + TreeData.D_SPAN + ' Do you wish to continue ?', MtConfirmation, [MbYes, MbNo], 0) = MrYes) Then
            FDAP309Requests.Showmodal;
        End;
    End;

  If Assigned(FDAP309Requests) Then
    FDAP309Requests.Free;
End;

// BSL - 15/03/2023 - ISC-547 New SMS in CRM providing additional support links post ASC decline
                   // Code optimization
Procedure TFRM_Tree.SendSMSMessage1Click(Sender: TObject);
Var
  bMobile,
  bFn    : String;
  i      : Integer;
Begin
  XNode    := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);
  Custid   := TreeData.D_Customer_Id;
  bMobile  := TreeData.D_Tel;
  bMobile  := StringReplace(bMobile, ' ', EmptyStr, [rfReplaceAll]);

  With Sender As TMenuItem do
    Begin
      i   := Tag;
      bFn := Hint;
    End;

  // Default
  If i <> 0 then
    Begin
      Main_Data_Module.TempQuery.Close;
      Main_Data_Module.TempQuery.SQL.Text := 'Select ' + bFn + '(' + CustId + ',' + IntToStr(i) + ') as Msg From Dual';
      Main_Data_Module.TempQuery.Open;

      If Main_Data_Module.TempQuery.FieldByName('Msg').AsString <> EmptyStr then
        FRM_SMS.Populate_SMS_To(EmptyStr, CustId, bMobile, Main_Data_Module.TempQuery.FieldByName('Msg').AsString, True)
      Else
        MessageDlg('No data was returned for this SMS request.', mtError, [mbOk], 0);
    End
  Else
    FRM_SMS.Populate_SMS_To(EmptyStr, CustId, bMobile, EmptyStr, False);
End;

procedure TFRM_TREE.SetasBill(Regid:string);
begin
 xnode:=treeview1.FocusedNode;
 nodeData := treeview1.GetNodeData(xnode);

 If Messagedlg('Are you sure you wish allow billing of this registration?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from crm.DONOTBILL_registrations where registration_id='+regid);
  try
   execute;
  except
  end;
 End;
 frm_login.mainsession.commit;
 // Refresh Parent Node
 if treeview1.Selected[xnode]=true then
 Begin
  treeview1.Expanded[xnode.parent]:=false;
  treeview1.Expanded[xnode.parent]:=true;
  //if node is first item, then no paretn so refresh span
  if treeview1.AbsoluteIndex(XNode)=0 then FRM_Main.SearchForSpan(nodedata.D_SPAN,0);
 end;

end;

procedure TFRM_TREE.RemoveET(Regid:string);
begin
 xnode:=treeview1.FocusedNode;
 nodeData := treeview1.GetNodeData(xnode);

 If Messagedlg('Are you sure you wish to Remove this Erroneous Transfer?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from crm.ET_registrations where registration_id='+regid+'');
  try
   execute;
  except
  end;
 End;
 frm_login.mainsession.Commit;

 // Refresh Parent Node
 if treeview1.Selected[xnode]=true then
 Begin
  treeview1.Expanded[xnode.parent]:=false;
  treeview1.Expanded[xnode.parent]:=true;
  //if node is first item, then no paretn so refresh span
  if treeview1.AbsoluteIndex(XNode)=0 then FRM_Main.SearchForSpan(nodedata.D_SPAN,0);
 end;



end;

procedure TFRM_Tree.SignupLetter1Click(Sender: TObject);
Var
REGID,AGID,ntext,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Application.CreateForm(TFRM_TEL_FAF, FRM_TEL_FAF);
 try
 FRM_TEL_FAF.tag:=0;
 FRM_TEL_FAF.SHOWMODAL;
 if FRM_TEL_FAF.TAG<>0 then
 Begin
  REGID:=TreeData.D_REGid;
  AGID:=TreeData.D_Agreement_id;
  Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
  ntext:=frm_common.moneyformat(frm_tel_faf.DDAmount.value)+' on '+frm_tel_faf.DDDate.Text;
  FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Letter_TELECOM.rpt',Ntext,'{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
 end;
 finally
  FRM_TEL_FAF.release;
 end;
end;

procedure TFRM_Tree.FriendsFamilyForm1Click(Sender: TObject);
Var
REGID,AGID,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 REGID:=TreeData.D_REGid;
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_FAF_TELECOMS.rpt','Telecoms FAF','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.InfoPackApplicationForm1Click(Sender: TObject);
Var
CUSTid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 CUSTID:=TreeData.D_CUSTOMER_id;
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\Info_letter_telecoms.rpt','Telecoms Info Pack','{CUSTOMER.CUSTOMER_ID} = '+CUSTid,'','PRINTER',CUSTID,'');
end;

procedure TFRM_Tree.d1Click(Sender: TObject);
var
agreement_id:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 agreement_id:=TreeData.D_agreement_id;

 if messagedlg('Are you sure you wish to Place System Account Reviewer on Hold?'+#13+
               'Agreement ID '+agreement_id,mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
 with main_data_module.updatequery do
 Begin
  close;
  sql.Clear;
  sql.add('Insert into crm.account_reviews_on_hold');
  sql.add('values ('+agreement_id+','''+userid+''',sysdate)');
  try
   execute;
   frm_login.mainsession.commit;
  except
   Messagedlg('System Account reviewer already on Hold',mtinformation,[mbok],0);
   exit;
  end;
 End;

 treeview1.Expanded[xnode]:=false;
 treeview1.Expanded[xnode]:=true;

end;

procedure TFRM_Tree.MenuItem5Click(Sender: TObject);
Var
Ag,QU,Act,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 nodeData := treeview1.GetNodeData(xnode);

 ACT:=nodedata.D_ACTIONED;
 AG:=nodedata.D_agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Ag);
 QU:=nodedata.D_PERIOD_id;

 if act='Y' then
 Begin
  Messagedlg('You cannot delete this reviewer as it as already been actioned.'+#13+
             'Any changes to DD schedule must be done manually.',mtinformation,[MBOK],0);
  exit;
 End;
 If Messagedlg('Are you sure you wish to delete this Review?'+#13+
               '** NOTE:**'+#13+
               'DELETE if you want account to be Re-Reviewed.'+#13+
               'Do NOT delete if you don''t want another review.',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;

 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('Delete from crm.account_reviews');
  sql.add('where agreement_id='+ag+' and Statement_reviewed='''+qu+'''');
  execute;
  // Add Note To Account, Reviwer Deleted
  close;
  sql.clear;
  sql.add('Insert into enquiry.enquiries values(NULL,'''+uppercase(userid)+''',sysdate,''14'',''301'',null,''Account Reviewer Deleted.'',null,''Y'',null,'''+uppercase(userid)+''',null,NULL,trunc(sysdate),null,'+CID+',null,'+frm_Common.NextNoteId+',''X'',null)');
  execute;
 End;
 FRM_Login.MainSession.commit;

 treeview1.expanded[xnode.parent]:=false;
 treeview1.expanded[xnode.parent]:=true;
end;

procedure TFRM_Tree.MovetoNewPPMAgreement1Click(Sender: TObject);
begin
 TransferSpan;
end;

procedure TFRM_Tree.GasReadRequired1Click(Sender: TObject);
Var
Agid,PremiseID,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 PREMISEID:=TreeData.D_Premise_id;
 FRM_Reports.PrintThisReport('CRM\Meter Readings\Read_Required_Gas.rpt','Meter Read Required','{SITE_ADDRESS.PREMISE_ID}='+premiseid+' and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.ElectricReadRequired1Click(Sender: TObject);
Var
Agid,PremiseID,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 PREMISEID:=TreeData.D_Premise_id;
 FRM_Reports.PrintThisReport('CRM\Meter Readings\Read_Required_Electric.rpt','Meter Read Required','{SITE_ADDRESS.PREMISE_ID}='+premiseid+' and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
end;


procedure TFRM_Tree.DualReadsRequired1Click(Sender: TObject);
Var
Agid,PremiseID,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 PREMISEID:=TreeData.D_Premise_id;
 FRM_Reports.PrintThisReport('CRM\Meter Readings\Read_Required_Prem.rpt','Meter Read Required','{SITE_ADDRESS.PREMISE_ID}='+premiseid+' and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.MenuItem6Click(Sender: TObject);
Var
Agid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;

 If Messagedlg('Are you sure you wish to take this account reviewer off hold?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;
  with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from crm.account_reviews_on_hold where agreement_id='+agid);
  execute;
 end;
 frm_login.mainsession.commit;

 treeview1.expanded[xnode.parent]:=false;
 treeview1.expanded[xnode.parent]:=true;

end;

procedure TFRM_Tree.RegistrationHistory1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 FRM_Dataflow_history.showhistory(TreeData.D_SPAN);
end;


procedure TFRM_Tree.G_RETClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
RemoveET(TreeData.D_REGid);
end;

procedure TFRM_Tree.E_RETClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 RemoveET(TreeData.D_REGid);
end;

procedure TFRM_Tree.CurrentTariffSheet1Click(Sender: TObject);
Var
Agid,PPLAN,trf:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AgID:=TreeData.D_Agreement_id;
 with productquery do
 Begin
  close;
  sql.clear;
  sql.Add('select PRICE_PLAN_ID');
  sql.add('from CRM.AGREEMENT_PRODUCTS');
  sql.add('where agreement_id='+agid);
  sql.add('order by DATE_SETUP desc');
  open;
  first;
 end;
 try
  pplan:=productquery.fields[0].text;
 except
  pplan:='';
 end;
 if (pplan='4') or (pplan='U') then Trf:='energysaverplus'
 else if (pplan='5') then Trf:='energysaver'
 else if (pplan='6') then Trf:='energysaverextra'
 else if (pplan='7') then Trf:='planetsaver'
 else if (pplan='S') then Trf:='utilitastandard'
 else
 Begin
  Messagedlg('Unsupported Price Book Tariff',mtinformation,[MBOK],0);
  exit;
 End;
 FRM_Reports.PrintThisReport('RATING_BILLING\PRICEBOOK.rpt','Utilita Tariff Book','{PRICE_BOOK_GAS.INC_OR_EX}=''Excluding VAT'' and {PRICE_BOOK_GAS.EFFECTIVE_FROM}=Date (2005,09 ,01 ) and {PRICE_BOOK_GAS.B5}=''U'' and {PRICE_BOOK_GAS.TARIFF_DESC}='''+trf+'''','P','','','');
end;

procedure TFRM_Tree.VacantPremiseProgrammedforDisconnection1Click(
  Sender: TObject);
Var
Agid,PremiseID,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 PREMISEID:=TreeData.D_Premise_id;
 FRM_Reports.PrintThisReport('CRM\Safety Cut Off.rpt','Safety Cut Off','{SITE_ADDRESS.PREMISE_ID}='+premiseid+' and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.raceExecutors2Click(Sender: TObject);
Var
Custid,Ahid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 custid:=TreeData.D_customer_id;
 ahid:=TreeData.D_Account_Holder_Id;
 // Check if Deceased First
 with main_data_module.generalquery do
 Begin
  close;
  deletevariables;
  declarevariable('AHID',otlong);
  sql.clear;
  sql.add('select account_holder_status_id from crm.account_holders where account_holder_id=:AHID');
  setvariable('AHID',ahid);
  open;
  deletevariables;
 End;
 if main_data_module.generalquery.Fields[0].text<>'3' then
 Begin
  Messagedlg('Letter can only be issued for Deceased Account Holders.',mtinformation,[MBOK],0);
  exit;
 End;
 FRM_Reports.PrintThisReport('CRM\Trace_Executors.rpt','Trace Executors','isnull({AGREEMENT_PREMISES.DATE_MOVED_OUT})=true and isnull({SERVICE.SERVICE_ID})=false and {ACCOUNT_HOLDERS.ACCOUNT_HOLDER_ID}= '+ahid,'','PRINTER',custid,'');
end;

procedure TFRM_Tree.EnquiryRecevied1Click(Sender: TObject);
Var
Custid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Custid:=TreeData.D_Customer_ID;
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\Enquiry_Received.rpt','Enquiry Recevied','{ACCOUNT_HOLDERS.CONTACT_ORDER}=1 and {CUSTOMER.CUSTOMER_ID}='+custid+'','','PRINTER',custid,'');
end;

procedure TFRM_Tree.raceExecutorsFollowUp1Click(Sender: TObject);
Var
Custid,Ahid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 custid:=TreeData.D_customer_id;
 ahid:=TreeData.D_Account_Holder_Id;
 // Check if Deceased First
 with main_data_module.generalquery do
 Begin
  close;
  deletevariables;
  declarevariable('AHID',otlong);
  sql.clear;
  sql.add('select account_holder_status_id from crm.account_holders where account_holder_id=:AHID');
  setvariable('AHID',ahid);
  open;
  deletevariables;
 End;
 if main_data_module.generalquery.Fields[0].text<>'3' then
 Begin
  Messagedlg('Letter can only be issued for Deceased Account Holders.',mtinformation,[MBOK],0);
  exit;
 End;
 FRM_Reports.PrintThisReport('CRM\Trace_Executors_Follow_Up.rpt','Trace Executors Follow Up','isnull({AGREEMENT_PREMISES.DATE_MOVED_OUT})=true and isnull({SERVICE.SERVICE_ID})=false and {ACCOUNT_HOLDERS.ACCOUNT_HOLDER_ID}= '+ahid,'','PRINTER',custid,'');
end;

procedure TFRM_Tree.SignupLetterVerbal1Click(Sender: TObject);
Var
AgId,CID:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AgID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Signup_Letter_Verbal.rpt','Signup Letter','{ACCOUNT_HOLDERS.CONTACT_ORDER} = 1 and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\SignUp_Application_Form_Verbal.rpt','Application Form','{ACCOUNT_HOLDERS.CONTACT_ORDER} = 1 and {AGREEMENTS.AGREEMENT_ID} = '+agid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.ShowAnnualUsagekWh2Click(Sender: TObject);
Var
X:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 X:=frm_common.GetKWH(TreeData.D_SPAN);
 if x='' then
 Begin
  Messagedlg('No Eac exists',mtinformation,[MBOK],0);
  exit;
 End;
 Messagedlg('Annual kWh usage = '+x,mtinformation,[MBOK],0);
 exit;
end;

procedure TFRM_Tree.DoNotBillDisconnectedDeEnergised1Click(
  Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 setasDNB(TreeData.D_REGid);
end;

procedure TFRM_Tree.DoNotBillDisconnectedDeEnergised2Click(
  Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
setasDNB(TreeData.D_REGid);
end;

procedure TFRM_Tree.CheckCommsClick(Sender: TObject);
begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  TFRM_CHECK_COMMS.StartModal(Self, StrToInt64(Treedata.d_customer_id), StrToInt64(TreeData.D_agreement_id), StrToInt64(TreeData.D_Premise_Id), PremAddr)
end;

procedure TFRM_Tree.CheckCVStatus1Click(Sender: TObject);
var
 CustomerID:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;
 CustomerID:=TreeData.D_Customer_ID;
 Application.CreateForm(TFRM_CV_ERRORS, FRM_CV_ERRORS);
 try
  FRM_CV_ERRORS.GetData(CustomerID);
  if FRM_CV_ERRORS.AFFECTED.recordcount<>0 then FRM_CV_ERRORS.showmodal
  else
  begin
   messagedlg('There are no Qualifying Calorific GAS refunds for this customer',mtconfirmation,[mbok],0);
  end;
 finally
  FRM_CV_ERRORS.release;
 end;

end;

{------------------------------------------------------------------------------}
function TFRM_Tree.CheckHotNoteOpen(aNID: string): boolean;
begin
  Result := gHotNoteList.CheckHotNote(aNID) > 0;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.ChequeReceived1Click(Sender: TObject);
var
  nodeData    : PMyRec;
  agreementId : Int64;
  oldCursor   : TCursor;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := Treeview1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  agreementId := StrToInt64(nodeData.D_Agreement_ID);

  oldCursor     := Screen.Cursor;
  Screen.Cursor := crHourGlass;
  try
    TFrm_Financial_History.Start(Self, ctChequeReceived, agreementId);

    // Refresh Agreement Tree
    Frm_Main.SearchForCust(IntToStr(GetCustomerIdFromAgreementId(agreementId)));
  finally
    Screen.Cursor := oldCursor;
  end;
end;

procedure TFRM_Tree.MenuItem7Click(Sender: TObject);
Var
Agid,reason:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;
 REASON:=TreeData.D_REASON;

 If Messagedlg('Are you sure you wish to REMOVE this account dispute?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;

 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from crm.agreements_in_dispute where agreement_id='+agid);
  sql.add('and dispute_reason_code='''+reason+'''');
  execute;
 end;
 frm_login.mainsession.commit;

   treeview1.expanded[xnode.parent]:=false;
   treeview1.expanded[xnode.parent]:=true;
end;

procedure TFRM_Tree.Gas1Click(Sender: TObject);
begin
 AddDispute('2','Gas Read in Dispute');
end;

procedure TFRM_Tree.Elec1Click(Sender: TObject);
begin
 AddDispute('1','Electric Read in Dispute');
end;

procedure TFRM_Tree.Dual1Click(Sender: TObject);
begin
 AddDispute('4','Gas & Electric Reads in Dispute');
end;

procedure TFRM_Tree.Surcharge1Click(Sender: TObject);
begin
 AddDispute('3','Surcharges');
end;

procedure TFRM_Tree.ET1Click(Sender: TObject);
begin
 AddDispute('5','Erroneous Transfer');
end;

procedure TFRM_Tree.AddDispute(Rcode,Rdesc:string);
Var
 AGID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
{ with main_data_module.generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select * from crm.agreements_in_dispute where agreement_id='+agid);
  open;
 End; }
 if Messagedlg('Are you sure you wish to add Account Dispute, reason '+rdesc,mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('insert into crm.agreements_in_dispute values('+agid+','''+rcode+''',''Y'','''+UPPERCASE(USERID)+''')');
  try
   execute;
  except
   Begin
    Messagedlg('Account already has a Dispute of this type recorded against it.',mtinformation,[MBOK],0);
    exit;
   End;
  end;
 End;
 frm_login.mainsession.commit;

 treeview1.Expanded[xnode]:=false;
 treeview1.Expanded[xnode]:=true;

End;

procedure TFRM_Tree.Correspondance1Click(Sender: TObject);
begin
 AddDispute('6','Correspondance');
end;

procedure TFRM_Tree.CosGainStatus(SPAN : String);
begin
   //See if there are any COs Record
  with main_data_module.tempquery do
  Begin
   close;
   deletevariables;
   declarevariable('SPAN',otstring);
   sql.clear;
   sql.add('select * from LIBERTY100.VW_LIB100_COS_GAIN_STATUS');
   sql.add('where cfg_state is null and is_live_span is not null and');
   sql.add('servicepointno=:SPAN');
   setvariable('SPAN',SPAN);
   open;
   deletevariables;
  end;
end;

procedure TFRM_Tree.ShowMopTree(Span:string);
Var
oldfilename,filename,Supp_mpid,Supp_name,Supp_ssd,Effective_from,Contract_ref,Service_ref,Service_level_ref,gsp_group:string;
contracterror:boolean;
begin
 MopTree.beginupdate;
 TreeUpdating:=true;
 moptree.clear;
 TabMop.tabvisible:=false;

 with MOPD0155 do
 begin
  close;
  setvariable('MPAN',SPAN);
  open;
 end;
 if MopD0155.Recordcount=0 then
 Begin
  moptree.endupdate;
  Treeupdating:=false;
  exit;
 end;

 TabMOP.tabvisible:=true;
 with MopD0155 do
 Begin
  oldFilename:='Lee';
  oldspan:='Lee';
  while not eof do
  Begin
   SPAN:=fields[0].text;
   FILENAME:=fields[19].text;
  // if (SPAN<>oldspan) or (Filename<>oldfilename) then
   Begin
    SUPP_NAME:=fields[27].text;
    SUPP_SSD:=fields[1].text;
    EFFECTIVE_FROM:=fields[16].text;

    SUPP_MPID:=fields[22].text;
    CONTRACT_REF:=fields[12].text;
    SERVICE_REF:=fields[17].text;
    SERVICE_LEVEL_REF:=fields[18].text;
    GSP_GROUP:=fields[21].text;

    if (Fields[24].text<>'') and (fields[24].text<>'D0011') then ContractError:=True
    else ContractError:=false;
    Begin

     MyNodeMopSpan:=MopTree.Addchild(nil);
     nodeData := MopTree.GetNodeData(MyNodeMopSpan);
     NodeData.caption := 'No Records Found';

     //MyNodeMopSpan:=Moptree.items.Add(nil, 'No Records Found');
     // Set Span Type
     if fields[24].text='' then NodeData.index:=4
     else if fields[24].text='D0011' then NodeData.index:=5
     else NodeData.index:=4;


     //MyNodeMopSpan.selectedindex:=MyNodeMopSpan.imageindex;
     NodeData.caption:='MPAN: '+SPAN+' - Meter Operator Service Account: Supplier '+Supp_mpid+' ('+Supp_name+')'+#10;
     if contracterror=false then
     Begin
      nodedata.caption:=nodedata.caption+'MO Appointment Effective From: '+Effective_From;
      if Fields[23].text<>'' then
      Begin
       nodedata.caption:=nodedata.caption+' -  '+Fields[23].text+' ('+Fields[26].text+'-'+fields[35].text+')';
       nodedata.fontcolor:=clmaroon;
       nodedata.fontBold:=true;
       nodedata.index:=4;
      end
      else
      nodedata.fontcolor:=clblack;
     end
     else
     Begin
      nodedata.caption:=nodedata.caption+' ** Invalid D0155 - '+fields[36].text+' **';
      nodedata.fontcolor:=clred;
     End;
     //mynodeMopSpan.selectedindex:=MyNodeMopSpan.imageindex;
     // MyNodeMopSite:=Moptree.items.Addchild(MyNodeMopSpan, 'Premises TEMP - ');

     nodedata.D_SPAN := SPAN;
     nodedata.D_FILENAME := FILENAME;
     nodedata.D_ETDMOA := fields[23].text;
     nodedata.D_ContractError := ContractError;

      MyNodeMopSite:=MopTree.Addchild(MyNodeMopSpan);
      nodeData := MopTree.GetNodeData(MyNodeMopSite);
      NodeData.caption := 'Premises TEMP - ';


         //MyNodeMopSpan.data:=MyRecPtr;
    end;
   end;
   oldspan:=span;
   oldfilename:=filename;
   next;
  End;
 end;
 moptree.endupdate;
 Treeupdating:=false;
end;

procedure TFRM_Tree.ShowMyUtilitaData1Click(Sender: TObject);
var
  c_type : Integer;
begin
  Xnode    := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(Xnode);

  FMyUtilita := TFMyUtilita.Create(Application);
  Try
    FMyUtilita.CustomerID.Text := TreeData.D_Customer_ID;
     c_type:=FRM_Common.getCustomerTypeId(TreeData.D_Customer_ID);
    If (c_type = 2) Or (c_type = 4) Then
      If FMyUtilita.CommercialSearch Then
        FMyUtilita.ShowModal
      Else
        Application.MessageBox('No My Utilita Data', 'Warning', MB_OK)
    Else If FMyUtilita.DomesticSearch Then
      FMyUtilita.ShowModal
    Else
      Application.MessageBox('No My Utilita Data', 'Warning', MB_OK);
  Finally
    FMyUtilita.Free;
  End;

End;

procedure TFRM_Tree.MyUtilitaWebFeaturesClick(Sender: TObject);
var
  c_type : Integer;
begin
  Xnode    := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(Xnode);

   // form used to check for my utilita data (not used here)
  FMyUtilita := TFMyUtilita.Create(Application);

  try
    FMyUtilita.CustomerID.Text := TreeData.D_Customer_ID;
    c_type:=FRM_Common.getCustomerTypeId(TreeData.D_Customer_ID);

    if (c_type = 2) Or (c_type = 4) then
    begin
      If FMyUtilita.CommercialSearch then
      begin
        // show web crm
        try
          frm_common.ShowWebCRM(UserID, TreeData.D_Customer_ID);
        except
          Application.MessageBox('Unable to connect to website', 'Error', MB_OK);
        end;
      end
      else
      begin
        Application.MessageBox('No My Utilita Data', 'Warning', MB_OK)
      end;
    end
    else
    begin
      if FMyUtilita.DomesticSearch then
      begin
        // show web crm
        try
          frm_common.ShowWebCRM(UserID, TreeData.D_Customer_ID);
        except
          Application.MessageBox('Unable to connect to website', 'Error', MB_OK);
        end;
      end
      else
      begin
        Application.MessageBox('No My Utilita Data', 'Warning', MB_OK);
      end;
    end;
  finally
    FMyUtilita.Free;
  end;
end;

Procedure TFRM_Tree.RefreshMopPremiseNode(xnode:pvirtualnode);
Var
sr,slr,gsp,efd,Span,SSD,oldfilename,filename,cust_name,custaddr,cref,scn:String;
ContractError:Boolean;
Begin

 mynodemopsite:=xnode;
 TreeData:= moptree.GetNodeData(mynodemopsite);
 SPAN:=TreeData.D_SPAN;
 SSD:=TreeData.D_SSD;
 scn:=TreeData.D_SUPPLIERNAME;
 cref:=TreeData.D_ContractRef;
 ContractError:=TreeData.D_ContractError;
 sr:=TreeData.D_SERVICEREF;
 slr:=TreeData.D_SERVICELEVELREF;
 gsp:=TreeData.D_GSP;
 efd:=TreeData.D_EFFECTIVEFROM;
 moptree.deletechildren(mynodemopsite);
 //moptree.selected.deletechildren;
 //mynodemopsite:=moptree.selected;

 // Add Agreement Details
 // Get Contract Details based on Flow Info Lookup


 MyNodeMopAgreement:=MopTree.Addchild(mynodemopsite);
 nodeData := Treeview1.GetNodeData(MyNodeMopAgreement);
 NodeData.caption := 'Super Customer Agreement: ** '+scn+' **';
 nodedata.index:=71;

// MyNodeMopAgreement:=Moptree.items.Addchild(MopTree.selected, 'Super Customer Agreement: ** '+(TreeData.D_SUPPLIERNAME)+' **');
 //MyNodeMopAgreement.imageindex:=71;
 if contracterror=true then
 Begin
  nodedata.Fontcolor:=clred;
  nodedata.index:=88;
  nodedata.caption:=nodedata.caption+'** INVALID CONTRACT DETAILS IN D0155 **';
 End;
 //MyNodeMopAgreement.selectedindex:=MyNodeMopAgreement.imageindex;

 {mynode1:=MopTree.items.AddChild(MyNodeMopAgreement,'Testing');
 mynode1.imageindex:=79;
 mynode1.selectedindex:=79;
 mynode1.font.color:=clpurple; }

 mynode1:=MopTree.Addchild(MyNodeMopAgreement);
 nodeData := Treeview1.GetNodeData(mynode1);
 NodeData.caption := 'Testing';
 nodedata.index:=79;
 nodedata.fontcolor:=clpurple;

 nodedata.caption:='Contract Ref: '+cref;
 nodedata.caption:=nodedata.caption+', Service Ref: '+sr;
 nodedata.caption:=nodedata.caption+', Service Level Ref: '+slr;
 nodedata.caption:=nodedata.caption+', GSP Group: '+gsp;
 nodedata.caption:=nodedata.caption+', Effective From: '+efd;
 nodedata.fontcolor:=clblack;
 if contracterror=true then
 Begin
  nodedata.fontcolor:=clred;
 End;

 // Get Customer Details
 // D0302

 // Show Latest Customer Details Held in TreeData.base on or Before SSD
 with MOP302 do
 begin
  close;
  setvariable('MPAN',SPAN);
  setvariable('SSD',SSD);
  open;
 end;

 if mop302.recordcount<>0 then
 Begin
  with Mop302 do
  Begin
   oldFilename:='Lee';
 //  while not eof do
   Begin
    Filename:=fields[2].text;
    if Filename<>oldfilename then
    Begin
     Cust_NAME:=fields[4].text;
     Begin
      // Display Tree
      //MyNodeMopCustomer:=Moptree.items.Addchild(MyNodeMopSite, 'Customer: '+Cust_Name);

       MyNodeMopCustomer:=MopTree.Addchild(MyNodeMopSite);
       nodeData := moptree.GetNodeData(MyNodeMopCustomer);
       NodeData.caption := 'Customer: '+Cust_Name;


      if ssd<>Mop302.fields[1].text then
      Begin
       nodedata.caption:=nodedata.caption+' ** Historic **';
       nodedata.Fontcolor:=clblue;
      End
      else nodedata.fontColor:=clblack;
      // Add Site Node
      CustAddr:='';
      if fields[14].text<>'' then CustAddr:=CustAddr+fields[14].text+', ';
      if fields[15].text<>'' then CustAddr:=CustAddr+fields[15].text+', ';
      if fields[16].text<>'' then CustAddr:=CustAddr+fields[16].text+', ';
      if fields[17].text<>'' then CustAddr:=CustAddr+fields[17].text+', ';
      if fields[18].text<>'' then CustAddr:=CustAddr+fields[18].text+', ';
      if fields[19].text<>'' then CustAddr:=CustAddr+fields[19].text+', ';
      if fields[20].text<>'' then CustAddr:=CustAddr+fields[20].text+', ';
      if fields[21].text<>'' then CustAddr:=CustAddr+fields[21].text+', ';
      if fields[22].text<>'' then CustAddr:=CustAddr+fields[22].text+', ';
      CustAddr:=CustAddr+fields[23].text+'';
      nodedata.caption:=nodedata.caption+#10+CustAddr;
      // SJ-BSL - 02/05/2021 - Replacing constant assignment by Global Variable.
      NodeData.Index := FCustIcon; // 84;
      //MyNodeMopCustomer.selectedindex:=MyNodeMopCustomer.imageindex;
      // Add Other Details
      // Additional Information
      if fields[5].text<>'' then
      Begin
       {mynode1:=MopTree.items.AddChild(MyNodeMopCustomer,'Additional Information - '+fields[5].text);
       mynode1.imageindex:=39; // Special Needs
       mynode1.selectedindex:=39;
       mynode1.font.color:=clpurple;
       mynode1.font.style:=[fsbold];}

       mynode1:=MopTree.Addchild(MyNodeMopCustomer);
       nodeData := moptree.GetNodeData(mynode1);
       NodeData.caption := 'Additional Information - '+fields[5].text;
       NodeData.index:=39;
       NodeData.fontcolor:=clpurple;
       NodeData.fontBold:=true;

      end;
      // Special Access
      if fields[8].text<>'' then
      Begin
       {mynode1:=MopTree.items.AddChild(MyNodeMopCustomer,'Special Access - '+fields[8].text);
       mynode1.imageindex:=8; // Special Access
       mynode1.selectedindex:=8;
       mynode1.font.color:=clpurple;
       mynode1.font.style:=[fsbold]; }

       mynode1:=MopTree.Addchild(MyNodeMopCustomer);
       nodeData := moptree.GetNodeData(mynode1);
       NodeData.caption := 'Special Access - '+fields[8].text;
       NodeData.index:=8;
       NodeData.fontcolor:=clpurple;
       NodeData.fontBold:=true;

      end;
      // Max Power Req
      if fields[12].text<>'' then
      Begin
       {mynode1:=MopTree.items.AddChild(MyNodeMopCustomer,'Maximum Power Requirement - '+fields[12].text);
       mynode1.imageindex:=17;
       mynode1.selectedindex:=17;
       mynode1.font.color:=clpurple;
       mynode1.font.style:=[fsbold]; }

       mynode1:=MopTree.Addchild(MyNodeMopCustomer);
       nodeData := moptree.GetNodeData(mynode1);
       NodeData.caption := 'Maximum Power Requirement - '+fields[12].text;
       NodeData.index:=17;
       NodeData.fontcolor:=clpurple;
       NodeData.fontBold:=true;
      end;
      // Password
      if fields[6].text<>'' then
      Begin
       {mynode1:=MopTree.items.AddChild(MyNodeMopCustomer,'Password-      '+fields[6].text+'      Effective From ('+Fields[7].text+')');
       mynode1.imageindex:=7;
       mynode1.selectedindex:=7;
       mynode1.font.color:=clpurple;
       mynode1.font.style:=[fsbold]; }

        mynode1:=MopTree.Addchild(MyNodeMopCustomer);
       nodeData := moptree.GetNodeData(mynode1);
       NodeData.caption := 'Password-      '+fields[6].text+'      Effective From ('+Fields[7].text+')';
       NodeData.index:=7;
       NodeData.fontcolor:=clpurple;
       NodeData.fontBold:=true;

      end;
      // Contact Details
      {premiseContactNode:=MopTree.items.AddChild(MyNodeMopCustomer,'Site Contact Name - '+fields[9].text);
      premiseContactNode.imageindex:=34;
      premiseContactNode.selectedindex:=34;  }

      premiseContactNode:=MopTree.Addchild(MyNodeMopCustomer);
       nodeData := moptree.GetNodeData(premiseContactNode);
       NodeData.caption := 'Site Contact Name - '+fields[9].text;
       NodeData.index:=34;


      if fields[10].text<>'' then
      Begin
       {premiseContactitemNode:=MopTree.items.AddChild(premisecontactnode,'Tel - '+Fields[10].text);
       premiseContactitemNode.imageindex:=14;
       premiseContactitemNode.selectedindex:=14; }

       premiseContactNode:=MopTree.Addchild(MyNodeMopCustomer);
       nodeData := moptree.GetNodeData(premiseContactNode);
       NodeData.caption := 'Tel - '+Fields[10].text;
       NodeData.index:=14;

      end;
      if Fields[11].text<>'' then
      Begin
      { premiseContactitemNode:=MopTree.items.AddChild(premisecontactnode,'Fax - '+Fields[11].text);
       premiseContactitemNode.imageindex:=11;
       premiseContactitemNode.selectedindex:=11; }

       premiseContactNode:=MopTree.Addchild(MyNodeMopCustomer);
       nodeData := moptree.GetNodeData(premiseContactNode);
       NodeData.caption := 'Fax - '+Fields[11].text;
       NodeData.index:=11;

      end;
      end;
    end;
    oldFilename:=Filename;
  //  next;
   End;
  end;
 end
 else
 begin
  {MyNodeMopCustomer:=Moptree.items.Addchild(MyNodeMopSite, 'Customer: Details not Known. No D0302s Received.');
  MyNodeMopCustomer.Font.color:=clblue;
  MyNodeMopCustomer.imageindex:=84;
  MyNodeMopCustomer.selectedindex:=MyNodeMopCustomer.imageindex; }

  MyNodeMopCustomer:=MopTree.Addchild(MyNodeMopSite);
  nodeData := moptree.GetNodeData(MyNodeMopCustomer);
  NodeData.caption := 'Customer: Details not Known. No D0302s Received.';
  // SJ-BSL - 02/05/2021 - Replacing constant assignment by Global Variable.
  NodeData.Index := FCustIcon; // 84;

 End;
End;

procedure TFRM_Tree.BuildElectricMeterNodeMOP(MPAN,ENDDATE:string);
var
TPR,metertype,enstatus,ssc,sscdesc,daterem,nsr,maketype:string;
//MyRecPtr: PMyRec;
Begin
   // Check for Meter Technical Details
 {myNodeMopSpan:=MopTree.selected;
 mpan:=nodedata.D_SPAN;
 ENDDATE:=nodedata.D_SPANEND;
 }

 if ENDDATE='' then ENDDATE:='10/10/2060';


  with MTDSMOP do
  begin
   close;
   setvariable('ENDDATE',ENDDATE);
   setvariable('MPAN',MPAN);
   open;
  end;

  // Only Do This Block If Meter Records Exist
  mtdsmop.First;
  if MTDSMOP.recordcount<>0 then
  Begin
   msid:='LEEOK';
   oldefsdmsmtd:='lee';
   oldmeterid:='';
   oldregister:='';
   while not MTDSMOP.eof do
   Begin
    // Build Tree Of MTDSMOP
   // Create Subtree of Effective From Dates
    efsdmsmtd:=MTDSMOP.fields[1].text;
    MeterType:=MTDSMOP.fields[23].text;
    EnStatus:=MTDSMOP.fields[2].text;
    SSC:=MTDSMOP.fields[5].text;
    SSCDesc:=MTDSMOP.fields[6].text;
    DateRem:=MTDSMOP.fields[37].text;
    NSR:=MTDSMOP.fields[36].text;
    Mregister:=MTDSMOP.fields[26].text;
    meterid:=MTDSMOP.fields[10].text;
    maketype:='';
    if mtdsmop.fields[14].text<>'' then maketype:=copy(MTDsmop.fields[14].text,1,3);

    if efsdmsmtd<>oldefsdmsmtd then
    Begin
     if oldefsdmsmtd='lee' then config:='Current Configuration'
     else config:='Previous Configuration';

     // Current Configuration
     if ((config='Previous Configuration') and (historic1.checked=true)) or (config<>'Previous Configuration') then
     Begin
      if (efsdmsmtd<>'') and (MeterType='') and (SSC='') then
      Begin
      { MeterConfigNode:=MopTree.items.AddChild(myNodeMopSpan,config+' - '+ENSTATUS+' - '+efsdmsmtd+' - MOP Reports No Meters on this Supply. Energisation Status = '+EnStatus);
       MeterConfigNode.font.color:=clred;
       MeterConfigNode.font.style:=[fsbold];
       MeterConfigNode.imageindex:=24;
       MeterConfigNode.selectedindex:=24;}

       MeterConfigNode:=MopTree.Addchild(myNodeMopSpan);
       nodeData := moptree.GetNodeData(MeterConfigNode);
       NodeData.caption := config+' - '+ENSTATUS+' - '+efsdmsmtd+' - MOP Reports No Meters on this Supply. Energisation Status = '+EnStatus;
       NodeData.fontcolor:=clred;
       NodeData.fontBold:=true;
       NodeData.index:=24;

      end

      else
      if MeterType='' then
      Begin
       {MeterConfigNode:=MopTree.items.AddChild(myNodeMopSpan,'Metering Configuration not Known (Missing / Incomplete meter technical details)');
       MeterConfigNode.font.color:=clred;
       MeterConfigNode.font.style:=[fsbold];
       MeterConfigNode.imageindex:=26;
       MeterConfigNode.selectedindex:=26;  }

        MeterConfigNode:=MopTree.Addchild(myNodeMopSpan);
       nodeData := moptree.GetNodeData(MeterConfigNode);
       NodeData.caption := 'Metering Configuration not Known (Missing / Incomplete meter technical details';
       NodeData.fontcolor:=clred;
       NodeData.fontBold:=true;
       NodeData.index:=26;

      end;

      if (SSC<>'') then
      Begin
       //MeterConfigNode:=MopTree.items.AddChild(myNodeMopSpan,config+' - '+ENSTATUS+' - '+efsdmsmtd+' - SSC ID ('+SSC+') - '+SSCDESC);

       MeterConfigNode:=MopTree.Addchild(myNodeMopSpan);
       nodeData := moptree.GetNodeData(MeterConfigNode);
       NodeData.caption := config+' - '+ENSTATUS+' - '+efsdmsmtd+' - SSC ID ('+SSC+') - '+SSCDESC;


       if config<>'Previous Configuration' then
       Begin
        NodeData.fontcolor:=clgreen;
        NodeData.fontBold:=true;
       end;
       NodeData.index:=27;
       //MeterConfigNode.selectedindex:=27;
       if ENSTATUS='D' then NodeData.fontcolor:=clpurple;
       oldmeterid:='lee';
      end
     end;
    end; // End Of Configuration Date

    // Do Meters
    if ((config='Previous Configuration') and (historic1.checked=true)) or (config<>'Previous Configuration') then
    Begin
     if meterid<>oldmeterid then
     Begin
      oldregister:='Lee';
      mtype:='';
      if MeterType='' then mtype:='*NO*';
      if (efsdmsmtd<>'') and (metertype='') then mtype:='*NO*';

      if mtype<>'*NO*' then
      Begin
       {if DateRem<>'' then Dateremoved:='    (Date Removed='+DateRem+')'
       else }
       dateremoved:='';
       if (NSR='Y') and (v_non.checked=false) then
       Begin
       //
       end
       else
       Begin
       // MeterNode:=MopTree.items.AddChild(MeterConfigNode,'NHH Meter ID-'+Meterid+dateremoved);

        MeterNode:=MopTree.Addchild(MeterConfigNode);
        nodeData := moptree.GetNodeData(MeterNode);
        NodeData.caption := 'NHH Meter ID-'+Meterid+dateremoved;
        nodedata.index:=17; // NHH Credit Meter
        nodedata.D_SPAN :=mpan;
        nodedata.M_METERID :=Meterid;
        nodedata.M_SERVICE :='0';

        //MeterNode.selectedindex:=17;
        if MeterType='N' then
        Begin
         nodedata.caption:='NHH Credit Meter ID-'+MeterID+dateremoved;
        end;
        if MeterType='S' then
        Begin
         nodedata.caption:='NHH Smart Card Meter ID-'+MeterID+dateremoved;
         nodedata.index:=22; // NHH Smart Card meter
        // MeterNode.selectedindex:=22;
        end;

         if (MeterType='S') and (mtdsmop.fields[20].text='R') then
        Begin
         nodedata.caption:='Remote Read Smart Meter ID-'+MeterID+dateremoved;
         nodedata.index:=205; // NHH Smart Card meter
         //MeterNode.selectedindex:=205;
        end;

        if (copy(MeterType,1,4)='RCAM') or (MakeType='PRI') or (copy(MeterType,1,3)='NSS') then
        Begin
         nodedata.caption:='Smart Meter ID-'+MeterID+dateremoved;
         nodedata.index:=314;
        end;

        if (copy(MeterType,1,2)='S1') then
        Begin
         ShowSmetsMeterCOmmsMop(MeterNode,MPAN,METERID,'0',dateremoved,'M');
        end;

        if (copy(MeterType,1,2)='S2')  then
        Begin
         nodedata.caption:='SMETS 2 Meter ID-'+MeterID+dateremoved;
         nodedata.index:=205; // SMETS 2 ICON
        end;

        if MeterType='T' then
        Begin
         nodedata.caption:='NHH Token Meter ID-'+MeterID+dateremoved;
         nodedata.index:=23; // NHH token Meter
         //MeterNode.selectedindex:=23;
        end;
        if MeterType='K' then
        Begin
         nodedata.caption:='NHH Key Meter ID-'+MeterID+dateremoved;
         nodedata.index:=21; // NHH key Meter
         //MeterNode.selectedindex:=21;
        end;
        if MeterType='H' then
        Begin
         nodedata.caption:='HH Meter ID-'+MeterID+dateremoved;
         nodedata.index:=9; // HH Meter
        // MeterNode.selectedindex:=9;
        end;
        if enstatus[1]='D' then       // De-Energised
        Begin
         nodedata.index:=195;
        // MeterNode.selectedindex:=195;
        End;
       end;
      end
      else
      Begin
       //MeterNode:=MopTree.items.AddChild(MeterConfigNode,'NHH Meter ID-'+Meterid+' (Missing D0150)');

       MeterNode:=MopTree.Addchild(MeterConfigNode);
       nodeData := moptree.GetNodeData(MeterNode);
       NodeData.caption := 'NHH Meter ID-'+Meterid+' (Missing D0150)';
       nodedata.D_SPAN :=mpan;
      nodedata.M_METERID :=Meterid;
      nodedata.M_SERVICE :='0';
       nodedata.index:=26; // NHH Credit Meter
       //MeterNode.selectedindex:=26;
       nodedata.fontcolor:=clred;
      End;
     end; // Change Of Meter
    end;  // End Of Meter Configuration Block

    // Have any Meters been Removed?
    if (efsdmsmtd<>'') and (MeterType='') and (SSC='') and (daterem<>'')then
    Begin
     //MeterNode:=MopTree.items.AddChild(MeterConfigNode,'Removed Meter -'+MTDSMOP.fields[39].text+'. Date Removed ('+DateRem+')');
       MeterNode:=MopTree.Addchild(MeterConfigNode);
       nodeData := moptree.GetNodeData(MeterNode);
       NodeData.caption := 'Removed Meter -'+MTDSMOP.fields[39].text+'. Date Removed ('+DateRem+')';
       nodedata.D_SPAN :=mpan;
      nodedata.M_METERID :=Meterid;
      nodedata.M_SERVICE :='0';
     nodedata.index:=24; // Removed Meter
     //MeterNode.selectedindex:=24;
    end;

    if ((config='Previous Configuration') and (historic1.checked=true)) or (config<>'Previous Configuration') then
    Begin
     if mregister<>oldregister then
     Begin
      if (NSR='Y') and (v_non.checked=false) then
      begin
      //
      end
      else
      Begin
       TPR:=MTDSMOP.fields[33].text;
       if tpr<>'' then
       Begin
        MeterRegisterNode:=MopTree.Addchild(MeterNode);
        nodeData := moptree.GetNodeData(MeterRegisterNode);
        NodeData.caption := mregister+' - TPR '+TPR+' ('+MTDSMOP.fields[34].text+') - '+MTDSMOP.fields[30].text;
        //MeterRegisterNode:=MopTree.items.AddChild(MeterNode,mregister+' - TPR '+TPR+' ('+MTDSMOP.fields[34].text+') - '+MTDSMOP.fields[30].text);
        NodeData.fontcolor:=clblack;
        if MTDSMOP.fields[30].text='' then NodeData.fontcolor:=clred;
       end
       else
       Begin
        if MeterID<>'' then
        Begin
         MeterRegisterNode:=MopTree.Addchild(MeterNode);
         nodeData := moptree.GetNodeData(MeterRegisterNode);
         NodeData.caption := mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * ) - '+MTDSMOP.fields[30].text;

         //MeterRegisterNode:=MopTree.items.AddChild(MeterNode,mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * ) - '+MTDSMOP.fields[30].text);
         NodeData.fontcolor:=clred;
        end;
       end;
       NodeData.Index:=Frm_Common.GetRegisterPic(MTDSMOP.fields[33].text);
       //MeterRegisterNode.selectedindex:=MeterRegisterNode.imageindex;


       nodedata.D_SPAN :=mpan;
       nodedata.M_EFSDMSMTD :=efsdmsmtd;
       nodedata.M_METERID :=Meterid;
       nodedata.M_REGISTERID :=mregister;
       if MeterType='H' then nodedata.M_HH_REGISTER :='H'
       else nodedata.M_HH_REGISTER :='N';
       try
       //MeterRegisterNode.data:=MyRecPtr;
       except
       end;

       if MTDSMOP.fields[29].text='RI' then
       Begin
        NodeData.index:=44;
       end;
       if NSR='Y' then
       Begin
        // non settlement register
        NodeData.Fontcolor:=clred;
        if v_non.checked then
        Begin
        NodeData.Index:=Frm_Common.GetNonRegisterPic(MTDSMOP.fields[34].text);
        //MeterRegisterNode.selectedindex:=MeterRegisterNode.imageindex;
        if MTDSMOP.fields[29].text='RI' then
         Begin
          NodeData.index:=45;
         end;
        end;
       end;
      end;
     end; // End Of Add Register
    end;  // End Of Configuration Block
    oldefsdmsmtd:=efsdmsmtd;
    oldmeterid:=meterid;
    oldregister:=mregister;
    MTDSMOP.next;
   end;
  end; // End Of Meter Strucutre Tree

 ////////// Build Tree Of Orphaned Registers //////////////
 With generalquery do
  Begin
   Close;
   deletevariables;
   declarevariable('MPAN',otstring);
   sql.clear;
   sql.add('select distinct R.mpancore,R.meterid,R.registerid,r.current_status');
   sql.add('from mopmgr.readings R,mopmgr.d0149a D, mopmgr.d0150_293 M');
   sql.add('where');
   sql.add('r.MPANCORE=:MPAN');
   sql.add('and r.mpancore=d.mpancore (+)');
   sql.add('and');
   sql.add('r.meterid=d.meterid (+)');
   sql.add('and');
   sql.add('r.registerid=d.registerid (+)');
   sql.add('and');
   sql.add('r.mpancore=M.mpancore (+)');
   sql.add('and');
   sql.add('r.meterid=M.meterid (+)');
   sql.add('and');
   sql.add('r.registerid=M.meter_register_id (+)');
   sql.add('and d.mpancore is null');
   sql.add('and m.mpancore is null');
   sql.add('and r.current_status<>''D''');
   sql.add('order by R.meterid,R.registerid');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
  end;
   // Only Do This Block If Meter Records Exist
  if generalquery.recordcount<>0 then
  Begin
   oldmeterid:='OldMeter';
   msid:='LEEOK';
   oldefsdmsmtd:='lee';

   MeterConfigNode:=MopTree.Addchild(myNodeMopSpan);
   nodeData := moptree.GetNodeData(MeterConfigNode);
   NodeData.caption := 'Register Readings (Orphans - No Mapping Details';


   //MeterConfigNode:=MopTree.items.AddChild(myNodeMopSpan,'Register Readings (Orphans - No Mapping Details)');
   NodeData.fontcolor:=clred;
   NodeData.fontBold:=true;
   NodeData.index:=26;
   //MeterConfigNode.selectedindex:=26;

   while not generalquery.eof do
   Begin
     // Do Meters
    Begin
     meterid:=generalquery.fields[1].text;
     if meterid<>oldmeterid then
     Begin
      oldregister:='Lee';
      mtype:='';
       Begin
        MeterNode:=MopTree.Addchild(MeterConfigNode);
        nodeData := moptree.GetNodeData(MeterNode);
        NodeData.caption := 'NHH Meter ID-'+Meterid+dateremoved;
        nodedata.D_SPAN :=mpan;
      nodedata.M_METERID :=Meterid;
      nodedata.M_SERVICE :='0';
        //MeterNode:=MopTree.items.AddChild(MeterConfigNode,'NHH Meter ID-'+Meterid+dateremoved);
        nodedata.index:=26; // NHH Meter
        //MeterNode.selectedindex:=26;
       end;
      end; // Change Of Meter
    end;  // End Of Configuration Block
    Begin
     Mregister:=generalquery.fields[2].text;
     if mregister<>oldregister then
     Begin
      Begin
       Begin
        MeterRegisterNode:=MopTree.Addchild(MeterNode);
        nodeData := moptree.GetNodeData(MeterRegisterNode);
        NodeData.caption := mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * )';

        //MeterRegisterNode:=MopTree.items.AddChild(MeterNode,mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * )');
        NodeData.fontcolor:=clred;
       end;
       NodeData.index:=29;

       nodedata.D_SPAN :=mpan;
       nodedata.M_EFSDMSMTD :=efsdmsmtd;
       nodedata.M_METERID :=MeterID;
       nodedata.M_REGISTERID :=mregister;
       //MeterRegisterNode.data:=MyRecPtr;

      end;
     end; // End Of Add Register
    end;  // End Of Configuration Block
    oldmeterid:=meterid;
    oldregister:=mregister;
    generalquery.next;
   end;
  end; // End Of Meter Strucutre Tree
end;

procedure TFRM_Tree.ViewDataflows1Click(Sender: TObject);
begin
 if treeupdating=true then exit;
 mpannode:=moptree.FocusedNode;
 TreeData:= Moptree.GetNodeData(mpannode);

 //mpannode:=MopTree.selected;
 mpan:=TreeData.D_SPAN;
 if not Assigned(FRM_DFLOW_HISTORY_MOP) then Application.CreateForm(TFRM_DFLOW_HISTORY_MOP, FRM_DFLOW_HISTORY_MOP);
 FRM_DFLOW_HISTORY_MOP.DflowQuery(MPAN,'');
 //FRM_DFLOW_HISTORY_MOP.Show_Mpan_Status(MPAN,'');
 if FRM_DFLOW_HISTORY_MOP.caption='' then
 Begin
  messagedlg('There is no Dataflow History for this MPAN',MTinformation,[MBOK],0);
  exit;
 end;
 FRM_DFLOW_HISTORY_MOP.show;
end;

procedure TFRM_Tree.BuildMOPNOTES(MPAN:string);
Var
//MyRecPtr: PMyRec;
 x:integer;
 num:string;
 maxgroup,noteid:integer;
Begin
 // Do Notes
 maxgroup:=10;
 // Get Notes for Customer
 x:=1;
 With GeneralQuery do
 Begin
  Close;
  deletevariables;
  declarevariable('MPAN',otstring);
  sql.clear;
  sql.add('SELECT R.description, E.comments_1,e.date_raised,e.due_date,E.mpancore,E.site_id,e.record_id');
  sql.add('FROM enquiry.enquiries E,enquiry.request_type R');
  sql.add('where E.MPANCORE=:MPAN');
  sql.add('and r.id=e.request_type');
  sql.add('and r.enqiury_or_note=''N''');
  sql.add('and request_type<>''307''');
  sql.add('and e.system_role=''M''');
  sql.add('order by 3 desc');
  setvariable('MPAN',MPAN);
  open;
  deletevariables;
  if generalquery.recordcount<>0 then
  Begin
  //  add first note
   Firstline:=frm_common.getFirstLine(generalquery.fields[1].text);
   desc:='(N0001) - '+generalquery.fields[2].text+' - '+generalquery.fields[0].text;
   if v_fullnotes.checked then desc:=desc+#10+generalquery.fields[1].text
   else desc:=desc+' '+firstline;

   NoteNode:=MopTree.Addchild(mynodeMopSpan);
   nodeData := moptree.GetNodeData(NoteNode);
   NodeData.caption := desc;

   //NoteNode:=MopTree.items.AddChild(mynodeMopSpan,desc);
   noteid:=53;
   if generalquery.fields[5].text<>'' then noteid:=52;
   if generalquery.fields[4].text<>'' then noteid:=51;
   NodeData.index:=noteid;


   nodedata.C_record_id :=generalquery.fields[6].text;

   //NoteNode.data:=MyRecPtr;

   NodeData.fontcolor:=clblack;
   generalquery.next;
   while not generalquery.eof do   // add sub notes
   Begin
    inc(x);
    if x<10000 then num:=inttostr(x);
    if x<1000 then num:='0'+inttostr(x);
    if x<100 then num:='00'+inttostr(x);
    if x<10 then num:='000'+inttostr(x);
    // Desc:=getcomments(fields[1].text);
    Firstline:=frm_common.getFirstLine(generalquery.fields[1].text);
    desc:='(N'+num+') - '+generalquery.fields[2].text+' - '+generalquery.fields[0].text;
    if v_fullnotes.checked then desc:=desc+#10+generalquery.fields[1].text
    else desc:=desc+' '+firstline;


    NotesubNode:=MopTree.Addchild(NoteNode);
    nodeData := moptree.GetNodeData(NotesubNode);
    NodeData.caption := desc;
    //NotesubNode:=MOPTREE.items.AddChild(NoteNode,desc);

    noteid:=53;
    if generalquery.fields[5].text<>'' then noteid:=52;
    if generalquery.fields[4].text<>'' then noteid:=51;
    NodeData.index:=noteid;


    nodedata.C_record_id :=generalquery.fields[6].text;
    //NotesubNode.data:=MyRecPtr;

    NodeData.fontcolor:=clblack;
    generalquery.next;
   end;
  end;
 end;

  // Get Outstanding Eqnuiries for Customer
 x:=0;
 With GeneralQuery do
 Begin
  Close;
  deletevariables;
  declarevariable('MPAN',otstring);
  sql.clear;
  sql.add('SELECT R.description, E.comments_1,e.date_raised,e.due_date,');
  sql.add('e.mpancore,e.site_id,e.record_id,e.owner,e.raised_by,r.enqiury_or_note FROM enquiry.enquiries E,enquiry.request_type R');
  sql.add('where E.MPANCORE=:MPAN');
  sql.add('and r.id=e.request_type');
  sql.add('and (r.enqiury_or_note=''E''');
  //sql.add('or r.enqiury_or_note=''S''');
  sql.add('or r.enqiury_or_note=''D'')');
  sql.add('and e.resolved=''N''');
  sql.add('and e.system_role=''M''');
  sql.add('order by E.date_raised desc');
  setvariable('MPAN',mpan);
  open;
  deletevariables;
  if generalquery.recordcount<>0 then
  Begin

   if generalquery.recordcount>maxgroup then
   Begin
    desc:='(Enquiries Outstanding for MPAN = '+inttostr(generalquery.recordcount)+')';

    NoteTopNode:=MopTree.Addchild(mynodeMopSpan);
    nodeData := moptree.GetNodeData(NoteTopNode);
    NodeData.caption := desc;

    //NoteTopNode:=MOPTREE.items.AddChild(mynodeMopSpan,desc);
    NodeData.index:=63;
   end;

   while not generalquery.eof do      // add sub notes
   Begin
    inc(x);
    if x<10000 then num:=inttostr(x);
    if x<1000 then num:='0'+inttostr(x);
    if x<100 then num:='00'+inttostr(x);
    if x<10 then num:='000'+inttostr(x);
    // Desc:=getcomments(fields[1].text);
    Firstline:=frm_common.getFirstLine(generalquery.fields[1].text);
    desc:='(E'+num+') - ';

    noteid:=56;
    if generalquery.fields[5].text<>'' then noteid:=55;
    if generalquery.fields[4].text<>'' then noteid:=54;


    if generalquery.fields[4].text<>'' then desc:=desc+generalquery.fields[4].text+' - ';
    desc:=desc+generalquery.fields[2].text+' - ';
    desc:=desc+generalquery.fields[0].text;
    if generalquery.fields[7].text<>'' then desc:=desc+' - (Owned by '+generalquery.fields[7].text+')';

  //  if v_fullnotes.checked then desc:=desc+#10+generalquery.fields[1].text
  //  else desc:=desc+' '+firstline;

    if generalquery.fields[9].text<>'D' then
    Begin
     if v_fullnotes.checked then desc:=desc+#10+generalquery.fields[1].text
     else desc:=desc+' '+firstline;
    end;

    if generalquery.fields[9].text='D' then
    Begin
     desc:='(E'+num+') - ';
     desc:=desc+generalquery.fields[2].text+' - ';
     desc:=desc+generalquery.fields[0].text+'. Document ID - ';
     desc:=desc+generalquery.fields[6].text;
     noteid:=128;
    End;

    if generalquery.recordcount<=maxgroup then
    begin
     NotesubNode:=MopTree.Addchild(mynodeMopSpan);
     nodeData := moptree.GetNodeData(NotesubNode);
     NodeData.caption := desc;
     //NotesubNode:=MOPTREE.items.AddChild(mynodeMopSpan,desc);
    end
    else
    begin
     NotesubNode:=MopTree.Addchild(NoteTopNode);
     nodeData := moptree.GetNodeData(NotesubNode);
     NodeData.caption := desc;
       //NotesubNode:=MOPTREE.items.AddChild(NoteTopNode,desc);
    end;

    NodeData.index:=noteid;

    NodeData.fontcolor:=clred;
    if generalquery.fields[3].text<>'' then
    Begin
     if strtodatetime(generalquery.fields[3].text)<(now) then
     Begin
      // overdue
      if noteid<>128 then
      Begin
       noteid:=62;
       if generalquery.fields[5].text<>'' then noteid:=61;
       if generalquery.fields[4].text<>'' then noteid:=60;
      end
      else noteid:=63;
      nodedata.index:=noteid;
     end;
    end;

    nodedata.C_record_id :=generalquery.fields[6].text;
    nodedata.C_Owner :=generalquery.fields[7].text;
    nodedata.C_Raised_by :=generalquery.fields[8].text;
    nodedata.C_Date_Raised :=generalquery.fields[2].text;
    nodedata.C_FirstLine := generalquery.Fields[1].Text;

    //NotesubNode.data:=MyRecPtr;
    nodedata.fontcolor:=clblack;
    generalquery.next;
   end;

  end;
 end;

  // Get Outstanding Service Orders
 x:=0;
 With GeneralQuery do
 Begin
  Close;
  deletevariables;
  declarevariable('MPAN',otstring);
  sql.clear;
  sql.add('SELECT R.description, S.comments_1,S.date_raised,S.due_date,');
  sql.add('S.mpancore,null,S.record_id,S.owner,S.raised_by,s.Resolved_status,S.DATE_SENT_TO_MAM FROM MOPMGR.SERVICE_ORDERS S,enquiry.request_type R');
  sql.add('where S.MPANCORE=:MPAN');
  sql.add('and r.id=S.request_type');
  sql.add('and S.resolved_Status in (''N'',''M'',''P'')');
  sql.add('order by s.date_raised desc');
  setvariable('MPAN',mpan);
  open;
  deletevariables;
  if generalquery.recordcount<>0 then
  Begin

   if generalquery.recordcount>MaxGroup then
   Begin
    desc:='(Service Orders Outstanding for MPAN = '+inttostr(generalquery.recordcount)+')';

     NoteTopNode:=MopTree.Addchild(mynodeMopSpan);
     nodeData := moptree.GetNodeData(NoteTopNode);
     NodeData.caption := desc;

    //NoteTopNode:=MOPTREE.items.AddChild(,desc);
    NodeData.index:=27;
   end;

   while not generalquery.eof do      // add sub notes
   Begin
    inc(x);
    if x<10000 then num:=inttostr(x);
    if x<1000 then num:='0'+inttostr(x);
    if x<100 then num:='00'+inttostr(x);
    if x<10 then num:='000'+inttostr(x);
    // Desc:=getcomments(fields[1].text);
    Firstline:=frm_common.getFirstLine(generalquery.fields[1].text);
    desc:='(S'+num+') - ';
    if generalquery.fields[4].text<>'' then desc:=desc+generalquery.fields[4].text+' - ';
    desc:=desc+generalquery.fields[2].text+' - ';
    desc:=desc+generalquery.fields[0].text;
    if generalquery.fields[7].text<>'' then desc:=desc+' - (Owned by '+generalquery.fields[7].text+')';
    if generalquery.fields[9].text='M' then desc:=desc+' - Job sent to MAM on '+generalquery.fields[10].text;
    if generalquery.fields[9].text='P' then desc:=desc+' - See Feedback Form '+generalquery.fields[10].text;
    desc:=desc+' (JOB-'+generalquery.fields[6].text+')';
    if v_fullnotes.checked then desc:=desc+#10+generalquery.fields[1].text
    else
    desc:=desc;
    //desc:=desc+' '+firstline;

    if generalquery.recordcount<=MaxGroup then
    begin
     NotesubNode:=MopTree.Addchild(mynodeMopSpan);
     nodeData := moptree.GetNodeData(NotesubNode);
     NodeData.caption := desc;
      //NotesubNode:=MOPTREE.items.AddChild(mynodeMopSpan,desc)
    end
    else
    begin
     NotesubNode:=MopTree.Addchild(NoteTopNode);
     nodeData := moptree.GetNodeData(NotesubNode);
     NodeData.caption := desc;
     //NotesubNode:=MOPTREE.items.AddChild(NoteTopNode,desc);
    end;

    noteid:=196;  // Normal Order
    if generalquery.fields[9].text='R' then noteid:=197; // Rejected
    if generalquery.fields[9].text='Y' then noteid:=198; // Resolved
    if generalquery.fields[9].text='M' then noteid:=199; // Mam
    if generalquery.fields[9].text='P' then noteid:=19; // Mam

    nodedata.index:=noteid;
    nodedata.fontcolor:=clblack;

    if (generalquery.fields[3].text<>'') and (strtodatetime(generalquery.fields[3].text)<(now)) then
    Begin
     nodedata.fontcolor:=clred;
    end;
    if generalquery.fields[9].text='M' then nodedata.fontcolor:=clblue;


    nodedata.d_span :=generalquery.fields[4].text;
    nodedata.C_record_id :=generalquery.fields[6].text;
    nodedata.C_Owner :=generalquery.fields[7].text;
    nodedata.C_Raised_by :=generalquery.fields[8].text;
    nodedata.C_Date_Raised :=generalquery.fields[2].text;

    //NotesubNode.data:=MyRecPtr;
    generalquery.next;
   end;
  end;
 end;

 try
  //mynodeMopSpan.selected:=true;
 except
 end;
end;

procedure TFRM_Tree.MenuItem13Click(Sender: TObject);
begin
 if treeupdating=true then exit;
 mpannode:=moptree.FocusedNode;
 TreeData:= Moptree.GetNodeData(mpannode);

// mpannode:=MopTree.selected;
 mpan:=TreeData.D_SPAN;
 if not Assigned(FRM_NHH_Metering_MOP) then Application.CreateForm(TFRM_NHH_Metering_MOP, FRM_NHH_Metering_MOP);
 FRM_nhh_metering_mop.show;
 FRM_nhh_metering_mop.getmeterdetails(mpan,'','','');
end;

procedure TFRM_Tree.MenuItem8Click(Sender: TObject);
var
MPAN:string;
begin
 xnode:=moptree.FocusedNode;
 TreeData:= moptree.GetNodeData(xnode);
 mpan:=TreeData.D_SPAN;
 FRM_Main.SearchForSpan(MPAN,1);
end;

procedure TFRM_Tree.MopTreeDblClick(Sender: TObject);
var
s,ts:string;
begin
  xnode:=moptree.FocusedNode;
  TreeData:= Moptree.GetNodeData(xnode);

  if assigned(Treedata)=false then exit;

  screen.cursor:=crhourglass;

  // Check if Note or Enquiry
  if (TreeData.index=51) or
    (TreeData.index=52) or
    (TreeData.index=53) or
    (TreeData.index=54) or
    (TreeData.index=55) or
    (TreeData.index=56) or
    (TreeData.index=57) or
    (TreeData.index=58) or
    (TreeData.index=59) or
    (TreeData.index=60) or
    (TreeData.index=61) or
    (TreeData.index=63) or
    (TreeData.index=128) or
    (TreeData.index=62) then
  begin
    S:=TreeData.C_record_id;
    DisplayOrder:=5;
    TREEENQUIRYRESOLVED:=false;
    ts := EmptyStr;

    if (TreeData.index=63) or (TreeData.index=128) then
    begin
      ts:=TreeData.c_firstline;
      if Messagedlg('Select YES to open document, NO to open enquiry.',mtconfirmation,[MBYES,MBNO],0)=mryes then
      begin
        // doc only
        //shellexecute(Handle,'open',pchar(ts),nil,nil,sw_shownormal);
        FRM_COMMON.ShowImageDoc(ts);
      end
      else
      begin
        // enquiry
        Show_Hot_Note(s);
      end;
    end
    else
    begin
      // Open Note or Enquiry
      Show_Hot_Note(s);
    end;

    if moptree.Selected[xnode.parent]=true then
    begin
      moptree.Selected[xnode.parent]:=false;
      moptree.Selected[xnode.parent]:=true;
    end;
  end
  else

  // Check if Note or Enquiry
  if (TreeData.index=196) or
    (TreeData.index=197) or
    (TreeData.index=198) or
    (TreeData.index=19)  or
    (TreeData.index=199) then
  begin
    frm_mop_service_orders.ordersquery.qbemode:=true;
    frm_mop_service_orders.ordersquery.active:=true;
    MOPORDER:='and s.record_id='''+Treedata.C_record_id+''' order by S.DATE_RAISED asc';
    frm_mop_service_orders.DefaultQuery;
    frm_mop_service_orders.ordersquery.qbemode:=true;
    frm_mop_service_orders.ordersquery.active:=true;
    frm_mop_service_orders. ordersquery.executeqbe;
    frm_mop_service_orders.BuildTree;
    MopOrderClosed:=false;
    frm_mop_service_orders.OpenOrder;

    if moptree.Selected[xnode.parent]=true then
    begin
      moptree.Selected[xnode.parent]:=false;
      moptree.Selected[xnode.parent]:=true;
    end;
  end;

  screen.cursor:=crdefault;
end;


Procedure TFRM_Tree.BuildMopSiteAddress(Mpancore,filename:string;contracterror:boolean);
Var
SiteAddr:string;
begin
 SiteAddr:='';
 with mopsite do
 Begin
  close;
  setvariable('MPAN',mpancore);
  setvariable('FILENAME',filename);
  open;
 end;
 if mopsite.fields[2].text<>''  then SiteAddr:=SiteAddr+mopsite.fields[2].text+', ';
 if mopsite.fields[3].text<>''  then SiteAddr:=SiteAddr+mopsite.fields[3].text+', ';
 if mopsite.fields[4].text<>''  then SiteAddr:=SiteAddr+mopsite.fields[4].text+', ';
 if mopsite.fields[5].text<>''  then SiteAddr:=SiteAddr+mopsite.fields[5].text+', ';
 if mopsite.fields[6].text<>''  then SiteAddr:=SiteAddr+mopsite.fields[6].text+', ';
 if mopsite.fields[7].text<>''  then SiteAddr:=SiteAddr+mopsite.fields[7].text+', ';
 if mopsite.fields[8].text<>''  then SiteAddr:=SiteAddr+mopsite.fields[8].text+', ';
 if mopsite.fields[9].text<>''  then SiteAddr:=SiteAddr+mopsite.fields[9].text+', ';
 if mopsite.fields[10].text<>'' then SiteAddr:=SiteAddr+mopsite.fields[10].text+', ';
 SiteAddr:=SiteAddr+mopsite.fields[11].text+'';

 MyNodeMopSite:=MopTree.Addchild(MyNodeMopSpan);
 nodeData := moptree.GetNodeData(MyNodeMopSite);
 NodeData.caption := 'Premises - '+SiteAddr;
 nodedata.index:=49;

 {MyNodeMopSite:=Moptree.items.Addchild(MyNodeMopSpan, 'Premises - '+SiteAddr);
 MyNodeMopSite.imageindex:=49;
 MyNodeMopSite.selectedindex:=MyNodeMopSite.imageindex;}


 nodedata.D_SPAN := mopsite.fields[0].text;
 nodedata.D_SSD := mopsite.fields[1].text;
 nodedata.D_SPANend := '';
 nodedata.D_ContractRef := mopsite.fields[12].text;
 nodedata.D_ServiceRef := mopsite.fields[17].text;
 nodedata.D_ServiceLevelRef := mopsite.fields[18].text;
 nodedata.D_GSP := mopsite.fields[21].text;;
 nodedata.D_EffectiveFrom := mopsite.fields[16].text;
 nodedata.D_SupplierMPID := mopsite.fields[22].text;
 nodedata.D_SupplierName := mopsite.fields[27].text;
 nodedata.D_ContractError := ContractError;
 //MyNodeMopSite.data:=MyRecPtr;
 MyNodeMopCustomer:=MopTree.Addchild(MyNodeMopSite);
 nodeData := moptree.GetNodeData(MyNodeMopCustomer);
 NodeData.caption := 'Customer';

 //MyNodeMopCustomer:=Moptree.items.Addchild(MyNodeMopSite, 'Customer');
End;

procedure TFRM_Tree.N60Click(Sender: TObject);
begin
 xnode:=moptree.FocusedNode;
 TreeData:= Moptree.GetNodeData(xnode);

 frm_enquiry_summary.setoptions('M');
 FRM_Enquiry_Note.tag:=1;
 FRM_ENQUIRY_summary.mpancore.text:=TreeData.D_SPAN;
 FRM_ENQUIRY_note.mpancore.text:=TreeData.D_SPAN;
 FRM_Enquiry_note.adddata('','','','','M');
  try
  FRM_Enquiry_Note.close;
 except
 end;
 FRM_Enquiry_Note.showModal;
 Begin
   moptree.expanded[xnode]:=false;
   moptree.expanded[xnode]:=true;
 end;
end;

procedure TFRM_Tree.AddEnquiry1Click(Sender: TObject);
begin
 xnode:=moptree.FocusedNode;
 TreeData:= Moptree.GetNodeData(xnode);

 frm_enquiry_summary.setoptions('M');
 FRM_Enquiry_Note.tag:=0;
 FRM_ENQUIRY_summary.mpancore.text:=TreeData.D_SPAN;
 FRM_ENQUIRY_note.mpancore.text:=TreeData.D_SPAN;
 FRM_Enquiry_note.adddata('','','','','M');
  try
  FRM_Enquiry_Note.close;
 except
 end;
 FRM_Enquiry_Note.showModal;
   moptree.expanded[xnode]:=false;
   moptree.expanded[xnode]:=true;
end;

procedure TFRM_Tree.Add1Click(Sender: TObject);
begin
 xnode:=moptree.FocusedNode;
 TreeData:= moptree.GetNodeData(xnode);
 Addscanneddoc((TreeData.D_SPAN),'','M');
   moptree.expanded[xnode]:=false;
   moptree.expanded[xnode]:=true;
end;

procedure TFRM_Tree.AquireAttachDocument2Click(Sender: TObject);
{Var
Custid,custname:string;}
begin
 {xnode:=moptree.FocusedNode;
 TreeData:= moptree.GetNodeData(xnode);

 Custid:=TreeData.D_SPAN;
 CustName:='';
 FRM_Twain.custlabel.caption:=custid;
 FRM_Twain.custname.caption:=custname;
  FRM_Twain.RoleLabel.caption:='M';
 Frm_Twain.showmodal;}
end;

procedure TFRM_Tree.ShowAllEnquiriesNotesDocs1Click(Sender: TObject);
Var
MPANCORE:string;
begin
 xnode:=moptree.FocusedNode;
 TreeData:= moptree.GetNodeData(xnode);

 if treeupdating=true then exit;
 Begin
  MPANCORE:=TreeData.D_SPAN;
  frm_enquiry_summary.setoptions('M');
  FRM_ENQUIRY_SUMMARY.MPANEnquiryNotes(MPANCORE);
  FRM_Enquiry_Summary.show;
  FRM_Enquiry_Summary.windowstate:=wsnormal;
 end;
end;

procedure TFRM_Tree.AddServiceOrder1Click(Sender: TObject);
begin
 xnode:=moptree.FocusedNode;
 TreeData:= Moptree.GetNodeData(xnode);

 FRM_MOP_SERVICE_ORDERS.MPANCORE.text:=TreeData.D_SPAN;
 FRM_MOP_SERVICE_ORDERS.AddNoteBtnClick(sender);

 moptree.expanded[xnode]:=false;
 moptree.expanded[xnode]:=true;

end;


procedure TFRM_Tree.ShowAllServiceOrders1Click(Sender: TObject);
Var
Mpan:string;
begin
 xnode:=moptree.FocusedNode;
 TreeData:= moptree.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 FRM_MOP_SERVICE_ORDERS.mpancore.text:=MPAN;
 if mpan<>'' then FRM_MOP_SERVICE_ORDERS.findbtn.click();
 FRM_MOP_SERVICE_ORDERS.show;
end;

procedure TFRM_Tree.M_AUDITClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Application.CreateForm(TFRM_CUSTOMER_ACCESS_LOG, FRM_CUSTOMER_ACCESS_LOG);
 try
 with frm_customer_access_log.auditquery do
 Begin
  close;
  setvariable('custid',Treedata.D_Customer_ID);
  open;
 End;
 FRM_Customer_access_log.showmodal;
 finally
  FRM_Customer_access_log.release;
 end;
end;

procedure TFRM_Tree.m_ShowStockOrdersClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 //--------by Maryam on 14/11/2016
 Application.CreateForm(TFrm_ShowStockOrders, Frm_ShowStockOrders);
 try
   Frm_ShowStockOrders.customer_id := Treedata.D_Customer_ID ;
   Frm_ShowStockOrders.DoQuery;
   if Frm_ShowStockOrders.order_query.recordcount=0 then
   begin
     Messagedlg('There are no Stock Item Orders for this customer.',mtinformation,[mbok],0);
   end
   else
   Frm_ShowStockOrders.showmodal;
 finally
  Frm_ShowStockOrders.release;
 end;
end;

procedure TFRM_Tree.m_OrderNewStockItem_CustomerClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 //--------by Maryam on 14/11/2016
 Application.CreateForm(TFrm_OrderNewStockItem, Frm_OrderNewStockItem);
 try
   Frm_OrderNewStockItem.level := 'CUSTOMER';//in customer level
   Frm_OrderNewStockItem.customer_id := Treedata.D_Customer_ID ;
   Frm_OrderNewStockItem.showmodal;
 finally
  Frm_OrderNewStockItem.release;
 end;
end;

procedure TFRM_Tree.m_OrderNewStockItem_agreementClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 //--------by Maryam on 14/11/2016
 Application.CreateForm(TFrm_OrderNewStockItem, Frm_OrderNewStockItem);
 try
   Frm_OrderNewStockItem.level := 'AGREEMENT' ;//in agreement level
   Frm_OrderNewStockItem.customer_id := Treedata.D_Customer_ID ;
   Frm_OrderNewStockItem.agreement_id := Treedata.D_Agreement_ID;
   Frm_OrderNewStockItem.showmodal;
 finally
  Frm_OrderNewStockItem.release;
 end;
end;

procedure TFRM_Tree.m_OrderNewStockItem_PremiseClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 //--------by Maryam on 14/11/2016
 Application.CreateForm(TFrm_OrderNewStockItem, Frm_OrderNewStockItem);
 try
   Frm_OrderNewStockItem.level := 'PREMISE';//in premise level
   Frm_OrderNewStockItem.customer_id := Treedata.D_Customer_ID ;
   Frm_OrderNewStockItem.agreement_id := Treedata.D_Agreement_ID ;
   Frm_OrderNewStockItem.premise_id := Treedata.D_Premise_Id ;
   Frm_OrderNewStockItem.showmodal;
 finally
  Frm_OrderNewStockItem.release;
 end;
end;

procedure TFRM_Tree.S_TransferCreditClick(Sender: TObject);
begin
//added by maryam on 23/11/2016
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

  Application.CreateForm(TFrm_TransferCredit, Frm_TransferCredit);
 try
  Frm_TransferCredit.premise_id := Treedata.D_Premise_id;
  Frm_TransferCredit.showmodal;
 finally
  Frm_TransferCredit.release;
 end;
end;

procedure TFRM_Tree.m_OrderNewStockItem_ElecClick(Sender: TObject);
var
agid,cid:string;
begin
 mpannode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(mpannode);
 agid:=TreeData.D_Agreement_ID;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 //--------by Maryam on 14/11/2016
 Application.CreateForm(TFrm_OrderNewStockItem, Frm_OrderNewStockItem);
 try
   Frm_OrderNewStockItem.level := 'SUPPLY ELEC SMETS';//in supply level
   Frm_OrderNewStockItem.customer_id := cid;
   Frm_OrderNewStockItem.agreement_id := agid ;
   Frm_OrderNewStockItem.premise_id := Treedata.D_Premise_Id ;
   Frm_OrderNewStockItem.supply_id := Treedata.D_SPAN ;
   Frm_OrderNewStockItem.showmodal;
 finally
  Frm_OrderNewStockItem.release;
 end;
end;

procedure TFRM_Tree.m_OrderNewStockItem_GasClick(Sender: TObject);
var
agid,cid:string;
begin
 mpannode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(mpannode);
  agid:=TreeData.D_Agreement_ID;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 //--------by Maryam on 14/11/2016
 Application.CreateForm(TFrm_OrderNewStockItem, Frm_OrderNewStockItem);
 try
   Frm_OrderNewStockItem.level := 'SUPPLY';//in supply level
   Frm_OrderNewStockItem.customer_id := cid;
   Frm_OrderNewStockItem.agreement_id := agid ;
   Frm_OrderNewStockItem.premise_id := Treedata.D_Premise_Id ;
   Frm_OrderNewStockItem.supply_id := Treedata.D_SPAN ;
   Frm_OrderNewStockItem.showmodal;
 finally
  Frm_OrderNewStockItem.release;
 end;
end;

procedure TFRM_Tree.IGTAdminCharge1Click(Sender: TObject);
Var
RegID,AGID,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 If messagedlg('Are you sure you wish to print an IGT charge letter for this MPRN?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('CRM\IGT\IGT_Letter.rpt','Gas IGT','{ACCOUNT_HOLDERS.CONTACT_ORDER}=1 and {SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.RegistrationHistory2Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
FRM_Gas_history.showhistory(TreeData.D_SPAN);
end;

Procedure TFRM_Tree.FormActivate(Sender: TObject);
Begin
  If ExternalAgent <> EmptyStr then
    Begin
      NEWBTN.Enabled := False;
      NEWBTN.Visible := False;
    End; // If

  N63.Visible     := CAUDIT;
  M_AUDIT.Visible := CAUDIT;
  M_AUDIT.Enabled := CAUDIT;

  // BSL-SJ - 11/08/2021 - CRM-511 - Control over Super Customer - Fixing Bug.
  SuperCustIcon := 0;
End; // Proc

procedure TFRM_Tree.RemoveDonNotBillAllowBilling1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 setasBill(TreeData.D_REGid);
end;

procedure TFRM_Tree.RemoveDoNotBillAllowBilling1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 setasBill(TreeData.D_REGid);
end;

procedure TFRM_Tree.AddOneOffCharge(desc: String);
var
  Span, Agreement_ID, Reg_id: String;
begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  Span := TreeData.D_SPAN;
  Agreement_ID := TreeData.D_AGREEMENT_ID;
  Reg_id := TreeData.D_REGID;
  FRM_ONE_OFF_CHARGE.ShowTheseDetails(Span, Reg_id, Agreement_ID, desc);
  FRM_ONE_OFF_CHARGE.ShowModal;
end;

procedure TFRM_Tree.AddOneOffCharge1Click(Sender: TObject);
begin
  AddOneOffCharge('G');
end;

procedure TFRM_Tree.AddOneOffCharge2Click(Sender: TObject);
begin
  AddOneOffCharge('E');
end;

procedure TFRM_Tree.AddOneOffCharge3Click(Sender: TObject);
begin
  AddOneOffCharge('T');
end;

procedure TFRM_Tree.AddOneOffCharge4Click(Sender: TObject);
begin
  AddOneOffCharge('B');
end;

procedure TFRM_Tree.RefereAFriend1Click(Sender: TObject);
Var
  CustomerID:string;
begin
  //added by maryam on 15/07/2016
xnode:=treeview1.FocusedNode;
TreeData := treeview1.GetNodeData(xnode);

if treeupdating=true then exit;
CustomerID:=TreeData.D_Customer_ID;

Application.CreateForm(Tfrm_Refer_Friend, frm_Refer_Friend);
frm_Refer_Friend.Caption := 'Refer a Friend - Customer id: '+ CustomerID;

try
  frm_Refer_Friend.RefreshAll(CustomerID);
  if ((frm_Refer_Friend.gbReferee.Visible )or
     (frm_Refer_Friend.gbReferrer.Visible ))
  then frm_Refer_Friend.showmodal
  else
  begin
   messagedlg('There is no data as Referer or Referee for this customer',mtconfirmation,[mbok],0);
  end;
finally
 frm_Refer_Friend.release;
end;

end;

procedure TFRM_Tree.RefreshAccountOrder(Custid:string);
var
x:integer;
begin
 // Now Refresh Primary Contacts. Dont Want deceased as primary
 with main_data_module.generalquery do
 Begin
  close;
  deletevariables;
  declarevariable('CID',otlong);
  sql.clear;
  sql.add('select account_holder_id,primary_contact,contact_order,account_holder_status_id');
  sql.add('from crm.account_holders where customer_id=:CID');
  sql.add('order by account_holder_status_id asc ,primary_contact desc, contact_order asc');
  setvariable('CID',custid);
  open;
  deletevariables;
 End;
 x:=1;

 main_data_module.updatescript.lines.clear;
 while not main_data_module.generalquery.eof do
 Begin
  with main_data_module.updatequery do
  begin
   close;
   sql.clear;
   sql.add('update crm.account_holders');
   sql.add('set contact_order='+inttostr(x));
   sql.add('where account_holder_id='+main_data_module.generalquery.fields[0].text);
   main_data_module.updatescript.lines.add(main_data_module.updatequery.sql.text+';');
  end;
  inc(x);
  main_data_module.generalquery.next;
 End;
 main_data_module.updatescript.execute;
 FRM_Login.MainSession.commit;
end;


procedure TFRM_Tree.PrepareInstall1Click(Sender: TObject);
var
eregid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;

 Messagedlg('Prepare Smart Meter Install functionality is no longer available.'+#13+
            'This feature has been DISABLED by ADMIN until further notice.',mterror,[mbok],0);
 exit;

 eregid:=TreeData.D_REGID;
 // First Check SPAN STATUS - Don Wnat to Reorder Duplicates.
 with main_data_module.tempquery do
 Begin
  deletevariables;
  close;
  sql.clear;
  sql.add('select order_status_id,span_start_date from crm.spans where registration_id='+eregid);
  open;
 end;

 if ((main_data_module.tempquery.Fields[0].text<>'21') and (main_data_module.tempquery.Fields[0].text<>'3')) then
 Begin
  Messagedlg('you cannot book/rebook a SMART Meter install unless the MPAN status is ORDER PLACED or SMART ORDER PLACED.',mterror,[mbok],0);
  exit;
 End;

 if main_data_module.tempquery.Fields[0].text='21' then
 Begin
  If Messagedlg('A SMART Meter Install has already been ordered for this MPAN for '+main_data_module.tempquery.Fields[1].text+#13+
                'Are you sure you wish to re-order another install?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;
 End;



 Application.CreateForm(TFRM_PREPARE_INSTALL, FRM_PREPARE_INSTALL);
 try
 frm_prepare_install.tag:=1;
 WIth FRM_PREPARE_INSTALL.spanquery do
 Begin
  close;
  setvariable('Registrationid',eregid);
  open;
 end;
 FRM_PREPARE_INSTALL.SHOWMODAL;
 finally
 FRM_PREPARE_INSTALL.RELEASE;
 end;
end;

procedure TFRM_Tree.MenuItem9Click(Sender: TObject);
var
  s:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 S:=TreeData.C_record_id;

  if MessageDlg('Are you sure you wish to Drop this note off the tree?'+#13+'This will clear the Validity Period of the Note',
    mtConfirmation, [mbYes, mbNo], 0) = mrno then exit;

 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('update enquiry.enquiries set expiry_code=''E'' where record_id='+s);
  execute;
 end;
 FRM_Login.MainSession.commit;
 treeview1.DeleteNode(xnode);
 treeview1.expanded[xnode.parent]:=false;
 treeview1.expanded[xnode.parent]:=true;
end;

procedure TFRM_Tree.COTMoveIn1Click(Sender: TObject);
Var
Custid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Custid:=TreeData.D_Customer_ID;
 FRM_Reports.PrintThisReport('CRM\cot_in.rpt','Customer Move In','{ACCOUNT_HOLDERS.CONTACT_ORDER}=1 and {CUSTOMER.CUSTOMER_ID}='+custid+'','','PRINTER',custid,'');
end;

procedure TFRM_Tree.SUnrestricted1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 DoSPANCHANGE('S',mpan);
end;

procedure TFRM_Tree.EEconomy71Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 DoSPANCHANGE('E',mpan);
end;

procedure TFRM_Tree.doSpanchange(Stype,span:string);
Var
ssd:string;
begin
 frm_date.efd.date:=now;
 frm_date.showmodal;
 if frm_Date.tag=0 then exit;
 ssd:=frm_date.efd.text;
 if Messagedlg('Are you sure you wish to change the SPAN type to '+stype+' for MPAN '+mpan+#13+'Any Other changes Effective on or after '+ssd+' will be removed.',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;
 // do change
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from crm.span_ssc_changes where span='''+mpan+''' and effective_from>=to_date('''+ssd+''',''DD/MM/YYYY'')');
  execute;
  close;
  sql.clear;
  sql.add('insert into crm.span_ssc_changes (SPAN,SPAN_TYPE,EFFECTIVE_FROM) values ('''+mpan+''','''+stype+''',to_date('''+ssd+''',''DD/MM/YYYY''))');
  execute
 End;
 try
  treeview1.Expanded[xnode]:=false;
  treeview1.Expanded[xnode]:=true;
 except
  Messagedlg('Refresh Tree To view Changes',mtinformation,[mbok],0);
 end;
end;

procedure TFRM_Tree.ReviveAgreement1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 frm_agreement.ReviveAgreements(TreeData.D_agreement_id);
end;

procedure TFRM_Tree.RiaseErroneousTransferRequest1Click(Sender: TObject);
Var
  regid,span,ssd:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 REGID:=TreeData.D_REGid;
 SPAN:=TreeData.D_SPAN;
 SSD:=TreeData.D_SSD;
 Application.CreateForm(TFRM_ELEC_ET, FRM_ELEC_ET);
 try
  FRM_ELEC_ET.clearfields;
  FRM_ELEC_ET.SETDefault(SPAN);
  FRM_ELEC_ET.ShowModal;
 finally
  FRM_ELEC_ET.release;
 end;
end;

procedure TFRM_Tree.PrepareInstall2Click(Sender: TObject);
var
eregid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;

 Messagedlg('Prepare Smart Meter Install functionality is no longer available.'+#13+
            'This feature has been DISABLED by ADMIN until further notice.',mterror,[mbok],0);
 exit;

 eregid:=TreeData.D_REGID;
 // First Check SPAN STATUS - Don Wnat to Reorder Duplicates.
 with main_data_module.tempquery do
 Begin
  deletevariables;
  close;
  sql.clear;
  sql.add('select order_status_id,span_start_date from crm.spans where registration_id='+eregid);
  open;
 end;

 if ((main_data_module.tempquery.Fields[0].text<>'21') and (main_data_module.tempquery.Fields[0].text<>'3')) then
 Begin
  Messagedlg('you cannot book/rebook a SMART Meter install unless the MPRN status is ORDER PLACED or SMART ORDER PLACED.',mterror,[mbok],0);
  exit;
 End;

 if main_data_module.tempquery.Fields[0].text='21' then
 Begin
  If Messagedlg('A SMART Meter Install has already been ordered for this MPRN for '+main_data_module.tempquery.Fields[1].text+#13+
                'Are you sure you wish to re-order another install?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;
 End;



 Application.CreateForm(TFRM_PREPARE_INSTALL, FRM_PREPARE_INSTALL);
 try
 frm_prepare_install.tag:=2;
 WIth FRM_PREPARE_INSTALL.spanquery do
 Begin
  close;
  setvariable('Registrationid',eregid);
  open;
 end;
 FRM_PREPARE_INSTALL.SHOWMODAL;
 finally
 FRM_PREPARE_INSTALL.RELEASE;
 end;

end;

procedure TFRM_Tree.FeedbackForm1Click(Sender: TObject);
Var
span,SSD,job_id:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 job_id:=(TreeData.C_record_id);
 with main_data_module.tempquery do
 Begin
  deletevariables;
  close;
  sql.clear;
  sql.add('select * from mopmgr.gsmi where job_id='+job_id);
  open;
 End;
 // if no Feedback Exists
 if main_data_module.tempquery.recordcount=0 then
 Begin
  SSD:=TreeData.D_SSD;
  span:=TreeData.d_sPAN;
  if ssd='' then ssd:='01/01/2090';

  Application.CreateForm(TFRM_GSMI, FRM_GSMI);
  try
   FRM_GSMI.jobno.caption:=job_ID;
   FRM_GSMI.setdetails(span,ssd);
   FRM_GSMI.showmodal;
  finally
   FRM_GSMI.release;
  end;

 end
 else
 Begin
  //If Feedback Exists
  Application.CreateForm(TFRM_GSMI, FRM_GSMI);
  try
   FRM_GSMI.getdetails(job_id);
   FRM_GSMI.showmodal;
   span:=frm_gsmi.MPAN.text;
  finally
   FRM_GSMI.release;
  end;
 end;
 FRM_Main.SearchForSpan(span,1);
end;

// BSL-SJ - 11/08/2021 - CRM-511 - Control over Super Customer.
Function TFRM_Tree.fGetCustIcon(aCustTypeId: Integer): Integer;
Begin
  qrGetCustIcon.Close;
  qrGetCustIcon.SetVariable('CustTypeID', aCustTypeId);
  qrGetCustIcon.Open;
  Result := qrGetCustIconICON_INDEX.AsInteger;
  qrGetCustIcon.Close;
End; // Funct

procedure TFRM_Tree.gsmiletterClick(Sender: TObject);
Var
AgId,title,salesref,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AgID:=TreeData.D_Agreement_id;

 with main_data_module.generalquery do
 Begin
  close;
  deletevariables;
  DeclareVariable('AID', otlong);
  sql.clear;
  sql.add('Select sales_reference,customer_id from crm.agreements where agreement_id=:AID');
  setvariable('AID',agid);
  open;
  deletevariables;
 End;
 salesref:=main_data_module.generalquery.Fields[0].text;
 cid:=main_data_module.GeneralQuery.Fields[1].Text;
 if copy(salesref,5,3)='DTD' then TITLE:='SIGNUP LETTER + WELCOME-PACK'
 else title:='SIGNUP LETTER';

 //FRM_Reports.PrintThisReport('MASTER\blank_agreement_template.rpt',''+TITLE+'','{LETTER_DETAILS.OUR_REF}=''CS-W-PP-2A'' and {CUSTOMER_MAILING.AGREEMENT_ID}='+agid,'','PRINTER',CID'');
 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\Signup_smart_install.rpt','Get Smart Signup Letter','{ACCOUNT_HOLDERS.CONTACT_ORDER} = 1 and {CUSTOMER.CUSTOMER_ID} = '+agid,'','PRINTER',CID,'');
end;



procedure TFRM_Tree.ObjectionReceivedgetSmart1Click(Sender: TObject);
Var
RegID,agid,cid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('GETSMART\get_smart_SignUp_Objection_received.rpt','Objection Received get Smart','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;


procedure TFRM_Tree.ObjectionReceivedgetSmart2Click(Sender: TObject);
Var
RegID,agid,cid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('GETSMART\get_smart_SignUp_Objection_received.rpt','Objection Received get Smart','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;


procedure TFRM_Tree.RejectionReceivedgetSmart1Click(Sender: TObject);
Var
RegID,agid,cid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('GETSMART\get_smart_SignUp_rejection_received.rpt','Rejection Received get Smart','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.RejectionReceivedgetSmart2Click(Sender: TObject);
Var
RegID,agid,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('GETSMART\get_smart_SignUp_rejection_received.rpt','Rejection Received get Smart','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;

Procedure TFRM_Tree.RelationshipRating1Click(Sender: TObject);
Begin
  TFrmRelationshipRatingChange.Launch(Treedata.D_Customer_ID);
  ViewCustomerTree1Click(Self);
End; // Proc

procedure TFRM_Tree.AddCustomerFLAG1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Raise_Hot_Note(4,Treedata.D_Customer_ID);
end;

procedure TFRM_Tree.ShowLossNotifications1Click(Sender: TObject);
var
mpan:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 Application.CreateForm(TFRM_Losses, FRM_Losses);
 frm_losses.lossdate.date:=now;
 frm_losses.btn_cust_search.click();
 frm_losses.btn_cust_search.click();
 frm_losses.db_SPAN.text:=mpan;
 frm_losses.lossdate.date:=strtodate('01/01/2007');
 frm_losses.btn_cust_search.click();
 FRM_LOSSES.showmodal;
 FRM_Losses.release;
end;

procedure TFRM_Tree.ShowLossNotifications2Click(Sender: TObject);
var
mpan:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 Application.CreateForm(TFRM_Losses, FRM_Losses);
 frm_losses.lossdate.date:=now;
 frm_losses.btn_cust_search.click();
 frm_losses.btn_cust_search.click();
 frm_losses.db_SPAN.text:=mpan;
 frm_losses.lossdate.date:=strtodate('01/01/2007');
 frm_losses.btn_cust_search.click();
 FRM_LOSSES.showmodal;
 FRM_Losses.release;
end;

procedure TFRM_Tree.MeterExchangeQuery1Click(Sender: TObject);
Var
RegID,agid,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('CRM\meter_exchange_form.rpt','Meter Exchange Query','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.MeterExchangeQuery2Click(Sender: TObject);
Var
RegID,agid,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);

 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('CRM\meter_exchange_form.rpt','Meter Exchange Query','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.E7Cancellation1Click(Sender: TObject);
Var
RegID,agid,CID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);

 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('GETSMART\GET_SMART_E7_CANX.rpt','get Smart e7 Cancellation','{SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',CID,'');
end;

procedure TFRM_Tree.EBSSPaymentClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
  try
   Application.CreateForm(TFRM_EBSSPayment, FRM_EBSSPayment);
   FRM_EBSSPayment.SPAN.Text:=TreeData.D_SPAN;
   FRM_EBSSPayment.REGID := Treedata.D_REGID;
   FRM_EBSSPayment.showmodal;
  finally
   FreeAndNil(FRM_EBSSPayment);
  end;
end;

procedure TFRM_Tree.ShowLibertyVends1Click(Sender: TObject);
var
cid,agid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_Agreement_id;
 cid:=frm_common.GetCustomerIdfromAgreementid(agid);

 Application.CreateForm(Tfrm_liberty_vend_codes, frm_liberty_vend_codes);
 try

   // Set Effective Dates based on Agreement
  frm_liberty_vend_codes.EFD.Text:=TreeData.D_agreement_Start_date;
  if frm_liberty_vend_codes.EFD.Text='' then frm_liberty_vend_codes.EFD.Text:='01/01/2000';

  frm_liberty_vend_codes.ETD.Text:=TreeData.D_agreement_end_date;
  if frm_liberty_vend_codes.ETD.Text='' then frm_liberty_vend_codes.ETD.Text:=datetostr(now+1);


  frm_liberty_vend_codes.removedcheck.checked:=false;
  with frm_liberty_vend_codes.meters do
  Begin
   close;
   sql:=frm_liberty_vend_codes.minimal_meters.sql;
   setvariable('agreement_id',agid);
   setvariable('custid',cid);
   open;

   // WRIKE 160077345: Bug - Liberty Vend Code Screen
   if recordcount=0 then
   begin
    close;
    setvariable('custid',FRM_common.GetLibertyIDfromCustomerId(CID));
    open;
   end;


   if recordcount<>0 then
   Begin
    frm_liberty_vend_codes.vends.close;
    frm_liberty_vend_codes.Showmodal;
   end
   else
   begin
    Messagedlg('No Liberty Vends Found. Check this is a NSS Metered account.',mtinformation,[mbok],0);

    exit;
   End;
  end;
 finally
  frm_liberty_vend_codes.release;
 end;
end;

procedure TFRM_Tree.PaypointAgencyLocator1Click(Sender: TObject);
var
  url : string;
begin
  url := WideString(FRM_Common.GETVALUE('PAYPOINT_LOCATOR'));
  if trim(url) <> EmptyStr  then
    ShellExecute(Handle, 'open', pchar(url), nil, nil, SW_SHOWNORMAL)
  else
    Messagedlg('Unable to open Paypoint Locator website.',mterror,[mbok],0);
end;

procedure TFRM_Tree.DirectSignupLetterGetSmart1Click(Sender: TObject);
Var
AgId,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AgID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);

 {with main_data_module.GeneralQuery do
 Begin
  close;
  sql.clear;
  sql.add('select * from crm.get_smart_installs_letters');
  sql.add('where customer_id='+agid);
  open;
 End;


 if main_data_module.generalquery.recordcount=0 then
 Begin
  Messagedlg('Unable to print Direct Get Smart Signup Letter.'+#13+
             'Have you prepared the Electric Install?',mterror,[mbok],0);
  exit;
 End; }

 FRM_Reports.PrintThisReport('GETSMART\GET_SMART_DIRECT_SIGNUP.rpt','Direct Get Smart Signup Letter','{ACCOUNT_HOLDERS.CONTACT_ORDER} = 1 and {CUSTOMER.CUSTOMER_ID} = '+cid,'','PRINTER',cid,'');
end;

function TFRM_Tree.GetCustomerPronoun(ContactID, CustomerID: string): string;
var
  vPrefID: Integer;
begin
  Result := '';
  with PronounQuery do
    begin
      Close;
      SQL.Clear;
      DeleteVariables;
      DeclareVariable('CONTACT_ID', otString);
      DeclareVariable('CUSTOMER_ID', otString);
      DeclareVariable('RESULT', otInteger);
      SQL.Add('begin');
      SQL.Add('WEB_API.pk_addressee_preferences.get_data(:CUSTOMER_ID, :CONTACT_ID, :RESULT);');
      SQL.Add('end;');
      SetVariable('CONTACT_ID', ContactID);
      SetVariable('CUSTOMER_ID', CustomerID);
      Open;
      vPrefID := GetVariable('RESULT');

      if vPrefID > 0 then
        begin
          Close;
          SQL.Clear;
          DeleteVariables;
          DeclareVariable('ID_PREF', otInteger);
          SetVariable('ID_PREF', vPrefID);
          SQL.Add('SELECT apl.DESCRIPTION FROM crm.ADDRESSEE_PREFERENCE_LOOKUP apl ');
          SQL.Add('WHERE apl.ID = :ID_PREF');
          Open;
          if RecordCount > 0 then
            Result := ' (' + FieldByName('DESCRIPTION').AsString + ')'
          else
            Result := '';
          Close;
        end;
    end;
end;

procedure TFRM_Tree.GetSmartRenewalLetter1Click(Sender: TObject);
Var
AgId,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AgID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);

 FRM_Reports.PrintThisReport('CRM\Sign Up Reports\GetSmart_Renewal.rpt','Get Smart Renewal','{ACCOUNT_HOLDERS.CONTACT_ORDER} = 1 and {CUSTOMER.CUSTOMER_ID} = '+cid,'','PRINTER',cid,'');
end;

procedure TFRM_Tree.ShowLegacyPrePayVends1Click(Sender: TObject);
Var
agid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;

 Application.CreateForm(Tfrm_legacy_vends, frm_legacy_vends);
 try
  with frm_legacy_vends.meters do
  Begin
   close;
   setvariable('agreement_id',agid);
   open;
   if recordcount<>0 then
   Begin
    frm_legacy_vends.vends.close;
    frm_legacy_vends.Showmodal;
   end
   else
   begin
    Messagedlg('No Legacy Vends Found. Check this is a PrePay account.',mtinformation,[mbok],0);

    exit;
   End;
  end;
 finally
  frm_legacy_vends.release;
 end;
end;

procedure TFRM_Tree.MenuItem11Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 FRM_Main.SearchForSupercust(TreeData.D_Customer_ID);
end;

procedure TFRM_Tree.DisagregatefromSuperCustomer1Click(Sender: TObject);
Var
custid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 custid:=TreeData.D_Customer_id;
 if Messagedlg('Are you sure you wish to de-aggregate this account from Super Customer?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from crm.customer_to_super_customer where customer_id='+custid);
  execute;
  frm_login.mainsession.commit;
  FRM_Common.SetAudit(CUSTID,'','','Customer De-aggregated');
 end;
 FRM_Main.SearchForcust(custid);
 xNode := Treeview1.GetFirst();
 treeview1.expanded[xnode]:=false;
 treeview1.expanded[xnode]:=true;
end;

procedure TFRM_Tree.AggretagewithSuperCustomer1Click(Sender: TObject);
Var
custid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 custid:=TreeData.D_Customer_id;
 Application.CreateForm(Tfrm_select_supercust,frm_select_supercust);
 try
  frm_select_supercust.supercust.close;
  frm_select_supercust.supercust.open;
  if frm_select_supercust.supercust.recordcount=0 then
  Begin
   Messagedlg('There are no super customers in the database.',mtinformation,[mbok],0);
   exit;
  end;
  frm_select_supercust.tag:=0;
  frm_select_supercust.showmodal;
  if frm_select_supercust.tag=2 then
  Begin
   with main_data_module.updatequery do
   Begin
    close;
    sql.clear;
    sql.add('delete from crm.customer_to_super_customer where customer_id='+custid);
    execute;
    close;
    sql.clear;
    sql.add('insert into crm.customer_to_super_customer values('+custid+','+frm_select_supercust.supercust.fields[0].text+')');
    execute;
    frm_login.mainsession.commit;
    FRM_Common.SetAudit(CUSTID,'','','Customer aggregated to Super Customer - '+frm_select_supercust.supercust.fields[0].text);
   end;
  end;
  finally
  FRM_select_SuperCust.release;
  FRM_Main.SearchForcust(custid);
  xNode := Treeview1.GetFirst();
  treeview1.expanded[xnode]:=false;
  treeview1.expanded[xnode]:=true;
 End;
end;

procedure TFRM_Tree.SuperCustomerStatement1Click(Sender: TObject);
Var
Custid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Custid:=TreeData.D_Customer_ID;
 FRM_Reports.PrintThisReport('RATING_BILLING\Group_statement.rpt','Super Customer Statement','{CUSTOMER_TO_SUPER_CUSTOMER.SUPER_CUSTOMER_ID}='+custid+'','','',custid,'');
end;

procedure TFRM_Tree.ChangeDefaultSpanType1Click(Sender: TObject);
var
  regid,span : string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 regid:=TreeData.D_REGID;
 span:=TreeData.D_SPAN;
 Application.CreateForm(TFRM_SPANTYPE, FRM_SPANTYPE);
 Try
  FRM_SPANTYPE.DataSetSpanType.close;
  FRM_SPANTYPE.DataSetSpanType.Open;
  FRM_SPANTYPE.registrationId := regid;
  FRM_SPANTYPE.span.text := Treedata.D_SPANDESC;
  FRM_SPANTYPE.Showmodal;
  if frm_spantype.tag=1 then
  Begin
    // Refresh Parent Node
   if treeview1.Selected[xnode]=true then
   Begin
    treeview1.Expanded[xnode.parent]:=false;
    treeview1.Expanded[xnode.parent]:=true;
    //if node is first item, then no paretn so refresh span
    if treeview1.AbsoluteIndex(XNode)=0 then FRM_Main.SearchForSpan(span,0);
   end;
  end;
 finally
  FRM_SPANTYPE.release;
 end;
end;

procedure TFRM_Tree.RejectIGT1Click(Sender: TObject);
Var
agid,regid,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 REGID:=TreeData.D_REGid;
 FRM_Reports.PrintThisReport('CRM\IGT\IGT_GET_SMART_rejection.rpt','IGT Reject Signup get Smart','{ACCOUNT_HOLDERS.CONTACT_ORDER}=1 and {SPANS.REGISTRATION_ID}='+REGid,'','PRINTER',cid,'');
end;

procedure TFRM_Tree.SetStartDate1Click(Sender: TObject);
var
v:integer;
regid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 regid:=TreeData.D_REGID;

 // Status mus be Order Placed first.
 v:=2;
 status:=TreeData.D_STATUS;
 if status<>'Order Placed' then
 Begin
  Messagedlg('Status must be Order Placed, in order to change start date.',mterror,[mbok],0);
  exit;
 end;
 Begin
  Application.CreateForm(Tfrm_datepicker, frm_datepicker);
  repeat
   FRM_DatePicker.tag:=0;
   frm_datepicker.DatePan.caption:='Please Select Supply Start Date';
   FRM_DatePicker.showmodal;
   if FRM_DatePicker.tag=0 then messagedlg('Please Select a SSD',mtinformation,[MBOK],0)
  until FRM_DatePicker.tag=1;
  ssd:=datetostr(frm_datepicker.cal1.date);
  frm_datepicker.release;
 End;

 if (strtodate(ssd)<date+2) or (SSD='') then
  Begin
   if messagedlg('Registration Date of '+ssd+' is in the PAST or too late for registration.'+#13+'Default to Earliset SSD. e.g. Today +2 for Elec, Today+16 for Gas.',mtconfirmation,[MByes,MBno],0)=mryes then
   SSD:=(datetostr(now+V)); // Earliest SSD
  End;

  if strtodate(SSD)>now+28 then
  Begin
   If Messagedlg('Registration date of '+ssd+' is more than 28 days in the future.'+#13+
                 'Continue?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;
  End;
  // do stuff
  if Messagedlg('Confirm Set SPAN start Date to '+ssd+'?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;


  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('Update crm.spans');
   sql.add('set span_start_date=to_date('''+SSD+''',''DD/MM/YYYY'')');
   sql.add('where registration_id='+regid);
   execute;
  end;
  frm_login.mainsession.commit;

     // Refresh Parent Node
   if treeview1.Selected[xnode]=true then
   Begin
    treeview1.Expanded[xnode.parent]:=false;
    treeview1.Expanded[xnode.parent]:=true;
    //if node is first item, then no paretn so refresh span
    if treeview1.AbsoluteIndex(XNode)=0 then FRM_Main.SearchForSpan(TreeData.D_SPAN,0);
   end;

end;

// BSL-SJ - 11/08/2021 - CRM-511 - Control over Super Customer - Fixing Bug.
Procedure TFRM_Tree.SetSuperCustIcon(Const Value: Integer);
Begin
  FSuperCustIcon := fGetCustIcon(Value);
End; // Proc

procedure TFRM_Tree.PC11Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 DoBillPCCHANGE('1',mpan);
end;

procedure TFRM_Tree.PopUpAccountPopup(Sender: TObject);
begin
  RemoveAccountHolder1.Visible := IsValidFinancial;
end;

procedure TFRM_Tree.RaiseFlag(Sender: TObject);
begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  with Main_Data_Module.UpdateQuery do
  begin
    Close;
    SQL.Clear;
    DeleteVariables;
    SQL.Add('BEGIN');
    SQL.Add('CRM.PR_ENQUIRY_INSERT(:AGREEMENTID, :FLAGID);');
    SQL.Add('END;');
    DeclareAndSet('AGREEMENTID', otFloat, StrToFloat(TreeData.D_Agreement_ID));
    DeclareAndSet('FLAGID', otInteger, TMenuItem(Sender).Tag);
    try
      Execute;
      FRM_Login.MainSession.Commit;
      MessageDlg('Flag [' + StringReplace(TMenuItem(Sender).Caption, '&', EmptyStr, []) + '] has been raised.', TMsgDlgType.mtInformation, [TMsgDlgBtn.mbOK], 0);
    except on E: exception do
      MessageDlg(E.ClassName + #13 + 'Error raised for flag [' + TMenuItem(Sender).Caption + '], with message:' + #13 + E.Message, TMsgDlgType.mtError, [TMsgDlgBtn.mbOK], 0);
    end;
  end;
end;

procedure TFRM_Tree.PopUpAgreementsPopup(Sender: TObject);
var
  Item: TMenuItem;
begin
  DeleteAgreementfromCRM1.Visible := IsValidFinancial;
  SendfurtherNOI1.Visible := CanSendFurtherNOI;

  SendExtracareletter1.Visible := IsPrepayAndLive;

  // check if the sub-menus are already created
  if (mnuRaiseFlag.Count > 0) and (mnuRaiseFlag.Items[0].Tag = 0) then
  begin
    mnuRaiseFlag.Remove(mnuRaiseFlag.Items[0]);

    try
      with GeneralQuery do
      begin
        Close;
        SQL.Clear;
        DeleteVariables;
        SQL.Add('SELECT UO.FLAG_ID, RT.DESCRIPTION AS FLAG_MENU_NAME');
        SQL.Add('FROM ENQUIRY.AGREEMENT_FLAGS_USER_OPTIONS UO');
        SQL.Add('INNER JOIN ENQUIRY.AGREEMENT_FLAGS AF ON AF.REQUEST_TYPE = UO.FLAG_ID');
        SQL.Add('INNER JOIN ENQUIRY.REQUEST_TYPE RT ON RT.ID = UO.FLAG_ID');
        SQL.Add('WHERE AF.CAN_USER_SELECT = ' + QuotedStr('Y'));
        SQL.Add('AND UO.USER_ID = :USERID');
        SQL.Add('ORDER BY UO.FLAG_ID');
        DeclareAndSet('USERID', otString, FRM_Login.edtUsername.Text);
        Open;

        while not Eof do
        begin
          Item := TMenuItem.Create(mnuRaiseFlag);
          Item.Tag := FieldByName('FLAG_ID').AsInteger;
          Item.Caption := FieldByName('FLAG_MENU_NAME').AsString;
          Item.ImageIndex := mnuRaiseFlag.ImageIndex;
          Item.OnClick := RaiseFlag;
          mnuRaiseFlag.Add(Item);
          Next;
        end;
      end;
    finally
      GeneralQuery.Close;
      mnuRaiseFlag.Visible := (mnuRaiseFlag.Count > 0);
    end;
  end;
end;

procedure TFRM_Tree.PopUpCustPopup(Sender: TObject);
const
  strSQL = 'SELECT C.CUSTOMER_TYPE_ID FROM CRM.CUSTOMER C WHERE C.CUSTOMER_ID = :p_customer_id';
begin
  xnode:=treeview1.FocusedNode;
  TreeView1.Expanded[xnode] := true;

  //BALDINOL - PT-731 - validation to not allow user to access the screen if the customer is not domestic
  nodeData := treeview1.GetNodeData(xnode);
  nodeData.D_Cust_Type :=
    gSqlUtil.SelectQueryInteger(strSQL, ['p_customer_id', otLong, StrToInt64(nodeData.D_Customer_Id)]);
  mniWinterWarmerFinancialAssistancePayment.Visible := (nodeData.D_Cust_Type = 1) and FundAvailab;
  /////////////////////////////////////////////////////////////////////////////////////////////////////////

  if Not TreeView1.Expanded[xnode] then
  begin
    TreeView1.Expanded[xnode] := false;
    Abort;
  end;
end;

function TFRM_Tree.FundAvailab: Boolean;
var
  vYear: integer;
  vSQL: String;
  vReturn: Variant;
begin
  vReturn := EmptyStr;
  vSQL := 'SELECT MAX(WYEAR) WYEAR FROM CRM.WHD_PROCESS';
  vYear := gSqlUtil.SelectQueryInteger(vSQL);

  vSQL := 'GES.pk_winter_warmer.pr_fund_availability(:p_scheme_year, :p_return_message)';
  try
    gSqlUtil.ExecProc(vSQL, TRANSACTION_NO,
          ['p_scheme_year',    otString,  pdInput,  vYear,
           'p_return_message', otString, pdOutput, @vReturn]);
  except
    On E: Exception do
      MessageDlg(E.message, mtError, [mbOk], 0);
  end;
  Result := Trim(vReturn) = 'Y';
end;

procedure TFRM_Tree.PopUpSmets_DCCPopup(Sender: TObject);
var
  bSpanObj: string;
  bMeterLoss: boolean;
begin
  xnode := treeview1.FocusedNode;
  nodeData := treeview1.GetNodeData(xnode);
  bSpanObj :=  nodedata.D_SPAN;

  bMeterLoss := isMeterLost(bSpanObj);
  Self.smets_vend_DCC.enabled := bMeterLoss;
  Self.smets_debt_DCC.enabled := bMeterLoss;
  Self.smets_loan_DCC.enabled := bMeterLoss;
end;

procedure TFRM_Tree.PowerPayEligibilityClick(Sender: TObject);
const
  cEligible = 'Eligible';
var
  vIsEligible: Boolean;
  vEligibilityDset: TOracleDataSet;
  vEligExceptionsSL: TStringList;
  vPowerPayEligibilityDialog: TAdvTaskDialogEx;
begin
  vEligExceptionsSL := TStringList.Create;
  vPowerPayEligibilityDialog := TAdvTaskDialogEx.Create(nil);
  vPowerPayEligibilityDialog.Title := 'Power Pay Eligibility';

  try
    vEligibilityDset := gSQLUtil.CreateCursor(
                          'crm.pk_smart_pay_eligibility.pr_crm_eligibility(:p_customer_id, :p_exceptions)',
                          TRANSACTION_NO,
                          ['p_customer_id', otLong,   CustId,
                           'p_exceptions',  otCursor, Null]);
    vIsEligible := vEligibilityDset.FieldByName('title').AsString = cEligible;

    if vIsEligible then
    begin
      vPowerPayEligibilityDialog.Instruction := 'This customer is eligible for Power Pay';
      vPowerPayEligibilityDialog.Icon := tiInformation;
    end
    else
    begin
      vEligibilityDset.First;
      while not vEligibilityDset.Eof do
      begin
        if (vEligExceptionsSL.Text = '') then
          vEligExceptionsSL.Add(UpperCase(vEligibilityDset.FieldByName('title').AsString))
        else
          vEligExceptionsSL.Add(#13 + UpperCase(vEligibilityDset.FieldByName('title').AsString));

        vEligExceptionsSL.Add(vEligibilityDset.FieldByName('message').AsString);
        vEligibilityDset.Next;
      end;
      vPowerPayEligibilityDialog.Instruction := 'This customer is not eligible for Power Pay';
      vPowerPayEligibilityDialog.Icon := tiError;
      vPowerPayEligibilityDialog.ExpandControlText := 'Click to hide the details';
      vPowerPayEligibilityDialog.CollapsControlText := 'Click to see the details';
      vPowerPayEligibilityDialog.ExpandedText := vEligExceptionsSL.Text;
    end;
    vPowerPayEligibilityDialog.Execute;
  finally
    vEligExceptionsSL.Free;
    FreeAndNil(vEligibilityDset);
    FreeAndNil(vPowerPayEligibilityDialog);
  end;
end;

procedure TFRM_Tree.PopUpElectricPopup(Sender: TObject);
var
  dccResult : Boolean;
  nodeData  : PMyRec;
begin

  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);

  if not Assigned(nodeData) then
    exit;

  dccResult := getMeterDCC(nodeData.D_Span);

  if not dccResult then
  begin
    S_COSLOSSE.Visible := true;
    S_COSLOSSE.enabled := true;
  end;

  if dccResult then
  begin
    S_COSLOSSE.Visible := false;
    S_COSLOSSE.enabled := false;
  end;
end;

procedure TFRM_Tree.PopUpGasPopup(Sender: TObject);
var
  dccResult : Boolean;
  nodeData  : PMyRec;
begin

  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);

  if not Assigned(nodeData) then
    exit;

  dccResult := getMeterDCC(nodeData.D_Span);

  if not dccResult then
  begin
    S_COSLOSSG.Visible := true;
    S_COSLOSSG.enabled := true;
  end;

  if dccResult then
  begin
    S_COSLOSSG.Visible := false;
    S_COSLOSSG.enabled := false;
  end;
end;

procedure TFRM_Tree.PopUpPremisePopup(Sender: TObject);
var
  sqlText         : string;
  vDeviceS2Enroll : Variant;

begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  try
    //Query to Check if Device Type S2 enrolled
    sqlText := 'ODS.pk_crmui_metering.pr_is_dcc_meter(';
    sqlText := sqlText + ':p_customer_id,';
    sqlText := sqlText + ':p_return)';

    gSqlUtil.ExecProc(sqlText, TRANSACTION_YES,
    ['p_customer_id'     , otString, pdInput , TreeData.D_Customer_Id,
     'p_return', otString, pdOutput, @vDeviceS2Enroll]);

     if vDeviceS2Enroll = 'Y' then CheckComms.Visible := true
     else CheckComms.Visible := false;

  except
    on e: Exception do
    begin
      raise Exception.Create('Error: Device type retrieval');
    end;

  end;

end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.PPMIDMessages1Click(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Message_History_Dcc.StartModal(Self, nodeData.D_Span);
end;

procedure TFRM_Tree.N021Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 DoBillPCCHANGE('2',mpan);
end;

procedure TFRM_Tree.N031Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 DoBillPCCHANGE('3',mpan);
end;

procedure TFRM_Tree.N041Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 DoBillPCCHANGE('4',mpan);
end;


procedure TFRM_Tree.doBillPCchange(PC,span:string);
Var
ssd:string;
begin
 frm_date.efd.date:=now;
 frm_date.showmodal;
 if frm_Date.tag=0 then exit;
 ssd:=frm_date.efd.text;
 if Messagedlg('Are you sure you wish to change the Billing Profile type to 0'+PC+' for MPAN '+mpan+#13+'Effective from '+ssd,mtconfirmation,[mbyes,mbno],0)<>mryes then exit;
 // do change
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from billing.mpan_bill_profile where mpancore='''+mpan+'''');
  sql.add('and effective_from>=to_date('''+ssd+''',''DD/MM/YYYY'')');
  execute;
  close;
  sql.clear;
  sql.add('insert into billing.mpan_bill_profile values ('''+mpan+''','+pc+',to_date('''+ssd+''',''DD/MM/YYYY''),'''+USERID+''',sysdate)');
  execute
 End;
 try
  treeview1.Expanded[xnode]:=false;
  treeview1.Expanded[xnode]:=true;
 except
  Messagedlg('Refresh Tree To view Changes',mtinformation,[mbok],0);
 end;
end;

procedure TFRM_Tree.N051Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 DoBillPCCHANGE('5',mpan);
end;

procedure TFRM_Tree.N061Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 DoBillPCCHANGE('6',mpan);
end;

procedure TFRM_Tree.N071Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 DoBillPCCHANGE('7',mpan);
end;

procedure TFRM_Tree.N081Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 mpan:=TreeData.D_SPAN;
 DoBillPCCHANGE('8',mpan);
end;

procedure TFRM_Tree.N1010Click(Sender: TObject);
begin
TREELIMIT:=10;
end;

procedure TFRM_Tree.N115Click(Sender: TObject);
begin
TREELIMIT:=1;
end;

procedure TFRM_Tree.N510Click(Sender: TObject);
begin
TREELIMIT:=5;
end;

procedure TFRM_Tree.AddAdditionalCharges1Click(Sender: TObject);
var
agreement_id:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Agreement_id:=TreeData.D_agreement_id;

 FRM_ONE_OFF_CHARGE.ShowTheseDetails('0','0',Agreement_id,'O');
 FRM_ONE_OFF_CHARGE.showmodal;
end;

procedure TFRM_Tree.DeleteAgreementfromCRM1Click(Sender: TObject);
Var
AGID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AGID:=TreeData.D_agreement_ID;
 if frm_common.deletecustomer(AGid,'2')=true then
 Begin
  try
   treeview1.Expanded[xnode.parent]:=false;
   treeview1.Expanded[xnode.parent]:=true;
  except
  end;
 end;
end;

function TFRM_Tree.IsSmartPay(aIdentifier: string): Boolean;
var
  SQLText: string;
  TempOraDS: TOracleDataSet;
begin
  SQLText := 'SELECT crm.pk_smart_pay_workflow.fn_is_smart_pay( ' +
             'p_identifier => ' + QuotedStr(aIdentifier) + ') ' +
             'FROM dual';
  Result := gSQLUtil.SelectQueryString(SQLText, []) = 'Y';
end;

procedure TFRM_Tree.IssueD0190Key1Click(Sender: TObject);
var
mpan:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 mpan:=TreeData.D_SPAN;
 Application.CreateForm(TFRM_D0190, FRM_D0190);
 try
  frm_d0190.tag:=0;
  frm_D0190.statusbar1.panels[1].text:='Issue Customer Key (MPAN'+mpan+')';
  FRM_D0190.Run_mpan_Procedure(mpan);
  if frm_d0190.MPANLIST.RecordCount<>0 then
  Begin
   frm_d0190.MPANlookup.KeyValue:=mpan;
   frm_d0190.ShowModal
  end
  else Messagedlg('You cannot Generate a Key request for this MPAN.',mtwarning,[mbok],0);
 finally
  FRM_D0190.release;
 end;

end;

procedure TFRM_Tree.IssueReplacementQuantumCard1Click(Sender: TObject);
var
mprn,agid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 mprn:=TreeData.D_SPAN;
 agid:=TreeData.D_agreement_id;
 Application.CreateForm(TFrm_GAS_QUANTUM, Frm_GAS_QUANTUM);
 try
  frm_gas_quantum.tag:=1;
  frm_gas_quantum.DoQueryMPRN(MPRN,agid);
  frm_GAS_QUANTUM.showmodal;
 finally
  frm_GAS_QUANTUM.release;
 end;
end;

procedure TFRM_Tree.G_RELClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 DoNotObject(TreeData.D_SPAN,Treedata.D_AGREEMENT_ID);
end;

procedure TFRM_TREE.DoNotObject(Span,AGid:string);
var
cid:string;
begin
 if messagedlg('Are you sure you wish to add a marker to NOT Object'+#13+
               'to any losses that come in for SPAN '+SPAN+'?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;

 //if frm_common.authoritycheck=false then exit;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('Insert into crm.spans_do_not_object values('''+SPAN+''','''+USERID+''',sysdate,null)');
  try
   execute;
   close;
   sql.clear;
   sql.add('Insert into enquiry.enquiries values('+SPAN+','''+USERID+''',sysdate,'+inttostr(14)+','+inttostr(50)+',null,''Accept LOSS. Do NOT Object'',null,''Y'',null,''SYSTEM'',sysdate,NULL,sysdate,null,'+cid+',null,'+frm_common.NextNoteId+',''X'',''3'')');
     if messagedlg('Do you wish to add an Accept Loss note to the account.?',mtconfirmation,[mbyes,mbno],0)=mryes then execute;
  except
   Messagedlg('There was a problem actioning this request. Maybe a Release has already been requested.',mterror,[mbok],0);
   exit;
  end;
 End;
 frm_login.mainsession.commit;
 try
  treeview1.expanded[xnode.parent]:=false;
  treeview1.expanded[xnode.parent]:=true;
 except
  Messagedlg('Refresh Tree to view changes',Mtinformation,[MBOK],0);
 end;
end;

procedure TFRM_Tree.E_RELClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 DoNotObject(TreeData.D_SPAN,Treedata.D_AGREEMENT_ID);
end;

procedure TFRM_Tree.E_RELOClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 RemoveRelease(TreeData.D_SPAN,Treedata.D_AGREEMENT_ID);
end;

procedure TFRM_Tree.G_RELOClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 RemoveRelease(TreeData.D_SPAN,Treedata.D_AGREEMENT_ID);
end;


procedure TFRM_TREE.RemoveRelease(Span,custid:string);
begin
 if messagedlg('Are you sure you wish to remove the DO NOT OBJECT Marker'+#13+
               ' for SPAN '+SPAN+'?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;


 with main_data_module.tempquery do
 Begin
  deletevariables;
  close;
  sql.clear;
  sql.add('select * from crm.spans_do_not_object where span='''+span+'''');
  open;
 End;
 if main_data_module.tempquery.recordcount=0 then
 Begin
  Messagedlg('No Release Marker exists for this SPAN.',mtwarning,[mbok],0);
  exit;
 End;

 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from crm.spans_do_not_object where SPAN='''+SPAN+'''');
  execute;
 End;
 frm_login.mainsession.commit;
 Messagedlg('DO NOT OBJECT Marker has been removed for this SPAN.',mtinformation,[mbok],0);
 try
  treeview1.expanded[xnode.parent]:=false;
  treeview1.expanded[xnode.parent]:=true;
 except
  Messagedlg('Refresh Tree to view changes',Mtinformation,[MBOK],0);
 end;
end;

procedure TFRM_Tree.RemoveSuppressDDMarker1Click(Sender: TObject);
var
agreement_id,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 agreement_id:=TreeData.D_Agreement_id;

 If Messagedlg('Are you sure you wish to remove this DD marker?'+#13+
               'Any future catchup DDs will be applied to the account.',mtconfirmation,[mbyes,mbno],0)<mryes then exit;


 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from crm.dd_suppress_catchups where agreement_id='''+Agreement_id+'''');
  execute;
 End;
  Cid:=frm_common.GetCustomerIdfromAgreementid(Agreement_id);

 FRM_Common.SetAudit(cid,agreement_id,'','Suppress Catch Up DD marker Removed');

 frm_login.mainsession.commit;
 Messagedlg('Suppress Catch Up DD Marker has been removed.',mtinformation,[mbok],0);
 treeview1.expanded[xnode.parent]:=false;
 treeview1.expanded[xnode.parent]:=true;
end;

procedure TFRM_Tree.SuppressCatchUpDDs1Click(Sender: TObject);
var
agreement_id,CID:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Agreement_ID:=TreeData.D_Agreement_ID;

 If Messagedlg('Are you sure you wish to add a marker to this account to Suppress Catchup DDs?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;


 with Main_data_module.tempquery do
 Begin
  deletevariables;
  close;
  sql.Clear;
  sql.add('select * from crm.dd_suppress_catchups where agreement_id='''+Agreement_id+'''');
  open;
 End;
 if Main_data_module.tempquery.recordcount<>0 then
 Begin
  Messagedlg('A Marker already exists on this account.',mtwarning,[mbok],0);
  exit;
 End;

 with main_data_module.updatequery do
 begin
  close;
  sql.clear;
  sql.add('insert into crm.dd_suppress_catchups values('+agreement_id+',''Marker Added'','''+uppercase(userid)+''',sysdate)');
  execute;
 end;
  Cid:=frm_common.GetCustomerIdfromAgreementid(Agreement_id);

 FRM_Common.SetAudit(CID,agreement_id,'','Suppress Catch Up DD marker Added');
 frm_login.mainsession.commit;
 Messagedlg('Suppress Catch Up DD Marker Added.',mtinformation,[mbok],0);
 try
  treeview1.expanded[xnode]:=false;
  treeview1.expanded[xnode]:=true;
 except
  Messagedlg('Refresh Tree to view changes',Mtinformation,[MBOK],0);
 end;
end;

procedure TFRM_Tree.Custom_SPANClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 FRM_Main.SearchForSpan(TreeData.D_SPAN,0);
end;

procedure TFRM_Tree.BuildCustomMeterNode(MPANNODE:PVirtualnode);
var
mpan,ENDDATE,TPR,metertype,enstatus,ssc,sscdesc,daterem,nsr,maketype:string;
//MyRecPtr: PMyRec;
Begin
 nodeData := treeview1.GetNodeData(mpannode);

   // Check for Meter Technical Details
 mpan:=nodedata.D_SPAN;


 ENDDATE:=nodedata.D_SPANEND;
 if ENDDATE='' then ENDDATE:='10/10/2060';
 //mpannode:=treeview1.selected;

  // Check if Customer Has Requested Single Rate Billing
  m_single.caption:='Default to Single Rate Billing';
  with main_data_module.generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('MPAN', otString);
   sql.clear;
   sql.add('select * from crm.mpans_single_rate_billing');
   sql.add('where mpancore=:mpan');
   sql.add('Order by effective_from desc');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
  End;
  if main_data_module.generalquery.recordcount<>0 then
  Begin
   m_single.caption:='Remove Single Rate Billing';
   efsdmsmtd:=main_data_module.generalquery.fields[1].text;

   MeterConfigNode:=Treeview1.Addchild(mpannode);
   nodeData := Treeview1.GetNodeData(MeterConfigNode);
   NodeData.caption := 'Customer Requests Single Rate Billing from '+efsdmsmtd;
   NodeData.fontcolor:=clpurple;
   NodeData.fontBold:=true;
   NodeData.index:=140;

   {MeterConfigNode:=Treeview1.items.AddChild(mpannode,'Customer Requests Single Rate Billing from '+efsdmsmtd);
   MeterConfigNode.font.color:=clpurple;
   MeterConfigNode.font.style:=[fsbold];
   MeterConfigNode.imageindex:=140;
   MeterConfigNode.selectedindex:=140; }
  End;

    // Check if Span has SSC change
  with main_data_module.generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('MPAN', otString);
   sql.clear;
   sql.add('select s.*,t.description from crm.span_ssc_changes s,crm.span_type t');
   sql.add('where s.span=:MPAN');
   sql.add('and s.span_type=t.span_type_id');
   sql.add('Order by s.effective_from desc');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
  End;
  while not main_data_module.generalquery.eof do
  Begin
   efsdmsmtd:=main_data_module.generalquery.fields[2].text;

   MeterConfigNode:=Treeview1.Addchild(mpannode);
   nodeData := Treeview1.GetNodeData(MeterConfigNode);
   NodeData.caption := efsdmsmtd+' Billing changed to '+main_data_module.generalquery.fields[3].text;
   NodeData.fontcolor:=clpurple;
   NodeData.fontBold:=true;
   NodeData.index:=26;

  { MeterConfigNode:=Treeview1.items.AddChild(mpannode,efsdmsmtd+' Billing changed to '+main_data_module.generalquery.fields[3].text);
   MeterConfigNode.font.color:=clpurple;
   MeterConfigNode.font.style:=[fsbold];
   MeterConfigNode.imageindex:=26;
   MeterConfigNode.selectedindex:=26;}
   main_data_module.generalquery.next;
  End;

    // Check for Billing Profile Change
  with main_data_module.generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('MPAN', otString);
   sql.clear;
   sql.add('select * from billing.mpan_bill_profile');
   sql.add('where mpancore=:MPAN');
   sql.add('Order by effective_from desc');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
  End;
  while not main_data_module.generalquery.eof do
  Begin
   efsdmsmtd:=main_data_module.generalquery.fields[2].text;

   pcConfigNode:=Treeview1.Addchild(mpannode);
   nodeData := Treeview1.GetNodeData(pcConfigNode);
   NodeData.caption := efsdmsmtd+' -  Billing Profile Set to 0'+main_data_module.generalquery.fields[1].text;
   NodeData.fontcolor:=clpurple;
   NodeData.fontBold:=true;
   NodeData.index:=44;

   {pcConfigNode:=Treeview1.items.AddChild(mpannode,efsdmsmtd+' -  Billing Profile Set to 0'+main_data_module.generalquery.fields[1].text);
   pcConfigNode.font.color:=clpurple;
   pcConfigNode.font.style:=[fsbold];
   pcConfigNode.imageindex:=44;
   pcConfigNode.selectedindex:=44; }
   main_data_module.generalquery.next;
  End;


  with mtdscustom do
  begin
   close;
   setvariable('ENDDATE',ENDDATE);
   setvariable('MPAN',MPAN);
   open;
  end;

  // Only Do This Block If Meter Records Exist
  if MTDscustom.recordcount<>0 then
  Begin
   msid:='LEEOK';
   oldefsdmsmtd:='lee';
   oldmeterid:='';
   oldregister:='';
   while not MTDscustom.eof do
   Begin
    // Build Tree Of MTDS
   // Create Subtree of Effective From Dates
    efsdmsmtd:=mtdscustom.fields[1].text;
    MeterType:=mtdscustom.fields[23].text;
    EnStatus:=mtdscustom.fields[2].text;
    SSC:=mtdscustom.fields[5].text;
    SSCDesc:=mtdscustom.fields[6].text;
    DateRem:=mtdscustom.fields[37].text;
    NSR:=mtdscustom.fields[36].text;
    Mregister:=mtdscustom.fields[26].text;
    meterid:=mtdscustom.fields[10].text;
    maketype:='';
    if mtdscustom.fields[14].text<>'' then maketype:=copy(mtdscustom.fields[14].text,1,3);
    if efsdmsmtd<>oldefsdmsmtd then
    Begin
     if oldefsdmsmtd='lee' then config:='Current Configuration'
     else config:='Previous Configuration';
     if ((config='Previous Configuration') and (historic1.checked=true)) or (config<>'Previous Configuration') then
     Begin
      if (efsdmsmtd<>'') and (MeterType='') and (SSC='') then
      Begin
       MeterConfigNode:=Treeview1.Addchild(mpannode);
       nodeData := Treeview1.GetNodeData(MeterConfigNode);
       NodeData.caption := config+' - '+efsdmsmtd+' - MOP Reports No Meters on this Supply. Energisation Status = '+EnStatus;
       NodeData.fontcolor:=clred;
       NodeData.fontBold:=true;
       NodeData.index:=24;


      { MeterConfigNode:=Treeview1.items.AddChild(mpannode,config+' - '+efsdmsmtd+' - MOP Reports No Meters on this Supply. Energisation Status = '+EnStatus);
       MeterConfigNode.font.color:=clred;
       MeterConfigNode.font.style:=[fsbold];
       MeterConfigNode.imageindex:=24;
       MeterConfigNode.selectedindex:=24; }
      end
      else
      if MeterType='' then
      Begin
       MeterConfigNode:=Treeview1.Addchild(mpannode);
       nodeData := Treeview1.GetNodeData(MeterConfigNode);
       NodeData.caption := 'Metering Configuration not Known (Missing / Incomplete meter technical details)';
       NodeData.fontcolor:=clred;
       NodeData.fontBold:=true;
       NodeData.index:=26;

       {MeterConfigNode:=Treeview1.items.AddChild(mpannode,'Metering Configuration not Known (Missing / Incomplete meter technical details)');
       MeterConfigNode.font.color:=clred;
       MeterConfigNode.font.style:=[fsbold];
       MeterConfigNode.imageindex:=26;
       MeterConfigNode.selectedindex:=26;}
      end;
      if (SSC<>'') then
      Begin
       MeterConfigNode:=Treeview1.Addchild(mpannode);
       nodeData := Treeview1.GetNodeData(MeterConfigNode);
       desc :=config+' - '+efsdmsmtd+' - SSC ID ('+SSC+') - '+SSCDESC;
      // MeterConfigNode:=Treeview1.items.AddChild(mpannode,config+' - '+efsdmsmtd+' - SSC ID ('+SSC+') - '+SSCDESC);

       if (MTDscustom.fields[4].text='') and (MTDscustom.fields[39].text='') and (MTDscustom.fields[13].text<>'Not Known') then
       begin
        desc:=desc+#10+'(*Warning: '+MTDscustom.fields[13].text+' *)';
       end;
       NodeData.caption:=desc;

       if config<>'Previous Configuration' then
       Begin
        NodeData.fontcolor:=clgreen;
        NodeData.fontBold:=true;
       end;
       NodeData.index:=27;
       oldmeterid:='lee';
      end;
     end;
    end; // End Of Configuration Date
     // Do Meters

    if ((config='Previous Configuration') and (historic1.checked=true)) or (config<>'Previous Configuration') then
    Begin
     if meterid<>oldmeterid then
     Begin
      oldregister:='Lee';
      mtype:='';
      if MeterType='' then mtype:='*NO*';
      if (efsdmsmtd<>'') and (metertype='') then mtype:='*NO*';

      if mtype<>'*NO*' then
      Begin
       {if DateRem<>'' then Dateremoved:='    (Date Removed='+DateRem+')'
       else}
       dateremoved:='';
       if (NSR='Y') and (v_non.checked=false) then
       Begin
       //
       end
       else
       Begin
        MeterNode:=Treeview1.Addchild(MeterConfigNode);
        nodeData := Treeview1.GetNodeData(MeterNode);
        NodeData.caption :='NHH Meter ID-'+Meterid+dateremoved;
        nodedata.D_SPAN :=mpan;
        nodedata.M_METERID :=Meterid;

        nodedata.Metertype := MeterType;

        // What is Service Type for SUb Meters?
        nodedata.M_SERVICE :='0';
        if copy(mpan,1,2)='98' then nodedata.M_SERVICE :='1';

        // for customer meters need to know if gas or elec


        //MeterNode:=Treeview1.items.AddChild(MeterConfigNode,'NHH Meter ID-'+Meterid+dateremoved);
        nodedata.index:=17; // NHH Credit Meter
         if MeterType='N' then
        Begin
         nodedata.caption:='NHH Credit Meter ID-'+MeterID+dateremoved;
        end;
        if MeterType='S' then
        Begin
         nodedata.caption:='NHH Smart Card Meter ID-'+MeterID+dateremoved;
         nodedata.index:=22; // NHH Smart Card meter
        end;

        if (MeterType='S') and (mtdscustom.fields[20].text='R') then
        Begin
         nodedata.caption:='Remote Read Smart Meter ID-'+MeterID+dateremoved;
         nodedata.index:=205; // NHH Smart Card meter
        end;

        if (copy(MeterType,1,4)='RCAM') or (MakeType='PRI') or (copy(MeterType,1,3)='NSS') then
        Begin
         nodedata.caption:='Smart Meter ID-'+MeterID+dateremoved;
         nodedata.index:=314;
        end;

        if MeterType = 'S1EA' then
        begin
          nodedata.caption:='SMETS1 E&A Smart Meter ID-' + MeterID + dateremoved;
          nodedata.index := 313;
        end
        else
        if (copy(MeterType,1,2)='S1') then
        Begin
         ShowSmetsMeterCOmmsSupplier(MeterNode,MPAN,METERID,'0',dateremoved,'X');
        end;

         if (copy(MeterType,1,2)='S2')  then
        Begin
         nodedata.caption:='SMETS 2 Meter ID-'+MeterID+dateremoved;
         nodedata.index:=205; // SMETS 2 ICON
        end;

        if MeterType='T' then
        Begin
         nodedata.caption:='NHH Token Meter ID-'+MeterID+dateremoved;
         nodedata.index:=23; // NHH token Meter
        end;
        if MeterType='K' then
        Begin
         nodedata.caption:='NHH Key Meter ID-'+MeterID+dateremoved;
         nodedata.index:=21; // NHH key Meter
        end;
        if MeterType='H' then
        Begin
         nodedata.caption:='HH Meter ID-'+MeterID+dateremoved;
         nodedata.index:=9; // HH Meter
        end;
        if (copy(mpan,1,2)='95') or (copy(mpan,1,2)='96') or (copy(mpan,1,2)='93') then
        Begin
         nodedata.caption:='Water Meter ID-'+MeterID+dateremoved;
         nodedata.index:=229; // Water Meter
        end;

        if dateremoved<>'' then
        Begin
         nodedata.index:=24; // Removed Meter
        end;
       end;
      end
      else
      Begin
       // A Meter from D0149 but No D0150
       MeterNode:=Treeview1.Addchild(MeterConfigNode);
       nodeData := Treeview1.GetNodeData(MeterNode);
       NodeData.caption :='NHH Meter ID-'+Meterid+' (Missing D0150)';
       nodedata.D_SPAN :=mpan;
      nodedata.M_METERID :=Meterid;
      nodedata.M_SERVICE :='0';
       //MeterNode:=Treeview1.items.AddChild(MeterConfigNode,'NHH Meter ID-'+Meterid+' (Missing D0150)');
       nodedata.index:=26; // NHH Credit Meter
       nodedata.fontcolor:=clred;
      End;
     end; // Change Of Meter
    end;  // End Of Configuration Block

        // Have any Meters been Removed?
    if (efsdmsmtd<>'') and (MeterType='') and (SSC='') and (daterem<>'')then
    Begin
     MeterNode:=Treeview1.Addchild(MeterConfigNode);
     nodeData := Treeview1.GetNodeData(MeterNode);
     NodeData.caption :='Removed Meter -'+mtdscustom.fields[39].text+'. Date Removed ('+DateRem+')';
     nodedata.D_SPAN :=mpan;
      nodedata.M_METERID :=Meterid;
      nodedata.M_SERVICE :='0';
     //MeterNode:=Treeview1.items.AddChild(MeterConfigNode,'Removed Meter -'+mtdscustom.fields[39].text+'. Date Removed ('+DateRem+')');
     nodedata.index:=24; // Removed Meter
    end;


     if ((config='Previous Configuration') and (historic1.checked=true)) or (config<>'Previous Configuration') then
    Begin
     if mregister<>oldregister then
     Begin
      if (NSR='Y') and (v_non.checked=false) then
      begin
      //
      end
      else
      Begin
       TPR:=mtdscustom.fields[33].text;  // TPR ID e.g. 00001
       if tpr<>'' then
       Begin
        MeterRegisterNode:=Treeview1.Addchild(MeterNode);
        nodeData := Treeview1.GetNodeData(MeterRegisterNode);
        NodeData.caption :=mregister+' - TPR '+TPR+' ('+mtdscustom.fields[38].text+') - '+mtdscustom.fields[30].text;

       // MeterRegisterNode:=Treeview1.items.AddChild(MeterNode,mregister+' - TPR '+TPR+' ('+mtdscustom.fields[38].text+') - '+mtdscustom.fields[30].text);
        NodeData.fontcolor:=clblack;
        if mtdscustom.fields[30].text='' then NodeData.fontcolor:=clred;
       end
       else
       Begin
        if MeterID<>'' then
        Begin
         MeterRegisterNode:=Treeview1.Addchild(MeterNode);
         nodeData := Treeview1.GetNodeData(MeterRegisterNode);
         NodeData.caption :=mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * ) - '+mtdscustom.fields[30].text;

         //MeterRegisterNode:=Treeview1.items.AddChild(MeterNode,mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * ) - '+mtdscustom.fields[30].text);
         NodeData.fontcolor:=clred;
        end;
       end;
       NodeData.Index:=Frm_Common.GetRegisterPic(mtdscustom.fields[33].text);
       //MeterRegisterNode.selectedindex:=MeterRegisterNode.imageindex;

       nodedata.D_SPAN :=mpan;
       nodedata.M_EFSDMSMTD :=efsdmsmtd;
       nodedata.M_METERID :=Meterid;
       nodedata.M_REGISTERID :=mregister;
       if MeterType='H' then nodedata.M_HH_REGISTER :='H'
       else nodedata.M_HH_REGISTER :='N';
       //MeterRegisterNode.data:=MyRecPtr;

       if mtdscustom.fields[29].text='RI' then
       Begin
        nodedata.index:=44;
       end;
       if mtdscustom.fields[29].text='AE' then
       Begin
        nodedata.index:=234;
       end;
       if NSR='Y' then
       Begin
        // non settlement register
       nodedata.Fontcolor:=clred;
        if v_non.checked then
        Begin
         nodedata.Index:=Frm_Common.GetNonRegisterPic(mtdscustom.fields[33].text);


        if mtdscustom.fields[29].text='RI' then
         Begin
          nodedata.index:=45;
         end;
        end;
       end;
      end;
     end; // End Of Add Register
    end;  // End Of Configuration Block
    oldefsdmsmtd:=efsdmsmtd;
    oldmeterid:=meterid;
    oldregister:=mregister;
    mtdscustom.next;
   end;
  end; // End Of Meter Strucutre Tree

 ////////// Build Tree Of Orphaned Registers //////////////
 With generalquery do
  Begin
   close;
   DeleteVariables;
   DeclareVariable('MPAN', otString);
   sql.clear;
   sql.add('select distinct R.mpancore,R.meterid,R.registerid,r.current_status');
   sql.add('from edmgr.readings R,edmgr.d0149a D, edmgr.d0150_293 M');
   sql.add('where');
   sql.add('r.MPANCORE=:MPAN');
   sql.add('and r.mpancore=d.mpancore (+)');
   sql.add('and');
   sql.add('r.meterid=d.meterid (+)');
   sql.add('and');
   sql.add('r.registerid=d.registerid (+)');
   sql.add('and');
   sql.add('r.mpancore=M.mpancore (+)');
   sql.add('and');
   sql.add('r.meterid=M.meterid (+)');
   sql.add('and');
   sql.add('r.registerid=M.meter_register_id (+)');
   sql.add('and d.mpancore is null');
   sql.add('and m.mpancore is null');
   sql.add('and r.current_status<>''D''');
  // sql.add('and r.rdngtype<>''W''');
   sql.add('order by R.meterid,R.registerid');
   setvariable('MPAN',mpan);
   open;
   deletevariables;
  end;
   // Only Do This Block If Meter Records Exist
  if generalquery.recordcount<>0 then
  Begin
   oldmeterid:='OldMeter';
   msid:='LEEOK';
   oldefsdmsmtd:='lee';

   MeterConfigNode:=Treeview1.Addchild(mpannode);
   nodeData := Treeview1.GetNodeData(MeterConfigNode);
   NodeData.caption :='Register Readings (Orphans - No Mapping Details)';


   //MeterConfigNode:=Treeview1.items.AddChild(mpannode,'Register Readings (Orphans - No Mapping Details)');
   NodeData.fontcolor:=clred;
   NodeData.fontBold:=true;
   NodeData.index:=26;
   //NodeData.selectedindex:=26;

   while not generalquery.eof do
   Begin
     // Do Meters
    Begin
     meterid:=generalquery.fields[1].text;
     if meterid<>oldmeterid then
     Begin
      oldregister:='Lee';
      mtype:='';
       Begin
        MeterNode:=Treeview1.Addchild(MeterConfigNode);
        nodeData := Treeview1.GetNodeData(MeterNode);
        NodeData.caption :='NHH Meter ID-'+Meterid+dateremoved;
        nodedata.D_SPAN :=mpan;
      nodedata.M_METERID :=Meterid;
      nodedata.M_SERVICE :='0';
        //MeterNode:=Treeview1.items.AddChild(MeterConfigNode,'NHH Meter ID-'+Meterid+dateremoved);
        nodedata.index:=26; // NHH Meter
        end;
      end; // Change Of Meter
    end;  // End Of Configuration Block
    Begin
     Mregister:=generalquery.fields[2].text;
     if mregister<>oldregister then
     Begin
      Begin
       Begin
        MeterRegisterNode:=Treeview1.Addchild(MeterNode);
        nodeData := Treeview1.GetNodeData(MeterRegisterNode);
        NodeData.caption :=mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * )';

        //MeterRegisterNode:=Treeview1.items.AddChild(MeterNode,mregister+' - TPR '+' ( * MISSING D0149 MAPPING DETAILS * )');
        NodeData.fontcolor:=clred;
       end;
       NodeData.index:=29;

       nodedata.D_SPAN :=mpan;
       nodedata.M_EFSDMSMTD :=efsdmsmtd;
       nodedata.M_METERID :=MeterID;
       nodedata.M_REGISTERID :=mregister;
       //MeterRegisterNode.data:=MyRecPtr;

      end;
     end; // End Of Add Register
    end;  // End Of Configuration Block
    oldmeterid:=meterid;
    oldregister:=mregister;
    generalquery.next;
   end;
  end; // End Of Meter Strucutre Tree
end;

procedure TFRM_Tree.BuildS1EnrolledNode(mpxn:string;spannode:PVirtualNode);
var
 DCCEnrolledNode : PVirtualNode;
 NodeCaption : PMyRec;
 dccDateStr  : string;
begin
  {Only shows the node if it is a DCC Managed Meter(S1EA or S2)}
  if isS1Enrolled(mpxn) then  {Enrolled by Utilita}
  begin
    with Generalquery do
    begin
      close;
      DeleteVariables;
      DeclareVariable('MPXN', otstring);
      sql.clear;
      sql.add('SELECT	seie.PROCESSED_DATE FROM ods.S1EA_INTEGRATION_EVENT seie WHERE seie.MPXN = :MPXN');
      setvariable('MPXN', mpxn);
      open;
      deletevariables;
    end;

    dccDateStr := FormatDateTime('dd/mm/yyyy',GeneralQuery.Fields[0].AsDateTime);

    DCCEnrolledNode := Treeview1.AddChild(spannode);
    NodeCaption := TreeView1.GetNodeData(DCCEnrolledNode);
    NodeCaption.Index := -1;
    NodeCaption.Caption := 'DCC Managed Effective ('+dccDateStr+')';

  end else if isCosGainEnrolled(mpxn) then {Came enrolled by another supplier via CosGain}
  begin
    with Generalquery do
    begin
      close;
      DeleteVariables;
      DeclareVariable('MPXN', otstring);
      sql.clear;
      sql.add('SELECT TRUNC(DMS.DEVICEINSTALLATIONDATE) EFSDMSMTD ');
      sql.add('FROM AUTOMATIONPRO.SMETS2_COS_GAINS SCG JOIN ODS.DCC_MPXN_STATUS DMS ON (SCG.MPXN = DMS.MPXN)');
      sql.Add('WHERE SCG.MPXN = :MPXN and  SCG.METER_TYPE = '+QuotedStr('S1')+'  and SCG.STATE = 1');
      setvariable('MPXN', mpxn);
      open;
      deletevariables;
    end;

    dccDateStr := FormatDateTime('dd/mm/yyyy',GeneralQuery.Fields[0].AsDateTime);

    DCCEnrolledNode := Treeview1.AddChild(spannode);
    NodeCaption := TreeView1.GetNodeData(DCCEnrolledNode);
    NodeCaption.Index := -1;
    NodeCaption.Caption := 'DCC Managed Effective ('+dccDateStr+')';
  end;
end;

procedure TFRM_Tree.SetCharge1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 SetGreenDeal(TreeData.D_REGID);
 ShowMessage('Refresh Tree to see any changes');
end;

procedure TFRM_TREE.SetGreenDeal(REGID:string);
begin
 Application.CreateForm(TFrm_Green_Deal, Frm_Green_Deal);
 try
  frm_Green_Deal.showmodal;
  if frm_green_deal.tag=1 then
  begin
   with main_data_module.updatequery do
   Begin
    close;
    sql.clear;
    sql.add('Update crm.spans set e_eac='+floattostr(frm_green_deal.cost.value));
    sql.add('where registration_id='+regid);
    execute;
   End;
   frm_login.mainsession.commit;
  End;
 finally
  frm_Green_Deal.release;
 end;
end;

procedure TFRM_Tree.RemoveCOTAsifNeverVacated1Click(Sender: TObject);
var
  Agreement_id, Premise_ID, CID, strSQL: string;
begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);
  Agreement_id := TreeData.D_Agreement_ID;
  Premise_ID := TreeData.D_Premise_Id;

  If MessageDlg('Are you sure you wish to UN VACATE this premise?', mtconfirmation, [MBYES, MBNO], 0) <> mryes then
    exit;
  if FRM_Common.authoritycheck = False then
    exit;

  strSQL := 'update crm.agreement_premises set date_moved_out = null';
  strSQL := strSQL + CRLF + 'where premise_id = :p_premise_id and agreement_id = :p_agreement_id';

  gSqlUtil.ExecSql(strSQL, TRANSACTION_NO,
     [
       'p_premise_id', otString, pdInput, Premise_ID
       ,'p_agreement_id', otString, pdInput, Agreement_id
     ]);

  // Need Something in Here to Unlock The Spans, if Manual Overrides
  strSQL := 'UPDATE crm.spans SET';
  strSQL := strSQL + CRLF + '	span_end_date = NULL, lock_end_date = NULL, lock_status = NULL';
  strSQL := strSQL + CRLF + '	, SPAN_END_REASON = NULL, order_status_id = :p_order_status_id';
  strSQL := strSQL + CRLF + 'WHERE service_id IN';
  strSQL := strSQL + CRLF + '	(';
  strSQL := strSQL + CRLF + '		SELECT service_id FROM crm.service';
  strSQL := strSQL + CRLF + '		WHERE premise_id = :p_premise_id AND agreement_id = :p_agreement_id';
  strSQL := strSQL + CRLF + '	)';

  gSqlUtil.ExecSql(strSQL, TRANSACTION_NO,
     [
       'p_order_status_id', otInteger, pdInput, 8
       ,'p_premise_id', otString, pdInput, Premise_ID
       ,'p_agreement_id', otString, pdInput, Agreement_id
     ]);

  frm_login.MainSession.commit;
  // Need Something in Here to Unlock The Spans, if Manual Overrides

  CID := FRM_Common.GetCustomerIdfromAgreementid(Agreement_id);
  FRM_Common.SetAudit(CID, Agreement_id, '', 'Premises ID ' + Premise_ID + ' Un Vacated');
  MessageDlg('Premises as been Un Vacated.' + #13 + 'Please check Customer mailing address is correct and remove any forwarding addresses if required.', mtinformation,
    [mbOk], 0);

  Treeview1.Expanded[XNode.parent] := False;
  Treeview1.Expanded[XNode.parent] := True;
end;

procedure TFRM_Tree.ChangeDateMovedOut1Click(Sender: TObject);
var
agreement_id,premise_id,dmo,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Agreement_id:=TreeData.D_agreement_id;
 Premise_id:=TreeData.D_Premise_id;

 repeat
  Application.CreateForm(Tfrm_datepicker, frm_datepicker);
  FRM_DatePicker.tag:=0;
  frm_datepicker.cal1.date:=now;
  frm_datepicker.caption:='Select Vacated Date';
  frm_datepicker.DatePan.caption:='Please Select Vacated Date';
  FRM_DatePicker.showmodal;
  if FRM_DatePicker.tag=0 then messagedlg('Please Select Vacated Date',mtinformation,[MBOK],0)
 until FRM_DatePicker.tag=1;
 dmo:=datetostr(frm_datepicker.cal1.date);
 frm_datepicker.release;

 If Messagedlg('Are you sure you wish to change the Vacated Date on this premise to '+dmo+'?',mtconfirmation,[MBYES,MBNO],0)<>mryes then exit;

 if frm_common.authoritycheck=false then exit;

 With Main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('update crm.agreement_premises set date_moved_out = to_date('''+dmo+''',''DD/MM/YYYY'')');
  sql.add('where premise_id='+premise_id+' and agreement_id='+agreement_id);
  execute;
 End;
 frm_login.mainsession.commit;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agreement_id);
 FRM_Common.SetAudit(CID,agreement_id,'','Premises ID '+premise_id+' COT Date Changed to '+dmo);
 Messagedlg('Vacated Date on this Premise successfully changed to '+dmo+'. Refresh Tree to view changes',mtinformation,[mbok],0);
end;

procedure TFRM_Tree.ChangeCOTMovedfinDate1Click(Sender: TObject);
Var
AgId,startdate,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 AgID:=TreeData.D_Agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 Startdate:=TreeData.D_Agreement_Start_date;

 repeat
  Application.CreateForm(Tfrm_datepicker, frm_datepicker);
  FRM_DatePicker.tag:=0;
  frm_datepicker.cal1.date:=now;
  frm_datepicker.caption:='Select Vacated Date';
  frm_datepicker.DatePan.caption:='Please Select Vacated Date';
  FRM_DatePicker.showmodal;
  if FRM_DatePicker.tag=0 then messagedlg('Please Select Vacated Date',mtinformation,[MBOK],0)
 until FRM_DatePicker.tag=1;
 startdate:=datetostr(frm_datepicker.cal1.date);
 frm_datepicker.release;

 if Messagedlg('Are you sure you wish to change the COT Moved In Date to '+startdate+#13+
               'This new date will be applied to Agreement, Products, Services and SPANS etc',Mtconfirmation,[MBYES,MBNO],0)<>MRyes then exit;

 if frm_common.authoritycheck=false then exit;

 // Now MAKE LIVE Sub Orders
 with main_data_module.updatequery do
 Begin
  close;
  deletevariables;
  DeclareVariable('AID', otlong);
  setvariable('AID',agid);
  DeclareVariable('SSD', otstring);
  setvariable('SSD',startdate);
  sql.clear;
  sql.add('Update crm.agreements');
  sql.add('set agreement_start_date=to_date(:SSD,''DD/MM/YYYY'')');
  sql.add('where agreement_id=:AID');
  execute;
  close;
  sql.clear;
  sql.add('Update crm.agreement_products');
  sql.add('set effective_from=to_date(:SSD,''DD/MM/YYYY'')');
  sql.add('where agreement_id=:AID');
  execute;
  close;
  sql.clear;
  sql.add('Update crm.service');
  sql.add('set start_date=to_date(:SSD,''DD/MM/YYYY'')');
  sql.add('where agreement_id=:AID');
  execute;
  close;
  sql.clear;
  sql.add('Update crm.spans');
  sql.add('set span_start_date=to_date(:SSD,''DD/MM/YYYY'')');
  sql.add('where (service_id) in');
  sql.add('(select service_id from crm.service');
  sql.add('where agreement_id=:AID)');
  execute;
  deletevariables;
 End;
 FRM_Common.SetAudit(CID,agreement_id,'','COT moved in Date Changed to '+startdate);
 frm_login.mainsession.commit;

 treeview1.Expanded[xnode]:=false;
 treeview1.Expanded[xnode]:=true;

end;

procedure TFRM_Tree.M_ALClick(Sender: TObject);
Var
custid,lib_cid:string;
updated:boolean;
begin

 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 custid:=TreeData.D_Customer_id;
 updated:=false;
 Application.CreateForm(Tfrm_liberty_custid,frm_liberty_custid);
 frm_liberty_custid.tag:=0;
 frm_liberty_custid.cid.text:='';
 frm_liberty_custid.showmodal;
 if frm_liberty_custid.tag=1 then
 Begin
  updated:=true;
  lib_cid:=frm_liberty_custid.cid.text;
  with main_data_module.updatequery do
  Begin
   close;
   sql.clear;
   sql.add('delete from crm.customer_to_liberty_customer where customer_id='+custid);
   execute;
   close;
   sql.clear;
   sql.add('insert into crm.customer_to_liberty_customer values('+custid+','+lib_cid+')');
   execute;
   close;
   sql.clear;
   sql.add('insert into crm.customer values('+lib_cid+',1,''Used In Liberty. CRM ID='+custid+''',0,null,null,null,null,null,sysdate,'''+USERID+''',1)');
   TRY
    execute;
   EXCEPT
   END;
   frm_login.mainsession.commit;
   FRM_Common.SetAudit(CUSTID,'','','Customer assigned to Liberty Customer - '+lib_cid);
  end;
 end;
 frm_liberty_custid.release;
 if updated=true then
 Begin
  FRM_Main.SearchForcust(custid);
  xNode := Treeview1.GetFirst();
  treeview1.expanded[xnode]:=false;
  treeview1.expanded[xnode]:=true;
 end;
End;

procedure TFRM_Tree.SmartMopClick(Sender: TObject);
var
  Agreement_id,premise_id, Cust_ID: string;

begin
  xnode:=treeview1.FocusedNode;
  TreeData:= treeview1.GetNodeData(xnode);
  Agreement_id:=TreeData.D_agreement_id;
  Premise_id:=TreeData.D_Premise_id;
  Cust_ID := TreeData.D_Customer_Id;

  if Cust_ID = EmptyStr then
  begin
    Cust_ID := custid;
  end;

  if (TreeData.index=142) or (TreeData.index=139) then
  begin
    Messagedlg('This action cannot be performed against a VACATED Property',mtwarning,[mbok],0);
    exit;
  end;

  if FRM_COMMON.GETVALUE('ENABLE_JBS_BOOKING')='D' then
  begin
    Messagedlg('This Feature has been Temporarily Disabled.'+#13+
               'Please Contact SmartMOp to book the job.',mtwarning,[MBOK],0);
    exit;
  end;

  if frm_common.isagreementlive(agreement_id)=false then
  begin
    ShowMessage('You must select a Premise on a LIVE agreement');
    exit;
  end;

  // see if live job already exists
  with main_data_module.TempQuery do
  begin
    Close;
    Deletevariables;
    DeclareVariable('CID',otlong);
    SQL.Clear;
    SQL.Add('SELECT * FROM SMIFF.WMOL_ACCOUNTS WHERE CUSTOMER_LIVE = :CID ');
    SetVariable('CID',Cust_ID);
    Open;
    DeleteVariables;
  end;

  if main_data_module.TempQuery.RecordCount <> 0 then
  begin
    Messagedlg('You are unable to book a job. A LIVE job already exists in JBS for this Customer. Please amend existing Job?',mterror,[mbok],0);
    exit;
  end;

  Application.CreateForm(TFrm_add2Jbs, Frm_add2jbs);
  try
    frm_add2jbs.clearfields;
    frm_add2jbs.agreementId := Agreement_id;
    frm_add2jbs.btnGenSuperAuthFlag := true;
    frm_add2jbs.CheckCommsRestrictFlag := true;
    frm_add2jbs.GetData(premise_id);
    frm_add2jbs.Position := poDesktopCenter;
    frm_add2jbs.showmodal;
  finally
    frm_add2jbs.release;
  end;

  treeview1.expanded[xnode]:=false;
  treeview1.expanded[xnode]:=true;
end;

procedure TFRM_Tree.custom_meterClick(Sender: TObject);
Var
regid,AGID,CID:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;

 mpan:=TreeData.D_SPAN;
 regid:=TreeData.D_REGID;
 AGID:=TreeData.D_agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 FRM_CUSTOM_METERING.show;
 FRM_CUSTOM_METERING.getmeterdetails(mpan,'','','',regid,CID);
end;

procedure TFRM_Tree.TransferSpan;
Var
Reg_ID,Premise_ID,service_id,agreement_id,span,cid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 // First Of all Get theexisting AGREEMENT_ID, PREMISE_ID & SERVICE ID for selected SPAN
 // Get Premise from Registration;
 reg_id:=TreeData.D_REGid;
 with main_data_module.generalquery do
 Begin
  close;
  deletevariables;
  declarevariable('REGID',otlong);
  sql.clear;
  sql.add('select SERVICE.PREMISE_ID,SERVICE.SERVICE_ID,SPANS.SPAN,SERVICE.AGREEMENT_ID');
  sql.add('FROM');
  sql.add('    CRM.SPANS SPANS,');
  sql.add('    CRM.SERVICE SERVICE,');
  sql.add('    CRM.SPAN_TYPE SPAN_TYPE,');
  sql.add('    CRM.SERVICE_TYPE SERVICE_TYPE');
  sql.add(' WHERE');
  sql.add('     SPANS.SERVICE_ID = SERVICE.SERVICE_ID AND');
  sql.add('     SPANS.SPAN_TYPE_ID = SPAN_TYPE.SPAN_TYPE_ID AND');
  sql.add('     SERVICE.SERVICE_TYPE_ID = SERVICE_TYPE.SERVICE_TYPE_ID');
  sql.add('     and SPANS.REGISTRATION_ID=:REGID');
  setvariable('REGID',REG_ID);
  open;
  deletevariables;
 End;
 PREMISE_ID:=main_data_module.generalquery.fields[0].text;
 SERVICE_ID:=main_data_module.generalquery.fields[1].text;
 SPAN:=main_data_module.GeneralQuery.fields[2].text;
 AGREEMENT_ID:=main_data_module.GeneralQuery.fields[3].text;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agreement_id);

 // Now Do we want to move to a NEW Agreement or an Existing Agreement
 FRM_TRANSFER_SPAN.l_SPAN.caption:=span;
 FRM_TRANSFER_SPAN.l_ag.caption:=agreement_id;
 FRM_TRANSFER_SPAN.l_prem.caption:=premise_id;
 FRM_TRANSFER_SPAN.l_serv.caption:=service_id;

 with FRM_TRANSFER_SPAN.agreements_current do
 begin
  close;
  sql.clear;
  sql.add('select * from crm.customer_agreement_products where agreement_id='+agreement_id);
  open;
 end;

 with FRM_TRANSFER_SPAN.agreements do
 begin
  close;
  sql.clear;
  sql.add('select * from crm.customer_agreement_products where customer_id='+cid+' and agreement_id<>'+agreement_id+' order by agreement_id');
  open;
 end;

 FRM_TRANSFER_SPAN.tag:=0;
 FRM_TRANSFER_SPAN.SPDESC.caption:=TreeData.caption;
 FRM_TRANSFER_SPAN.SPDESC.font.color:=TreeData.fontcolor;
 FRM_TRANSFER_SPAN.SpanPic.picture.assign(nil);
 dm_images.largeimages.GetIcon(TreeData.index,FRM_TRANSFER_SPAN.Spanpic.Picture.Icon);
 FRM_TRANSFER_SPAN.SHOWMODAL;
 if FRM_TRANSFER_SPAN.tag=1 then
 begin
  Messagedlg('Span has been Transferred'+#13+
            'Please Check Product Details are Correct on All Agreements.',mtinformation,[MBOK],0);
 FRM_Main.SearchForcust(cid);
 end;

end;

procedure TFRM_Tree.ransfertoNeworExistingAgreement1Click(Sender: TObject);
begin
 TransferSpan;
end;

procedure TFRM_Tree.AdditionalCharges1Click(Sender: TObject);
var
fuel_type,span,agreement_id,reg_id:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Reg_id:=TreeData.D_REGID;

 fuel_type:='';
 SPAN:=TreeData.D_SPAN;
 if copy(span,1,2)='92' then fuel_type:='X'; //Generation
 if copy(span,1,2)='93' then fuel_type:='S'; // Sewerage
 if copy(span,1,2)='94' then fuel_type:='g'; // Green Deal
 if copy(span,1,2)='95' then fuel_type:='w'; // Water
 if copy(span,1,2)='96' then fuel_type:='W'; // Grey Water
 if copy(span,1,2)='97' then fuel_type:='H'; // Heat
 if copy(span,1,2)='98' then fuel_type:='g'; // Sub G
 if copy(span,1,2)='99' then fuel_type:='e'; // Sub E
 Agreement_id:=TreeData.D_AGREEMENT_ID;

 FRM_ONE_OFF_CHARGE.ShowTheseDetails(Span,reg_id,agreement_id,fuel_type);
 FRM_ONE_OFF_CHARGE.showmodal;
end;

procedure TFRM_Tree.BACS1Click(Sender: TObject);
begin
//  jsilva - 24/09/2018
  if not Assigned(fDDMain) then
    fDDMain := TfDDMain.Create(Self);
  Try
    nodeData := treeview1.GetNodeData(xnode);
    fDDMain.clearAllFilter;
    fDDMain.edReference.Text:= nodedata.D_agreement_id;
    fDDMain.EditID.Text:= nodedata.D_agreement_id;
    fDDMain.pcDDMain.ActivePageIndex := 0;
    if Main.DDICUser then fDDMain.UnlockScreen
    else fDDMain.LockScreen;
    fDDMain.Show;
    fDDMain.BringToFront;
  finally
  end;
end;

procedure TFRM_TREE.BEN_TEMP_SPANS_SPLIT;
Var
totrec:integer;
Premise_ID,service_id,agreement_id,span,CID:String;
begin
 // First Of all Get theexisting AGREEMENT_ID, PREMISE_ID & SERVICE ID for selected SPAN
 // Get Premise from Registration;
 repeat
 with main_data_module.generalquery do
 Begin
  close;
  sql.clear;
  sql.add('select PREMISE_ID,SERVICE_ID,SPAN,AGREEMENT_ID');
  sql.add('FROM CRM.TEMP_BEN_WATER_SPANS_TO_SPLIT');
  sql.add('order by span');
  open;
 End;
 totrec:=main_data_module.generalquery.recordcount;
 if totrec=0 then exit;
 caption:=inttostr(totrec);
 PREMISE_ID:=main_data_module.generalquery.fields[0].text;
 SERVICE_ID:=main_data_module.generalquery.fields[1].text;
 SPAN:=main_data_module.GeneralQuery.fields[2].text;
 AGREEMENT_ID:=main_data_module.GeneralQuery.fields[3].text;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agreement_id);

 // Now Do we want to move to a NEW Agreement or an Existing Agreement
 FRM_TRANSFER_SPAN.l_SPAN.caption:=span;
 FRM_TRANSFER_SPAN.l_ag.caption:=agreement_id;
 FRM_TRANSFER_SPAN.l_prem.caption:=premise_id;
 FRM_TRANSFER_SPAN.l_serv.caption:=service_id;

 with FRM_TRANSFER_SPAN.agreements_current do
 begin
  close;
  sql.clear;
  sql.add('select * from crm.customer_agreement_products where agreement_id='+agreement_id);
  open;
 end;

 with FRM_TRANSFER_SPAN.agreements do
 begin
  close;
  sql.clear;
  sql.add('select * from crm.customer_agreement_products where customer_id='+cid+' and agreement_id<>'+agreement_id+' order by agreement_id');
  open;
 end;

 FRM_TRANSFER_SPAN.tag:=0;
 FRM_TRANSFER_SPAN.SPDESC.caption:=span;
 FRM_TRANSFER_SPAN.SHOW;
 FRM_TRANSFER_SPAN.NEWBTN.CLICK;
 FRM_TRANSFER_SPAN.CLOSEBTN.CLICK;
 until totrec=0;

end;
procedure TFRM_Tree.ransferSPANS1Click(Sender: TObject);
begin
BEN_TEMP_SPANS_SPLIT;
end;

procedure TFRM_Tree.T_CCClick(Sender: TObject);
var
agid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Agid:=TreeData.D_Agreement_ID;
 FRM_CREDIT_CONTROL_FOLLOW_UP.GetLatestRecord(agid);
 FRM_CREDIT_CONTROL_FOLLOW_UP.agid.caption:=agid;
 FRM_CREDIT_CONTROL_FOLLOW_UP.showmodal;
end;

Procedure TFRM_Tree.FMSClick(Sender: TObject);
var
    bAgreement_Id,
    bPremise_Id : string;
Begin
   XNode         := Treeview1.FocusedNode;
   TreeData         := Treeview1.GetNodeData(xnode);
   bAgreement_Id := Treedata.D_Agreement_ID;
   bPremise_Id   := Treedata.D_Premise_Id;

   If (TreeData.Index in [142, 139]) then
    Begin
     Messagedlg('This action cannot be performed against a VACATED Property',mtwarning,[mbok],0);
     exit;
    End;

   if FRM_Common.isagreementlive(bAgreement_Id) then
     TFRM_FMS.Launch(bAgreement_Id, bPremise_Id)
   else
     ShowMessage('You must select a Premise on a LIVE agreement');
 End; // If


procedure TFRM_Tree.CustomerLetters2Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 FRM_AVAILABLE_LETTERS.CID.caption:=TreeData.D_customer_id;
 FRM_AVAILABLE_LETTERS.AID.caption:='';
 FRM_AVAILABLE_LETTERS.PID.caption:='' ;
 FRM_AVAILABLE_LETTERS.RID.caption:='' ;
 FRM_AVAILABLE_LETTERS.SHOWLETTERS('''Customer''','Customer');
 FRM_AVAILABLE_LETTERS.Span := FRM_Main_Search.CustomerQuery.Fields[46].AsString;
 FRM_AVAILABLE_LETTERS.showmodal;
end;

procedure TFRM_Tree.Letters1NClick(Sender: TObject);
var
AGID,CID:STRING;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 FRM_AVAILABLE_LETTERS.CID.caption:=CID;
 FRM_AVAILABLE_LETTERS.AID.caption:=AGID;
 FRM_AVAILABLE_LETTERS.PID.caption:='';
 FRM_AVAILABLE_LETTERS.RID.caption:='';
 FRM_AVAILABLE_LETTERS.SHOWLETTERS('''Customer'',''Agreement''','Agreement');
 FRM_AVAILABLE_LETTERS.Span := FRM_Main_Search.CustomerQuery.Fields[46].AsString;
 FRM_AVAILABLE_LETTERS.showmodal;
end;

procedure TFRM_Tree.Letters2NClick(Sender: TObject);
var
agid,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 FRM_AVAILABLE_LETTERS.CID.caption:=cid;
 FRM_AVAILABLE_LETTERS.AID.caption:=agid;
 FRM_AVAILABLE_LETTERS.PID.caption:=TreeData.D_Premise_id;
 FRM_AVAILABLE_LETTERS.RID.caption:='';
 FRM_AVAILABLE_LETTERS.SHOWLETTERS('''Customer'',''Agreement'',''Premise''','Premise');
 FRM_AVAILABLE_LETTERS.Span := FRM_Main_Search.CustomerQuery.Fields[46].AsString;
 FRM_AVAILABLE_LETTERS.showmodal;
end;

procedure TFRM_Tree.LoanAmount1Click(Sender: TObject);
begin
  DoSmetsCreditDCC(5);
end;

procedure TFRM_Tree.L_G_NClick(Sender: TObject);
var
agid,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 FRM_AVAILABLE_LETTERS.CID.caption:=cid;
 FRM_AVAILABLE_LETTERS.AID.caption:=agid;
 FRM_AVAILABLE_LETTERS.PID.caption:=TreeData.D_Premise_id;
 FRM_AVAILABLE_LETTERS.RID.caption:=TreeData.D_regid;
 FRM_AVAILABLE_LETTERS.SHOWLETTERS('''Customer'',''Agreement'',''Premise'',''Supply-Gas''','Supply-Gas');
 FRM_AVAILABLE_LETTERS.Span := FRM_Main_Search.CustomerQuery.Fields[46].AsString;
 FRM_AVAILABLE_LETTERS.showmodal;
end;

procedure TFRM_Tree.L_E_NClick(Sender: TObject);
var
agid,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 FRM_AVAILABLE_LETTERS.CID.caption:=cid;
 FRM_AVAILABLE_LETTERS.AID.caption:=agid;
 FRM_AVAILABLE_LETTERS.PID.caption:=TreeData.D_Premise_id;
 FRM_AVAILABLE_LETTERS.RID.caption:=TreeData.D_regid;
 FRM_AVAILABLE_LETTERS.SHOWLETTERS('''Customer'',''Agreement'',''Premise'',''Supply-Elec''','Supply-Elec');
 FRM_AVAILABLE_LETTERS.Span := FRM_Main_Search.CustomerQuery.Fields[46].AsString;
 FRM_AVAILABLE_LETTERS.showmodal;
end;

procedure TFRM_Tree.VendPaymentStatementReport1Click(Sender: TObject);
var
regid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 REGID:=TreeData.D_REGid;
 Application.CreateForm(TFRM_VEND_PAYMENT_SUMMARY, FRM_VEND_PAYMENT_SUMMARY);
 FRM_VEND_PAYMENT_SUMMARY.GetMeters(Regid);
 FRM_VEND_PAYMENT_SUMMARY.showmodal;
 FRM_VEND_PAYMENT_SUMMARY.release;
end;

procedure TFRM_Tree.Custom_lettersClick(Sender: TObject);
var
supplytype,spt,agid,cid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AGID:=TreeData.D_agreement_id;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 SPT:='O';
 FRM_AVAILABLE_LETTERS.CID.caption:=cid;
 FRM_AVAILABLE_LETTERS.AID.caption:=agid;
 FRM_AVAILABLE_LETTERS.PID.caption:=TreeData.D_Premise_id;
 FRM_AVAILABLE_LETTERS.RID.caption:=TreeData.D_regid;
 with main_data_module.tempquery do
 begin
  deletevariables;
  close;
  sql.clear;
  sql.add('select st.service_type from crm.span_type st, crm.spans s where s.registration_id='+(TreeData.D_regid));
  sql.add('and st.span_type_id=s.span_type_id');
  open;
 end;
 if  main_data_module.tempquery.recordcount<>0 then spt:=main_data_module.tempquery.fields[0].text;
 if spt='H' then supplytype:='Supply-Heat';
 if spt='W' then supplytype:='Supply-Water';
 if spt='E' then supplytype:='Supply-Elec';
 if spt='G' then supplytype:='Supply-Gas';
 if spt='O' then supplytype:='Supply-Other';
 FRM_AVAILABLE_LETTERS.SHOWLETTERS('''Customer'',''Agreement'',''Premise'','''+supplytype+'''',supplytype);
 FRM_AVAILABLE_LETTERS.Span := FRM_Main_Search.CustomerQuery.Fields[46].AsString;
 FRM_AVAILABLE_LETTERS.showmodal;
end;

procedure TFRM_TREE.AddDapMarker(Span,AGid:string);
var
cid:string;
begin
  xnode:=treeview1.FocusedNode;
 if messagedlg('Are you sure you wish to add a Debt Assignment Process Marker to this Supply?'+#13+
               'Any losses that come in for SPAN '+SPAN+' will be automatically Objected.'+#13+
               'If Supply is NOT live, you will not be able to order it until marker is removed.'
               ,mtconfirmation,[mbyes,mbno],0)<>mryes then exit;

 if frm_common.Superauthoritycheck=false then exit;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('Insert into crm.spans_dap_marker values('''+SPAN+''','''+USERID+''',sysdate)');
  try
   execute;
   FRM_Common.SetAudit(cid,AGid,span,'Debt Assignment Process Marker Added');
   except
   Messagedlg('There was a problem actioning this request.',mterror,[mbok],0);
   exit;
  end;
 End;
 frm_login.mainsession.commit;

     // Refresh Parent Node
   if treeview1.Selected[xnode]=true then
   Begin
    treeview1.Expanded[xnode.parent]:=false;
    treeview1.Expanded[xnode.parent]:=true;
    //if node is first item, then no paretn so refresh span
    if treeview1.AbsoluteIndex(XNode)=0 then FRM_Main.SearchForSpan(nodedata.D_SPAN,0);
   end;

end;

procedure TFRM_TREE.RemoveDapMarker(Span,AGid:string);
var
cid:string;
begin
 xnode:=treeview1.FocusedNode;
 nodeData := treeview1.GetNodeData(xnode);

 if messagedlg('Are you sure you wish to REMOVE the Debt Assignment Process Marker from this Supply?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;
 if frm_common.Superauthoritycheck=false then exit;
 Cid:=frm_common.GetCustomerIdfromAgreementid(Agid);
 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from crm.spans_dap_marker where span='''+span+'''');
  try
   execute;
   FRM_Common.SetAudit(cid,AGid,span,'Debt Assignment Process Marker Removed');
   except
   Messagedlg('There was a problem actioning this request.',mterror,[mbok],0);
   exit;
  end;
 End;
 frm_login.mainsession.commit;

     // Refresh Parent Node
   if treeview1.Selected[xnode]=true then
   Begin
    treeview1.Expanded[xnode.parent]:=false;
    treeview1.Expanded[xnode.parent]:=true;
    //if node is first item, then no paretn so refresh span
    if treeview1.AbsoluteIndex(XNode)=0 then FRM_Main.SearchForSpan(nodedata.D_SPAN,0);
   end;

end;


procedure TFRM_Tree.G_DebtClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AddDapMarker(TreeData.D_SPAN,Treedata.D_AGREEMENT_ID);
end;

procedure TFRM_Tree.E_DebtClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 AddDapMarker(TreeData.D_SPAN,Treedata.D_AGREEMENT_ID);
end;

procedure TFRM_Tree.G_DebtRClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 RemoveDapMarker(TreeData.D_SPAN,Treedata.D_AGREEMENT_ID);
end;

procedure TFRM_Tree.E_DebtRClick(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 RemoveDapMarker(TreeData.D_SPAN,Treedata.D_AGREEMENT_ID);
end;

procedure TFRM_Tree.E_ErroneousClick(Sender: TObject);
begin
 if Messagedlg('Your are about to remove a complaint. Are you sure this is correct?',mtconfirmation,[mbyes,mbno],0)<> mryes then exit;

end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.Smets_AdminClick(Sender: TObject);
var
  nodeData           : PMyRec;
  service            : integer;
  agreementId        : Int64;
  smetsFormId        : integer;
  smetsCosGainFormId : integer;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  service := StrtoInt(nodeData.M_Service);

  if isDCCMeter then
  begin
    // DCC meter

    // ISC-722 (Anna): in this case field number must be used as this field position can
    // apparently contain other field name than "agreement_id"
    agreementId := StrToInt64Def(Frm_Main_Search.CustomerQuery.Fields[23].AsString, 0);

    TFrm_Smets_Dcc.ShowSmets(Self, service, nodeData.D_Span, agreementId);
  end
  else
  begin
    TFrm_Smets.ShowSmets(service, nodeData.D_Span, smetsFormId, smetsCosGainFormId);
  end;
end;

procedure TFRM_Tree.Smets_vend_addClick(Sender: TObject);
begin
 DoSmets2Credit(0);
end;

procedure TFRM_Tree.smets_vend_deductClick(Sender: TObject);
begin
DoSmets2Credit(1);
end;

procedure TFRM_Tree.Smets_vend_setClick(Sender: TObject);
begin
DoSmets2Credit(2);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.SM_EClick(Sender: TObject);
var
  nodeData           : PMyRec;
  agreementId        : Int64;
  smetsFormId        : integer;
  smetsCosGainFormId : integer;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  if nodeData.D_SpanE = '' then
    exit;

  // show DCC electric meter screen if DCC managed meter;
  if nodeData.D_SpanDCC_E then
  begin
    // ISC-722 (Anna): in this case field number must be used as this field position can
    // apparently contain other field name than "agreement_id"
    agreementId := StrToInt64Def(FRM_Main_Search.CustomerQuery.Fields[23].AsString, 0);

    TFrm_Smets_Dcc.ShowSmets(Self, SERVICE_ELECTRICITY, nodeData.D_SpanE, agreementId);
  end
  else
  begin
    TFrm_Smets.ShowSmets(SERVICE_ELECTRICITY, nodeData.D_SpanE, smetsFormId, smetsCosGainFormId);
  end;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.SM_GClick(Sender: TObject);
var
  nodeData           : PMyRec;
  agreementId        : Int64;
  smetsFormId        : integer;
  smetsCosGainFormId : integer;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  if nodeData.D_SpanG = '' then
    exit;

  if nodeData.D_SpanDCC_G then
  begin
    // ISC-722 (Anna): in this case field number must be used as this field position can
    // apparently contain other field name than "agreement_id"
    agreementId := StrToInt64Def(FRM_Main_Search.CustomerQuery.Fields[23].AsString, 0);

    TFrm_Smets_Dcc.ShowSmets(Self, SERVICE_GAS, nodeData.D_SpanG, agreementId);
  end
  else
  begin
    TFrm_Smets.ShowSmets(SERVICE_GAS, nodeData.D_SpanG, smetsFormId, smetsCosGainFormId);
  end;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.DoSmets2Debt(aMode: integer);
var
  nodeData  : PMyRec;
  service   : integer;
  suspended : TDate;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  service := StrToInt(nodeData.m_Service);

  // if DO_SMETS_SUPPLIER_CHECK(SPAN) = false then Exit;

  if GetMeterMode(nodeData.D_Span) <> '1' then
  begin
    Messagedlg('This option is only allowed if Current Meter is in PRE-PAYMENT Mode', mtError, [mbOk], 0);
    exit;
  end;

  if GetPaymentCardNo(nodeData.D_Span, EmptyStr) = EmptyStr then
  begin
    Messagedlg('This option is NOT allowed. No Topup Card is Registered', mtError, [mbOk], 0);
    exit;
  end;

  suspended := Check_For_Suspended_Debt_Dcc(nodeData.D_Span);
  if suspended > 0 then
  begin
    if Messagedlg('There is a suspended Debt on this account. ' +
                  'This must be Cancelled before any Debt Changes can be made.' + #13 + #13 +
                  'Do you wish to Remove this Suspended Debt?',
                  mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
      exit;

    try
      RemoveSuspendedDebtDCC(nodeData.D_Span, suspended);
    except
      on e:Exception do
      begin
        MessageDlg(e.Message, mtError, [mbOK], 0);
        exit;
      end;
    end;
  end;

  TFrm_Smets_Manage_Debt_Dcc.StartModal(
    Self,
    rqMpanTree,
    service,
    nodeData.D_Span,
    nodeData.M_MeterId,
    aMode,
    false);
end;

{------------------------------------------------------------------------------}
procedure TFRM_TREE.DoSmets2Credit(aMode: integer);
var
  nodeData      : PMyRec;
  service       : integer;
  agreementNode : PMyRec;
  agreementId   : Int64;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  service := StrToInt(nodeData.M_Service);

  // CR-511 (William): Implemented to prevent access to free vend options with RP flags,
  // the value returned by FRM_Main_Search.CustomerQuery in the case of smets1 is always blank
  agreementNode := TreeView1.GetNodeData(mynodeagreement);
  if Assigned(agreementNode) then
    agreementId := StrToInt64Def(agreementNode.D_Agreement_Id, -1)
  else
    agreementId := -1;

  if IsDccMeter then
  begin

    if TSmetsUtil.HasRpFlags(TCrmUtil.GetCustomerIdfromAgreementid(agreementId)) and (not Frm_Common.SuperAuthorityCheck) then
    begin
      MessageDlg('Super Authority check has failed', mtError, [mbOk], 0);
      exit;
    end;

    if GetMeterMode(nodeData.D_Span) <> '1' then
    begin
      MessageDlg('This option is only allowed if Current SMETS2 Meter is in PRE-PAYMENT Mode.', mtError, [mbOk], 0);
      exit;
    end;

    TFrm_Smets_Manage_Credit_Dcc.StartModal(Self, service, nodeData.D_Span, nodeData.M_MeterId, aMode);
  end
  else
  begin
    if not Do_Smets_Supplier_Check(nodeData.D_Span) then
      exit;

    if TSmetsUtil.HasRpFlags(TCrmUtil.GetCustomerIdfromAgreementid(agreementId)) and (not Frm_Common.SuperAuthorityCheck) then
    begin
      MessageDlg('Super Authority check has failed', mtError, [mbOk], 0);
      exit;
    end;

    if CheckSelectedMigration(nodeData.D_Span) then
    begin
      MessageDlg('Unable to action request. Meter is on 24/7 Friendly Credit during the migration window', mtError, [mbOk], 0);
      exit;
    end;

    if GetSmetsMeterMode(nodeData.D_Span) <> '1' then
    begin
      MessageDlg('This option is only allowed if Current SMETS Meter is in PRE-PAYMENT Mode', mtError, [mbOk], 0);
      exit;
    end;

    if GetSmetsPaymentCardno(nodeData.D_Span, '') = '' then
    begin
      MessageDlg('This option is NOT allowed. No Topup Card is Registered', mtError, [mbOk], 0);
      exit;
    end;

    TFrm_Smets_Manage_Credit.StartModal(Self, service, nodeData.D_Span, nodeData.M_MeterId, aMode);
  end;
end;

{------------------------------------------------------------------------------}
procedure TFRM_TREE.DoSmetsDebt(aMode: integer);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  if not Do_Smets_Supplier_Check(nodeData.D_Span) then
    exit;

  if CheckSelectedMigration(nodeData.D_Span) then
  begin
    MessageDlg('Unable to action request. Meter is on 24/7 Friendly Credit during the migration window.', mtError, [mbOk], 0);
    exit;
  end;

  if GetSmetsMeterMode(nodeData.D_Span) <> '1' then
  begin
    MessageDlg('This option is only allowed if Current SMETS Meter is in PRE-PAYMENT Mode.', mtError, [mbOk], 0);
    exit;
  end;

  if GetSmetsPaymentCardno(nodeData.D_Span, '' ) = '' then
  begin
    MessageDlg('This option is NOT allowed. No Topup Card is Registered.', mtError, [mbOk], 0);
    exit;
  end;

  TFrm_Smets_Manage_Debt.StartModal(
    Self,
    StrToInt(nodeData.M_Service),
    nodeData.D_Span,
    nodeData.M_MeterId,
    aMode,
    false,
    null,
    null);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.smets_debt_addClick(Sender: TObject);
begin
  DoSmetsDebt(DEBT_OPTION_ADD);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.Smets_Debt_Add_DCCClick(Sender: TObject);
begin
  DoSmets2Debt(0);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.smets_debt_deductClick(Sender: TObject);
begin
  DoSmetsDebt(DEBT_OPTION_REDUCE);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.Smets_Debt_Deduct_DCCClick(Sender: TObject);
begin
  DoSmets2Debt(1);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.Smets_debt_SetClick(Sender: TObject);
begin
  DoSmetsDebt(DEBT_OPTION_SET);
end;

procedure TFRM_Tree.Smets_ReadClick(Sender: TObject);
Var
Span,Meter,service:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Span:=TreeData.D_SPAN;
  if DO_SMETS_SUPPLIER_CHECK(SPAN)=false then Exit;
 Meter:=TreeData.m_meterid;
 service:=TreeData.m_service;

 DoFUllSnapShot(SPAN,METER,SERVICE,'Y','','','');

end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.Smets_txtClick(Sender: TObject);
Var
Span,Meter,service:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Span:=TreeData.D_SPAN;
  if DO_SMETS_SUPPLIER_CHECK(SPAN)=false then Exit;
 Meter:=TreeData.m_meterid;
 service:=TreeData.m_service;
 if GetSmetsWanStatus(Span,'')<>'ON' then
 Begin
  Messagedlg('Unable to action request. There does not appear to be any active Comms to this Supply',mtwarning,[mbok],0);
  exit;
 end;

 TFrm_Smets_TextMsg.StartModal(Self, StrToInt(TreeData.M_Service), TreeData.D_Span, TreeData.M_Meterid);
end;

procedure TFRM_Tree.smets_ihd_replaceClick(Sender: TObject);
Var
Span,Meter,service:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Span:=TreeData.D_SPAN;
  if DO_SMETS_SUPPLIER_CHECK(SPAN)=false then Exit;
 Meter:=TreeData.m_meterid;
 service:=TreeData.m_service;
  Application.CreateForm(TFRM_SMETS_REMOVE_DEVICE, FRM_SMETS_REMOVE_DEVICE);
 try
  FRM_SMETS_REMOVE_DEVICE.refreshdata(SPAN,service,Meter,'Y');
  FRM_SMETS_REMOVE_DEVICE.IHD_GROUP.visible:=true;
  FRM_SMETS_REMOVE_DEVICE.HAN_GROUP.visible:=true;
  FRM_SMETS_REMOVE_DEVICE.height:=340;
  FRM_SMETS_REMOVE_DEVICE.checknow.checked:=true;
  FRM_SMETS_REMOVE_DEVICE.caption:='Replace IHD';
  FRM_SMETS_REMOVE_DEVICE.showmodal;
 finally
  FRM_SMETS_REMOVE_DEVICE.release;
 end;
end;

procedure TFRM_Tree.smets_ihd_removeClick(Sender: TObject);
Var
Span,Meter,service:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Span:=TreeData.D_SPAN;
  if DO_SMETS_SUPPLIER_CHECK(SPAN)=false then Exit;
 Meter:=TreeData.m_meterid;
 service:=TreeData.m_service;
 Application.CreateForm(TFRM_SMETS_REMOVE_DEVICE, FRM_SMETS_REMOVE_DEVICE);
 try
  FRM_SMETS_REMOVE_DEVICE.refreshdata(SPAN,service,Meter,'Y');
  FRM_SMETS_REMOVE_DEVICE.IHD_GROUP.visible:=true;
  FRM_SMETS_REMOVE_DEVICE.HAN_GROUP.visible:=false;
  FRM_SMETS_REMOVE_DEVICE.checknow.checked:=true;
  //FRM_SMETS_REMOVE_DEVICE.duration.value:=1;
  FRM_SMETS_REMOVE_DEVICE.Caption:='Remove IHD';
  FRM_SMETS_REMOVE_DEVICE.height:=238;
  FRM_SMETS_REMOVE_DEVICE.showmodal;
 finally
  FRM_SMETS_REMOVE_DEVICE.release;
 end;
end;

procedure TFRM_Tree.smets_ihd_addClick(Sender: TObject);
Var
Span,Meter,service:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Span:=TreeData.D_SPAN;
  if DO_SMETS_SUPPLIER_CHECK(SPAN)=false then Exit;
 Meter:=TreeData.m_meterid;
 service:=TreeData.m_service;
 Application.CreateForm(TFRM_SMETS_REMOVE_DEVICE, FRM_SMETS_REMOVE_DEVICE);
 try
  FRM_SMETS_REMOVE_DEVICE.refreshdata(SPAN,service,Meter,'Y');
  FRM_SMETS_REMOVE_DEVICE.IHD_GROUP.visible:=false;
  FRM_SMETS_REMOVE_DEVICE.HAN_GROUP.visible:=true;
  FRM_SMETS_REMOVE_DEVICE.checknow.checked:=true;
  FRM_SMETS_REMOVE_DEVICE.Caption:='Add IHD';
//  FRM_SMETS_REMOVE_DEVICE.macaddress.text:='123';
  FRM_SMETS_REMOVE_DEVICE.height:=278;
  FRM_SMETS_REMOVE_DEVICE.showmodal;
 finally
  FRM_SMETS_REMOVE_DEVICE.release;
 end;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.smets_ihd_pinClick(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(Treeview1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(Treeview1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Ihd_pin.StartModal(Self, nodeData.D_Span);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.DataflowHisopry1Click(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Flow_History.Start(Self, nodeData.D_Span, false);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.DataflowHistory2Click(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Flow_History_Dcc.Start(Self, StrToInt(nodeData.M_Service), nodeData.D_Span, false);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.VendHistory1Click(Sender: TObject);
var
  dataNode : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  // Close Window First, in case no records returned.
  if IsDccMeter then
    TFrm_Smets_Vends_Dcc.Start(Self, StrToInt(nodeData.M_Service), nodeData.D_Span)
  else
    TFrm_Smets_Vends.Start(Self, StrToInt(nodeData.M_Service), nodeData.D_Span);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.VendHistory2Click(Sender: TObject);
var
  dataNode : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Vends_Dcc.Start(Self, StrToInt(nodeData.M_Service), nodeData.D_Span);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.CommsData1Click(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Comms.Start(Self, StrToInt(nodeData.M_Service), nodeData.D_Span);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.MeterReadings2Click(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Readings.Start(Self, nodeData.D_Span);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.mnuProfileDataDCCClick(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Profile.Start(Self, StrToInt(nodeData.M_Service), nodeData.D_Span, nodeData.M_MeterId);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.Alarms1Click(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := treeview1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Alarms.Start(Self, nodeData.D_Span);
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.AlertsDCCClick(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := treeview1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Alarms_Dcc.Start(Self, nodeData.D_Span);
end;

procedure TFRM_Tree.AnnulElectricityContract1Click(Sender: TObject);
var
  SPAN : String;
begin
  if treeupdating then
  begin
    exit;
  end;

  mpannode:=treeview1.FocusedNode;
  nodeData := treeview1.GetNodeData(mpannode);

  SPAN  := nodedata.D_SPAN;

  AnnulmentFrm := TAnnulmentFrm.Create(self, SPAN, Elec);
  AnnulmentFrm.PopulateAnnulmentReasons;
  AnnulmentFrm.ShowModal;
  AnnulmentFrm.Free;
end;

procedure TFRM_Tree.AnnulGasContractClick(Sender: TObject);
var
  SPAN : String;
begin
  if treeupdating then
  begin
    exit;
  end;

  mpannode:=treeview1.FocusedNode;
  nodeData := treeview1.GetNodeData(mpannode);

  SPAN  := nodedata.D_SPAN;

  AnnulmentFrm := TAnnulmentFrm.Create(self, SPAN, Gas);
  AnnulmentFrm.PopulateAnnulmentReasons;
  AnnulmentFrm.ShowModal;
  AnnulmentFrm.Free;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.extMEssageHistory1Click(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData:= treeview1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Txt_History.Start(Self, nodeData.D_Span);
end;

procedure TFRM_Tree.SHowSmetsMeterCommsSupplier(xnode:Pvirtualnode;SPAN,Meterid,service,dateremoved,role:string);
Var
LastReaDReqDate,baldate,xx:String;
amount1 : Currency;
begin
 LastReaDReqDate:='15';   // Last 15 mins for Elec
 if service='1' then LastReaDReqDate:='30';  //Lsat 30 mins for Gas
  // Get Lastest Smets TreeData.
  // Read Meter


 startsmets(SPAN); // Get Latest Snapshot

 if AUTOREADMETER=true then
 Begin
  DoFUllSnapShot(SPAN,METERID,SERVICE,'',LastReaDReqDate,'','N');
 end;
 nodeData := treeview1.GetNodeData(xnode);

 xx:=NodeData.M_METERID;
 xx:=Nodedata.m_service;

 // now we need more info on the Meter Node to display
 // SUpply Status and Comms Status

         with main_data_module.smets_current_values do
         Begin
          close;
          if USE_DEDICATED_STAGING_CONNECTION=TRUE then SESSION:=FRM_LOGIN.STAGINGSESSION
          ELSE SESSION:=FRM_LOGIN.MainSession;

          deletevariables;
          declarevariable('SPAN',otstring);
          sql.clear;
          IF USE_DEDICATED_STAGING_CONNECTION=TRUE  THEN sql.add('select * from LIBERTY100.sn_current_values where servicepointno=:SPAN')
          ELSE sql.add('select * from LIBERTY100.sn_current_values@STAGINGDB where servicepointno=:SPAN');
          setvariable('SPAN',span);
          open;
          deletevariables;
         end;
         if (main_data_module.smets_current_values.recordcount<>0) and
            (main_data_module.smets_current_values.fields[5].text=meterid) then
            Begin
             Commsstatus:=main_data_module.smets_current_values.fields[64].text;
             if commsstatus='ON' then hi:=243
             else if commsstatus='OFF' then hi:=241
             else hi:=242;
             SupplyStatus:=main_data_module.smets_current_values.fields[66].text;
             if Supplystatus='ON' then si:=243
             else if supplyStatus='OFF' then si:=241
             else si:=242;

             desc:='Smart Meter ID - ('+main_data_module.smets_current_values.fields[13].text+') - '+MeterID+dateremoved;
             if commsstatus<>'ON' then desc:=desc+', - Remote Comms is '+commsstatus;
             if Supplystatus<>'ON' then desc:=desc+', - Supply is '+Supplystatus;

             // If balance is within last 30mins, then display.
             // Format To Currency Value.
             try
              baldate:=copy(main_data_module.smets_current_values.fields[55].text,1,19);
              if strtodatetime(baldate)>now -(30/1440) then
              begin
               amount1:=main_data_module.smets_current_values.fields[54].value;
               desc:=desc+' - Balance '+CurrToStrF(amount1,ffCurrency,2)+' ('+copy(baldate,length(baldate)-7,8)+')';
              end
             except
             end;

             nodedata.caption:=desc;
             nodedata.index:=239; // Liberty Smets 1 Meter
             nodedata.fontcolor:=clblack;
             // SHow Gray if NoCOmms or Supply
             if (commsstatus<>'ON') or (SupplyStatus<>'ON') then
             Begin
              nodedata.index:=240;
              nodedata.fontcolor:=clred;
             end;

             if commsstatus<>'ON' then
             begin
              MeterCommsNOde:=Treeview1.Addchild(MeterNode);
              nodeData := Treeview1.GetNodeData(MeterCommsNOde);
              NodeData.caption := 'Remote Comms Status: '+CommsStatus;
              nodedata.index:=hi;

              {MeterCommsNOde:=Treeview1.items.AddChild(MeterNode,'Comms Status: '+CommsStatus);
              MeterCommsNOde.selectedindex:=hi;
              MeterCommsNode.imageindex:=hi;}
             end;

             if supplystatus<>'ON' then
             Begin
              MeterSupplyNOde:=Treeview1.Addchild(MeterNode);
              nodeData := Treeview1.GetNodeData(MeterSupplyNOde);
              NodeData.caption := 'Supply Status: '+SupplyStatus;
              nodedata.index:=si;

             { MeterSupplyNOde:=Treeview1.items.AddChild(MeterNode,'Supply Status: '+SupplyStatus);
              MeterSupplyNode.imageindex:=si;
              MeterSupplyNode.selectedindex:=si;  }
             end;
            end
            else
            Begin
             //What if an old Meter
             desc:='Smart Meter ID - '+MeterID+dateremoved;
             nodedata.caption:=desc;
             nodedata.index:=239; // Liberty Smets 1 Meter
             nodedata.fontcolor:=clblack;
           end;
  endsmets;
end;


procedure TFRM_Tree.ShowSmartPayScreen(aCustId: Int64);
begin
  TFRM_POWER_PAY_HISTORY.Launch(aCustId, FRM_main_search.customercontacts.FieldByName('CUSTOMER_TYPE_ID').AsInteger,
    IsSmartPay(IntToStr(aCustId)), ((hidecheck.Visible = true) and (hidecheck.checked = true)));
end;

procedure TFRM_Tree.SHowSmetsMeterCommsMop(xnode:Pvirtualnode;SPAN,Meterid,service,dateremoved,role:string);
Var
LastReaDReqDate,baldate:String;
amount1 : Currency;
begin
 LastReaDReqDate:='15';   // Last 15 mins for Elec
 if service='1' then LastReaDReqDate:='30';  //Lsat 30 mins for Gas
  // Get Lastest Smets TreeData.
  // Read Meter


 startsmets(SPAN); // Get Latest Snapshot

 if AUTOREADMETER=true then
 Begin
  DoFUllSnapShot(SPAN,METERID,SERVICE,'',LastReaDReqDate,'','N');
 end;
 nodeData := MopTree.GetNodeData(xnode);

 // now we need more info on the Meter Node to display
 // SUpply Status and Comms Status

         with main_data_module.smets_current_values do
         Begin
          close;
          if USE_DEDICATED_STAGING_CONNECTION=TRUE then SESSION:=FRM_LOGIN.STAGINGSESSION
          ELSE SESSION:=FRM_LOGIN.MainSession;

          deletevariables;
          declarevariable('SPAN',otstring);
          sql.clear;
          IF USE_DEDICATED_STAGING_CONNECTION=TRUE  THEN sql.add('select * from LIBERTY100.sn_current_values where servicepointno=:SPAN')
          ELSE sql.add('select * from LIBERTY100.sn_current_values@STAGINGDB where servicepointno=:SPAN');
          setvariable('SPAN',span);
          open;
          deletevariables;
         end;
         if (main_data_module.smets_current_values.recordcount<>0) and
            (main_data_module.smets_current_values.fields[5].text=meterid) then
            Begin
             Commsstatus:=main_data_module.smets_current_values.fields[64].text;
             if commsstatus='ON' then hi:=243
             else if commsstatus='OFF' then hi:=241
             else hi:=242;
             SupplyStatus:=main_data_module.smets_current_values.fields[66].text;
             if Supplystatus='ON' then si:=243
             else if supplyStatus='OFF' then si:=241
             else si:=242;

             desc:='Smart Meter ID - ('+main_data_module.smets_current_values.fields[13].text+') - '+MeterID+dateremoved;
             if commsstatus<>'ON' then desc:=desc+', - Remote Comms is '+commsstatus;
             if Supplystatus<>'ON' then desc:=desc+', - Supply is '+Supplystatus;

             // If balance is within last 30mins, then display.
             // Format To Currency Value.
             try
              baldate:=copy(main_data_module.smets_current_values.fields[55].text,1,19);
              if strtodatetime(baldate)>now -(30/1440) then
              begin
               amount1:=main_data_module.smets_current_values.fields[54].value;
               desc:=desc+' - Balance '+CurrToStrF(amount1,ffCurrency,2)+' ('+copy(baldate,length(baldate)-7,8)+')';
              end
             except
             end;

             nodedata.caption:=desc;
             nodedata.index:=239; // Liberty Smets 1 Meter
             nodedata.fontcolor:=clblack;
             // SHow Gray if NoCOmms or Supply
             if (commsstatus<>'ON') or (SupplyStatus<>'ON') then
             Begin
              nodedata.index:=240;
              nodedata.fontcolor:=clred;
             end;

             if commsstatus<>'ON' then
             begin
              MeterCommsNOde:=MopTree.Addchild(MeterNode);
              nodeData := MopTree.GetNodeData(MeterCommsNOde);
              NodeData.caption := 'Remote Comms Status: '+CommsStatus;
              nodedata.index:=hi;

              {MeterCommsNOde:=Treeview1.items.AddChild(MeterNode,'Comms Status: '+CommsStatus);
              MeterCommsNOde.selectedindex:=hi;
              MeterCommsNode.imageindex:=hi;}
             end;

             if supplystatus<>'ON' then
             Begin
              MeterSupplyNOde:=MopTree.Addchild(MeterNode);
              nodeData := MopTree.GetNodeData(MeterSupplyNOde);
              NodeData.caption := 'Supply Status: '+SupplyStatus;
              nodedata.index:=si;

             { MeterSupplyNOde:=Treeview1.items.AddChild(MeterNode,'Supply Status: '+SupplyStatus);
              MeterSupplyNode.imageindex:=si;
              MeterSupplyNode.selectedindex:=si;  }
             end;
            end
            else
            Begin
             //What if an old Meter
             desc:='Smart Meter ID - '+MeterID+dateremoved;
             nodedata.caption:=desc;
             nodedata.index:=239; // Liberty Smets 1 Meter
             nodedata.fontcolor:=clblack;
           end;
  endsmets;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.ProfileData1Click(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Profile.Start(Self, StrToInt(nodeData.M_Service), nodeData.D_Span, nodeData.M_MeterId);
end;

procedure TFRM_Tree.COTWizard1Click(Sender: TObject);
Var
CustomerID,Premise_ID:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 Agreement_id:=TreeData.D_agreement_id;
 Customerid:=TreeData.D_Customer_id;
 Premise_id:=TreeData.D_Premise_id;
 Application.CreateForm(TFRM_COT_WIZARD, FRM_COT_WIZARD);
 try
 // FRM_FRM_COT_WIZARD.Premiseid.text:=Premise_ID;
 // FRM_FRM_COT_WIZARD.Agreementid.text:=Agreement_ID;
 FRM_COT_WIZARD.tag:=0;
 FRM_COT_WIZARD.ShowModal;
 if FRM_COT_WIZARD.tag=2 then
 Begin
  Messagedlg('COT Actions complete. Please Refresh Customer Tree',MTInformation,[MBOK],0);
 end;
 finally
 FRM_COT_WIZARD.release;
 end;

end;



procedure TFRM_Tree.CreateAgreementttoPay1Click(Sender: TObject);
begin
  xnode:=treeview1.FocusedNode;
  TreeData:= treeview1.GetNodeData(xnode);

  FrmArgToPay := TFrmArgToPay.Create(self);
  try
    with FrmArgToPay Do
    begin
      Customer_Id := TreeData.D_customer_id;
      Agreement_Id := TreeData.D_Agreement_id;
      Customer_Name := frm_common.GetCustomerNameFromId(TreeData.D_customer_id);
      Username := UserID;
      PageControl1.Activepage := tsCreatePlan;
      RefreshData;
      if (HasErrors = false) and CanUserAccess then
      begin
        CostEstimator;
        ShowModal;
      end;
    end;
  finally
    FrmArgToPay.Free;
  end;
end;

procedure TFRM_Tree.Treeview1GetImageIndex(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Kind: TVTImageKind; Column: TColumnIndex;
  var Ghosted: Boolean; var ImageIndex: TImageIndex);
begin
if not (Kind in [ikNormal, ikSelected]) then Exit;
  TreeData:= Sender.GetNodeData(Node);
  if Assigned(Treedata) then imageindex := Treedata.index;
end;

procedure TFRM_Tree.Treeview1GetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
    var CellText: String);
begin
  TreeData:= sender.GetNodeData(Node);
  if Assigned(Treedata) then
  Begin
   CellText := Treedata.caption;
  End;
end;

procedure TFRM_Tree.Treeview1PaintText(Sender: TBaseVirtualTree;
  const TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  TextType: TVSTTextType);
var
dbt:string;
begin
 TreeData:= sender.GetNodeData(Node);
 TargetCanvas.Font.Color:=clblack;
 TargetCanvas.Font.style:=[];
 TargetCanvas.Font.name:='Tahoma';
 if Assigned(Treedata) then
 Begin
  TargetCanvas.Font.Color := Treedata.fontcolor;
  TargetCanvas.Font.name := Treedata.fontname;
  dbt:=TreeData.D_CDEBT;
  if Treedata.fontbold=true then TargetCanvas.Font.style := Targetcanvas.Font.Style+[fsbold];
  if Treedata.fontunderline=true then TargetCanvas.Font.style := Targetcanvas.Font.Style+[fsunderline];
 end;

 if treeview1.Selected[node] then
 //TargetCanvas.Font.Color := clHighlightText;
  BEGIN
   TargetCanvas.Font.Color :=styleservices.GetSYSTEMColor(cLHIGHLIGHTTEXT);
  END;
  //Node.Align := 20; // Alignment of expand/collapse button nearly at the top of the node.
  if (copy(TreeData.caption,1,8)='Customer') and (dbt<>'') then treeview1.NodeHeight[Node] := 44
  else if copy(TreeData.caption,1,9)='Agreement' then treeview1.NodeHeight[Node] := 54
  else if copy(TreeData.caption,1,6)='Legacy' then treeview1.NodeHeight[Node] := 44
  else if copy(TreeData.caption,1,7)='(JBS ID' then treeview1.NodeHeight[Node] := 44
  else
    treeview1.NodeHeight[Node] := 34;
end;

Procedure TFRM_Tree.Treeview1Change(Sender: TBaseVirtualTree;
                                    Node  : PVirtualNode);
Var Sp, Et: Boolean;
Begin
  TreeData := Treeview1.GetNodeData(Node);
  // BSL - 12/05/2021 - Code Optimization.
  // BSL - 29/06/2021-CRM-510-Super Customer menu pops up for the last commercial account under the Super customer profile.
                           // Happens when menu is accessed first time after that shows correct menu
  If Assigned(Treedata) then
    Begin //exit;
      Treeview1.PopupMenu := Nil;

      If Copy(TreeData.Caption, 1, 8) = 'Customer' then
        Treeview1.PopupMenu := PopupCust
      Else
        If Copy(TreeData.Caption, 1, 14) = 'Super Customer' then
          Treeview1.PopupMenu := PopUpSuperCust
        Else
          If Copy(TreeData.Caption, 1, 8) = 'Prospect' then
            Treeview1.PopupMenu := PopUpProspect
          Else
            If Copy(TreeData.Caption, 1, 16) = 'Account Reviewed' then
              Treeview1.PopupMenu := PopUpReviewer
            Else
              If Copy(TreeData.Caption, 1, 22) = 'System Reviewer Placed' then
                Treeview1.PopupMenu := OnHoldPopUp
              Else
                If Copy(TreeData.Caption, 1, 14) = 'Account Holder' then
                  Treeview1.PopupMenu := PopUpAccount
                Else
                  If     (Copy(TreeData.Caption, 1, 14) = 'Tel No: Mobile')
                     or ((Copy(TreeData.Caption, 1, 14) = 'Tel No: Day - ') and (Copy(TreeData.Caption, 15, 2) = '07'))
                     or ((Copy(TreeData.Caption, 1, 14) = 'Tel No: eve - ') and (Copy(TreeData.Caption, 15, 2) = '07')) then
                    Treeview1.PopupMenu := PopupMobile
                  Else
                    If Copy(TreeData.Caption, 1, 23) = 'Unactioned Pending Loss' then
                      Treeview1.PopupMenu := Popup_Losses
                    Else
                      // BSL - 15/12/2014 - Change JBS from Web to CRM.
                      If Copy(TreeData.Caption, 1, 8) = '(JBS ID:' then
                        Treeview1.PopupMenu := pupRescheduling
                      Else
                        If Copy(TreeData.Caption, 1, 2) = '(N' then
                          Treeview1.PopupMenu := PopupNote
                        Else
                          If Copy(TreeData.Caption, 1, 9) = 'Agreement' then
                            Treeview1.PopupMenu := PopupAgreements
                          Else
                            If Copy(TreeData.Caption, 1, 23) = 'DD Catch Ups Suppressed' then
                              Treeview1.PopupMenu := PopupSuppress
                            Else
                              If Copy(TreeData.Caption, 1, 21) = '** Account In Dispute' then
                                Treeview1.PopupMenu := PopUpDispute
                              Else
                                If Copy(TreeData.Caption, 1, 8) = 'Premises' then
                                  Begin
                                    Sp                  := TreeData.Index = 110;
                                    Treeview1.PopupMenu := PopupPremise;
                                    P1.Enabled          := Not Sp;
                                    P1.Visible          := Not Sp;
                                    P3.Enabled          := Not Sp;
                                    P3.Visible          := Not Sp;
                                  End // If
                                Else
                                  If    (Copy(TreeData.Caption, 1, 13) = 'Statement For')
                                     or (Copy(TreeData.Caption, 1, 14) = '* FINAL BILL *')
                                     or (Copy(TreeData.Caption, 1,  8) = '* BILL *')
                                     or (Copy(TreeData.Caption, 1, 15) = 'Posted Invoices') then
                                    Treeview1.PopupMenu := PopUpRated
                                  Else
                                    If (UpperCase(Copy(TreeData.Caption, 1, 26)) = UpperCase('My Utilita Savings Balance')) then
                                    begin
                                      SavingsTransMenuItem.Caption := 'Savings Transactions';
                                      Treeview1.PopupMenu := PopUpSmartPay;
                                    end
                                      Else If (UpperCase(Copy(TreeData.Caption, 1, 9)) = Uppercase('Power Pot')) then
                                      begin
                                        SavingsTransMenuItem.Caption := 'Power Pot Transactions';
                                        Treeview1.PopupMenu := PopUpSmartPay;
                                      end;


      // Wrike Ecards
      //if Treedata.index=20 then  treeview1.popupmenu:=popupemail;

      // Check if Deceased
      if Treedata.index=43 then
      Begin
       A_dec.enabled:=false;
       a_dec.visible:=false;
       a_rev.enabled:=true;
       a_rev.visible:=true;
      end
      else
      Begin
       A_dec.enabled:=true;
       a_dec.visible:=true;
       a_rev.enabled:=false;
       a_rev.visible:=false;
      end;

      // Determine Which Span Node is Selected
      try

       SPANTYPE:=TreeData.D_SPANTYPE;
       if (SPANTYPE='G') or (SPANTYPE='C') then
       Begin
        treeview1.popupmenu:=PopupGAS;
        ET:=TreeData.D_ET;
        if et=true then
        begin
         g_set.Visible:=false;
         g_ret.visible:=true;
        end
        else
        begin
         g_set.Visible:=true;
         g_ret.visible:=false;
        end;
        g_set.enabled:=g_set.visible;
        g_ret.enabled:=g_ret.visible;
       end;
       if (SPANTYPE='E') or (SPANTYPE='F') then
       Begin
        treeview1.popupmenu:=PopupELECTRIC;
        ET:=TreeData.D_ET;
        if et=true then
        begin
         e_set.Visible:=false;
         e_ret.visible:=true;
        end
        else
        begin
         e_set.Visible:=true;
         e_ret.visible:=false;
        end;
        e_set.enabled:=e_set.visible;
        e_ret.enabled:=e_ret.visible;
       end;
       if (SPANTYPE='T') or (SPANTYPE='J') then treeview1.popupmenu:=PopupTELECOM;
       if (SPANTYPE='Y') then treeview1.popupmenu:=PopupBroadBand;

       if (SPANTYPE='A')
       or (SPANTYPE='B')
       or (SPANTYPE='H')
       or (SPANTYPE='I')
       or (SPANTYPE='K')
       or (SPANTYPE='L')
       or (SPANTYPE='U')
       or (SPANTYPE='V')
       or (SPANTYPE='W')
       or (SPANTYPE='X')
       or (SPANTYPE='M')
       or (SPANTYPE='R')
       or (SPANTYPE='6')
       or (SPANTYPE='S') then
       Begin
        custom_meter.visible:=true;
        treeview1.popupmenu:=PopupCustom;
        if spantype='M' then custom_meter.visible:=false;
        custom_meter.enabled:=custom_meter.visible;
       end;

      except
      end;

      // BSL - 30/07/2021 - Code Optimization
      // BSL - 29/06/2021-CRM-510-Super Customer menu pops up for the last commercial account under the Super customer profile.
                               // Happens when menu is accessed first time after that shows correct menu
      If (TreeData.Index = 54)  or
         (TreeData.Index = 55)  or
         (TreeData.Index = 56)  or
         (TreeData.Index = 60)  or
         (TreeData.Index = 61)  or
         (TreeData.Index = 62)  or
         (TreeData.Index = 63)  or
         (TreeData.Index = 263) or
         (TreeData.Index = 264) or
         (TreeData.Index = 265) or
         (TreeData.Index = 266) or
         (TreeData.Index = 267) or
         (TreeData.Index = 268) then
      Begin
        Treeview1.PopupMenu     := PopupEnquiry;
        E_TakeOwnership.Enabled := Treedata.C_Owner = EmptyStr;
      End; // If

//      if (TreeData.index=263) or
//         (TreeData.index=264) or
//         (TreeData.index=265) or
//         (TreeData.index=266) or
//         (TreeData.index=267) or
//         (TreeData.index=268) then
//      Begin
//       treeview1.popupmenu:=PopupEnquiry;
//       if Treedata.C_Owner<>'' then E_Takeownership.enabled:=false
//       else e_takeownership.enabled:=true;
//
//      end;

      case TreeData.Index of
        239,
        240 : begin
                treeview1.popupmenu := PopupSmets;
              end;
        205,
        313,
        321 : begin
                treeview1.popupmenu := PopupSMETS_DCC;
                PopupSMETS_DCC.Items[0].ImageIndex := TreeData.Index;
              end;
      end;

    End; // If
End; // Proc

procedure TFRM_Tree.FormCreate(Sender: TObject);
begin
 treeview1.OnGetText := treeview1GetText;
 treeview1.NodeDataSize := SizeOf(TMyRec);

 moptree.OnGetText := MopTreeGetText;
 moptree.NodeDataSize := SizeOf(TMyRec);

 showfeedback:=true;
 fPremiseDcc:= false;

 fAgreementsToReassign := TAgreementCheckedNodeManager.Create(treeView1);
end;

procedure TFRM_Tree.FormDestroy(Sender: TObject);
begin
  fAgreementsToReassign.Free;
end;

procedure TFRM_Tree.FormShow(Sender: TObject);
begin
  fCustomerNote := EmptyStr;
  fCustomerAccountHolder := EmptyStr;
  fCustomerPriorityNotification := EmptyStr;
  mniChangeOfTenancy.Enabled := TCrmUtil.HasUserFeature(UserId, USER_FEATURE__ALLOW_CHANGE_OF_TENANCY);
end;

// BSL - 29/06/2021-CRM-510-Super Customer menu pops up for the last commercial account under the Super customer profile.
                         // Happens when menu is accessed first time after that shows correct menu
Procedure TFRM_Tree.Treeview1Click(Sender: TObject);
Begin
 //treeview1.popupmenu:=nil;
  XNode    := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  // DANBYT - 04/12/2024 - CRMX-164 - Ensure CustId persists even if TreeData.D_Customer_Id is empty
  PersistStringValue(CustID,TreeData.D_Customer_Id);

  If Assigned(TreeData) then
    Begin
      // Customer
      FRM_Main_Search.pUpdateCustomer(TreeData.D_Customer_Id);

      If Copy(TreeData.Caption, 1, 8) = 'Customer' then
        Treeview1.PopupMenu := PopupCust
      Else
        If Copy(TreeData.Caption, 1, 14) = 'Super Customer' then
          Treeview1.PopupMenu := PopUpSuperCust;
    End; // If
End; // Proc

{------------------------------------------------------------------------------}
procedure TFRM_Tree.Treeview1Expanding(Sender: TBaseVirtualTree; Node: PVirtualNode; var Allowed: Boolean);
var
  s                : string;
  dataProtectionOK : boolean;
begin
  TreeData := Treeview1.GetNodeData(Node);

  if Assigned(TreeData) then
  begin
    fIsExpanding := true;
    try
      s                := Copy(TreeData.Caption, 1, 19);
      spantype         := TreeData.D_spantype;

      // DANBYT - 04/12/2024 - CRMX-164 - Ensure CustId persists even if TreeData.D_Customer_Id
      // is empty. Before this fix if a user navigated to a Meter node from the tree top and then
      // back to the tree top again the CustID would have been set to empty at the Meter node
      PersistStringValue(CustID,TreeData.D_Customer_Id);

      dataProtectionOK := true;

      if (Copy(TreeData.Caption, 1, 8) = 'Customer') and (CustId <> '') then
        ExpandCustomerNode(Sender, Node, dataProtectionOK)
      else if Copy(TreeData.caption, 1, 9) = 'Agreement' then
        RefreshAgreementNode(node)      // Agreements
      else if Copy(TreeData.caption, 1, 13) = 'List Of Sites' then
        RefreshAgentPremiseNode         // Sale Agents Premise
      else if Copy(TreeData.caption, 1, 8) = 'Premises' then
        RefreshpremiseNode(node);

      if dataProtectionOK then
      begin
        if (spantype = 'G') or (spantype = 'C') or (spantype = 'E') or (spantype = 'F') or
           (spantype = 'T') or (spantype = 'J') or (spantype = 'Y') or
           (spantype = 'A') or (spantype = 'B') or (spantype = 'H') or (spantype = 'I') or
           (spantype = 'K') or (spantype = 'L') or (spantype = 'U') or (spantype = 'V') or
           (spantype = 'W') or (spantype = 'X') or (spantype = 'R') or (spantype = 'S') or
           (spantype = '6') or (spantype = '7') or (spantype = 'M') then
          ExpandSpanNode(Node);

        ShowCursor(true);

        if ShowReassign and (Copy(TreeData.caption, 1, 8) = 'Customer') then
          doExpandAgreements(node);

        UAandFPs1.Visible := (agreement_id = fSuspenseAGID) and (main.isFinance);
        UAdivider.Visible := UAandFPs1.Visible;
      end;
    finally
      fIsExpanding := false;
    end;
  end;
end;

procedure TFRM_Tree.MopTreePaintText(Sender: TBaseVirtualTree;
  const TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  TextType: TVSTTextType);
begin
 TreeData:= sender.GetNodeData(Node);
 TargetCanvas.Font.Color:=clblack;
 TargetCanvas.Font.style:=[];
 TargetCanvas.Font.name:='Tahoma';
 if Assigned(Treedata) then
 Begin
  TargetCanvas.Font.Color := Treedata.fontcolor;
  TargetCanvas.Font.name := Treedata.fontname;
  if Treedata.fontbold=true then TargetCanvas.Font.style := Targetcanvas.Font.Style+[fsbold];
  if Treedata.fontunderline=true then TargetCanvas.Font.style := Targetcanvas.Font.Style+[fsunderline];
 end;
 if moptree.Selected[node] then  TargetCanvas.Font.Color :=styleservices.GetSYSTEMColor(cLHIGHLIGHTTEXT);
end;

procedure TFRM_Tree.MopTreeInitNode(Sender: TBaseVirtualTree; ParentNode, Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
begin
  //with treeview1 do
  with moptree do
  begin
    TreeData:= GetNodeData(Node);
  end;
 // Node.Align := 20; // Alignment of expand/collapse button nearly at the top of the node.
  if copy(TreeData.caption,1,9)='Agreement' then moptree.NodeHeight[Node] := 54
  else moptree.NodeHeight[Node] := 34;
  Include(InitialStates, ivsMultiline);
end;

procedure TFRM_Tree.MopTreeGetImageIndex(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Kind: TVTImageKind; Column: TColumnIndex;
  var Ghosted: Boolean; var ImageIndex: TImageIndex);
begin
if not (Kind in [ikNormal, ikSelected]) then Exit;
  TreeData:= Sender.GetNodeData(Node);
  if Assigned(Treedata) then imageindex := Treedata.index;
end;

procedure TFRM_Tree.MopTreeGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
var CellText: String);
begin
  TreeData:= sender.GetNodeData(Node);
  if Assigned(Treedata) then
  Begin
   CellText := Treedata.caption;
  End;
end;

procedure TFRM_Tree.MopTreeExpanding(Sender: TBaseVirtualTree;
  Node: PVirtualNode; var Allowed: Boolean);
var
span,filename,ETDMOA:string;
contracterror:boolean;
begin
 mynodecustomer:=node;
 TreeData:= moptree.GetNodeData(mynodecustomer);
 FILENAME:=TreeData.D_FILENAME;
 Contracterror:=TreeData.D_ContractError;
 ETDMOA:=TreeData.D_ETDMOA;
 span:=TreeData.D_SPAN;

 //mynodecustomer:=node;
 //mynodecustomer.selected:=true;

 if Copy(TreeData.caption,1,4)='MPAN' then
 Begin
  treeupdating:=true;
  moptree.deletechildren(node);
  // Add Notes/Enquires
  BuildMOPNOTES(SPAN);
  // Add Meter Node
  buildElectricMeterNodeMOP(SPAN,ETDMOA);
  // Add Site Node
  BuildMopSiteAddress(Span,filename,contracterror);
 end

 else
 // Premise
 if Copy(TreeData.caption,1,8)='Premises' then
 Begin
  treeupdating:=true;
  RefreshMopPremiseNode(node);
 end;

 treeupdating:=false;
end;

procedure TFRM_Tree.MopTreeChange(Sender: TBaseVirtualTree;
  Node: PVirtualNode);
begin
 xnode:=moptree.FocusedNode;
 TreeData:= Moptree.GetNodeData(xnode);
 if assigned(Treedata)=false then exit;

   // Account Review   on hold
 //if treeupdating=true then exit;
 MopTree.popupmenu:=nil;

 if Copy(TreeData.caption,1,5)='MPAN:' then
 Begin
   MopTree.popupmenu:=PopUpElectricMOP;
 end
 else
 if Copy(TreeData.caption,1,2)='(S' then
 Begin
   MopTree.popupmenu:=PopUpSO;
 end;
end;

procedure TFRM_Tree.Treeview1InitNode(Sender: TBaseVirtualTree; ParentNode, Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
begin
  try
    with Treeview1 do
    begin
      TreeData := GetNodeData(Node);
    end;
    if (Copy(TreeData.caption,1,9) = 'Agreement') and (SHOWREASSIGN = true) and (Copy(TreeData.caption,44,4)<>'Term')
 then
    begin
      node.CheckType := virtualtrees.ctCheckBox;
      if (fAgreementsToReassign.IndexOf[TreeData.D_Agreement_ID] >= 0) then
        node.CheckState := csCheckedNormal;
    end;
  except
  end;
  Include(InitialStates, ivsMultiline);
end;

procedure TFRM_Tree.Smets_LoanClick(Sender: TObject);
begin
 DoSmets2Credit(5);
end;

procedure TFRM_Tree.UAandFPs1Click(Sender: TObject);
begin
 Application.CreateForm(TFRM_Status_Change, FRM_Status_Change);
 try
  FRM_Status_Change.UserID := UserID;
  FRM_Status_Change.showmodal;
 finally
  FRM_Status_Change.release;
 end;
end;

procedure TFRM_Tree.UnAllocateLibertyCUstomerID1Click(Sender: TObject);
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 custid:=TreeData.D_Customer_id;
 with main_Data_module.tempquery do
 Begin
  deletevariables;
  close;
  sql.clear;
  sql.add('select liberty_customer_id from crm.customer_to_liberty_customer where customer_id='+custid);
  open;

 End;
 if main_Data_module.tempquery.recordcount=0 then
 Begin
  Messagedlg('There is currently No Liberty Customer ID allocated to the account?',mtinformation,[mbok],0);
  exit;
 end;

 if Messagedlg('Are you sure you wish to Un-Allocate Liberty Customer ID '+main_Data_module.tempquery.fields[0].text+' from this account?',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;

 with main_data_module.updatequery do
 Begin
  close;
  sql.clear;
  sql.add('delete from crm.customer_to_liberty_customer where customer_id='+custid);
  execute;
 end;
 FRM_Common.SetAudit(CUSTID,'','','Customer Un-assigned from Liberty Customer - '+main_Data_module.tempquery.fields[0].text);
 frm_login.mainsession.commit;
 FRM_Main.SearchForcust(custid);
 xNode := Treeview1.GetFirst();
 treeview1.expanded[xnode]:=false;
 treeview1.expanded[xnode]:=true;
end;

procedure TFRM_Tree.SPANOverride3Click(Sender: TObject);
begin
DoSpanOverride;
end;




procedure TFRM_Tree.doCopyNotestoAnotherCustomer(copyTo: string);
var
	CustomerID:string;
begin
 	xnode:=treeview1.FocusedNode;
 	TreeData:= treeview1.GetNodeData(xnode);

 	if treeupdating=true then
  	exit;

 	CustomerID:=TreeData.D_Customer_ID;
 	Application.CreateForm(TFRM_COPY_NOTES, FRM_COPY_NOTES);
 	try
  	frm_Main.InitialiseUnfocusedSelectionColour(FRM_COPY_NOTES.NotesTree);

  	frm_copy_notes.enquiries.close;
    frm_copy_notes.enquiries.setvariable('CID',customerid);
  	frm_copy_notes.enquiries.open;
  	frm_copy_notes.newcid.text:=copyTo;
//  	frm_copy_notes.btn_copy.enabled:=false;
    frm_copy_notes.btn_check.Click;
  	if frm_copy_notes.enquiries.recordcount<>0 then
  	begin
   		FRM_COPY_NOTES.buildtree;
   		FRM_COPY_NOTES.showmodal
  	end
  	else
    	ShowMessage('There are no Enquiries, notes or Docs for this customer');
 	finally
  	FRM_COPY_NOTES.release;
 	end;
end;

procedure TFRM_Tree.CopyBillClick(Sender: TObject);
begin
 Application.CreateForm(TFRM_Super_Cust_Bill, FRM_Super_Cust_Bill);
 try
   FRM_Super_Cust_Bill.SuperCustID := StrToInt64(Treedata.D_Customer_ID);
  FRM_Super_Cust_Bill.showmodal;
 finally
  FRM_Super_Cust_Bill.release;
 end;
end;

procedure TFRM_Tree.CopyNotestoAnotherCustomer1Click(Sender: TObject);
var
	CustomerID:string;
begin
 	xnode:=treeview1.FocusedNode;
 	TreeData:= treeview1.GetNodeData(xnode);

 	if treeupdating=true then
  	exit;

 	CustomerID:=TreeData.D_Customer_ID;
 	Application.CreateForm(TFRM_COPY_NOTES, FRM_COPY_NOTES);
 	try
  	frm_Main.InitialiseUnfocusedSelectionColour(FRM_COPY_NOTES.NotesTree);

  	frm_copy_notes.enquiries.close;
    frm_copy_notes.enquiries.setvariable('CID',customerid);
    if CustomerID = intToStr(3896597383) then
    begin
      with frm_copy_notes.Enquiries do
     begin
       close;
       DeleteVariables;
       sql.Clear;
       sql.Add('SELECT');
       sql.Add('E.MPANCORE, E.RAISED_BY, E.DATE_RAISED, C.DESCRIPTION "C.DESCRIPTION", E.DUE_DATE, R.DESCRIPTION "R.DESCRIPTION", E.COMMENTS_1, E.COMMENTS_2,');
       sql.Add('E.RESOLVED, E.RESOLVED_COMMENTS, E.RESOLVED_BY, E.DATE_RESOLVED, E.OWNER,E.ROWID, E.TIMED_OUT, R.ENQIURY_OR_NOTE, E.CUSTOMER_ID "E.CUSTOMER_ID",');
       sql.Add('E.SITE_ID, CU.LEGAL_ENTITY_NAME, S.PREMISE_POSTCODE, E.RECORD_ID "E.RECORD_ID", R.ICON_INDEX, E.SYSTEM_ROLE, E.EXPIRY_CODE, R.DEPT_ID,');
       sql.Add('D.DESCRIPTION "D.DESCRIPTION", E.CONTACT_TYPE, E.REQUEST_TYPE, SC.DESCRIPTION "SC.DESCRIPTION", EC.RAG_STATUS, CS.DESCRIPTION "CS.DESCRIPTION"');
       sql.Add('FROM');
       sql.Add('ENQUIRY.ENQUIRIES E, ENQUIRY.CONTACT_TYPE C, ENQUIRY.REQUEST_TYPE R, ENQUIRY.DEPARTMENTS D, ENQUIRY.SUB_CATEGORY SC, CRM.CUSTOMER CU,');
       sql.Add('CRM.PREMISES S, ENQUIRY.COMPLAINT_DETAIL EC, ENQUIRY.COMPLAINT_STATUS cs');
       sql.Add('WHERE');
       sql.Add('R.DEPT_ID = D.ID AND');
       sql.Add('E.CONTACT_TYPE = C.ID AND');
       sql.Add('E.REQUEST_TYPE = R.ID AND');
       sql.Add('E.CUSTOMER_ID = CU.CUSTOMER_ID (+) AND');
       sql.Add('E.SITE_ID = S.PREMISE_ID (+) AND');
       sql.Add('E.RECORD_ID = EC.RECORD_ID (+) AND');
       sql.Add('EC.SUB_CATEGORY_ID = SC.SUB_CATEGORY_ID (+) AND');
       sql.Add('EC.COMPLAINT_STATUS_ID = CS.COMPLAINT_STATUS_ID (+) AND');
       sql.Add('E.DATE_RAISED >= TRUNC(SYSDATE-90) AND');
       sql.Add('E.CUSTOMER_ID = 3896597383');
       sql.Add('ORDER BY E.DATE_RAISED DESC');
       open;
       DeleteVariables;
      end;
    end

    else
    begin
  	  frm_copy_notes.enquiries.open;
    end;

  	frm_copy_notes.newcid.text:='';
  	frm_copy_notes.btn_copy.enabled:=false;
  	if frm_copy_notes.enquiries.recordcount<>0 then
  	begin
   		FRM_COPY_NOTES.buildtree;
   		FRM_COPY_NOTES.showmodal
  	end
  	else
    	ShowMessage('There are no Enquiries, notes or Docs for this customer');
 	finally
  	FRM_COPY_NOTES.release;
 	end;
end;

procedure TFRM_Tree.NoInstallTariffManagment1Click(Sender: TObject);
var
agreement_id,issuereason,issuereason1:string;
clickedok:boolean;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Agreement_id:=TreeData.D_agreement_id;

 Application.CreateForm(TFRM_NO_INST_PROCESS, FRM_NO_INST_PROCESS);
 try
  FRM_NO_INST_PROCESS.agreements.close;
  FRM_NO_INST_PROCESS.agreements.setvariable('agreement_id',agreement_id);
  FRM_NO_INST_PROCESS.agreements.open;
  if FRM_NO_INST_PROCESS.agreements.recordcount<>0 then FRM_NO_INST_PROCESS.showmodal
  else
  begin
   if messagedlg('Are you sure you wish to add this agreement to the No Install Tariff Process?'+#13+
                 'This will cause a letter to be issued on the next batch run and the tariff will be updated automtically 30 days later.',mtconfirmation,[mbyes,mbno],0)<>mryes then exit;

   IssueReason:='';
   repeat
    ClickedOK := InputQuery('Reason For Change of Tariff', 'Please Enter a Reason for requesting this Tariff Change', IssueReason);
    if not ClickedOK then exit;
   until (clickedok=false) or (IssueReason<>'');
   IssueReason1:=stringreplace(IssueReason,'''','''''',[rfreplaceall]);

   if frm_common.authoritycheck=false then exit;

   with main_data_module.updatequery do
   Begin
    close;
    sql.clear;
    sql.add('insert into crm.no_install_tariff_process values('+agreement_id+',sysdate,user,'''+issuereason1+''',1,null,null,sysdate,user)');
    execute;
   end;
   frm_login.mainsession.commit;
   FRM_NO_INST_PROCESS.agreements.close;
   FRM_NO_INST_PROCESS.agreements.open;
   if FRM_NO_INST_PROCESS.agreements.recordcount<>0 then FRM_NO_INST_PROCESS.showmodal
  end;
 finally
  FRM_NO_INST_PROCESS.release;
 end;
end;

procedure TFRM_Tree.ShowScript1Click(Sender: TObject);
Var
priority,InObjPeriod,ref,purpose,SUMMARY,MSG:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 Priority:=TreeData.D_priority;
 InObjPeriod:=TreeData.D_InObjPeriod;

 // Need Something here to pick up Ref 'In OBJ' if still in objection Window.

 if InObJPeriod<>'Y' then
 Begin
  if priority='T' then ref:='Priority 1';
  if priority='U' then ref:='Priority 2';
  if priority='V' then ref:='Priority 3';
  if priority='W' then ref:='Priority 4';
 end
 else ref:='In OBJ';

 with main_data_module.tempquery do
 Begin
  deletevariables;
  declarevariable('REF',otstring);
  close;
  sql.clear;
  sql.add('Select * from crm.losses_retention_scripts where reference=:REF');
  setvariable('REF',ref);
  open;
  deletevariables;
 end;
 if main_data_module.tempquery.recordcount<>0 then
 Begin
  PURPOSE:=main_data_module.tempquery.fields[3].text;
  MSG:=main_data_module.tempquery.fields[2].text;
  SUMMARY:=main_data_module.tempquery.fields[1].text;
  FRM_HELPER_MESSAGE.SHOW_DETAILS(REF,PURPOSE,SUMMARY,MSG,'');
  FRM_HELPER_MESSAGE.SHOW;
 end
 else
 begin
  FRM_HELPER_MESSAGE.close;
 end;
end;

procedure TFRM_Tree.ShowSMSLog1Click(Sender: TObject);
Var
Customerid:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;
 CustomerID:=TreeData.D_Customer_ID;
 Application.CreateForm(TFRM_SMS_HISTORY, FRM_SMS_HISTORY);
 try
  FRM_SMS_HISTORY.GetMessagelist(CustomerID);
  if main_data_module.tempquery.recordcount<>0 then FRM_SMS_HISTORY.showmodal
  else
  begin
   messagedlg('There is no SMS history for this customer',mtconfirmation,[mbok],0);
  end;
 finally
  FRM_SMS_HISTORY.release;
 end;
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.S_COSLOSSEClick(Sender: TObject);
Var
  regId              : string;
  span               : string;
  etd                : string;
  spanType           : string;
  service            : integer;
  smetsFormId        : integer;
  smetsCosGainFormId : integer;
begin
  XNode    := TreeView1.FocusedNode;
  TreeData := TreeView1.GetNodeData(XNode);
  regId    := TreeData.D_REGid;
  span     := TreeData.D_SPAN;
  etd      := TreeData.D_SPANEND;
  spanType := TreeData.D_SPANTYPE;
  service  := 0;                   // Default to ELEC

  if (spanType = 'A') or (spanType = 'G') then
    service := 1;    // Gas

  if etd = '' then
  begin
    if MessageDlg('End Date is BLANK. Default to Today''s Date?', mtConfirmation, [mbYes, mbNo],0) <> mrYes then
      exit;

    etd := DateToStr(now);
  end;

  CosGainStatus(SPAN);

  if main_data_module.tempquery.recordcount <> 0 then
  begin
    if main_data_module.tempquery.fields[15].Text <> '-2' then
    begin
      // Cos In Progess and user doesnt have permission to Reorder
      if not USER_FEATURE__Reorder_Smets_Cos then
      begin
        Messagedlg('You cannot initiate a COS LOSS request for this supply.' + #13 + 'A COS request is already in progress.', mtWarning, [mbOK], 0);

        TFrm_Smets.ShowSmets(service, span, smetsFormId, smetsCosGainFormId);

        exit;
      end
      else
      begin
       // Cos In Progresss, user can force through another Request
        if MessageDlg('A COS Request is already in progress.'+#13+'Do you wish to Initiate another COS LOSS Request?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
        begin
          TFrm_Smets.ShowSmets(service, span, smetsFormId, smetsCosGainFormId);

          exit;
        end;
      end;
    end;
  end;

  if MessageDlg('Confirm you wish to Initiate a SMETS COS LOSS request for supply number: '+span+#13+'COS END DATE '+etd, mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
    exit;

  if not Frm_Common.SuperAuthorityCheck then
    exit;

  if not SupplierSwitchLoss(IntToStr(service), span, etd) then
  begin
    Messagedlg('There was an error submitting this request', mtError, [mbOK], 0);
  end
  else
  begin
    Messagedlg('COS request has been initiated.', mtInformation, [mbOK], 0);
    TFrm_Smets.ShowSmets(service, span, smetsFormId, smetsCosGainFormId);
  end;
end;

procedure TFRM_Tree.InitiateCOSRequestGAIN1Click(Sender: TObject);
Var
  span,ssd,spantype,service:string;
begin
  xnode:=treeview1.FocusedNode;
  TreeData:= treeview1.GetNodeData(xnode);
  SPAN:=TreeData.D_SPAN;
  SSD:=TreeData.D_SSD;

  MeterTypeSwitch(span, ssd);
end;

procedure TFRM_Tree.MeterTypeSwitch(span, ssd: String);
var
  isDCCMeterStg: string;
begin

  if (USER_FEATURE__REORDER_SMETS_COS) then

  begin
    if (frm_common.Superauthoritycheck = false) then
    begin
      Messagedlg('Your user is not authorized to request COS GAIN' ,mtWarning,[mbOK],0);
      exit;
    end;
  end
  else
  begin
    Messagedlg('Your user is not authorized to request COS GAIN' ,mtWarning,[mbOK],0);
    exit;
  end;

  if strtodate(ssd) < now then ssd := datetostr(now);

  if MessageDlg('Confirm you wish to Initiate a SMETS COS request for supply number: '+SPAN+#13+
            'COS will be effective from '+SSD,mtConfirmation,[mbyes,mbNo],0)<>mryes
    then Exit;

  if isDCCMeter() then
    isDCCMeterStg := 'Y'
  else
    isDCCMeterStg := 'N';

  Messagedlg('COS request has been initiated.',mtInformation,[mbOK],0);

  try
    MeterTypeSwitchPerform(span, ssd, userID, isDCCMeterStg);

    Messagedlg('COS request done successfully.',mtInformation,[mbOK],0);
  except
    Messagedlg('There was an error submitting this request',mtError,[mbOK],0);
    exit;
  end;
end;

procedure TFRM_Tree.GovernmentEnergySchemesScreen(xSchemes: String);
var
  bOpenGES: Boolean;
begin
  // BSL - 10/12/2021 - BRM-1513 CORE customers status not showing in GES
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);

  bOpenGES := True;
  if not TreeUpdating then
  begin
    Application.CreateForm(TFRM_GOV, FRM_GOV);
    try
      FRM_GOV.pCustomerName := TreeData.D_Customer_Name;
      FRM_GOV.pSchemes := xSchemes;
      FRM_GOV.RefreshAll(TreeData.D_Customer_Id);

      if xSchemes = 'GER' then
      begin
        if not FRM_GOV.TabSheet_GER.TabVisible then
        begin
          MessageDlg('There is no information for "Government Electricity Rebate"', mtInformation, [mbOK], 0);
          FRM_GOV.Close;
          bOpenGES := False;
        end;
      end
      else
      if xSchemes = 'WHD' then
      begin
        if not FRM_GOV.TabSheet_WHD.TabVisible then
        begin
          MessageDlg('There is no information for "Warm Home Discount"', mtInformation, [mbOK], 0);
          FRM_GOV.Close;
          bOpenGES := False;
        end;
      end
      else
      if xSchemes = 'SLC14' then
      begin
        if not FRM_GOV.TabSheet_SLC14.TabVisible then
        begin
          MessageDlg('There is no information for "SLC14 Objection Refund Process"', mtInformation, [mbOK], 0);
          FRM_GOV.Close;
          bOpenGES := False;
        end;
      end
      else
      if xSchemes = 'WWP' then
      begin
        if not FRM_GOV.tabWWP.TabVisible then
        begin
          MessageDlg('There is no information for "Winter Warmer Payment"', mtInformation, [mbOK], 0);
          FRM_GOV.Close;
          bOpenGES := False;
        end;
      end
      else
      if xSchemes = 'EBSS' then
      begin
        if not FRM_GOV.TabEBSS.TabVisible then
        begin
          MessageDlg('There is no information for "Energy Bills Support Scheme"', mtInformation, [mbOK], 0);
          FRM_GOV.Close;
          bOpenGES := False;
        end;
      end
      else
      if xSchemes = 'AFP' then
      begin
        if not FRM_GOV.tsAFP.TabVisible then
        begin
          MessageDlg('There is no information for "Alternative Fuel Payment"', mtInformation, [mbOK], 0);
          FRM_GOV.Close;
          bOpenGES := False;
        end;
      end;

      if bOpenGES then
        FRM_GOV.ShowModal;
    finally
      FRM_GOV.Release;
    end;
  end; // If
end;

procedure TFRM_Tree.HH1Click(Sender: TObject);
Var
CustID:String;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);
 FRM_Enquiry_Note.tag:=3;
 Custid:=TreeData.D_Customer_ID;
  //*************************
  //changed by maryam on 05/05/2016 for HH requested by Rosie & Martin
  //************************
 Application.CreateForm(TFRM_HHAuditLog, FRM_HHAuditLog);
 try
  FRM_HHAuditLog.custid := CustID;
  FRM_HHAuditLog.ShowModal;
 finally
  FRM_HHAuditLog.release;
 end;



 treeview1.expanded[xnode]:=false;
 treeview1.expanded[xnode]:=true;
end;

procedure TFRM_Tree.StartComplaintsProcess(Customerid,method:string);
var
  msg,msg1,msg2,rpt:string;
  result:Integer;
  selection,our_ref:String;
begin

  // D-direct usually on Phone
  // I = Indeirect

  our_ref:='CS-25';
  rpt:='MASTER\CS-25_Complaints_procedure.rpt';
  selection:='{LETTER_DETAILS.OUR_REF}='''+our_ref+'''  and {CUSTOMER_MAILING.CUSTOMER_ID}='+Customerid;
  if method='D' then
  begin
   // usally Triggered from Enquiry, Typically Customer is on th ePHone with Customer

   //msg:='Would you like me to provide you with information on our complaints handling procedure?';
   msg:='';
   msg2:='';
   //msg1:='Do you wish to read out the complaints procedure VERBALLY?';
   //msg2:='Would you like us to send you a copy of our complaints procedure?';
   msg2:='';
  end
  else
  Begin
   // Deafalt Action When Triggered From Customer Tree
   msg:='Do you wish to send a copy of our complaints handling procedure?';
   msg1:='How would you like this information provided?';
   msg2:='Do you wish to send a copy of our complaints procedure in the post?';
  end;

 If msg<>'' then
 Begin
  if MessageDlg(msg,mtconfirmation,[mbyes,mbno],0)<>mryes then Exit;
 end;

 if msg='' then
 begin
  if msg1 <> '' then
  begin
    if messagedlg(msg1,mtconfirmation,[mbyes,mbno],0)=mryes then
    begin
     try
      FRM_HELPER_MESSAGE.GET_DETAILS('DISSATISFIED');
      FRM_HELPER_MESSAGE.SHOWMODAL;
     except
      Messagedlg('This Feature is not available on this operating system',mterror,[mbok],0);
     end;
    end;
  end;
 end
 else
 begin
 with CreateMessageDialog(msg1+#13+'N.B. Verbally will take approximately Three and a half minutes (Must be advised to Customer)',mtConfirmation,[mbyes,mbno,mbcancel])
 do
 try
  TButton(FindComponent('Yes')).Caption :='POST';
  TButton(FindComponent('No')).Caption :='VERBALLY';
  if msg<>'' then TButton(FindComponent('Cancel')).Caption :='CANCEL';
  Position := poScreenCenter;
  Result := ShowModal;
  if result=6 then
  begin
     FRM_Reports.PrintThisReport(RPT,'Complaints Procedure',selection,'','PRINTQUEUE',CUSTOMERID,'');
  end;
  if result=7 then
  begin
   try
    FRM_HELPER_MESSAGE.GET_DETAILS('DISSATISFIED');
    FRM_HELPER_MESSAGE.SHOWMODAL;
   except
    Messagedlg('This Feature is not available on this operating system',mterror,[mbok],0);

   end;
   if msg2<>'' then
   Begin
    if Messagedlg(msg2,mtConfirmation,[mbyes,mbno],0)<>mryes then exit;
    FRM_Reports.PrintThisReport(RPT,'Complaints Procedure',selection,'','PRINTQUEUE',CUSTOMERID,'');
   end;

  end;

 finally
  Free;
 end;
 end;
end;

procedure TFRM_Tree.CustomerDIsSatiisfied1Click(Sender: TObject);
var
  customerid:string;
begin
 xnode:=treeview1.FocusedNode;
 TreeData:= treeview1.GetNodeData(xnode);

 if treeupdating=true then exit;
 CustomerID:=TreeData.D_Customer_ID;
 StartComplaintsProcess(customerid,'I')
end;

function TFRM_Tree.CustomerHasInvoluntaryModeChangeFlag(
  const aCustomerId: Int64): Boolean;
const
  SQL = 'select count(1) from enquiry.enquiries where customer_id = :customerID and request_type = :requestType and resolved = :resolved';
begin
  Result := gSqlUtil.SelectQueryInteger(SQL, ['customerId' , otLong   , aCustomerId,
                                              'requestType', otInteger, 10593,
                                              'resolved'   , otString , 'N']) > 0;
end;

{
  Re-assign agreements between Customers
  Wrike Ref 151273193
}
//******************************************************************************
procedure TFRM_Tree.doHideAgreementReassignment;
//******************************************************************************
begin
  ShowReassign := false;
  grpbxReassignAgreements.Visible := false;
  Panel1.Visible := true;
 // M_Reassign.Enabled := true;
  fAgreementsToReassign.Clear;
end;

//******************************************************************************
procedure TFRM_Tree.doShowAgreementReassignment;
//******************************************************************************
begin
  ShowReassign := true;
  grpbxReassignAgreements.Visible := true;
  edtAssignee.Clear;
  Panel1.Visible := false;
  //M_Reassign.Enabled := false;
  // Check to see if treeview is holding multiple customers and if so then need to
  // ForcePremiseExpansion next time showing agreement reassignment
  //fInitialPremiseExpansion := fInitialPremiseExpansion or (fCustomersOnTreeview.FindTree > 1);
end;

//******************************************************************************
procedure TFRM_Tree.refreshCustomerTree(node: PVirtualNode; nag: Boolean; expand: Boolean);
//******************************************************************************
var
  nagging: Boolean;
begin
  nagging := dpaNag;
  dpaNag := nag;

  fAssignFromCustomerId := fAgreementsToReassign.CustomerId[node];
  FRM_Main.SearchForCust(fAssignFromCustomerId);

  TreeView1.Expanded[node] := false;

  if expand then
  begin
    TreeView1.Expanded[node] := true;
  //doExpandAgreements(node);    //INVOKED IN TREEVIEWEXPANDING
  end;

  dpaNag := nagging;
end;

//******************************************************************************

//Procedure TFRM_Tree.mnuChangeRelationshipRatingClick(Sender: TObject);
//Begin
 // TFrmRelationshipRatingChange.Launch(Treedata.D_Customer_ID);
 // ViewCustomerTree1Click(Self);
//End; //Proc

procedure TFRM_Tree.mnuCreditCheckClick(Sender: TObject);
begin
  TCreditCheckForm.Launch(TreeData.D_Customer_Id);
end;

procedure TFRM_Tree.mniAlternativeFuelPaymentClick(Sender: TObject);
begin
  GovernmentEnergySchemesScreen('AFP');
end;

procedure TFRM_Tree.mniChangeOfTenancyClick(Sender: TObject);
begin
  TFrmChangeOfTenancy.StartModal(Self, StrToInt64(TreeData.D_Customer_Id), CustName);
end;

procedure TFRM_Tree.mniEnergyBillsSupportSchemeClick(Sender: TObject);
begin
  GovernmentEnergySchemesScreen('EBSS');
end;

procedure TFRM_Tree.mniTransferCreditFromSavingsToAgreementClick(Sender: TObject);
begin
  TFRM_TransferCreditFromSavingsToAgreement.StartModal(Self, StrToInt64(custid),
    FRM_main_search.customercontacts.FieldByName('CUSTOMER_TYPE_ID').AsInteger,
    'credit', ((hidecheck.Visible = true) and (hidecheck.checked = true))
  );
  treeview1.expanded[xnode.parent] := false;
  treeview1.expanded[xnode.parent] := true;
end;

procedure TFRM_Tree.mniTransferCreditFromSavingsToMeterClick(Sender: TObject);
begin
  TTransferCreditFromSavingsToMeter.StartModal(Self, StrToInt64(custid),
    ((hidecheck.Visible = true) and (hidecheck.checked = true)));
  treeview1.expanded[xnode.parent] := false;
  treeview1.expanded[xnode.parent] := true;
end;

procedure TFRM_Tree.mniTransferDebitFromSavingsToAgreementClick(Sender: TObject);
begin
  TFRM_TransferCreditFromSavingsToAgreement.StartModal(Self, StrToInt64(custid),
    FRM_main_search.customercontacts.FieldByName('CUSTOMER_TYPE_ID').AsInteger,
    'debit', ((hidecheck.Visible = true) and (hidecheck.checked = true))
  );
  treeview1.expanded[xnode.parent] := false;
  treeview1.expanded[xnode.parent] := true;
end;

procedure TFRM_Tree.mniWarmHomeDiscountClick(Sender: TObject);
begin
  GovernmentEnergySchemesScreen('WHD');
end;

procedure TFRM_Tree.mniWinterWarmerFinancialAssistancePaymentClick(Sender: TObject);
begin
  XNode := Treeview1.FocusedNode;
  TreeData := Treeview1.GetNodeData(XNode);
  Application.CreateForm(TFRM_WinterWarmerFinancialAssistancePayment, FRM_WinterWarmerFinancialAssistancePayment);
  try
    FRM_WinterWarmerFinancialAssistancePayment.pCustomerID := StrToInt64(TreeData.D_Customer_Id);
    FRM_WinterWarmerFinancialAssistancePayment.pCustomerName := TreeData.D_Customer_Name;
    FRM_WinterWarmerFinancialAssistancePayment.ShowModal;
  finally
    FRM_WinterWarmerFinancialAssistancePayment.Release;
  end;
end;

procedure TFRM_Tree.mniWinterWarmerPaymentClick(Sender: TObject);
begin
  GovernmentEnergySchemesScreen('WWP');
end;

{------------------------------------------------------------------------------}
procedure TFRM_Tree.mnuDCCMeterReadingsClick(Sender: TObject);
var
  nodeData : PMyRec;
begin
  if not Assigned(TreeView1.FocusedNode) then
    exit;

  nodeData := TreeView1.GetNodeData(TreeView1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  TFrm_Smets_Readings_Dcc.Start(Self, nodeData.D_Span);
end;

procedure TFRM_Tree.mnuDUoSInvoicingClick(Sender: TObject);
begin
 //added by maryam on wrike ticket 171744089
 Frm_DUoSPassCharges := TFrm_DUoSPassCharges.Create(self);
 Frm_DUoSPassCharges.DUOSQuery.SetVariable('MPAN',TreeData.D_SPAN);
 Frm_DUoSPassCharges.DUOSQuery.SetVariable('REG_ID',TreeData.D_regid);
 Frm_DUoSPassCharges.ShowModal;
end;

procedure TFRM_Tree.mnuIncomeAndExpenditureWebFormClick(Sender: TObject);
var
  sCustomerId, sPostCode, sURL: string;
begin
  XNode := Treeview1.FocusedNode;
  NodeData := Treeview1.GetNodeData(XNode);
  sCustomerId := NodeData.D_Customer_Id;
  sPostCode := StringReplace(
                 gSqlUtil.SelectQueryString('SELECT PR.PREMISE_POSTCODE' +
                                            '  FROM CRM.PREMISES PR' +
                                            '  JOIN CRM.CUSTOMER CU ON CU.PRIMARY_MAILING_ADDRESS_ID = PR.PREMISE_ID' +
                                            ' WHERE CU.CUSTOMER_ID = :CUSTOMER_ID',
                                            ['CUSTOMER_ID', otString, sCustomerId]),
                 ' ', EmptyStr, [rfReplaceAll]);
  sURL := Format(Trim(FRM_Common.GETVALUE('INCOME_AND_EXPENDITURE_FORM_URL')), [sCustomerId, sPostCode]);

  if (sCustomerId <> EmptyStr) and (sPostCode <> EmptyStr) and (sURL <> EmptyStr) then
    ShellExecute(Handle, 'open', PChar(sURL), nil, nil, SW_SHOWNORMAL);
end;

procedure TFRM_Tree.mnuRefreshPriorityNotificationClick(Sender: TObject);
begin
  Application.CreateForm(TFRM_Notifications, FRM_Notifications);
  try
    FRM_Notifications.RefreshData(CustId, UserID);
    FRM_Notifications.GetNotifications(CustId, TreeData.D_Customer_Name, UserID, False);
  finally
    FreeAndNil(FRM_Notifications);
  end;
end;

procedure TFRM_Tree.mnuViewPriorityNotificationClick(Sender: TObject);
begin
  Application.CreateForm(TFRM_Notifications, FRM_Notifications);
  try
    FRM_Notifications.GetNotifications(CustId, TreeData.D_Customer_Name, UserID, Sender = mnuViewPriorityNotification);
  finally
    FreeAndNil(FRM_Notifications);
  end;
end;

procedure TFRM_Tree.M_ReassignClick(Sender: TObject);
//******************************************************************************
begin
  doShowAgreementReassignment;
  if TreeView1.Expanded[TreeView1.FocusedNode] then
    refreshCustomerTree(TreeView1.FocusedNode, false, true)
  else
    refreshCustomerTree(TreeView1.FocusedNode, dpaNag, true);
end;

//******************************************************************************
procedure TFRM_Tree.btnHideReassignAgreementPanelClick(Sender: TObject);
//******************************************************************************
begin
  doHideAgreementReassignment;
  refreshCustomerTree(Treeview1.GetFirst, false, false);
end;

//******************************************************************************
procedure TFRM_Tree.btnReassignClick(Sender: TObject);
//******************************************************************************
begin
  fAgreementsToReassign.FindAgreementsToReassign(TreeView1.GetFirst(false));
  processAgreements;
end;

//******************************************************************************
procedure TFRM_Tree.processAgreements;
//******************************************************************************
var
  prompt: string;
  i: integer;
begin
  if fAgreementsToReassign.Count > 0 then
  begin
    prompt := 'The following agreement(s) shall be reassigned to Customer ' + edtAssignee.Text + ': ' + lblAssigneeName.Caption + #10#13;

    for i := 0 to fAgreementsToReassign.Count - 1 do
      prompt := prompt + #10#13 + fAgreementsToReassign.AgreementId[i];

    if MessageDlg(prompt, mtConfirmation, [mbOk, mbCancel], 0) = mrOK then
    begin
      reassignAgreements;

      if messageDlg('Would you like to transfer Notes to ' + lblAssigneeName.Caption + '?', mtInformation, [mbYes, mbNo], 0) = mrYes then
        doCopyNotestoAnotherCustomer(edtAssignee.Text);

      doHideAgreementReassignment;
      refreshCustomerTree(Treeview1.GetFirst, false, false);
    end;
  end
  else
  begin
    prompt := 'No agreements have been selected.' + #10#13;
    prompt := prompt + 'Please select one or more agreements to reassign to Customer ' + edtAssignee.Text + ': ' + lblAssigneeName.Caption;
    MessageDlg(prompt, mtWarning, [mbOk], 0);
  end;
end;

//******************************************************************************
procedure TFRM_Tree.reassignAgreements;
//******************************************************************************
var
  i: integer;
begin
  edtAssignee.Enabled := false;

  for i := 0 to fAgreementsToReassign.Count - 1 do
    FRM_COMMON.Execute_Oracle_Procedure(getSQLMoveAgreement(edtAssignee.Text, fAgreementsToReassign.AgreementId[i]));

  fAgreementsToReassign.Clear;
  edtAssignee.Enabled := true;
end;

//******************************************************************************
function TFRM_Tree.getSQLMoveAgreement(customer, agreement: string): string;
//******************************************************************************
begin
  result := 'crm.pk_utilities.pr_move_agreement_to_customer(' + customer + ',' + agreement + ')';
end;

// BSL-SJ - 11/08/2021 - CRM-511 - Control over Super Customer - Fixing Bug.
Function TFRM_Tree.GetSuperCustIcon: Integer;
Begin
  Result := FSuperCustIcon;
End; // Funct

//******************************************************************************
procedure TFRM_Tree.edtAssigneeChange(Sender: TObject);
//******************************************************************************
begin
  btnReassign.Enabled := (validCustomerId(edtAssignee.Text,fAssignFromCustomerId) and (fAssignFromCustomerId <> trim(edtAssignee.Text)));
end;

//******************************************************************************
function TFRM_Tree.validCustomerId(id,existingid: string): Boolean;
//******************************************************************************
var
  qry: TOracleDataSet;
begin
  result := false;
  lblAssigneeName.Caption := '';

  if length(id) <> 10 then
    exit;

  qry := main_data_module.tempquery;
  qry.Close;
  qry.DeleteVariables;
  qry.SQL.clear;
  qry.declarevariable('CID',otstring);
  qry.declarevariable('EID',otstring);
  qry.setvariable('CID',id);
  qry.setvariable('EID',Existingid);
  qry.SQL.Add('select legal_entity_name from crm.customer where customer_id= :CID');
  // Only bring back customers of same type
  qry.SQL.Add('and customer_id <> :EID and customer_type_id in (select customer_type_id from crm.customer where customer_id= :EID )');
  qry.Open;
  qry.DeleteVariables;

  result := qry.RecordCount <> 0;
  if result then
  begin
    lblAssigneeName.Caption := qry.FieldByName('legal_entity_name').AsString;
    if fAssignFromCustomerId = trim(edtAssignee.Text) then
      lblAssigneeName.Caption := lblAssigneeName.Caption + ' CANNOT ASSIGN TO SAME CUSTOMER';
  end;
end;

//******************************************************************************
function TFRM_Tree.isAgreementNode(node: PVirtualNode): Boolean;
//******************************************************************************
var
  data: pmyrec;
begin
  data := TreeView1.GetNodeData(node);
  result := copy(data.Caption, 1, length(CAgreementNode)) = CAgreementNode;
end;

function TFRM_Tree.isS1Enrolled(mpxn: string): Boolean;
begin
  with Generalquery do
  Begin
    close;
    DeleteVariables;
    DeclareVariable('MPXN', otstring);
    sql.clear;
    sql.add('SELECT count(seie.mpxn) FROM ods.S1EA_INTEGRATION_EVENT seie');
    sql.add('WHERE seie.mpxn = :MPXN and seie.processed = '+QuotedStr('P'));
    setvariable('MPXN', mpxn);
    open;
    deletevariables;
  End;

  result := generalquery.Fields[0].AsInteger > 0;
end;

function TFRM_Tree.isCosGainEnrolled(mpxn: string): Boolean;
begin
  with Generalquery do
  Begin
    close;
    DeleteVariables;
    DeclareVariable('MPXN', otstring);
    sql.clear;
    sql.add('SELECT count(scg.mpxn) from AUTOMATIONPRO.SMETS2_COS_GAINS scg');
    sql.add('WHERE scg.mpxn = :MPXN and SCG.METER_TYPE = '+QuotedStr('S1')+' and SCG.STATE = 1');
    setvariable('MPXN', mpxn);
    open;
    deletevariables;
  End;

  result := generalquery.Fields[0].AsInteger > 0;
end;

function TFRM_Tree.isDCCMeter: Boolean;
begin
  Result := (((TreeData.m_service = '0') and TreeData.D_SpanDCC_E and (treeview1.popupmenu=PopupSmets_DCC)) or ((TreeData.m_service = '1') and TreeData.D_SpanDCC_G and (treeview1.popupmenu=PopupSmets_DCC)));
end;

function TFRM_Tree.IsPrepayAndLive: Boolean;
var
  qryIsPrepayAndLive : TOracleDataSet;
begin
  xnode:=treeview1.FocusedNode;
  TreeData:= treeview1.GetNodeData(xnode);

  qryIsPrepayAndLive := TOracleDataSet.Create(Self);
  qryIsPrepayAndLive.Session := FRM_Login.MainSession;

  try
    try
      with qryIsPrepayAndLive do
      begin
        Close;
        DeleteVariables;
        DeclareAndSet('p_is_latest_product', otString, 'Y');
        DeclareAndSet('p_agreement_status_id', otInteger, 1);
        DeclareAndSet('p_payment_plan_id', otString, 'P');
        DeclareAndSet('p_agreement_id', otString, TreeData.D_Agreement_ID);
        SQL.Clear;
        SQL.Add('SELECT COUNT (1) AS isPrepayAndLive');
        SQL.Add('FROM crm.agreements ag JOIN crm.agreement_products ap ON ap.agreement_id = ag.agreement_id');
        SQL.Add('WHERE ap.is_latest_product = :p_is_latest_product');
        SQL.Add('AND ag.agreement_status_id = :p_agreement_status_id');
        SQL.Add('AND ap.payment_plan_id = :p_payment_plan_id');
        SQL.Add('AND ag.agreement_id = :p_agreement_id');
        Open;
        Result := (FieldByName('isPrepayAndLive').AsInteger = 1)
      end;
    finally
      FreeAndNil(qryIsPrepayAndLive);
    end;
  except
    on e: Exception do
    begin
      MessageDlg('An error has occurred while retrieving information from agreement.',
                 mtError, [mbOK], 0);
    end;
  end;
end;

function TFRM_Tree.getMeterDCC(mpxn:string): Boolean;
var
bMeterType:string;
begin
  try
    with main_data_module.tempquery do
    begin
      close;
      deletevariables;
      declarevariable('MPXN', otstring);
      sql.clear;
      sql.add('select case when smets_meter_version is null then' + QuotedStr('S2'));
      sql.add('else smets_meter_version end as version from ods.dcc_mpxn_status ');
      sql.add('WHERE');
      sql.add('MPXN=:MPXN');
      setvariable('MPXN', mpxn);
      open;
      deletevariables;
      bMeterType := main_data_module.tempquery.fields[0].text;
    end;
  finally
    if ((bMeterType = 'S2') or (bMeterType = 'S1EA')) then
      result := true
    else result := false;
  end;
end;

//******************************************************************************
procedure TFRM_Tree.doExpandNode(node: PVirtualNode);
//******************************************************************************
begin
  TreeView1.Expanded[node] := true;
end;

//******************************************************************************
procedure TFRM_Tree.doExpandAgreements(node: PVirtualNode);
//******************************************************************************
begin
  fAgreementsToReassign.Iterator.FindNode(TreeView1, node, isAgreementNode, doExpandNode);

//																		                       Anon Method Ref    Anon Method Directly implemented as aparamater
//  fAgreementsToReassign.iterator.FindNode(TreeView1, node, isAgreementNode, procedure(node: PVirtualNode) begin TreeView1.Expanded[node] := true; end;);
end;

//******************************************************************************
procedure TFRM_Tree.HideAgreementReassignment;
//******************************************************************************
begin
  doHideAgreementReassignment;
end;

//******************************************************************************
procedure TFRM_Tree.Treeview1Collapsing(Sender: TBaseVirtualTree; Node: PVirtualNode; var Allowed: Boolean);
//******************************************************************************
begin
  fAgreementsToReassign.FindNodes(TreeView1.GetFirst(false));
end;
//
{ END
  Re-assign agreements between Customers
  Wrike Ref 151273193
}
// BSL - 15/03/2023 - ISC-547 New SMS in CRM providing additional support links post ASC decline
                   // Code optimization
Procedure TFRM_TREE.BuildSMSMenu;
Var bItem: TMenuItem;
Begin
  Main_Data_Module.GeneralQuery.Close;
  Main_Data_Module.GeneralQuery.SQL.Text := 'Select * From CRM.SMS_Lookup_List Where Enabled = ''Y'' Order by Sort_Order';
  Main_Data_Module.GeneralQuery.Open;

  While Not Main_Data_Module.GeneralQuery.Eof do
    Begin
      bItem            := TMenuItem.Create(PopUpMobile);
      bItem.Caption    := 'SMS - ' + Main_Data_Module.GeneralQuery.Fields[1].Text;
      bItem.ImageIndex := Main_Data_Module.GeneralQuery.FieldByName('Image_Index').AsInteger;
      bItem.Tag        := Main_Data_Module.GeneralQuery.FieldByName('Template_Id').AsInteger;

      Try
        bItem.Hint := Main_Data_Module.GeneralQuery.FieldByName('Call_This_Function').AsString;
      Except
        bItem.Hint := 'CRM.FN_RET_SMS_TEXT';
      End;

      bItem.OnClick := SendSMSMessage1Click;
      PopUpMobile.Items.Add(bItem);
      Main_Data_Module.GeneralQuery.Next;
    End;
End; // Proc

{------------------------------------------------------------------------------}
procedure TFRM_TREE.DoSmetsCreditDCC(AMode: Integer);
var
  nodeData : PMyRec;
  service  : integer;
begin
  if not Assigned(Treeview1.FocusedNode) then
    exit;

  nodeData := Treeview1.GetNodeData(Treeview1.FocusedNode);
  if not Assigned(nodeData) then
    exit;

  service := StrToInt(nodeData.M_Service);

  TFrm_Smets_Manage_Credit_Dcc.StartModal(Self, service, nodeData.D_Span, nodeData.M_MeterId, aMode);
end;
{------------------------------------------------------------------------------}

procedure TFRM_TREE.PersistStringValue(var aValue: string; aNewValue: string);
begin
  // DANBYT - 04/12/2024 - CRMX-164 - Ensure value persists when new value is blank
  if not aNewValue.IsEmpty then aValue := aNewValue;
end;

end.