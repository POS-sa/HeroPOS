Attribute VB_Name = "modDeclarations"

'Global API Declarations check if will work in Vista Business sp1


'Remove SendKeys calls and replace them with API code.
'Use the HKEY_CURRENT_USER in the registry for the settings of your application. Do not write to HKEY_LOCAL_MACHINE.
'If you are using ADO use 2.8 and above in your application.
'If you are using XML use XML version 3.0 and above in your application
'If you are using the PlaySound API, if your wave files are not PCM but mpeg layer-3 make sure your mpeg wave files are Stereo and not Mono.
'If you need the Printer Setup dialog either using the Common Dialog Control, or by using the API functions, the dialog will not return the correct number of copies. The dialog will always return 1 on Vista. The way around this bug is to create your own Printer Setup dialog box and when you get the correct number of copies you will have to send to the printer multiple times to print out multiple copies. Here is a link that discusses this issue in details.
'Relocate settings files, data files etc into "Common Files" (C:\Users\Public) instead of "Program Files". You should use the API calls to locate these folders because the folders are in different paths for different machines and OSs. Here is a link that discusses this issue in details.
'Per-user settings should be in a separate file located under "Application Data" and this should also be requested of the OS in the same manner.
'For "Common Files" ask for ssfCOMMONDATA (or CSIDL_COMMON_APPDATA).
'For "Application Data" ask for ssfAPPDATA (or CSIDL_APPDATA). To properly use these filesystem locations you are supposed to create a subdirectory for your "company name" and under that another for your "application name." Then put your settings or data under that.
'Any working "document" files that are meant to be found and manipulated by the user (i.e. via Explorer) should be placed into CSIDL_PERSONAL ("My Documents") or CSIDL_COMMON_DOCUMENTS ("All Users\Documents").
'DeleteSetting no longer works without a key. e.g. DeleteSetting "Mytestprogram, "General" fails to delete anything and gives an error. but DeleteSetting "Mytestprogram, "General","keyname" works fine. It seems that key is no longer Optional in: DeleteSetting appname, section[, key] as in documentation.


Public Declare Function LoadModule Lib "kernel32" (ByVal lpModuleName As String, lpParameterBlock As Any) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" _
(ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public lpVolumeSerialNumber As Long
Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type
Public SYSTEM_INFO As SYSTEM_INFO
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)


'*****************************************************
'User Type Block
Type UserRecord
   User_Number As Integer
   Name As String * 50
   FirstName As String * 20
   LastName As String * 50
   Password As String * 10
   LastUser As Integer
   Reservations As Boolean
   Rooms As Boolean
   Guests As Boolean
   Checkin As Boolean
   Checkout As Boolean
   Users As Boolean
   Reports As Boolean
   Settings As Boolean
   Sales As Boolean
   Inventory As Boolean
   uType As Integer
   Cash_Sales As Boolean
   Cheque_Sales As Boolean
   Card_Sales As Boolean
   Charge_Sales As Boolean
   Loyalty_Sales As Boolean
   Item_Corrects As Boolean
   Voids As Boolean
   Returns As Boolean
   Ullages As Boolean
   Over_Tender As Boolean
   Disc_Perc As Boolean
   Disc_Amt As Boolean
   Payouts As Boolean
   Pickups As Boolean
   Loans As Boolean
   Receive_Acc As Boolean
   Split_Tender As Boolean
   Buffer_Print As Boolean
   Reprint As Boolean
   Total_Clear As Boolean
   Trans_Store As Boolean
   Trans_Clear As Boolean
   Transfers As Boolean
   Overides As Boolean
   Search As Boolean
   Wage As Double
   Comm1 As Double
   Comm2 As Double
   Logged_in As Boolean
   All_Tables As Boolean
   App_Exit As Boolean
   Draw_Cash As Boolean
   Draw_Cheque As Boolean
   Draw_Card As Boolean
   Draw_Charge As Boolean
   Draw_Loyalty As Boolean
   System_Service As Boolean
   Service_Charge As Boolean
   Quotes As Boolean
   Drawer_No As Integer
   Com_Calc As Integer
   Bar_Cash As Integer
   No_Sales As Integer
   Owner_Transfer As Boolean
End Type
Public UserRecord As UserRecord
'*****************************************************
'Connection Stings and Server Variables
Public cnnMain As ADODB.Connection
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public rs4 As New ADODB.Recordset
Public rs5 As New ADODB.Recordset
Public rs6 As New ADODB.Recordset

Public us As New ADODB.Recordset
Public us1 As New ADODB.Recordset
Type Server
    SQL_Name As String * 20
    SQL_User As String * 20
    SQL_Password As String * 20
    SQL_Database As String * 20
End Type
Public Server As Server
'*****************************************************
'LogFiles
Type LogFiles
   MainLog As String * 50
   ErrorLog As String * 50
End Type
Public LogFiles As LogFiles
'*****************************************************
'LogFiles
Type KeyType
    ClearKey As Integer
    InputKey  As Integer
    FunctionKey As Integer
    ItemizerKey As Integer
    FinalizationKey As Integer
End Type
Public KeyType As KeyType
'*****************************************************
'TillModes
Type TillMode
    StartMode As Integer
    Inputmode As Integer
    FinMode As Integer
    NewMode As Integer
    TenderMode As Integer
    CashupMode As Integer
End Type
Public TillMode As TillMode
'*****************************************************
'TillData
Type TillData
    Creditlimit As Double
    Creditbalance As Double
    ProductCode As String
    DeptNo As String
    Description As String
    ShortDesc As String
    Qty As String
    TaxRate As Double
    TaxType As Integer
    UnitSize As String
    Cost As Double
    Price As String
    Keystring As String
    ExtraFunc As String
    Change As Double
    Tendered As Double
    DocNo As Double
    TransNo As Double
    KeyReg As String
    SaleTotal As Double
    VoidTotal As Double
    VoidCount As Integer
    ReturnTotal As Double
    ReturnCount As Integer
    UllageTotal As Double
    UllageCount As Integer
    PriceOveride As Integer
    Kitchen1 As String
    Kitchen2 As String
    Recipe As Integer
    Cashup_No As Integer
    Cash As Double
    Card As Double
    Cheque As Double
    CashCol As Double
    CardCol As Double
    ChequeCol As Double
    Charge As Double
    Loyalty As Double
    TaxTotal As Double
    TaxableSales As Double
    NonTaxableSales As Double
    CollectedTax As Double
    CalculatedTax As Double
    Corrects As Double
    CorrectCount As Integer
    TableNo As Single
    TabName As String
    TabNo As Single
    Covers As Integer
    Tipp As Double
    TippCount As Integer
    ShortTender As Boolean
    UserOveride As Single
    Discount As Double
    DiscountVal As Double
    TotDiscount As Double
    TotDiscountVal As Double
    TotDiscountCount As Double
    TotDiscountValCount As Double
    Room_No As Integer
    Account_No As String
    Res_No As Long
    Weight As Double
    Deposit As Integer
    Print_Count As Integer
    Table_Name As String
    Prev_Doc_No As Integer
End Type
Public TillData As TillData
Type Devices
    Drawer1KickString As String
    Drawer2KickString As String
    TwoDrawer As Integer
    ScaleModel As String
    ScalePort As String
    ScaleSet As String
    DisplayModel As String
    DisplayPort As String
    DisplaySet As String
    Barcode_Height  As Integer
    Label_Width As Integer
    Label_Printer As String
    Label_Height  As Integer

End Type
Public Devices As Devices
Type Cost_Code
    One As String
    Two As String
    Three As String
    Four As String
    Five As String
    Six As String
    Seven As String
    Eight As String
    Nine As String
    Ten As String
End Type
Public Cost_Code As Cost_Code
Type Workstation
    Disc10 As Integer
    Disc20 As Integer
    Disc30 As Integer
    Disc40 As Integer
    Disc50 As Integer
    Disc60 As Integer
    Disc70 As Integer
    Disc80 As Integer
    Disc90 As Integer
    DiscFree As Integer
End Type
Public Workstation As Workstation
'*****************************************************
Type OrientStructure
   Orientation As Long
   Pad As String * 16
End Type
'Declare Function Escape% Lib "GDI" (ByVal hDc%, ByVal nEsc%, ByVal nLen%, lpData As OrientStructure, lpOut As Any)

'Global Variables
Public gblApp_Name As String * 50
Public CurrentKey As Integer
Public Dates(6) As Date
Public GridXS As Integer
Public GridX As Integer
Public GridY As Integer
Public StartTime As String * 5
Public StopTime As String * 5
Public DateSelect As Date
Public KeyRegister As String
Public GlobalMode As Integer
Public ProductFilter(2) As String * 100
Public Workstation_No As Integer
Public Workstation_Name As String * 20
Public WorkstationSOH As Integer
Public Time_Start As Date
Public Time_Stop As Date
Public Panel_no As Integer
Public Branch_No As Integer
Public Branch_Name As String
Public Branch_Address  As String
Public Dept_Order As Integer
Public Location_No As Integer
Public System_Access As Integer
Public QCash As Integer
Public System_Service As Integer
Public Slip_Printer As String
Public Slip_PrinterPort As String
Public Slip_Printer_Type As Integer
Public Kitchen_Printer_No As Integer
Public Vat_No As String
Public Comp_Name As String
Public GQAnswer As String
Public Logo_File As String
Public Kitchen_Con As Integer
Public Branch_Type As Integer
Public ImPrinting
Public Tel_Ex As Integer
Public Tel_Ex_Dir  As String
Public PrintZeroItems As Integer
Public PrintVoids As Integer
Public Finalizing As Boolean
Public Inside As Boolean
Public Swiss_Round As Double
Public VoidReasons As Integer
Public PrintSlipTransfers As Integer
Public rowStart As Integer
Public colStart As Integer
Public colStop As Integer
Public ReplicationServ As Integer
Public AskLog As Integer
Public StockBarcode As Integer
Public Selender As Date
Public PrintBarStock As Integer
Public HappyHour As Integer
Public HappyHourPrice As Integer
Public HappyHour1 As Integer
Public HappyHourPrice1 As Integer
Public PayoutPrint As Integer
Public ChargePrint As Integer
Public ChargeSlip As Integer
Public RAPrint As Integer
Public CodePrefix
Public TradePrint
Public Member_No
Public Zero_Print
Public Senttofinalize As Boolean
Public Barcodematch As Boolean
Public Successful As Boolean
Public Resevationsenabled As Boolean ' splashbutton 1
Public Minitradeanalysis As Boolean ' splashbutton 1
Public Newidea As Boolean ' splashbutton 2
Public Redprintdept As String
Public Redprintdept2 As String
Public Priceonkitchenprint As Integer
Public Conversion_description As String
Public Conversion_Rate As Double
Public LastTable As Double
Public LastTab As Double
Public Last_KeyString As String
Public Process_Running As Boolean 'Use to prevent a process interrupting a current one


Public LineToVoid As Integer
Public Reprintplease As Boolean
Public Thediscounttotal As Double
Public Reprintdiscount As Integer
Public Redprintenabled As Integer
Public Subscriptloaded As Boolean
Public Maindate As Variant
Public Wholecode As String
Public Checksummedbarcodeean13 As String
' WWWSERVERSETTINGS
Public WWWServerName As String
Public WWWUserName As String
Public WWWPassword As String
Public WWWDatabase As String

'***************************


