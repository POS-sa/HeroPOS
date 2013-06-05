Attribute VB_Name = "modMain"
'check payout time date
'payout and grv reports

'********************************************************************
' Service Charge authentication
'ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Service_Charge = 1"
'                    If rs.RecordCount > 0 Then
'
'                        frmValidate.Tag = "1"
'                    Else
'                        frmValidate.Tag = "0"
'                    End If
'                    rs.Close
'********************************************************************
 


Private Sub Main()
    If App.PrevInstance = True Then
        Shell "C:\WINDOWS\system32\Taskmgr.exe"
        MsgBox "You already have and instance of HeroPOS running on this Computer." & Chr(13) & "Exit HeroPOS or Press CTRL - ALT - DEL and End the Task", vbCritical, "HeroPOS"
        
        End
    End If
    'If Date > "2007-03-01" Then End
    serial = Str(GetSerialNumber("C:\"))
    ImPrinting = False
    KeyType.ClearKey = 0
    KeyType.InputKey = 1
    KeyType.FunctionKey = 2
    KeyType.ItemizerKey = 3
    KeyType.FinalizationKey = 4
    Kitchen_Printer_No = 0
    TillMode.StartMode = 1
    TillMode.Inputmode = 2
    TillMode.NewMode = 3
    TillMode.FinMode = 4
    TillMode.TenderMode = 5
    TillMode.CashupMode = 6
    gblApp_Name = "HeroPOS"
    Maindate = DateValue(Date)
    Dim fso As New FileSystemObject
    
    
    If fso.FolderExists(App.Path & "\Logs") = False Then
    fso.CreateFolder (App.Path & "\Logs")
    If fso.FolderExists(App.Path & "\PDFReports") = False Then
    fso.CreateFolder (App.Path & "\PDFReports")
    End If
    filenum = FreeFile
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    Close filenum
    Open Trim(App.Path) & "\Logs\MainLog.log" For Append As filenum
    Close filenum
    Open Trim(App.Path) & "\Logs\ErrorLog.log" For Append As filenum
    Close filenum
    Open Trim(App.Path) & "\Logs\KeyLog.log" For Append As filenum
    Close filenum
    Open Trim(App.Path) & "\Logs\Delete.txt" For Append As filenum
    Close filenum
    
    End If
    
    
    
    SaveSetting Trim(gblApp_Name), "Logs", "Main_Log", App.Path & "\Logs"
    SaveSetting Trim(gblApp_Name), "Logs", "Error_Log", App.Path & "\Logs"
    Server.SQL_Name = GetSetting(appname:=Trim(gblApp_Name), Section:="Server", key:="Server")
    Server.SQL_Database = GetSetting(appname:=Trim(gblApp_Name), Section:="Server", key:="SQL_Database")
    Server.SQL_User = GetSetting(appname:=Trim(gblApp_Name), Section:="Server", key:="SQL_User", Default:="sa")
    Server.SQL_Password = GetSetting(appname:=Trim(gblApp_Name), Section:="Server", key:="SQL_Password")
    Slip_Printer = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Slip_Printer")
    Slip_Printer_Type = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Slip_Printer_Type", Default:=0)
    Slip_PrinterPort = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Slip_PrinterPort", Default:=0)
    Location_No = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Location", Default:=1)
    WorkstationSOH = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Stock_on_Hand", Default:=0)
    Kitchen_Printer_No = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Kitchen_Printer", Default:=0)
    Kitchen_Con = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Kitchen_Con", Default:=0)
    Tel_Ex = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Ex_Activated", Default:=0)
    Tel_Ex_Dir = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Exchange", Default:="C:\")
    PrintZeroItems = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="PrintZero", Default:=0)
    PrintVoids = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="PrintVoid", Default:=0)
    PrintSlipTransfers = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="PrintSlipTransfers", Default:=0)
    PrintBarStock = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="PrintBarStock", Default:=0)
    Devices.Drawer1KickString = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Drawer_One_KickString", Default:="<Not Installed>")
    Devices.Drawer2KickString = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Drawer_Two_KickString", Default:="<Not Installed>")
    Devices.TwoDrawer = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Use_Both_Drawers", Default:=0)
    Devices.ScaleModel = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Scale_Model", Default:="<Not Installed>")
    Devices.ScalePort = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Scale_Port", Default:="<Not Set>")
    Devices.ScaleSet = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Scale_Settings", Default:="<Not Set>")
    Devices.DisplayModel = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Display_Model", Default:="<Not Installed>")
    Devices.DisplayPort = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Display_Port", Default:="<Not Set>")
    Devices.DisplaySet = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Display_Settings", Default:="<Not Set>")
    Devices.Label_Printer = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Label_Printer", Default:="<Not Set>")
    Devices.Barcode_Height = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Label_Height", Default:=0)
    Devices.Label_Width = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Label_Width", Default:=0)
    Workstation.Disc10 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc10", Default:=1)
    Workstation.Disc20 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc20", Default:=1)
    Workstation.Disc30 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc30", Default:=1)
    Workstation.Disc40 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc40", Default:=1)
    Workstation.Disc50 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc50", Default:=1)
    Workstation.Disc60 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc60", Default:=1)
    Workstation.Disc70 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc70", Default:=1)
    Workstation.Disc80 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc80", Default:=1)
    Workstation.Disc90 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc90", Default:=1)
    Workstation.DiscFree = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="DiscFree", Default:=1)
    ReplicationServ = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="ReplicationServ", Default:=0)
    AskLog = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="AskLoc", Default:=0)
    Redprintenabled = GetSetting(appname:=Trim(gblApp_Name), Section:="Redprint", key:="Redprintenabled", Default:=0)
    Priceonkitchenprint = GetSetting(appname:=Trim(gblApp_Name), Section:="PriceonKitchen", key:="Priceonkitchenprint", Default:=0)
    If Priceonkitchenprint = 0 Then
    SaveSetting Trim(gblApp_Name), "PriceonKitchen", "Priceonkitchenprint", "0"
    End If
    'Label_Height
  
    'Command.Timeout.
    
    If Trim(Server.SQL_Name) <> "" Then
        Openconnection 240, Server.SQL_Name, Server.SQL_User, Server.SQL_Password, Server.SQL_Database
    End If
    'Get Splash Screen Display values from the registery
    
   
    Select Case GetSetting(appname:=Trim(gblApp_Name), Section:="Show_Splash", key:="Value", Default:="-1")
        Case error
        shw = "Startup"
        frmSerials.Show
        '                   Startup.Show
        Case -1
            'If starting for the first time write defaults to the registery
            SaveSetting Trim(gblApp_Name), "Show_Splash", "Value", 1
            SaveSetting Trim(gblApp_Name), "Server", "Server", ""
            SaveSetting Trim(gblApp_Name), "Server", "SQL_User", "sa"
            SaveSetting Trim(gblApp_Name), "Server", "SQL_Password", ""
            SaveSetting Trim(gblApp_Name), "Server", "SQL_Database", ""
            SaveSetting Trim(gblApp_Name), "Logs", "Main_Log", App.Path & "\Logs"
            SaveSetting Trim(gblApp_Name), "Logs", "Error_Log", App.Path & "\Logs"
            SaveSetting Trim(gblApp_Name), "Redprint", "Redprintenabled", "0"
            SaveSetting Trim(gblApp_Name), "PriceonKitchen", "Priceonkitchenprint", "0"
            
            shw = "Main"
            frmSerials.Show
'           frmSplash.Show
        Startup.Show
        Case 0
            'Do not display splash if Show_Slash value is set to Zero
            
            shw = "Main"
            frmSerials.Show
        
            '                               frmMain.Show
        Case 1
            'Display splash if Show_Slash value is set to one
            
            shw = "Splash"
            frmSerials.Show
            
            '                               frmSplash.Show
            If Redprintenabled = 1 Then
            ActiveReadServer " Select Redprintdepartment from Xtra"
            If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
            Redprintdept = rs.Fields("Redprintdepartment")
            rs.MoveNext
            Next i
            End If
            End If
            End Select
    
    SafetyCode "Specials"
    
 
End Sub
Public Sub Openconnection(Timeout As Integer, ServerName As String, UserName As String, Password As String, Database As String)
    On Error GoTo trap
    ' Open a connection without using a Data Source Name (DSN).
    Set cnnMain = New ADODB.Connection
    cnnMain.ConnectionString = "driver={SQL Server};server=" & Trim(ServerName) & ";uid=" & Trim(UserName) & ";pwd=" & Trim(Password) & ";database=" & Trim(Database)
    cnnMain.ConnectionTimeout = Timeout
    cnnMain.CommandTimeout = 260
    
    cnnMain.Open
    On Error GoTo 0
    Exit Sub
trap:
'    For Each Form In Forms
'        If Form.Name = "frmServerErr" Then
'            Stop
'        End If
'    Next
    Load frmServerErr
    frmServerErr.lblCap.Caption = err.Description
    frmServerErr.Show vbModal
    End
End Sub

Public Sub Openconnectionwww(Timeout As Integer, ServerName As String, UserName As String, Password As String, Database As String)
    On Error GoTo trap
    ' Open a connection without using a Data Source Name (DSN).
    Set cnnMain = New ADODB.Connection
    cnnMain.ConnectionString = "driver={SQL Server};server=" & Trim(WWWServerName) & ";uid=" & Trim(WWWUserName) & ";pwd=" & Trim(WWWPassword) & ";database=" & Trim(WWWDatabase)
    cnnMain.ConnectionTimeout = 60
    cnnMain.Open
    On Error GoTo 0
    Exit Sub
trap:
'    For Each Form In Forms
'        If Form.Name = "frmServerErr" Then
'            Stop
'        End If
'    Next
    Load frmServerErr
    frmServerErr.lblCap.Caption = err.Description
    frmServerErr.Show vbModal
    End
End Sub




Public Sub ActiveReadServer(Query As String)
    'Debug.Print Query
    On Error GoTo trap
    If rs.State = 1 Then rs.Close
    rs.Open Query, cnnMain, adOpenStatic, adLockReadOnly
    On Error GoTo 0
    Exit Sub
trap:
    filenum = FreeFile
    Close filenum
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    If LOF(filenum) > 144000 Then
        Close filenum
        Kill Trim(App.Path) & "\Logs\ServerErrors.log"
        GoTo trap
    End If
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, err.Description
    Print #filenum, Query
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
    MsgBox err.Description, vbInformation
    Resume Next
End Sub
Public Sub ActiveReadServer1(Query As String)
    On Error GoTo trap
    If rs1.State = 1 Then rs1.Close
    rs1.Open Query, cnnMain, adOpenStatic, adLockReadOnly
    On Error GoTo 0
    Exit Sub
trap:
    filenum = FreeFile
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    If LOF(filenum) > 144000 Then
        Close filenum
        Kill Trim(App.Path) & "\Logs\ServerErrors.log"
        GoTo trap
    End If
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, Query
    Print #filenum, "*********************************************************"
    'Close before reopening in another mode.
    Close filenum
    'MsgBox Err.Description, vbInformation
    Resume Next
End Sub
Public Sub ActiveReadServer2(Query As String)
    On Error GoTo trap
    If rs2.State = 1 Then rs2.Close
    rs2.Open Query, cnnMain, adOpenStatic, adLockReadOnly
    On Error GoTo 0
    Exit Sub
trap:
    filenum = FreeFile
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    If LOF(filenum) > 144000 Then
        Close filenum
        Kill Trim(App.Path) & "\Logs\ServerErrors.log"
        GoTo trap
    End If
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, Query
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
    'MsgBox Err.Description, vbInformation
    Resume Next
End Sub
Public Sub ActiveReadServer3(Query As String)
    On Error GoTo trap
    If rs3.State = 1 Then rs3.Close
    rs3.Open Query, cnnMain, adOpenForwardOnly, adLockReadOnly
    On Error GoTo 0
    Debug.Print Query
    Exit Sub
trap:
    filenum = FreeFile
    Close filenum
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, err.Description
    Print #filenum, Query
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
    'MsgBox Err.Description, vbInformation
    Resume Next
End Sub
Public Sub ActiveReadServer4(Query As String)
    On Error GoTo trap
    If rs4.State = 1 Then rs4.Close
    rs4.Open Query, cnnMain, adOpenStatic, adLockReadOnly
    On Error GoTo 0
    Exit Sub
trap:
    filenum = FreeFile
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    If LOF(filenum) > 144000 Then
        Close filenum
        Kill Trim(App.Path) & "\Logs\ServerErrors.log"
        GoTo trap
    End If
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, Query
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
    'MsgBox Err.Description, vbInformation
    Resume Next
End Sub
Public Sub ActiveReadServer5(Query As String)
    On Error GoTo trap
    If rs5.State = 1 Then rs5.Close
    rs5.Open Query, Mid(cnnMain, 1, InStrRev(cnnMain, ";") - 1), adOpenStatic, adLockReadOnly
    On Error GoTo 0
    Exit Sub
trap:
    filenum = FreeFile
    Close filenum
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    If LOF(filenum) > 144000 Then
        Close filenum
        Kill Trim(App.Path) & "\Logs\ServerErrors.log"
        GoTo trap
    End If
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, err.Description
    Print #filenum, Query
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
    'MsgBox Err.Description, vbInformation
    Resume Next
End Sub

Public Sub ActiveReadServer6(Query As String)
    On Error GoTo trap
    If rs6.State = 1 Then rs6.Close
    rs6.Open Query, cnnMain, adOpenStatic, adLockReadOnly
    On Error GoTo 0
    Exit Sub
trap:
    filenum = FreeFile
    Close filenum
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    If LOF(filenum) > 144000 Then
        Close filenum
        Kill Trim(App.Path) & "\Logs\ServerErrors.log"
        GoTo trap
    End If
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, err.Description
    Print #filenum, Query
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
    'MsgBox Err.Description, vbInformation
    Resume Next
End Sub






Public Sub ActiveUpdateServer(Query)
    On Error GoTo trap
    us.Open Query, cnnMain, , adLockBatchOptimistic
    On Error GoTo 0
    Exit Sub
trap:
    filenum = FreeFile
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    If LOF(filenum) > 144000 Then
        Close filenum
        Kill Trim(App.Path) & "\Logs\ServerErrors.log"
        GoTo trap
    End If
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, Query & error
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
    MsgBox err.Description, vbInformation
    Resume Next
End Sub
Public Sub ActiveUpdateServer1(Query)
    On Error GoTo trap
    us1.Open Query, cnnMain, , adLockBatchOptimistic
    On Error GoTo 0
    Exit Sub
trap:
    filenum = FreeFile
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    If LOF(filenum) > 144000 Then
        Close filenum
        Kill Trim(App.Path) & "\Logs\ServerErrors.log"
        GoTo trap
    End If
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, Query
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
    MsgBox err.Description, vbInformation
    Resume Next
End Sub
Public Sub WriteMainLog()
    On Error GoTo trap
top:
    filenum = FreeFile
    Open Trim(LogFiles.MainLog) & "\MainLog.log" For Append As filenum
    If LOF(filenum) > 144000 Then
        Close filenum
        Kill Trim(LogFiles.MainLog) & "\MainLog.log"
        GoTo top
    End If
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
    On Error GoTo 0
    Exit Sub
trap:
    Select Case err.Number
        Case 76
            MkDir Trim(LogFiles.MainLog)
            ErrorLog err.Description, err.Number, " - Log File Paths Created"
    End Select
    Resume
End Sub
Public Sub ErrorLog(ErrDesription, ErrorNumber, UserText)
    filenum = FreeFile
    Open Trim(LogFiles.ErrorLog) & "\ErrorLog.log" For Append As filenum
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, "Description: " & err.Description & " - " & UserText
    Print #filenum, "Error Number: " & err.Number
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
End Sub
Public Sub SafetyCode(Location)
    Select Case Location
        Case "Debtors"
            ActiveUpdateServer "Delete from Debtors where Debtor_No not in (Select Debtor_No from Debtors)"
        
        
        Case "Specials"
            ActiveUpdateServer "Delete from Appros where Product_Code not in (Select product_code from Products)"
            DoEvents
            ActiveUpdateServer "Update Specials set Active = 0 where StartDate > '" & Date & "'"
            DoEvents
            ActiveUpdateServer "Update Specials set Active = 1 where StartDate = '" & Date & "'"
            DoEvents
            ActiveUpdateServer "Update Specials set Active = 1 where StartDate < '" & Date & "'"
            DoEvents
            ActiveUpdateServer "Update Specials set Active = 0 where StopDate < '" & Date & "'"
            DoEvents
            'ActiveUpdateServer "Update Specials set Active = 1 where StartDate > Getdate() and dateadd(d,1,StopDate) > getdate()"
        
        
        Case "Users"
            ActiveUpdateServer "Delete from Users where User_Name = 'New User'"
        Case "SOH"
            ActiveUpdateServer "Delete from Pack_Links where Product_Code not in (Select Product_Code from Products)"
            DoEvents
            ActiveUpdateServer "Delete from Quantities where Product_Code not in (Select Product_Code from Products)"
            DoEvents
            ActiveUpdateServer "Delete from Quantities where Stock_on_Hand <> 0 and Product_Code in (Select Product_Code from Products where Stock_Item = 0)"
            DoEvents
            ActiveReadServer "Select Product_Code,sum(Stock_on_Hand) as Stock_on_Hand, Location_No from Quantities where product_code in" & _
            " (Select Product_Code from Quantities group by Location_No,Product_Code having count(Product_Code)>1)" & _
            " Group by product_Code,Location_No Having Count(Location_No) > 1"
            While Not rs.EOF
                ActiveUpdateServer "Delete from Quantities where Product_Code = '" & rs.Fields("Product_Code") & "' and Location_No = " & rs.Fields("Location_No")
                ActiveUpdateServer " Insert into Quantities (Product_Code,Location_No,Stock_on_Hand) values ('" & rs.Fields("Product_Code") & "'," & rs.Fields("Location_No") & "," & rs.Fields("Stock_on_Hand") & ")"
                rs.MoveNext
            Wend
            rs.Close
            DoEvents
            ActiveUpdateServer "Delete from Quantities where Location_No not in (Select Location_No from Locations)"
        Case "Suppliers"
            ActiveUpdateServer "Delete from Supplier_Links where Product_Code not in (Select Product_Code from Products)"
            DoEvents
        Case "Recipe"
            ActiveUpdateServer "Delete from Recipes where Line_Code not in (Select Product_Code from Products) and line_Code <> '0'"
            DoEvents
            ActiveUpdateServer "Delete from Recipes where Product_Code not in (Select Product_Code from Products) and line_Code <> '0'"
            DoEvents
            ActiveUpdateServer "Update Recipes set Description=(select CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size)" & _
            " + Unit_of_Measure END from Products where Products.Product_Code=Recipes.Line_Code)+' ,' + Line_Code from Recipes Where" & _
            " Description <> (select CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size)" & _
            "+ Unit_of_Measure END from Products where Products.Product_Code=Recipes.Line_Code)+' ,' + Line_Code"
    End Select
End Sub
Public Sub CreateLocation(LocationNo)
    If frmLocations.cmdUp1.Caption = "6" Then
        frmLocations.picTopFrame1.top = 4320
        frmLocations.picSeperate1.top = 4200
        frmLocations.grdLoc.Height = 3960
        frmLocations.grdLoc.top = 4890
        frmLocations.cmdUp1.Caption = "5"
    End If
    If LocationNo = -1 Then
        ActiveReadServer "Select isnull(max(location_No),0)+1 as NewLc from locations"
        Location_No = rs.Fields("NewLc")
        rs.Close
    Else
        Location_No = LocationNo
    End If
    frmLocations.grdLoc.Rows = frmLocations.grdLoc.Rows + 1
    ActiveUpdateServer "INSERT INTO Locations(Location_No,Loc_Name, Loc_Type,Stock_Take)" & _
    " VALUES(" & Location_No & ",'New Location',0,0)"
    ActiveReadServer "Select * from locations where  Location_No =  " & Location_No
    If rs.RecordCount > 0 Then
        frmLocations.grdLoc.TextMatrix(frmLocations.grdLoc.Rows - 1, 0) = rs.Fields("Location_No")
        frmLocations.grdLoc.TextMatrix(frmLocations.grdLoc.Rows - 1, 1) = rs.Fields("Loc_Name")
        frmLocations.grdLoc.TextMatrix(frmLocations.grdLoc.Rows - 1, 2) = "Sales Location"
        frmLocations.grdLoc.TextMatrix(frmLocations.grdLoc.Rows - 1, 3) = "Enter Count "
    End If
    rs.Close
    For i = 0 To 2
        frmLocations.chkPanels(i).Value = 1
        frmLocations.chkPanels(i).Tag = "1"
    Next i
    frmLocations.grdLoc.Row = frmLocations.grdLoc.Rows - 1
    frmLocations.txtLocNum.SetFocus
    frmMain.Toolbar1.Buttons(2).Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = True
End Sub
Public Sub UpdateAveCost(Product_Code, NewCost, Quantity, Landed_Cost)
    On Error Resume Next
    ActiveReadServer "SELECT Products.Unit_Size,Products.Unit_of_Measure,Products.Product_Code, Products.Landed_Cost, Products.Ave_Cost, SUM(ISNULL(Quantities.Stock_on_Hand, 0)) " & _
    " AS Stock_on_Hand FROM Products LEFT OUTER JOIN Quantities ON Products.Product_Code = Quantities.Product_Code" & _
    " GROUP BY Products.Unit_Size,Products.Unit_of_Measure, Products.Product_Code, Products.Landed_Cost, Products.Ave_Cost" & _
    " Having(Products.Product_Code = '" & Product_Code & "')"
    NEWQTY = Quantity
    If Landed_Cost = 0 Then
        Land = NewCost
        If NewCost = 0 Then
            Land = rs.Fields("Landed_Cost")
        End If
    Else
        Land = rs.Fields("Landed_Cost")
    End If
    If rs.RecordCount > 0 Then
        SOH = rs.Fields("Stock_on_Hand")
        If SOH < 0 Then SOH = 0
        AveCost = rs.Fields("Ave_Cost")
        NewAverage = ((SOH * AveCost) + (NEWQTY * Land)) / (SOH + NEWQTY)
        Unit_of_Measure = rs.Fields("Unit_of_Measure")
        unit_Size = rs.Fields("Unit_Size")
    Else
        NewAverage = Land
    End If
    rs.Close
    ActiveUpdateServer "Update Products set Landed_Cost = " & Land & ",Ave_Cost = " & NewAverage & " where Product_Code = '" & Product_Code & "'"
    ActiveReadServer "Select * from Recipes where Line_Code = '" & Product_Code & "'"
    While Not rs.EOF
        Factor1 = 1
        Select Case UCase(Unit_of_Measure & " to " & rs.Fields("Unit_of_Measure"))
            Case "ML TO SINGLE TOT"
                Select Case unit_Size
                    Case "1000"
                        Factor1 = 1000 / 25
                        resCost = NewAverage / Factor1
                    Case "750"
                        Factor1 = 750 / 25
                        resCost = NewAverage / Factor1
                    Case "500"
                        Factor1 = 500 / 25
                        resCost = NewAverage / Factor1
                End Select
            Case "ML TO DOUBLE TOT"
                Select Case unit_Size
                    Case "1000"
                        Factor1 = 1000 / 50
                        resCost = NewAverage / Factor1
                    Case "750"
                        Factor1 = 750 / 50
                        resCost = NewAverage / Factor1
                    Case "500"
                        Factor1 = 500 / 50
                        resCost = NewAverage / Factor1
                End Select
            Case "ML TO LT"
            Case "LT TO ML"
                Factor1 = (unit_Size * 1000) / rs.Fields("Qty_Used")
                resCost = NewAverage / Factor1
            Case "KG TO G"
                Factor1 = (unit_Size * 1000) / rs.Fields("Qty_Used")
                resCost = NewAverage / Factor1
            Case "G TO KG"
            Case "ML TO ML", "G TO G", "LT TO LT", "KG TO KG"
                Factor1 = unit_Size / rs.Fields("Qty_Used")
                resCost = NewAverage / Factor1
            Case Else
                resCost = NewAverage * rs.Fields("Qty_Used")
        End Select
        ActiveUpdateServer "Update Recipes set Cost = " & resCost & " where Line_No = " & rs.Fields("Line_No")
        DoEvents
        ActiveReadServer1 "Select * from Products where Recipe_Item = 1 and Product_Code= '" & rs.Fields("Product_Code") & "'"
        While Not rs1.EOF
            ActiveUpdateServer "Update Products set Landed_Cost = (Select sum(Cost) as Ave_Cost from Recipes where Line_Type not in (6,7) and Product_code = '" & rs.Fields("Product_Code") & "'),Ave_Cost = (Select sum(Cost) as Ave_Cost from Recipes where Line_Type not in (6,7) and Product_code = '" & rs.Fields("Product_Code") & "') where Product_Code = '" & rs.Fields("Product_Code") & "'"
            rs1.MoveNext
        Wend
        rs1.Close
        rs.MoveNext
    Wend
    rs.Close
    DoEvents
    On Error GoTo 0
End Sub
Public Function GetSerialNumber(strDrive As String) As Long
Dim SerialNum As Long
Dim Res As Long
Dim Temp1 As String
Dim Temp2 As String
  
Temp1 = String$(255, Chr$(0))
Temp2 = String$(255, Chr$(0))
Res = GetVolumeInformation(strDrive, Temp1, Len(Temp1), SerialNum, 0, 0, _
Temp2, Len(Temp2))
  
' this will be the value returned by the function
GetSerialNumber = SerialNum

End Function
Public Function xfile(datestarter As String, dateender As String)
Dim xfiler As String
xfiler = Right(datestarter, 2) + Right(dateender, 2) + Mid(dateender, 4, 2) + Mid(dateender, 6, 2)
xfile = xxfiler
End Function
Function Append_EAN_Checksum(RawString As String)
Dim Position As Integer
Dim CheckSum As Integer

CheckSum = 0
For Position = 2 To 12 Step 2
      CheckSum = CheckSum + Val(Mid$(RawString, Position, 1))
Next Position
CheckSum = CheckSum * 3
For Position = 1 To 11 Step 2
     CheckSum = CheckSum + Val(Mid$(RawString, Position, 1))
Next Position
CheckSum = CheckSum Mod 10
CheckSum = 10 - CheckSum
If CheckSum = 10 Then
     CheckSum = 0
End If
Append_EAN_Checksum = RawString & Format$(CheckSum, "0")
End Function


Public Sub Backupdb()
'To use with DMO
'ActiveUpdateServer "declare @Path varchar(500) ," & _
'   "     @DBName varchar(128)" & _
'"   select @DBName = '" & Server.SQL_Database & "'" & _
'"   select @Path = 'c:\backups\' " & _
'"   declare     @FileName varchar(4000) " & _
'"       select @FileName = @Path + @DBName + '_Full_' " & _
'"                            + convert(varchar(8),getdate(),112) + '_' " & _
'"                            + replace(convert(varchar(8),getdate(),108),':','') " & _
'"                            + '.bak' " & _
'"            backup database @DBName " & _
' "               to disk = @FileName "
Dim dbname As String
dbname = Trim(Server.SQL_Database)
ActiveUpdateServer "update FULLBACKUP set @DBName = " & "'" & dbname & "'"
i = cnnMain.Execute("FULLBACKUP")
'i% = cnnMain.ExecuteSQL("FULLBACKUP")



End Sub
