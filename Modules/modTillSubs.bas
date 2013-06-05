Attribute VB_Name = "modTillSubs"


 
Public Sub PrintRecipe(Plu_No)
On Error GoTo trap
With frmSlipDetails
    PrintErr = 0
    Slip_Port = ""
    
    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then  ' Kotie 17-03 2013
        If InStr(Trim(Slip_Printer), "\\") = 0 Then
            If Slip_Port = "" Then
                Open "\\" & Comp_Name & "\" & Slip_Printer For Output As #filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        Else
            If Slip_Port = "" Then
                Open Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
        If Slip_Port <> "" Then
            If UCase(Left(Slip_Port, 2)) = "NE" Then
                Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(69) & Chr(1);
    End If
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "PRODUCT RECIPE"
    Print #filenum, TillData.Description
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    ActiveReadServer2 "Select * from Recipes where Product_Code = '" & TillData.ProductCode & "'" '
    While Not rs2.EOF
        Print #filenum, Trim(Mid(rs2.Fields("Description"), InStrRev(rs2.Fields("Description"), ",") + 1)) & " - " & Trim(Mid(rs2.Fields("Description"), 1, InStrRev(rs2.Fields("Description"), ",") - 1))
        Print #filenum, "Using " & rs2.Fields("Qty_Used") & " (" & rs2.Fields("Unit_of_Measure") & ")"
        rs2.MoveNext
        If rs2.EOF = False Then
            If Slip_Printer_Type = 0 Then
                Print #filenum, Chr(27) & Chr(77) & Chr(49);
                Print #filenum, String(40, "-")
            Else
                Print #filenum, String(33, "-")
            End If
        End If
    Wend
    If rs2.RecordCount = 0 Then
        Print #filenum, "NO RECIPE LOADED"
    End If
    rs2.Close
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "-")
    Else
        Print #filenum, String(33, "-")
    End If
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "PREPARATION METHOD"
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    ActiveReadServer2 "Select * from Preparations where Product_Code = '" & TillData.ProductCode & "'"
    If rs2.RecordCount > 0 Then
        Print #filenum, rs2.Fields("Prep_Method")
    End If
    rs2.Close
    For i = 1 To .grdFoot.Rows - 1
        If Trim(.grdFoot.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdFoot.TextMatrix(i, 1))
                Case "Line Feeds"
                    Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdFoot.TextMatrix(i, 0)));
                Case Else
                    Select Case .grdFoot.TextMatrix(i, 2)
                        Case "Left": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdFoot.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    Print #filenum, Chr(29) & "V" & Chr(49);
    Close #filenum
End With
On Error GoTo 0
Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        Dim x As Printer
        For Each x In Printers
            If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = x.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = err.Description
    DoEvents
    frmError.Show vbModal
    Close #filenum
    On Error GoTo 0
End Sub
Public Sub Print_Payout(Res_No, Room_No, Account_No, Payment, Action, Deposit_No, Tender_Type)
On Error GoTo trap
With frmSlipDetails
    PrintErr = 0
    Slip_Port = ""
    
    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then  ' Kotie 17-03 2013
        If InStr(Trim(Slip_Printer), "\\") = 0 Then
            If Slip_Port = "" Then
                Open "\\" & Comp_Name & "\" & Slip_Printer For Output As #filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        Else
            If Slip_Port = "" Then
                Open Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
        If Slip_Port <> "" Then
            If UCase(Left(Slip_Port, 2)) = "NE" Then
                Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(69) & Chr(1);
    End If
    For i = 1 To .grdHead.Rows - 1
        If Trim(.grdHead.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdHead.TextMatrix(i, 1))
                Case "Line Feeds"
                    If Slip_Printer_Type = 0 Then
                        Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdHead.TextMatrix(i, 0)));
                    End If
                Case Else
                    Select Case .grdHead.TextMatrix(i, 2)
                        Case "Left"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdHead.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "CASH PAYOUT"
    If Trim(frmSales.lblPayReason) <> "" Then Print #filenum, frmSales.lblPayReason
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "Cash Payout : " & String(18 - Len(Format(Payment, "0.00")), " ") & Format(Payment, "0.00")
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, "Date: " & Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
    Print #filenum, "Payment No: " & Format(Deposit_No, "000000")
    
    If TillData.Account_No <> "" Then
        ActiveReadServer "Select * from Suppliers where Supplier_No = '" & frmPayout.Tag & "'"
        If rs.RecordCount > 0 Then
            Print #filenum, "Payed Out to: " & Format(frmPayout.Tag, "000000") & " - " & rs.Fields("Supplier_Name")
        End If
        rs.Close
    End If
    For i = 1 To .grdFoot.Rows - 1
        If Trim(.grdFoot.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdFoot.TextMatrix(i, 1))
                Case "Line Feeds"
                    Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdFoot.TextMatrix(i, 0)));
                Case Else
                    Select Case .grdFoot.TextMatrix(i, 2)
                        Case "Left": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdFoot.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    Print #filenum, Chr(29) & "V" & Chr(49);
    Close #filenum
End With
On Error GoTo 0
Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        Dim x As Printer
        For Each x In Printers
            If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = x.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    frmSales.Tag = "1"
    Load frmError
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = err.Description
    DoEvents
    Select Case Panel_no
        Case 0: frmSales.Tag = "1"
    End Select
    frmError.Show vbModal
    Close #filenum
    On Error GoTo 0
End Sub
Public Sub ReadScale()
    stime = Timer
    On Error Resume Next
    TillData.Weight = 0
    Select Case Panel_no
        Case 0
            With frmSales
                .Tag = "1"
                .MSComm1.CommPort = Right(Devices.ScalePort, 1)
                .MSComm1.InBufferSize = 20
                .MSComm1.Settings = Devices.ScaleSet
                .MSComm1.InputLen = 0
                .MSComm1.Inputmode = comInputModeText
                If .MSComm1.PortOpen = False Then .MSComm1.PortOpen = True
                .MSComm1.Output = Chr(5)
                While Timer - stime < 0.2: Wend
                buffer$ = ""
                i = 0
                Do
                    i = i + 1
                    buffer$ = buffer$ & .MSComm1.Input
                Loop Until InStr(buffer$, Chr(30)) Or i = 20
                TillData.Weight = Val(Mid(buffer$, 1, Len(buffer$) - 1)) / 1000
                DoEvents
                If .MSComm1.PortOpen = True Then .MSComm1.PortOpen = False
            End With
        Case 1
            With frmSales1
                .Tag = "1"
                .MSComm1.CommPort = Right(Devices.ScalePort, 1)
                .MSComm1.InBufferSize = 20
                .MSComm1.Settings = Devices.ScaleSet
                .MSComm1.InputLen = 0
                .MSComm1.Inputmode = comInputModeText
                If .MSComm1.PortOpen = False Then .MSComm1.PortOpen = True
                .MSComm1.Output = Chr(5)
                While Timer - stime < 0.2: Wend
                buffer$ = ""
                i = 0
                Do
                    i = i + 1
                    buffer$ = buffer$ & .MSComm1.Input
                Loop Until InStr(buffer$, Chr(30)) Or i = 20
                TillData.Weight = Val(Mid(buffer$, 1, Len(buffer$) - 1)) / 1000
                DoEvents
                If .MSComm1.PortOpen = True Then .MSComm1.PortOpen = False
            End With
        Case 2
            With frmBar
                .Tag = "1"
                .MSComm1.CommPort = Right(Devices.ScalePort, 1)
                .MSComm1.InBufferSize = 20
                .MSComm1.Settings = Devices.ScaleSet
                .MSComm1.InputLen = 0
                .MSComm1.Inputmode = comInputModeText
                If .MSComm1.PortOpen = False Then .MSComm1.PortOpen = True
                .MSComm1.Output = Chr(5)
                While Timer - stime < 0.2: Wend
                buffer$ = ""
                i = 0
                Do
                    i = i + 1
                    buffer$ = buffer$ & .MSComm1.Input
                Loop Until InStr(buffer$, Chr(30)) Or i = 20
                TillData.Weight = Val(Mid(buffer$, 1, Len(buffer$) - 1)) / 1000
                DoEvents
                If .MSComm1.PortOpen = True Then .MSComm1.PortOpen = False
            End With
    End Select
    On Error GoTo 0
End Sub
Public Sub Print_Receive_on_Account(Res_No, Room_No, Account_No, Payment, Action, Deposit_No, Tender_Type)
On Error GoTo trap
With frmSlipDetails
    PrintErr = 0
    Slip_Port = ""
    
    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then  ' Kotie 17-03 2013
        If InStr(Trim(Slip_Printer), "\\") = 0 Then
            If Slip_Port = "" Then
                Open "\\" & Comp_Name & "\" & Slip_Printer For Output As #filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        Else
            If Slip_Port = "" Then
                Open Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
        If Slip_Port <> "" Then
            If UCase(Left(Slip_Port, 2)) = "NE" Then
                Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(69) & Chr(1);
    End If
    For i = 1 To .grdHead.Rows - 1
        If Trim(.grdHead.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdHead.TextMatrix(i, 1))
                Case "Line Feeds"
                    If Slip_Printer_Type = 0 Then
                        Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdHead.TextMatrix(i, 0)));
                    End If
                Case Else
                    Select Case .grdHead.TextMatrix(i, 2)
                        Case "Left"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdHead.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Action = 0 Then
        Print #filenum, "DEPOSIT RECEIVED"
    Else
        Print #filenum, "PAYMENT RECEIVED"
    End If
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If TillData.Change < 0 Then
        Print #filenum, Tender_Type & " Tendered : " & String(18 - Len(Format(Payment + (TillData.Change * -1), "0.00")), " ") & Format(Payment + (TillData.Change * -1), "0.00")
        Print #filenum, " Change : " & String(18 - Len(Format(TillData.Change, "0.00")), " ") & Format(TillData.Change, "0.00")
        Print #filenum, Tender_Type & " Payment : " & String(18 - Len(Format(Payment, "0.00")), " ") & Format(Payment, "0.00")
    Else
        Print #filenum, Tender_Type & " Payment : " & String(18 - Len(Format(Payment, "0.00")), " ") & Format(Payment, "0.00")
    End If
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, "Date: " & Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
    If Action = 0 Then
        Print #filenum, "Deposit No: " & Format(Deposit_No, "000000")
    Else
        Print #filenum, "Receipt No: " & Format(Deposit_No, "000000")
    End If
    
    
    If TillData.TotDiscountVal + TillData.TotDiscount <> 0 Then
        Print #filenum, "Discount Value: " & Format(TillData.TotDiscountVal + TillData.TotDiscount, "0.00")
    End If
    
    
    
    If TillData.Room_No <> 0 Then
        Print #filenum, "Charged to Room: " & Format(TillData.Room_No, "000000")
    End If
    If TillData.Account_No <> "" Then
        ActiveReadServer "Select * from Debtors where Debtor_No = '" & TillData.Account_No & "'"
        If rs.RecordCount > 0 Then
            Print #filenum, "Charged to: " & Format(TillData.Account_No, "000000") & " - " & rs.Fields("Debtor_Name")
            Print #filenum, "Balance: " & Format(rs.Fields("Balance"), "00.00")
        
        End If
        rs.Close
    End If
    For i = 1 To .grdFoot.Rows - 1
        If Trim(.grdFoot.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdFoot.TextMatrix(i, 1))
                Case "Line Feeds"
                    Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdFoot.TextMatrix(i, 0)));
                Case Else
                    Select Case .grdFoot.TextMatrix(i, 2)
                        Case "Left": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdFoot.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    Print #filenum, Chr(29) & "V" & Chr(49);
    Close #filenum
End With
On Error GoTo 0
Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        Dim x As Printer
        For Each x In Printers
            If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = x.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = err.Description
    DoEvents
    frmError.Show vbModal
    Close #filenum
    On Error GoTo 0
End Sub
Public Sub DrawerKick(Keystring$)
    Dim KickString As Variant
    Dim x As Printer
    On Error GoTo trap
    If Panel_no = 2 Then frmBar.Label1 = ""
    DoEvents
    PrintErr = 0
    Slip_Port = ""
    If Trim(Slip_Printer) = "" Or Slip_Printer = "<None>" Then
        If Panel_no = 2 Then frmBar.Label1 = "No Slip Printer"
        Exit Sub
    End If
    filenum = FreeFile
    Close #filenum
    For Each x In Printers
    
   
        If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
            Slip_Port = x.Port
            Exit For
        End If
    
  
    
'    If UCase(x.DeviceName) = UCase(Slip_Printer) Then
'    Slip_Port = (x.Port)
'    Exit For
'    End If

    
    
    Next
    If Slip_Port = "" Then
        On Error GoTo 0
        Exit Sub
    End If
    Open Slip_Port For Output As #filenum
    DoEvents
    If PrintErr = 1 Then
        Open Slip_Printer For Output As #filenum
        DoEvents
    End If
    Print #filenum, Chr(27) & Chr(64);
    Select Case UserRecord.Drawer_No
    Case 1
        KickString = Split(Devices.Drawer1KickString, ",")
        If UBound(KickString) = 0 Then Print #filenum, Chr(KickString(0));
        If UBound(KickString) = 1 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1));
        If UBound(KickString) = 2 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1)) & Chr(KickString(2));
        If UBound(KickString) = 3 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1)) & Chr(KickString(2)) & Chr(KickString(3));
    Case 2
        KickString = Split(Devices.Drawer2KickString, ",")
        If UBound(KickString) = 0 Then Print #filenum, Chr(KickString(0));
        If UBound(KickString) = 1 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1));
        If UBound(KickString) = 2 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1)) & Chr(KickString(2));
        If UBound(KickString) = 3 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1)) & Chr(KickString(2)) & Chr(KickString(3));
    Case Else
         KickString = Split(Devices.Drawer1KickString, ",")
        If UBound(KickString) = 0 Then Print #filenum, Chr(KickString(0));
        If UBound(KickString) = 1 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1));
        If UBound(KickString) = 2 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1)) & Chr(KickString(2));
        If UBound(KickString) = 3 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1)) & Chr(KickString(2)) & Chr(KickString(3));
    End Select
    Print #filenum, ""
    Close #filenum
    DoEvents
    On Error GoTo 0
    Exit Sub
trap:
    If PrintErr = 1 Then
        Close #filenum
        On Error GoTo 0
        Exit Sub
    End If
    PrintErr = 1
    'If Panel_No = 2 Then frmBar.Label1 = "Drawer Error"
    Close #filenum
    Resume Next
End Sub
Public Sub KitchenPrint()
    Dim Print_Price As Variant
    On Error GoTo trap
    NextSet = False
Redprintenabled = GetSetting(appname:=Trim(gblApp_Name), Section:="Redprint", key:="Redprintenabled", Default:=0)
  
  
  If Redprintenabled = 1 Then
  ActiveReadServer2 " Select * from xtra"
  If rs2.RecordCount > 0 Then
    Redprintdept = rs2.Fields("Redprintdepartment")
    End If
    If rs2.RecordCount > 1 Then
    rs2.MoveNext
    Redprintdept2 = rs2.Fields("Redprintdepartment")
    rs2.Close
    End If

End If


Repeat:
    On Error GoTo trap
    PrintErr = 0
    Slip_Port = ""
    Static FormLoaded As Boolean
    If FormLoaded = False Then
        Load frmPrint
        FormLoaded = True
    End If
    'frmPrint.Show 0, frmSales1
    If NextSet = True Then
        frmPrint.LoadLines2
    Else
        frmPrint.LoadLines
    End If
   
   
    CurrentPrinter = ""
    filenum = 0
    NewPrinter = False
    
  '*****************************************************************************************
    For i = 0 To frmPrint.grdPrint.Rows - 1
top:
        If UCase(CurrentPrinter) <> UCase(frmPrint.grdPrint.TextMatrix(i, 3)) Then
            If Trim(CurrentPrinter) <> "" Then
                NewPrinter = True
                Slip_Port = ""
            End If
            If filenum <> 0 Then
                Print #filenum, Chr(27) & Chr(33) & Chr(0);
                Print #filenum, String(33, "=")
                Print #filenum, Chr(27) & Chr(33) & Chr(16);
                Print #filenum, Chr(27) & Chr(100) & Chr(7);
                Print #filenum, ""
                Print #filenum, ""
                Print #filenum, ""
                Print #filenum, ""
                Print #filenum, Chr(29) & "V" & Chr(49);
                Close #filenum
                PrintErr = 0
            End If
            CurrentPrinter = frmPrint.grdPrint.TextMatrix(i, 3)
            filenum = FreeFile
            On Error GoTo trap
            If InStr(Trim(CurrentPrinter), "\\") = 0 Then
top1:
                On Error GoTo trap
                If Slip_Port = "" Then
                    If InStr(Trim(CurrentPrinter), "\\") = 0 Then
                        Open "\\" & Comp_Name & "\" & CurrentPrinter For Output As filenum2
                    Else
                        Open CurrentPrinter For Output As filenum
                    End If
                Else
                    Open Slip_Port For Output As filenum
                End If
                If Slip_Port <> "" Then
                    If UCase(Left(Slip_Port, 2)) = "NE" Then
                        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
                    Else
                        If Slip_Port = "FILE:" Then
                            Open "C:\" & x.DeviceName & ".txt" For Output As filenum
                        Else
                            Open Slip_Port For Output As filenum
                        End If
                    End If
                End If
                On Error Resume Next
                If TillData.TabNo <> 0 Then
                    Print #filenum, Chr(27) & Chr(97) & Chr(49);
                    Print #filenum, Chr(27) & Chr(33) & Chr(16);
                    Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    Print #filenum, "Tab: " & TillData.TabName
                    Print #filenum, "Barman: " & Trim(UserRecord.Name)
                    Print #filenum, Format(Date, "dd MMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Print #filenum, String(33, "=")
                    Print #filenum, Chr(27) & Chr(33) & Chr(16);
                    
                End If
                If TillData.TableNo <> 0 Then
                    Print #filenum, Chr(27) & Chr(97) & Chr(49);
                    Print #filenum, Chr(27) & Chr(33) & Chr(16);
                    Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If TillData.TableNo <> 9999 Then
                    Print #filenum, "Table No: " & TillData.TableNo
                    Else
                    Print #filenum, "Table No: Training Table 9999"
                    End If
                    Print #filenum, "Waitron: " & Trim(UserRecord.Name)
                    Print #filenum, Format(Date, "dd MMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
                    Print #filenum, "Docno: " & Trim(TillData.DocNo)
                    Print #filenum, "Pax: " & TillData.Covers
                    If NextSet = True Then Print #filenum, String(33, "=")
                    If NextSet = True Then Print #filenum, "Information Print"
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Print #filenum, String(33, "=")
                    Print #filenum, Chr(27) & Chr(33) & Chr(16);
                End If
                If TillData.TabNo = 0 And TillData.TableNo = 0 Then
                    Print #filenum, Chr(27) & Chr(97) & Chr(49);
                    Print #filenum, Chr(27) & Chr(33) & Chr(16);
                    Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    Print #filenum, "Waitron: " & Trim(UserRecord.Name)
                    Print #filenum, Format(Date, "dd MMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Print #filenum, String(33, "=")
                    Print #filenum, Chr(27) & Chr(33) & Chr(16);
                End If
                GoTo top
            Else
                GoTo top1
            End If
        End If
        On Error Resume Next
        If frmPrint.grdPrint.TextMatrix(i, 8) <> "" And frmPrint.grdPrint.TextMatrix(i, 1) <> "" Then
            ActiveReadServer "Select Dept_Name from Departments where Department_No = '" & frmPrint.grdPrint.TextMatrix(i, 8) & "'"
            Print #filenum, Chr(27) & Chr(64);
            Print #filenum, Chr(27) & Chr(69) & Chr(1);
            Print #filenum, Chr(27) & Chr(97) & Chr(49);
            Print #filenum, Chr(27) & Chr(33) & Chr(16);
            Print #filenum, Chr(27) & Chr(69) & Chr(48);
            Print #filenum, rs.Fields("Dept_Name")
            Print #filenum, Chr(27) & Chr(97) & Chr(49);
            saveDept = rs.Fields("Dept_Name")
            Print #filenum, String(33, "-")
            If TillData.TableNo = 9999 Then
            Print #filenum, "*Training Table please Ignore*"
            Print #filenum, "*Do not prepare Food or Drinks*"
                    End If
            Print #filenum, Chr(27) & Chr(97) & Chr(48);
            NewPrinter = False
            rs.Close
        Else
            If NewPrinter = True Then
                Print #filenum, Chr(27) & Chr(64);
                Print #filenum, Chr(27) & Chr(69) & Chr(1);
                Print #filenum, Chr(27) & Chr(97) & Chr(49);
                Print #filenum, Chr(27) & Chr(33) & Chr(16);
                Print #filenum, Chr(27) & Chr(69) & Chr(48);
                Print #filenum, saveDept
                Print #filenum, Chr(27) & Chr(97) & Chr(49);
                Print #filenum, String(33, "-")
                Print #filenum, Chr(27) & Chr(97) & Chr(48);
                NewPrinter = False
            End If
        End If
        Print #filenum, Chr(27) & Chr(97) & Chr(48);
        
        
        If Left(frmPrint.grdPrint.TextMatrix(i, 2), 5) = "    >" Then
            If Slip_Printer_Type = 0 Then
            'Print #filenum,
            End If
            If Redprintenabled = 1 Then
                Print #filenum, Chr$(&H1B); "r"; Chr$(49); ' Red
            End If
            'Print #filenum, String(5 - Len(frmPrint.grdPrint.TextMatrix(i, 1)), " ") & frmPrint.grdPrint.TextMatrix(i, 1) & " X " & Trim(frmPrint.grdPrint.TextMatrix(i, 2))
            Print #filenum, Trim(frmPrint.grdPrint.TextMatrix(i, 2))
            'Print #filenum, " "               'Space after Kitchen Message  Kotie 14-03-2013
            
            If Redprintenabled = 1 Then
                Print #filenum, Chr$(&H1B); "r"; Chr$(48); ' Black
            End If
            need_blank_line = True
        Else
            If need_blank_line = True Then Print #filenum, " "
            need_blank_line = False
            YYYY = Trim(frmPrint.grdPrint.TextMatrix(i, 2))
            ActiveReadServer2 " Select Department_no from Products where Description = '" & YYYY & " ' "
            xxxxx = rs2.Fields("Department_no")
            If Redprintdept <> xxxxx Or Redprintdept2 <> xxxxx Then
            
            If Redprintenabled = 1 Then
            Print #filenum, Chr$(&H1B); "r"; Chr$(48); ' Black
            End If
            'Print #filenum, frmPrint.grdPrint.TextMatrix(i, 1) & " X " & Trim(frmPrint.grdPrint.TextMatrix(i, 2))
            'End If
            End If
            If Redprintdept = xxxxx Or Redprintdept2 = xxxxx Then  'Or xxxxx = "F-1"
            If Redprintenabled = 1 Then
            Print #filenum, Chr$(&H1B); "r"; Chr$(49); ' Red
            End If
            End If
            '***********************   Price for kitchenprinting
            If Priceonkitchenprint = 1 Then
                        ActiveReadServer3 " Select Selling_Price from Products where (Description = '" & YYYY & " ') "
            Print_Price = Format(Val(frmPrint.grdPrint.TextMatrix(i, 1) * rs3.Fields("Selling_Price")), "00.00")
            
            Print #filenum, Format(frmPrint.grdPrint.TextMatrix(i, 1), "0.000") & " X " & Trim(frmPrint.grdPrint.TextMatrix(i, 2)) & " @  " & Format(Print_Price, "00.00")
            
            rs3.Close
            End If
            
            
            
            
            
            
            '**********************   Close Price for kitchenprinting
            
            
            If Priceonkitchenprint = 0 Then
            Print #filenum, frmPrint.grdPrint.TextMatrix(i, 1) & " X " & Trim(frmPrint.grdPrint.TextMatrix(i, 2))
            End If
            
            
            If Redprintenabled = 1 Then
            Print #filenum, Chr$(&H1B); "r"; Chr$(48); ' Black
            End If
            
            rs2.Close
        
        End If
    Next i
    
    Print #filenum, "Order no: " & Trim(TillData.DocNo)
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, String(33, "=")
    
    If Priceonkitchenprint = 0 Then
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    Print #filenum, Chr(27) & Chr(100) & Chr(7);
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, Chr(29) & "V" & Chr(49);
    End If
    If Priceonkitchenprint = 1 Then
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    'Print #filenum, Chr(27) & Chr(100) & Chr(7);
    
    Print #filenum, "Spicer: "
            Print #filenum, Chr(27) & Chr(97) & Chr(49);
            Print #filenum, String(33, "-")
             Print #filenum, Chr(27) & Chr(97) & Chr(48);
            Print #filenum, "Griller: "
            Print #filenum, Chr(27) & Chr(97) & Chr(49);
            Print #filenum, String(33, "-")
            Print #filenum, ""
            Print #filenum, ""
            Print #filenum, ""
            Print #filenum, ""
            Print #filenum, Chr(29) & "V" & Chr(49);
    End If
    
    
    Close #filenum
    If Kitchen_Printer_No = 2 And NextSet = False Then
        NextSet = True
        GoTo Repeat
    End If
    Select Case Panel_no
        Case 0
             With frmSales
                For i = 1 To .grdMain.Rows - 1
                    .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = &HC0FFC0
                    .grdMain.TextMatrix(i, 14) = Chr(187)
                Next i
            End With
        Case 1
             With frmSales1
                For i = 1 To .grdMain.Rows - 1
                    .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = &HC0FFC0
                    .grdMain.TextMatrix(i, 14) = Chr(187)
                Next i
            End With
        Case 2
            With frmBar
                For i = 1 To .grdMain.Rows - 1
                    .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = &HC0FFC0
                    .grdMain.TextMatrix(i, 14) = Chr(187)
                Next i
            End With
    End Select
    
    NextSet = False
    On Error GoTo 0
    Exit Sub
trap:
    If frmPrint.grdPrint.Rows = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    If PrintErr = 0 And (err.Number = 76 Or err.Number = 52) Then
        PrintErr = 1
        For Each x In Printers
            If UCase(x.DeviceName) = UCase(Trim(Mid(CurrentPrinter, (InStrRev(CurrentPrinter, "\") + 1)))) Then
                Slip_Port = x.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    frmError.Caption = " Printer Error - " & CurrentPrinter
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = err.Description
    DoEvents
    filenum1 = FreeFile
    Close filenum1
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum1
    Print #filenum1, "User Logged On"
    Print #filenum1, "User Number: " & UserRecord.User_Number
    Print #filenum1, "User Name: " & UserRecord.Name
    Print #filenum1, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum1, err.Description
    Print #filenum1, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum1
    frmError.Show vbModal
    Close #filenum
    On Error GoTo 0
End Sub
Public Sub Key_Function(Keystring$)
     
  
    Dim ii As Integer
    'Reprint with discount
    If Keystring$ <> "Reprint" Then
        If Reprintdiscount > 1 Then Reprintdiscount = 0
        Thediscounttotal = 0
    End If
    If Finalizing = True Then Exit Sub
    If Keystring <> "CL" Then
        If Validatator(Keystring) = False Then
            DisplayErr "Invalid Key Pressed"
            Exit Sub
        End If
    End If
    Barcodematch = False
    Screen.MousePointer = 11
    Select Case Panel_no
        Case 0
            If frmSales.grdMain.Cell(flexcpBackColor, frmSales.grdMain.Rows - 1, 1, frmSales.grdMain.Rows - 1, 1) = &HC0FFFF And Keystring <> "Reprint" Then
                If TillData.ShortTender = False Then
                Thediscounttotal = 0
                    If GlobalMode <> TillMode.FinMode Then
                        If GlobalMode <> TillMode.StartMode Then
                            TillData.ReturnTotal = 0
                            TillData.UllageTotal = 0
                            TillData.VoidTotal = 0
                            TillData.Tendered = 0
                            TillData.Cash = 0
                            TillData.Card = 0
                            TillData.Cheque = 0
                            TillData.Charge = 0
                            TillData.Loyalty = 0
                            TillData.TaxTotal = 0
                            TillData.TaxableSales = 0
                            TillData.NonTaxableSales = 0
                            TillData.CollectedTax = 0
                            TillData.CalculatedTax = 0
                            TillData.Corrects = 0
                            TillData.TabNo = 0
                            TillData.TabName = ""
                            TillData.CorrectCount = 0
                            TillData.VoidCount = 0
                            TillData.ReturnCount = 0
                            TillData.UllageCount = 0
                            TillData.Tipp = 0
                            TillData.TippCount = 0
                            TillData.ShortTender = False
                            TillData.UserOveride = 0
                            TillData.Discount = 0
                            TillData.DiscountVal = 0
                            TillData.TotDiscount = 0
                            TillData.TotDiscountVal = 0
                            TillData.TotDiscountCount = 0
                            TillData.TotDiscountValCount = 0
                            TillData.Account_No = ""
                            TillData.Room_No = 0
                            TillData.Res_No = 0
                            frmSales.lblDebtor = ""
                            frmSales.cmdFancy(4).Caption = "Discount"
                            Finalizing = False
                            UpdateDisplay KeyType.ClearKey
                            frmSales.grdMain.Rows = 1
                            If Keystring = "Plu" Then
                                GlobalMode = TillMode.Inputmode
                            Else
                                KeyRegister = ""
                            End If
                            Select Case Panel_no
                                Case 0: frmSales.grdMain.ColHidden(14) = True
                                Case 1: frmSales1.grdMain.ColHidden(14) = True
                                Case 2: frmBar.grdMain.ColHidden(14) = True
                            End Select
                        End If
                    End If
                End If
            End If
        
        Case 1
            If frmSales1.grdMain.Cell(flexcpBackColor, frmSales1.grdMain.Rows - 1, 1, frmSales1.grdMain.Rows - 1, 1) = &HC0FFFF And Keystring <> "Reprint" Then
                If TillData.ShortTender = False Then
                    If GlobalMode <> TillMode.FinMode Then
                        If GlobalMode <> TillMode.StartMode Then
                            TillData.ReturnTotal = 0
                            TillData.UllageTotal = 0
                            TillData.VoidTotal = 0
                            TillData.Tendered = 0
                            TillData.Cash = 0
                            TillData.Card = 0
                            TillData.Cheque = 0
                            TillData.Charge = 0
                            TillData.Loyalty = 0
                            TillData.TaxTotal = 0
                            TillData.TaxableSales = 0
                            TillData.NonTaxableSales = 0
                            TillData.CollectedTax = 0
                            TillData.CalculatedTax = 0
                            TillData.Corrects = 0
                            TillData.TabNo = 0
                            TillData.TabName = ""
                            TillData.CorrectCount = 0
                            TillData.VoidCount = 0
                            TillData.ReturnCount = 0
                            TillData.UllageCount = 0
                            TillData.Tipp = 0
                            TillData.TippCount = 0
                            TillData.ShortTender = False
                            TillData.UserOveride = 0
                            TillData.Discount = 0
                            TillData.DiscountVal = 0
                            TillData.TotDiscount = 0
                            TillData.TotDiscountVal = 0
                            TillData.TotDiscountCount = 0
                            TillData.TotDiscountValCount = 0
                            TillData.Account_No = ""
                            TillData.Room_No = 0
                            TillData.Res_No = 0
                            Finalizing = False
                            UpdateDisplay KeyType.ClearKey
                            frmSales1.grdMain.Rows = 1
                            If Keystring = "Plu" Then
                                GlobalMode = TillMode.Inputmode
                            Else
                                KeyRegister = ""
                            End If
                            Select Case Panel_no
                                Case 0: frmSales.grdMain.ColHidden(14) = True
                                Case 1: frmSales1.grdMain.ColHidden(14) = True
                                Case 2: frmBar.grdMain.ColHidden(14) = True
                            End Select
                        End If
                    End If
                End If
            End If
        Case 2
            If frmBar.grdMain.Cell(flexcpBackColor, frmBar.grdMain.Rows - 1, 1, frmBar.grdMain.Rows - 1, 1) = &HC0FFFF And Keystring <> "Reprint" Then
                If TillData.ShortTender = False Then
                    If GlobalMode <> TillMode.FinMode Then
                        If GlobalMode <> TillMode.StartMode Then
                            frmBar.grdMain.Rows = 1
                            TillData.ReturnTotal = 0
                            TillData.UllageTotal = 0
                            TillData.VoidTotal = 0
                            TillData.Tendered = 0
                            TillData.Cash = 0
                            TillData.Card = 0
                            TillData.Cheque = 0
                            TillData.Charge = 0
                            TillData.Loyalty = 0
                            TillData.TaxTotal = 0
                            TillData.TaxableSales = 0
                            TillData.NonTaxableSales = 0
                            TillData.CollectedTax = 0
                            TillData.CalculatedTax = 0
                            TillData.Corrects = 0
                            TillData.TabNo = 0
                            TillData.TabName = ""
                            TillData.CorrectCount = 0
                            TillData.VoidCount = 0
                            TillData.ReturnCount = 0
                            TillData.UllageCount = 0
                            TillData.Tipp = 0
                            TillData.TippCount = 0
                            TillData.ShortTender = False
                            TillData.UserOveride = 0
                            TillData.Discount = 0
                            TillData.DiscountVal = 0
                            TillData.TotDiscount = 0
                            TillData.TotDiscountVal = 0
                            TillData.TotDiscountCount = 0
                            TillData.TotDiscountValCount = 0
                            TillData.Account_No = ""
                            TillData.Room_No = 0
                            TillData.Res_No = 0
                            Finalizing = False
                            UpdateDisplay KeyType.ClearKey
                            If Keystring = "Plu" Then
                                GlobalMode = TillMode.Inputmode
                            Else
                                KeyRegister = ""
                            End If
                            Select Case Panel_no
                                Case 0: frmSales.grdMain.ColHidden(14) = True
                                Case 1: frmSales1.grdMain.ColHidden(14) = True
                                Case 2: frmBar.grdMain.ColHidden(14) = True
                            End Select
                        End If
                    End If
                End If
            End If
    End Select
    If (GlobalMode = TillMode.FinMode Or GlobalMode = TillMode.StartMode) And Keystring <> "Reprint" Then
        TillData.ReturnTotal = 0
        TillData.UllageTotal = 0
        TillData.VoidTotal = 0
        TillData.Tendered = 0
        TillData.Cash = 0
        TillData.Card = 0
        TillData.Cheque = 0
        TillData.Charge = 0
        TillData.Loyalty = 0
        TillData.TaxTotal = 0
        TillData.TaxableSales = 0
        TillData.NonTaxableSales = 0
        TillData.CollectedTax = 0
        TillData.CalculatedTax = 0
        TillData.Corrects = 0
        TillData.TabNo = 0
        TillData.TabName = ""
        TillData.CorrectCount = 0
        TillData.VoidCount = 0
        TillData.ReturnCount = 0
        TillData.UllageCount = 0
        TillData.Tipp = 0
        TillData.TippCount = 0
        TillData.ShortTender = False
        TillData.UserOveride = 0
        TillData.Discount = 0
        TillData.DiscountVal = 0
        TillData.TotDiscount = 0
        TillData.TotDiscountVal = 0
        TillData.TotDiscountCount = 0
        TillData.TotDiscountValCount = 0
        If Member_No = 0 Then
            TillData.Account_No = ""
        End If
        TillData.Room_No = 0
        TillData.Res_No = 0
        If Panel_no = 0 Then frmSales.cmdFancy(4).Caption = "Member No"
        Finalizing = False
        UpdateDisplay KeyType.ClearKey
        If Keystring = "Plu" Then
            GlobalMode = TillMode.Inputmode
        Else
            KeyRegister = ""
        End If
        Select Case Panel_no
            Case 0: frmSales.grdMain.ColHidden(14) = True
            Case 1: frmSales1.grdMain.ColHidden(14) = True
            Case 2: frmBar.grdMain.ColHidden(14) = True
        End Select
    End If
    Select Case Keystring
        Case "Pay Out"
            Payouts
        Case "R/A"
            TillData.Account_No = ""
            Load frmCharge
            frmCharge.lblHeading.Caption = "Please Select a Room or Debtor"
            frmCharge.lblHeading.Tag = "Please Select a Room or Debtor"
            frmCharge.Show vbModal
            If frmCharge.Tag = "" Or frmCharge.Tag = "Cancel" Then
                Unload frmCharge
                
               Screen.MousePointer = 1
                Exit Sub
            Else
                If InStr(frmCharge.Tag, "Debtor") <> 0 Or InStr(frmCharge.Tag, "Staff") <> 0 Or InStr(frmCharge.Tag, "Management") <> 0 Or InStr(frmCharge.Tag, "Travel") <> 0 Or InStr(frmCharge.Tag, "Member") <> 0 Then
                    TillData.Account_No = Trim(Mid(frmCharge.Tag, 1, InStr(frmCharge.Tag, ">") - 1))
                    Tender_Type = Trim(Mid(frmCharge.Tag, InStr(frmCharge.Tag, "|") + 1))
                    Deposit = Trim(Mid(frmCharge.Tag, InStrRev(frmCharge.Tag, ">") + 1))
                    Deposit = Val(Trim(Mid(Deposit, 1, InStr(Deposit, "|") - 1)))
                Else
                    TillData.Room_No = Trim(Mid(frmCharge.Tag, 1, InStr(frmCharge.Tag, ">") - 1))
                    Tender_Type = Trim(Mid(frmCharge.Tag, InStr(frmCharge.Tag, "|") + 1))
                    Deposit = Trim(Mid(frmCharge.Tag, InStrRev(frmCharge.Tag, ">") + 1))
                    Deposit = Val(Trim(Mid(Deposit, 1, InStr(Deposit, "|") - 1)))
                End If
                If TillData.Account_No <> "" Then
                    ActiveReadServer1 "Select isnull(max(Payment_No),0)+1 as Receipt_No from Debtor_Accounts where Transaction_Type = 'Receipt'"
                    Deposit_No = rs1.Fields("Receipt_No")
                    rs1.Close
                    
                    Balance = 0
                    ActiveReadServer "Select Balance from Debtors where Debtor_No = '" & TillData.Account_No & "'"
                    If rs.RecordCount > 0 Then
                        Balance = rs.Fields("Balance")
                    End If
                    rs.Close
                    TillData.Change = 0
' Margate
'                    If Balance - Deposit < 0 Then
'                        TillData.Change = Balance - Deposit
'                        If TillData.Change > 0 Then
'                            Deposit = Deposit + TillData.Change
'                        End If
'                    End If
                    
                    ActiveUpdateServer "INSERT INTO [Debtor_Accounts]([User_No],[Date_Time],[Transaction_Type], [Payment_No], [Account_No], [Debit], [Credit], [Balance],[Tender_Type])" & _
                    "VALUES(" & UserRecord.User_Number & ",Getdate(),'Receipt'," & Deposit_No & ",'" & TillData.Account_No & "',0," & Deposit & "," & Balance + (Deposit * -1) & ",'" & Tender_Type & "')"
                    DoEvents
                    
                    ActiveUpdateServer "Update Debtors set Balance=Balance - " & Deposit & " where Debtor_No='" & TillData.Account_No & "'"
                    
                    TillData.Cashup_No = 0
                    ActiveReadServer "Select * from Counters where User_no= " & UserRecord.User_Number & " and Finalized= 0"
                    If rs.RecordCount > 0 Then
                        TillData.Cashup_No = rs.Fields("Cashup_No")
                    Else
                        ActiveReadServer1 "Select isnull(max(Cashup_No),0)+1 as Cashup_No from Counters"
                        TillData.Cashup_No = rs1.Fields("Cashup_No")
                        rs1.Close
                        ActiveReadServer1 "Select Function_Key,Date_Time from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                        "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                        If rs1.RecordCount > 0 Then
                            If rs1.Fields("Function_Key") = 3 Then
                                ClockinTime = rs1.Fields("Date_Time") & ""
                            End If
                        End If
                        rs1.Close
                        ActiveUpdateServer "Insert into Counters (User_No,Cashup_No,Finalized,Counted,Shift_Start) values (" & UserRecord.User_Number & ", " & TillData.Cashup_No & ",0,0,'" & ClockinTime & "')"
                    End If
                    rs.Close
            
                    ActiveUpdateServer "Update Counters set " & _
                    "ReceivedonAccount_Value = isnull(ReceivedonAccount_Value,0) +" & (Deposit) & _
                    ",ReceivedonAccount_Qty=isnull(ReceivedonAccount_Qty,0) +" & 1 & _
                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                    DoEvents
                    
                    With frmSales
                        ActiveReadServer "Select Debtor_Name from Debtors where Debtor_No = '" & TillData.Account_No & "'"
                        If rs.RecordCount > 0 Then
                            Debtor_Name = rs.Fields("Debtor_Name")
                        End If
                        rs.Close
                        .grdMain.Rows = 4
                        .grdMain.TextMatrix(1, 0) = ""
                        .grdMain.TextMatrix(1, 1) = Tender_Type & " Payment Received"
                        .grdMain.TextMatrix(1, 2) = Format(Deposit, "0.00")
                        .grdMain.TextMatrix(2, 1) = "Balance Account No - " & TillData.Account_No
                        .grdMain.TextMatrix(2, 2) = Format(Balance - Deposit, "0.00")
                        .grdMain.Cell(flexcpBackColor, 2, 0, 2, 2) = &HC0FFC0
                        .grdMain.TextMatrix(3, 1) = Tender_Type
                        .grdMain.TextMatrix(3, 2) = Format(Deposit, "0.00")
                        .grdMain.Cell(flexcpBackColor, 3, 0, 3, 2) = &HC0FFFF
                        .lblKeyRegister = "Payment Received on Account No: " & TillData.Account_No & " - " & Debtor_Name
                        If TillData.Change < 0 Then
                            .lblTender.Caption = Format(TillData.Change * -1, "0.00")
                            .lblCash = "Change"
                        Else
                            .lblTender.Caption = Format(Deposit, "0.00")
                            .lblCash = Tender_Type
                        End If
                        If Balance - Deposit < 0 Then
                            
                        End If
                    End With
                    Print_Receive_on_Account 0, 0, TillData.Account_No, Deposit, 1, Deposit_No, Tender_Type
                    DoEvents
                    GlobalMode = TillMode.FinMode
                    Unload frmCharge
                    Screen.MousePointer = 1
                   
                    Exit Sub
                End If
                If TillData.Room_No <> 0 Then
                    ActiveReadServer1 "Select isnull(max(Invoice_No),0)+1 as Receipt_No from Room_Accounts where Transaction_Type = 'Receipt'"
                    Deposit_No = rs1.Fields("Receipt_No")
                    rs1.Close
                    Balance = 0
                    ActiveReadServer "Select * from Room_Accounts where Res_No = '" & TillData.Res_No & "' order by Line_No"
                    While Not rs.EOF
                        Balance = Balance + (rs.Fields("Debit") - rs.Fields("Credit"))
                        rs.MoveNext
                    Wend
                    rs.Close
                    TillData.Change = 0
                    If Balance - Deposit < 0 Then
                        TillData.Change = Balance - Deposit
                        Deposit = Deposit + TillData.Change
                    End If
                    
                    ActiveUpdateServer "INSERT INTO [Room_Accounts]([User_No],[Date_Time],[Transaction_Type], [Invoice_No], [Account_No], [Res_No], [Debit], [Credit], [Balance],[Tender_Type])" & _
                    "VALUES(" & UserRecord.User_Number & ",Getdate(),'Receipt'," & Deposit_No & ",'" & TillData.Room_No & "','" & TillData.Res_No & "',0," & Deposit & "," & Balance + (Deposit * -1) & ",'" & Tender_Type & " ')"
                    DoEvents
                    
                    TillData.Cashup_No = 0
                    ActiveReadServer "Select * from Counters where User_no= " & UserRecord.User_Number & " and Finalized= 0"
                    If rs.RecordCount > 0 Then
                        TillData.Cashup_No = rs.Fields("Cashup_No")
                    Else
                        ActiveReadServer1 "Select isnull(max(Cashup_No),0)+1 as Cashup_No from Counters"
                        TillData.Cashup_No = rs1.Fields("Cashup_No")
                        rs1.Close
                        ActiveReadServer1 "Select Function_Key,Date_Time from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                        "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                        If rs1.RecordCount > 0 Then
                            If rs1.Fields("Function_Key") = 3 Then
                                ClockinTime = rs1.Fields("Date_Time") & ""
                            End If
                        End If
                        rs1.Close
                        ActiveUpdateServer "Insert into Counters (User_No,Cashup_No,Finalized,Counted,Shift_Start) values (" & UserRecord.User_Number & ", " & TillData.Cashup_No & ",0,0,'" & ClockinTime & "')"
                    End If
                    rs.Close
            
                    ActiveUpdateServer "Update Counters set " & _
                    "ReceivedonAccount_Value = isnull(ReceivedonAccount_Value,0) +" & (Deposit) & _
                    ",ReceivedonAccount_Qty=isnull(ReceivedonAccount_Qty,0) +" & 1 & _
                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                    DoEvents
                    
                    With frmSales
                        ActiveReadServer "Select Title +' ' + First_Name +' ' + Last_Name as Debtor_Name from Reservations where Res_No = " & TillData.Res_No
                        If rs.RecordCount > 0 Then
                            Debtor_Name = rs.Fields("Debtor_Name")
                        End If
                        rs.Close
                        .grdMain.Rows = 4
                        .grdMain.TextMatrix(1, 0) = ""
                        .grdMain.TextMatrix(1, 1) = Tender_Type & " Payment Received"
                        .grdMain.TextMatrix(1, 2) = Format(Deposit, "0.00")
                        .grdMain.TextMatrix(2, 1) = "Balance Room No - " & TillData.Room_No
                        .grdMain.TextMatrix(2, 2) = Format(Balance - Deposit, "0.00")
                        .grdMain.Cell(flexcpBackColor, 2, 0, 2, 2) = &HC0FFC0
                        .grdMain.TextMatrix(3, 1) = Tender_Type
                        .grdMain.TextMatrix(3, 2) = Format(Deposit, "0.00")
                        .grdMain.Cell(flexcpBackColor, 3, 0, 3, 2) = &HC0FFFF
                        .lblKeyRegister = "Payment Received on Room No: " & TillData.Room_No & " - " & Debtor_Name
                        If TillData.Change < 0 Then
                            .lblTender.Caption = Format(TillData.Change * -1, "0.00")
                            .lblCash = "Change"
                        Else
                            .lblTender.Caption = Format(Deposit, "0.00")
                            .lblCash = Tender_Type
                        End If
                        If Balance - Deposit < 0 Then
                            
                        End If
                    End With
                    Print_Receive_on_Account TillData.Res_No, TillData.Room_No, "", Deposit, 1, Deposit_No, Tender_Type
                    DoEvents
                    GlobalMode = TillMode.FinMode
                    Unload frmCharge
                    If Panel_no = 0 Then frmSales.cmdFancy(4).Caption = "Member No"
                    If Panel_no = 0 Then frmSales.lblDebtor.Caption = ""
                    Screen.MousePointer = 1
                   
                    Exit Sub
                End If
            End If
        Case "Member No"
            GlobalMode = TillMode.Inputmode
            KeyRegister = KeyRegister & " (Load Member Number) "
            UpdateDisplay KeyType.FunctionKey
        Case "P/R"
            GlobalMode = TillMode.Inputmode
            KeyRegister = KeyRegister & " (Print Recipe) "
            UpdateDisplay KeyType.FunctionKey
        Case "Discount"
            frmSales.Tag = "1"
            Screen.MousePointer = 1
            Load frmDiscount
            frmDiscount.Show vbModal
            DoEvents
            frmSales.Tag = ""
            If TillData.Discount = 0 And TillData.DiscountVal = 0 Then
                TillData.ExtraFunc = ""
                Screen.MousePointer = 1
                Exit Sub
            End If
            Select Case TillData.ExtraFunc
                Case "Sale Discount"
                    With frmSales
                        If TillData.Discount = -10 Then
                            With frmSales
                                For i = 1 To .grdMain.Rows - 1
                                    If .grdMain.ValueMatrix(i, 0) <> 0 And .grdMain.TextMatrix(i, 8) = "" Or .grdMain.TextMatrix(i, 8) = "Return Item" Then
                                        Discount1 = 0
                                        TillData.DiscountVal = .grdMain.ValueMatrix(i, 2) - ((.grdMain.ValueMatrix(i, 4) * .grdMain.ValueMatrix(i, 0) * ((100 + .grdMain.ValueMatrix(i, 5)) / 100)) * ((100 + 10) / 100))
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 19) = TillData.DiscountVal
                                        TillData.TotDiscountCount = TillData.TotDiscountCount + 1
                                        Discount1 = TillData.DiscountVal
                                        TaxPortion = TillData.DiscountVal - (TillData.DiscountVal / ((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                        .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) - TillData.DiscountVal, "0.00")
                                        TillData.TotDiscount = TillData.TotDiscount + Discount1
                                        TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                                        If TillData.TaxRate <> 0 Then
                                            TillData.TaxableSales = TillData.TaxableSales - Discount1
                                            TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                                        Else
                                            TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                        End If
                                        Sale_Total = 0
                                         Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                                        'For ib = 1 To .grdMain.Rows - 1
                                        '    If .grdMain.TextMatrix(ib, 8) <> "Corr" Then
                                        '        If .grdMain.TextMatrix(ib, 3) <> "Subtotal" Then
                                        '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(ib, 2))
                                        '        End If
                                        '    End If
                                        'Next ib
                                        'If Val(Swiss_Round) <> 0 Then
                                        '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                                        '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                                        '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                                        '            Sale_Total = Sale_Total - q
                                        '        End If
                                        '    End If
                                        'End If
                                        Sale_Total = TillData.SaleTotal
                                        .lblTender.Caption = Format(Sale_Total, "0.00")
                                        
                                    End If
                                Next i
                                .lblKeyRegister = "Selling at Cost + 10%"
                            End With
                            Screen.MousePointer = 1
                            Exit Sub
                        End If
                        If TillData.Discount <> 0 Then
                            .lblKeyRegister.Caption = "Sale Discount - " & Format(TillData.Discount, "0.00") & "%"
                            TillData.TotDiscountCount = TillData.TotDiscountCount + 1
                            For i = .grdMain.Rows - 1 To 1 Step -1
                                If .grdMain.ValueMatrix(i, 2) <> 0 And .grdMain.TextMatrix(i, 8) = "" Then
                                    Discount1 = 0
                                    .grdMain.TextMatrix(i, 18) = TillData.Discount
                                    If TillData.Discount <> 0 Then
                                        Discount1 = .grdMain.TextMatrix(i, 2) - Val(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100))
                                        TaxPortion = (.grdMain.TextMatrix(i, 2) - Val(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100))) - (.grdMain.TextMatrix(i, 2) - Val(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100))) / (((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                        .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100), "0.00")
                                        TillData.TotDiscount = TillData.TotDiscount + Discount1
                                   
                                    Thediscounttotal = TillData.TotDiscount
                                    End If
                                   
                                    
                                    TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                                    If TillData.TaxRate <> 0 Then
                                        TillData.TaxableSales = TillData.TaxableSales - Discount1
                                        TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                                    Else
                                        TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                    End If
                                End If
                            Next i
                        Else
                            .lblKeyRegister.Caption = "Sale Discount Amount - " & Format(TillData.DiscountVal, "0.00")
                            TillData.Discount = (TillData.DiscountVal / TillData.SaleTotal) * 100
                            TillData.TotDiscountValCount = TillData.TotDiscountValCount + 1
                            For i = .grdMain.Rows - 1 To 1 Step -1
                                If .grdMain.ValueMatrix(i, 2) <> 0 And .grdMain.TextMatrix(i, 8) = "" Then
                                    Discount1 = 0
                                    .grdMain.TextMatrix(i, 18) = TillData.Discount
                                    If TillData.Discount <> 0 And i <> 1 Then
                                        Discount1 = .grdMain.TextMatrix(i, 2) - Val(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100))
                                        TaxPortion = (.grdMain.TextMatrix(i, 2) - Val(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100))) - (.grdMain.TextMatrix(i, 2) - Val(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100))) / (((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                        .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100), "0.00")
                                        TillData.TotDiscount = TillData.TotDiscount + Discount1
                                        TillData.DiscountVal = TillData.DiscountVal - Round(Discount1, 2)
                                    Else
                                        If .grdMain.Rows - 1 = 1 Then
                                            Discount1 = .grdMain.TextMatrix(i, 2) - Val(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100))
                                            TaxPortion = (.grdMain.TextMatrix(i, 2) - Val(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100))) - (.grdMain.TextMatrix(i, 2) - Val(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100))) / (((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                            .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) * ((100 - TillData.Discount) / 100), "0.00")
                                            TillData.TotDiscount = TillData.TotDiscount + Discount1
                                        Else
                                            Discount1 = TillData.DiscountVal
                                            TaxPortion = TillData.DiscountVal - (TillData.DiscountVal / ((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                            .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) - TillData.DiscountVal, "0.00")
                                            TillData.TotDiscount = TillData.TotDiscount + Discount1
                                        End If
                                    End If
                                    TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                                    If TillData.TaxRate <> 0 Then
                                        TillData.TaxableSales = TillData.TaxableSales - Discount1
                                        TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                                    Else
                                        TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                    End If
                                End If
                            Next i
                        End If
                        Sale_Total = 0
                         Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                        '
                        'For i = 1 To .grdMain.Rows - 1
                        '    If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                        '        If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                        '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                        '        End If
                        '    End If
                        'Next i
                        'If Val(Swiss_Round) <> 0 Then
                        '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                        '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                        '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                        '            Sale_Total = Sale_Total - q
                        '        End If
                        '    End If
                        'End If
                        Sale_Total = TillData.SaleTotal
                        .lblTender.Caption = Format(Sale_Total, "0.00")
                        
                        .lblCash.Caption = "Subtotal"
                        If .grdMain.Rows > 11 Then .grdMain.TopRow = .grdMain.Row - 11
                    End With
                    'Kotie 14-03-2013 22:06
                    ' Changed discounts to be done on any item
                Case "Line Discount"
                    With frmSales
                        Discount1 = 0
                        '.grdMain.Row = .grdMain.Rows - 1
                        .grdMain.TextMatrix(.grdMain.RowSel - 1, 18) = TillData.Discount
                        .grdMain.TextMatrix(.grdMain.RowSel - 1, 19) = TillData.DiscountVal
                        If TillData.Discount <> 0 Then
                            TillData.TotDiscountCount = TillData.TotDiscountCount + 1
                            .lblKeyRegister.Caption = "Line Discount - " & Format(TillData.Discount, "0.00") & "%"
                            i = .grdMain.RowSel   'Kotie
                            'i = .grdMain.Rows - 1
                            'While .grdMain.ValueMatrix(i, 2) = 0
                            '    i = i - 1
                            'Wend
                            Discount1 = Val(.grdMain.TextMatrix(.grdMain.RowSel, 2)) - Val(.grdMain.TextMatrix(.grdMain.RowSel, 2) * ((100 - TillData.Discount) / 100))
                            TaxPortion = (.grdMain.TextMatrix(.grdMain.RowSel, 2) - Val(.grdMain.TextMatrix(.grdMain.RowSel, 2) * ((100 - TillData.Discount) / 100))) - (.grdMain.TextMatrix(.grdMain.RowSel, 2) - Val(.grdMain.TextMatrix(.grdMain.RowSel, 2) * ((100 - TillData.Discount) / 100))) / (((100 + Val(.grdMain.TextMatrix(.grdMain.RowSel, 5))) / 100))
                            .grdMain.TextMatrix(.grdMain.RowSel, 2) = Format(.grdMain.TextMatrix(.grdMain.RowSel, 2) * ((100 - TillData.Discount) / 100), "0.00")
                            TillData.TotDiscount = TillData.TotDiscount + Discount1
                        Else
                            TillData.TotDiscountValCount = TillData.TotDiscountValCount + 1
                            TillData.Discount = (TillData.DiscountVal / TillData.SaleTotal) * 100
                            .lblKeyRegister.Caption = "Line Discount - " & Format(TillData.DiscountVal, "0.00")
                            Discount1 = TillData.DiscountVal
                            i = .grdMain.RowSel ' Kotie
                            'i = .grdMain.Rows - 1
                            'While .grdMain.ValueMatrix(i, 2) = 0
                            '    i = i - 1
                            'Wend
                            TaxPortion = TillData.DiscountVal - (TillData.DiscountVal / ((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                            .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) - TillData.DiscountVal, "0.00")
                            TillData.TotDiscount = TillData.TotDiscount + Discount1
                        End If
                        TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                        If TillData.TaxRate <> 0 Then
                            TillData.TaxableSales = TillData.TaxableSales - Discount1
                            TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                        Else
                            TillData.NonTaxableSales = TillData.NonTaxableSales - Val(.grdMain.TextMatrix(i, 2)) * Val(TillData.Qty)
                        End If
                        Sale_Total = 0
                       ' For i = 1 To .grdMain.Rows - 1
                       '     If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                       '         If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                       '             sale_total = Val(sale_total) + Val(.grdMain.TextMatrix(i, 2))
                       '         End If
                       '     End If
                       ' Next i
                       ' If Val(Swiss_Round) <> 0 Then
                       '     If Str(sale_total / Swiss_Round) <> Str(Round(sale_total / Swiss_Round, 0)) Then
                       '         If Right(Swiss_Round, 1) <> Right(sale_total, 1) Then
                       '             q = Round((sale_total / Swiss_Round - Int(sale_total / Swiss_Round)) * Swiss_Round, 2)
                       '             sale_total = sale_total - q
                       '         End If
                       '     End If
                       ' End If
                        Update_Sale_Total (1)
                       Sale_Total = TillData.SaleTotal
                        .lblTender.Caption = Format(Sale_Total, "0.00")
                        .lblCash.Caption = "Subtotal"
                        If .grdMain.Rows > 11 Then .grdMain.TopRow = .grdMain.Row - 11
                    End With
            End Select
        Case "Transfer Items"
            Load frmItemTransfer
            frmItemTransfer.Show vbModal
            If Panel_no = 1 Then
                frmSales1.LoadOldTable TillData.TableNo
                If frmSales1.grdMain.Rows = 1 Then
                    TillData.TableNo = 0
                    frmInput.Show
                End If
            End If
            If Panel_no = 2 Then
                frmBar.LoadOldTab TillData.TabNo
                If frmBar.grdMain.Rows = 1 Then
                    TillData.TabNo = 0
                    TillData.TabName = 0
                    frmBar.lblTab = ""
                End If
            End If
        Case "Split Bill"
            Load frmSplit
            frmSplit.Show vbModal
        Case "Service Charge"
            DoEvents
            Load frmTipp
            frmTipp.Tag = "In Sale"
            frmSales1.Tag = "1"
            frmTipp.Show vbModal
            DoEvents
            If TillData.Tipp <> 0 Then
                TillData.Keystring = Keystring
                KeyRegister = TillData.Tipp & " <" & Keystring & ">"
                UpdateDisplay KeyType.ItemizerKey
            End If
        Case "Reprint"
            PrintSlip Keystring
        Case "Kitchen Message"
            kittMess
        Case "Print Bill"
            KitchenPrint
            ActiveUpdateServer ("Insert into Print_Journal (User_No,Doc_no,Doc_Type,DateTimePrinted, Table_no)VALUES(" & UserRecord.User_Number & "," & TillData.DocNo & ",'Bill Print', getdate(),'" & TillData.TableNo & "')")
            PrintSlip Keystring
            Select Case Panel_no
                Case 1
                    If TillData.TableNo <> 0 Then
                        GoTo PlaceOrder
                    End If
                Case 2
                    If TillData.TabNo <> 0 Then
                        GoTo AddtoTab
                    End If
            End Select
        Case "View Tables"
            frmInput.Show
            DoEvents
            Screen.MousePointer = 1
            Exit Sub
        Case "Place Order"



PlaceOrder:
            With frmSales1
                KitchenPrint
                DoEvents
                ActiveReadServer "Select * from Table_Listing_View where Table_No = " & TillData.TableNo
                If rs.RecordCount > 0 Then
                    User_No = rs.Fields("User_No")
                Else
                    User_No = UserRecord.User_Number
                End If
                rs.Close
                ActiveUpdateServer "Delete from Table_Listing where Table_No= " & TillData.TableNo
                For i = 1 To .grdMain.Rows - 1
                    ActiveUpdateServer "INSERT INTO [Table_Listing]([Table_No],[Covers], [User_No], [Workstation_No], [Qty],[Short_Desc], [Line_Total], [KeyString], [Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2], [Price_Override], [Printed],[Keyregister],[Doc_No],[Locked],[User_Overide],Discount_Amt,Dicount_Value,Member_No, Table_name)" & _
                    " VALUES('" & TillData.TableNo & "','" & TillData.Covers & "','" & User_No & "','" & Workstation_No & "','" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & _
                    .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & Format(.grdMain.TextMatrix(i, 4), "00.00") & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "','" & .grdMain.TextMatrix(i, 8) & "','" & .grdMain.TextMatrix(i, 9) & "','" & .grdMain.TextMatrix(i, 10) & "','" & .grdMain.TextMatrix(i, 11) & "','" & .grdMain.TextMatrix(i, 12) & "','" & .grdMain.TextMatrix(i, 13) & "','" & .grdMain.TextMatrix(i, 14) & "','" & .grdMain.TextMatrix(i, 7) & "'," & TillData.DocNo & ",0,'" & .grdMain.ValueMatrix(i, 17) & "'," & .grdMain.ValueMatrix(i, 18) & "," & .grdMain.ValueMatrix(i, 19) & ",'" & TillData.Account_No & "','" & TillData.Table_Name & "')"
                    DoEvents
                Next i
                .lblTable = ""
                .grdMain.Rows = 1
                
                .lblCash.Caption = ""
                .lblTender.Caption = "0.00"
                .lblKeyRegister = " Order Placed for Table No: " & TillData.TableNo
                .cmdDept(6).Caption = "No Sale"
                TillData.DocNo = 0
                TillData.TableNo = 0
                TillData.Table_Name = ""
                TillData.Account_No = ""
                TillData.Covers = 0
                TillData.TotDiscount = 0
                TillData.TotDiscountVal = 0
                TillData.TotDiscountCount = 0
                TillData.TotDiscountValCount = 0
                GlobalMode = TillMode.FinMode
                DoEvents
                Select Case UserRecord.Logged_in
                    Case False
                        frmSales1.KeyPreview = False
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
                        frmSplash.Show
                        KeyCode = 0
                        KeyRegister = ""
                        frmSales1.picHoldFocus.Tag = "1"
                        .lblKeyRegister = ""
                        DoEvents
                       'Unload frmSales1
                        frmSales1.Hide
                        Screen.MousePointer = 1
                        Exit Sub
                    Case Else
                        frmInput.Show
                End Select
            End With




        Case "Create Tab"
            With frmBar
                If TillData.TabNo = 0 Then
                    Screen.MousePointer = 11
                    Load frmKeyBoard
                    frmKeyBoard.Tag = "Tabs"
                    DoEvents
                    frmBar.Tag = "1"
                    frmKeyBoard.Show vbModal
                    Select Case frmBar.lblKeyRegister.Caption
                        Case ""
                            Screen.MousePointer = 1
                            Exit Sub
                        Case Else
                            ActiveReadServer1 "Select isnull(Max(Tab_No),0)+1 as Tab_No from Tab_Listing as Tab_No"
                            If rs1.RecordCount > 0 Then
                                TillData.TabNo = rs1.Fields("Tab_No")
                                TillData.TabName = frmBar.lblKeyRegister.Caption
                                frmBar.lblTab = "Tab: " & TillData.TabName
                                frmBar.cmdKey(4).Caption = "No Sale"
                            End If
                            rs1.Close
                    End Select
                    KitchenPrint
                    frmBar.Tag = ""
                    DoEvents
                    ActiveUpdateServer "Delete from Tab_Listing where Tab_No= " & TillData.TabNo
                    For i = 1 To .grdMain.Rows - 1
                        ActiveUpdateServer "INSERT INTO [Tab_Listing]([Tab_No],Tab_Name,[Covers], [User_No], [Workstation_No], [Qty],[Short_Desc], [Line_Total], [KeyString], [Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2], [Price_Override], [Printed],[Keyregister],[Doc_No],[Locked],[User_Overide],Discount_Amt,Dicount_Value,Member_No)" & _
                        " VALUES('" & TillData.TabNo & "','" & TillData.TabName & "','" & TillData.Covers & "','" & UserRecord.User_Number & "','" & Workstation_No & "','" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & _
                        .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & Format(.grdMain.TextMatrix(i, 4), "00.00") & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "','" & .grdMain.TextMatrix(i, 8) & "','" & .grdMain.TextMatrix(i, 9) & "','" & .grdMain.TextMatrix(i, 10) & "','" & .grdMain.TextMatrix(i, 11) & "','" & .grdMain.TextMatrix(i, 12) & "','" & .grdMain.TextMatrix(i, 13) & "','" & .grdMain.TextMatrix(i, 14) & "','" & .grdMain.TextMatrix(i, 7) & "'," & TillData.DocNo & ",0,'" & .grdMain.ValueMatrix(i, 17) & "'," & .grdMain.ValueMatrix(i, 18) & "," & .grdMain.ValueMatrix(i, 19) & ",'" & TillData.Account_No & "')"
                        DoEvents
                    Next i
                    .lblTab = ""
                    .grdMain.Rows = 1
                    .lblCash.Caption = ""
                    .lblTender.Caption = "0.00"
                    .lblKeyRegister = " Order Placed for Tab: " & TillData.TabName
                    TillData.DocNo = 0
                    TillData.TabNo = 0
                    TillData.Covers = 0
                    TillData.TabName = ""
                    TillData.Account_No = ""
                    GlobalMode = TillMode.FinMode
                    .cmdFancy(3).Caption = "Create Tab"
                    DoEvents
                End If
            End With
        Case "Add to Tab"
AddtoTab:
            With frmBar
                KitchenPrint
                DoEvents
                ActiveReadServer "Select * from Tab_Listing_View where Tab_No = " & TillData.TabNo
                If rs.RecordCount > 0 Then
                    User_No = rs.Fields("User_No")
                Else
                    User_No = UserRecord.User_Number
                End If
                rs.Close
                ActiveUpdateServer "Delete from Tab_Listing where Tab_No= " & TillData.TabNo
                For i = 1 To .grdMain.Rows - 1
                    ActiveUpdateServer "INSERT INTO [Tab_Listing]([Tab_No],Tab_Name,[Covers], [User_No], [Workstation_No], [Qty],[Short_Desc], [Line_Total], [KeyString], [Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2], [Price_Override], [Printed],[Keyregister],[Doc_No],[Locked],[User_Overide],Discount_Amt,Dicount_Value,Member_No)" & _
                    " VALUES('" & TillData.TabNo & "','" & TillData.TabName & "','" & TillData.Covers & "','" & User_No & "','" & Workstation_No & "','" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & _
                    .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & Format(.grdMain.TextMatrix(i, 4), "00.00") & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "','" & .grdMain.TextMatrix(i, 8) & "','" & .grdMain.TextMatrix(i, 9) & "','" & .grdMain.TextMatrix(i, 10) & "','" & .grdMain.TextMatrix(i, 11) & "','" & .grdMain.TextMatrix(i, 12) & "','" & .grdMain.TextMatrix(i, 13) & "','" & .grdMain.TextMatrix(i, 14) & "','" & .grdMain.TextMatrix(i, 7) & "'," & TillData.DocNo & ",0,'" & .grdMain.ValueMatrix(i, 17) & "'," & .grdMain.ValueMatrix(i, 18) & "," & .grdMain.ValueMatrix(i, 19) & ",'" & TillData.Account_No & "')"
                    DoEvents
                Next i
                .lblTab = ""
                .grdMain.Rows = 1
                .lblCash.Caption = ""
                .lblTender.Caption = "0.00"
                .lblKeyRegister = " Order Placed for Tab: " & TillData.TabName
                TillData.DocNo = 0
                TillData.TabNo = 0
                TillData.Account_No = ""
                TillData.Covers = 0
                TillData.TabName = ""
                GlobalMode = TillMode.FinMode
                .cmdFancy(3).Caption = "Create Tab"
                .picSlip.Visible = False
                DoEvents
            End With
        Case "Close Table"
            If frmSales1.grdMain.Rows = 1 Then
                TillData.TableNo = 0
                frmInput.Show
                Screen.MousePointer = 1
                Exit Sub
            Else
                Screen.MousePointer = 11
                frmSales.Show
                DoEvents
                frmSales1.Hide
                Screen.MousePointer = 0
                With frmSales
                    .grdMain.Rows = frmSales1.grdMain.Rows
                    .grdMain.ColHidden(14) = frmSales1.grdMain.ColHidden(14)
                    For i = 1 To frmSales1.grdMain.Rows - 1
                        For b = 0 To frmSales1.grdMain.Cols - 1
                            .grdMain.TextMatrix(i, b) = frmSales1.grdMain.TextMatrix(i, b)
                        Next b
                        .grdMain.Cell(flexcpBackColor, i, 0, i, 2) = frmSales1.grdMain.Cell(flexcpBackColor, i, 0, i, 0)
                        .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = frmSales1.grdMain.Cell(flexcpBackColor, i, 14, i, 14)
                        .grdMain.Cell(flexcpForeColor, i, 0, i, 2) = frmSales1.grdMain.Cell(flexcpForeColor, i, 0, i, 0)
                        If frmSales1.grdMain.Cell(flexcpFontStrikethru, i, 0, i, 0) = True Then
                            .grdMain.Cell(flexcpFontStrikethru, i, 0, i, 2) = True
                        Else
                            .grdMain.Cell(flexcpFontStrikethru, i, 0, i, 2) = False
                        End If
                        If frmSales1.grdMain.Cell(flexcpFontBold, i, 0, i, 0) = True Then
                            .grdMain.Cell(flexcpFontBold, i, 0, i, 2) = True
                        Else
                            .grdMain.Cell(flexcpFontBold, i, 0, i, 2) = False
                        End If
                    Next i
                    .lblKeyRegister.Caption = frmSales1.lblKeyRegister.Caption
                    .grdMain.HighLight = flexHighlightWithFocus
                    .lblCash.Caption = frmSales1.lblCash.Caption
                    .lblTender = frmSales1.lblTender.Caption
                    If TillData.DocNo <> 0 Then
                        .cmdInput(14).Caption = "Corr"
                        .grdMain.HighLight = flexHighlightAlways
                    Else
                        .grdMain.HighLight = flexHighlightWithFocus
                        .cmdInput(14).Caption = "No Sale"
                    End If
                    .grdMain.Row = frmSales1.grdMain.Row
                    .grdMain.ShowCell frmSales1.grdMain.Row, 0
                End With
            End If
        Case "Close Tab"
            If frmBar.grdMain.Rows = 1 Then
                TillData.TabNo = 0
                TillData.TabName = 0
                frmBar.lblTab = ""
                Screen.MousePointer = 1
                Exit Sub
            Else
                Screen.MousePointer = 11
                frmSales.Show
                DoEvents
                frmBar.Hide
                Screen.MousePointer = 0
                With frmSales
                    .grdMain.Rows = frmBar.grdMain.Rows
                    .grdMain.ColHidden(14) = frmBar.grdMain.ColHidden(14)
                    For i = 1 To frmBar.grdMain.Rows - 1
                        For b = 0 To frmBar.grdMain.Cols - 1
                            .grdMain.TextMatrix(i, b) = frmBar.grdMain.TextMatrix(i, b)
                        Next b
                        .grdMain.Cell(flexcpBackColor, i, 0, i, 2) = frmBar.grdMain.Cell(flexcpBackColor, i, 0, i, 0)
                        .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = frmBar.grdMain.Cell(flexcpBackColor, i, 14, i, 14)
                        .grdMain.Cell(flexcpForeColor, i, 0, i, 2) = frmBar.grdMain.Cell(flexcpForeColor, i, 0, i, 0)
                        If frmBar.grdMain.Cell(flexcpFontStrikethru, i, 0, i, 0) = True Then
                            .grdMain.Cell(flexcpFontStrikethru, i, 0, i, 2) = True
                        Else
                            .grdMain.Cell(flexcpFontStrikethru, i, 0, i, 2) = False
                        End If
                        If frmBar.grdMain.Cell(flexcpFontBold, i, 0, i, 0) = True Then
                            .grdMain.Cell(flexcpFontBold, i, 0, i, 2) = True
                        Else
                            .grdMain.Cell(flexcpFontBold, i, 0, i, 2) = False
                        End If
                    Next i
                    .lblKeyRegister.Caption = frmBar.lblKeyRegister.Caption
                    .grdMain.HighLight = flexHighlightWithFocus
                    .lblCash.Caption = frmBar.lblCash.Caption
                    .lblTender = frmBar.lblTender.Caption
                    If TillData.DocNo <> 0 Then
                        .cmdInput(14).Caption = "Corr"
                        .grdMain.HighLight = flexHighlightAlways
                    Else
                        .grdMain.HighLight = flexHighlightWithFocus
                        .cmdInput(14).Caption = "No Sale"
                    End If
                    .grdMain.Row = frmBar.grdMain.Row
                    .grdMain.ShowCell frmBar.grdMain.Row, 0
                End With
            End If
        Case "Pickup Tab"
            If TillData.TabNo = 0 Then
                frmBar.cmdFancy(3).Caption = "Create Tab"
            Else
                frmBar.picSlip.Visible = True
            End If
            frmInput1.Show
        Case "New Table"
            frmInput.Show
        Case "x"
            GlobalMode = TillMode.Inputmode
            If Right(KeyRegister, 2) = Chr$(215) & " " Then
                KeyRegister = KeyRegister & "*"
            Else
                KeyRegister = KeyRegister & " " & Chr$(215) & " "
            End If
            UpdateDisplay KeyType.InputKey
        Case "*"
            If KeyRegister = " (Return Item) " Then
                KeyRegister = " (Return Item) *"
            End If
            GlobalMode = TillMode.Inputmode
            UpdateDisplay KeyType.InputKey
        Case "00"
            Select Case GlobalMode
                Case TillMode.StartMode
                    GlobalMode = TillMode.Inputmode
                    KeyRegister = ""
                    KeyRegister = KeyRegister & Keystring
                Case TillMode.Inputmode, TillMode.TenderMode
                    KeyRegister = KeyRegister & Keystring
            
        
            End Select
            UpdateDisplay KeyType.InputKey
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", " "
            Select Case GlobalMode
                Case TillMode.StartMode, TillMode.FinMode
                    GlobalMode = TillMode.Inputmode
                    KeyRegister = ""
                    KeyRegister = KeyRegister & Keystring
                Case TillMode.Inputmode, TillMode.TenderMode
                    KeyRegister = KeyRegister & Keystring
            End Select
            UpdateDisplay KeyType.InputKey
        Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
            Select Case GlobalMode
                Case TillMode.StartMode, TillMode.FinMode
                    GlobalMode = TillMode.Inputmode
                    KeyRegister = ""
                    KeyRegister = KeyRegister & Keystring
                Case TillMode.Inputmode, TillMode.TenderMode
                    KeyRegister = KeyRegister & Keystring
            End Select
            UpdateDisplay KeyType.InputKey
        Case "X 2", "X 3", "X 4", "X 5", "X 6", "X 10", "X 12", "X 20", "X 30"
            Select Case GlobalMode
                Case TillMode.StartMode, TillMode.FinMode
                    GlobalMode = TillMode.Inputmode
                    KeyRegister = ""
                    KeyRegister = Trim(Mid(Keystring, 2)) & " " & Chr(215) & " "
                Case TillMode.Inputmode, TillMode.TenderMode
                    KeyRegister = Trim(Mid(Keystring, 2)) & " " & Chr(215) & " "
            End Select
            UpdateDisplay KeyType.InputKey
        Case "CL"
            If TillData.DocNo = 0 Then GlobalMode = TillMode.StartMode
            KeyRegister = ""
            UpdateDisplay KeyType.FunctionKey
        
        Case "Plu"
            If InStr(KeyRegister, "*") <> 0 And Panel_no = 0 Then
                With frmSales
                    .grdFind.Rows = 1
                    frmSales.KeyPreview = False
                    GlobalMode = TillMode.Inputmode
                    If InStr(.lblKeyRegister, "(") = 0 Then
                        If InStr(.lblKeyRegister, Chr(215)) = 0 Then
                            .lblKeyRegister = ""
                        End If
                    End If
                    .adoData.ConnectionString = cnnMain.ConnectionString
                    .adoData.CursorLocation = adUseServer
                    .adoData.CursorType = adOpenStatic
                    .adoData.LockType = adLockReadOnly
                    .adoData.RecordSource = "Select Product_Code,Description,Department,SOH,Landed_Cost,Tax_Rate,Selling_Price from Product_List where Sales_Item=1  and description like '%" & Replace(Mid(KeyRegister, InStr(KeyRegister, "*")), "*", "") & "%' order by Description"
                    .adoData.Refresh
                    If .adoData.Recordset.RecordCount = 0 Then
                        .adoData.RecordSource = "Select Product_Code,Description,Department,SOH,Landed_Cost,Tax_Rate,Selling_Price from Product_List where Sales_Item=1  and Product_Code= '" & Replace(Mid(KeyRegister, InStr(KeyRegister, "*")), "*", "") & "' order by Description"
                        .adoData.Refresh
                    End If
                    .grdFind.ColHidden(4) = True
                    .cmdArr(2).Tag = ""
                    .grdFind.RowHeight(0) = 740
                    .grdFind.TextMatrix(0, 0) = " Product Code"
                    .grdFind.TextMatrix(0, 1) = "Description"
                    .grdFind.TextMatrix(0, 2) = "Department"
                    .grdFind.TextMatrix(0, 3) = "Stock on Hand "
                    .grdFind.TextMatrix(0, 4) = "Landed Cost "
                    .grdFind.TextMatrix(0, 5) = "Tax Rate "
                    .grdFind.TextMatrix(0, 6) = "Price (Incl) "
                    .grdFind.ColAlignment(0) = flexAlignLeftCenter
                    .grdFind.ColAlignment(1) = flexAlignLeftCenter
                    .grdFind.ColAlignment(2) = flexAlignLeftCenter
                    .grdFind.ColAlignment(3) = flexAlignRightCenter
                    .grdFind.ColAlignment(4) = flexAlignRightCenter
                    .grdFind.ColAlignment(5) = flexAlignRightCenter
                    .grdFind.ColAlignment(6) = flexAlignRightCenter
                    .grdFind.ColWidth(0) = .grdFind.Width * 0.12
                    .grdFind.ColWidth(1) = .grdFind.Width * 0.33
                    .grdFind.ColWidth(2) = .grdFind.Width * 0.25
                    .grdFind.ColWidth(3) = .grdFind.Width * 0.1
                    .grdFind.ColWidth(4) = .grdFind.Width * 0.1
                    .grdFind.ColWidth(5) = .grdFind.Width * 0.1
                    .grdFind.ColWidth(6) = .grdFind.Width * 0.15
                    .grdFind.ColFormat(6) = "0.00"
                    .grdFind.Col = 1
                    If .grdFind.Rows <> 1 Then .grdFind.Row = 1
                    .grdDept.Rows = 0
                    For i = 0 To 6
                        .cmdDeptStrip(i).Value = 0
                    Next i
                    .cmdDeptStrip(0).Caption = ""
                    .cmdDeptStrip(0).Picture = ""
                    DoEvents
                    ActiveReadServer "Select * from Departments_Panel1"
                    i = -1
                    b = 0
                    While Not rs.EOF
                        i = i + 1
                        .grdDept.Rows = .grdDept.Rows + 1
                        If i < 7 And Not rs.EOF Then
                            .cmdDeptStrip(i).Caption = Replace(rs.Fields("Dept_Name"), "&", "&&")
                            .cmdDeptStrip(i).Tag = rs.Fields("Department_no")
                            If .cmdDeptStrip(i).Visible = False Then .cmdDeptStrip(i).Visible = True
                            .grdDept.Row = .grdDept.Rows - 1
                            .grdDept.TextMatrix(.grdDept.Rows - 1, 0) = Replace(rs.Fields("Dept_Name"), "&", "&&")
                            .grdDept.TextMatrix(.grdDept.Rows - 1, 1) = rs.Fields("Department_No")
                        Else
                            .grdDept.TextMatrix(.grdDept.Rows - 1, 0) = Replace(rs.Fields("Dept_Name"), "&", "&&")
                            .grdDept.TextMatrix(.grdDept.Rows - 1, 1) = rs.Fields("Department_No")
                        End If
                        rs.MoveNext
                    Wend
                    rs.Close
                    For b = i + 1 To .cmdDeptStrip.Count - 1
                       .cmdDeptStrip(b).Caption = "1"
                       .cmdDeptStrip(b).Tag = ""
                       .cmdDeptStrip(b).Visible = False
                    Next b
                    .picSearch.Height = 10935
                    DoEvents
                    .picSearch.Visible = True
                    DoEvents
                    .grdFind.SetFocus
                    On Error GoTo 0
                    Screen.MousePointer = 1
                    Exit Sub
                End With
            End If
            
            
            '******************* Scaleitem embedded price or weight code ************************
            If Len(KeyRegister) = 13 Then  ' No 1
            Wholecode = ""
            Wholecode = KeyRegister
            Checksummedbarcodeean13 = Append_EAN_Checksum(Left(KeyRegister, 12))
            If Wholecode = Checksummedbarcodeean13 Then
            GoTo Goahead
             
            Else
            GoTo Alldone
            End If
Goahead:
            For ii = 20 To 29
            If Left(KeyRegister, 2) = ii Then ' No 2
            If Mid(KeyRegister, 3, 1) <> "0" Then ' No 3
            
            KeyRegister = Mid(KeyRegister, 3, 4) ' 4 characters selected
            
            GoTo Endofembeddedsearch
            Else
            If Mid(KeyRegister, 4, 1) <> "0" Then ' No 4
            KeyRegister = Mid(KeyRegister, 4, 3) ' 3 characters selected
            
            GoTo Endofembeddedsearch
            Else
            If Mid(KeyRegister, 5, 1) <> "0" Then ' No 5
            KeyRegister = Mid(KeyRegister, 5, 2) ' 2 characters selected
            
            GoTo Endofembeddedsearch
            Else
             If Mid(KeyRegister, 6, 1) <> "0" Then ' No 6
            KeyRegister = Mid(KeyRegister, 6, 1) ' 1 character selected
            GoTo Endofembeddedsearch
                End If ' No 6
                End If ' No 5
                End If ' No 4
                End If ' No 3
                End If ' No 2
                Next ii
            
            
            
            
            
            GoTo Endofembeddedsearch
            End If ' No 1
            GoTo Alldone:
         
Endofembeddedsearch:
' Decide if it is price or weight embedded   Wholecode
 
 
 
 'Weight embedded *************************************************************
 ActiveReadServer "Select * from Products where Product_Code = '" & KeyRegister & "'"
    If rs.RecordCount > 0 Then
    'check if this correlates with this items scale prefix
    If rs.Fields("Scale_item") = 1 And rs.Fields("Scaleitemtype") = "Weight" And rs.Fields("Scale_Prefix") = ii Then
    Barcodematch = True
    Dim unformattedweight As String, formattedweight As String
    unformattedweight = Mid(Wholecode, 8, 5)
    unformattedweight = Val(unformattedweight) / 1000
    unformattedweight = Format(unformattedweight, "##.###")
    formattedweight = unformattedweight & " " & Chr(215) & " " & KeyRegister
    KeyRegister = formattedweight
     GoTo Alldone
    ' Keystring = "Plu"
        
     Else
     Barcodematch = False
     End If
     
        End If

'Price embedded *************************************************************
If rs.RecordCount > 0 Then
'check if this correlates with this items scale prefix
If rs.Fields("Scale_item") = 1 And rs.Fields("Scaleitemtype") = "Price" And rs.Fields("Scale_Prefix") = ii Then
Barcodematch = True
  Dim UnformattedPrice As String, FormattedPrice As String

    UnformattedPrice = Mid(Wholecode, 8, 5)
    UnformattedPrice = (UnformattedPrice / rs.Fields("Selling_Price"))
    UnformattedPrice = Val(UnformattedPrice) / 100
    UnformattedPrice = Format(UnformattedPrice, "##.####")
   
'   If Val(Swiss_Round) <> 0 Then
'                                If Str(UnformattedPrice / Swiss_Round) <> Str(Round(UnformattedPrice / Swiss_Round, 0)) Then
'                                    If Right(Swiss_Round, 1) <> Right(UnformattedPrice, 1) Then
'                                        q = Round((UnformattedPrice / Swiss_Round - Int(UnformattedPrice / Swiss_Round)) * Swiss_Round, 2)
'                                        UnformattedPrice = UnformattedPrice - q
'                                    End If
'                                End If
'                            End If
   
  FormattedPrice = UnformattedPrice & " " & Chr(215) & " " & KeyRegister
  KeyRegister = FormattedPrice
'
'
    GoTo Alldone
   Else
   Barcodematch = False
   End If
   
        End If



If Barcodematch = False Then KeyRegister = Wholecode

Alldone:


   
            ' example KeyRegister = "22 x 1 <Plu>" , Keystring = "Plu"
            KeyRegister = KeyRegister & " <" & Keystring & ">"
            
            GetData Keystring ' Normal code or instruction

            If InStr(KeyRegister, "Print Recipe") <> 0 Then
                PrintRecipe TillData.ProductCode
                KeyRegister = ""
                Select Case Panel_no
                    Case 1
                        frmSales1.lblKeyRegister = ""
                    Case 2
                        frmBar.lblKeyRegister = ""
                End Select
                Screen.MousePointer = 1
                Exit Sub
            End If
            If TillData.ProductCode <> "" Then
                GetDocNumbers Keystring
                If TillData.Recipe = 1 And InStr(KeyRegister, "(Void)") = 0 Then
                    Load frmRecipe
                    Select Case Panel_no
                        Case 0
                            frmRecipe.top = frmSales.Picture1.top
                            frmRecipe.Left = frmSales.Picture1.Left
                            If frmRecipe.Height > 5400 Then
                                frmRecipe.Height = 7180
                                frmRecipe.Width = 6040
                            End If
                            For i = 0 To 6
                                frmRecipe.cmdCook(i).Width = frmRecipe.Width
                            Next i
                            frmSales.grdMenu.Rows = 0
                            ActiveReadServer "Select * from recipes where Product_Code= '" & TillData.ProductCode & "' order by Line_No"
                            While Not rs.EOF
                                frmSales.grdMenu.Rows = frmSales.grdMenu.Rows + 1
                                If rs.Fields("line_Type") <> 4 Then
                                    frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Rows - 1, 0) = rs.Fields("line_no")
                                    frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Rows - 1, 1) = rs.Fields("line_code")
                                    frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Rows - 1, 2) = rs.Fields("line_type")
                                    frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Rows - 1, 3) = Replace(rs.Fields("Description"), "&", "&&")
                                End If
                                rs.MoveNext
                            Wend
                            rs.Close
                            If frmSales.grdMenu.Rows > 0 Then frmSales.grdMenu.Row = 0
                        Case 1
                            frmRecipe.top = frmSales1.cmdPlu(0).top - 30
                            frmRecipe.Left = frmSales1.cmdPlu(0).Left - 30
                            If frmRecipe.Height > 5400 Then
                                frmRecipe.Height = 6810
                                frmRecipe.Width = 5880
                            End If
                            For i = 0 To 6
                                frmRecipe.cmdCook(i).Width = frmRecipe.Width
                            Next i
                            frmSales1.grdMenu.Rows = 0
                            ActiveReadServer "Select * from recipes where Product_Code= '" & TillData.ProductCode & "' order by Line_No"
                            While Not rs.EOF
                                If rs.Fields("line_Type") <> 4 Then
                                    frmSales1.grdMenu.Rows = frmSales1.grdMenu.Rows + 1
                                    frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Rows - 1, 0) = rs.Fields("line_no")
                                    frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Rows - 1, 1) = rs.Fields("line_code")
                                    frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Rows - 1, 2) = rs.Fields("line_type")
                                    frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Rows - 1, 3) = Replace(rs.Fields("Description"), "&", "&&")
                                End If
                                rs.MoveNext
                            Wend
                            rs.Close
                            If frmSales1.grdMenu.Rows > 0 Then frmSales1.grdMenu.Row = 0
                        Case 2
                            frmRecipe.top = frmBar.cmdPlu(0).top - 30
                            frmRecipe.Left = frmBar.cmdPlu(0).Left - 30
                            If frmRecipe.Height > 5400 Then
                                frmRecipe.Height = 6810
                                frmRecipe.Width = 5880
                            End If
                            For i = 0 To 6
                                frmRecipe.cmdCook(i).Width = frmRecipe.Width
                            Next i
                            frmBar.grdMenu.Rows = 0
                            ActiveReadServer "Select * from recipes where Product_Code= '" & TillData.ProductCode & "' order by Line_No"
                            While Not rs.EOF
                                If rs.Fields("line_Type") <> 4 Then
                                    frmBar.grdMenu.Rows = frmBar.grdMenu.Rows + 1
                                    frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Rows - 1, 0) = rs.Fields("line_no")
                                    frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Rows - 1, 1) = rs.Fields("line_code")
                                    frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Rows - 1, 2) = rs.Fields("line_type")
                                    frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Rows - 1, 3) = Replace(rs.Fields("Description"), "&", "&&")
                                End If
                                rs.MoveNext
                            Wend
                            rs.Close
                            If frmBar.grdMenu.Rows > 0 Then frmBar.grdMenu.Row = 0
                    End Select
                    frmRecipe.Tag = "Sale"
                    Select Case Panel_no
                        Case 0
                            If frmSales.grdMenu.Rows > 0 Then
                                frmRecipe.Show vbModal
                            End If
                        Case 1
                            If frmSales1.grdMenu.Rows > 0 Then
                                frmRecipe.Show vbModal
                            End If
                        Case 2
                            If frmBar.grdMenu.Rows > 0 Then
                                frmRecipe.Show vbModal
                            End If
                    End Select
                    frmRecipe.Tag = ""
                End If
                UpdateDisplay KeyType.ItemizerKey
            End If
        Case "Dept"
            KeyRegister = KeyRegister & " <" & Keystring & ">"
            GetData Keystring
            If TillData.DeptNo <> "" Then
                GetDocNumbers Keystring
                UpdateDisplay KeyType.ItemizerKey
            End If
        Case "Corr"
            If TillData.Tipp <> 0 Then
                TillData.Tipp = 0
                Select Case Panel_no
                    Case 0
                        frmSales.grdMain.RemoveItem (frmSales.grdMain.Rows - 1)
                        KeyRegister = "Corr - Service Charge"
                        KeyRegister = ""
                        With frmSales
                            Sale_Total = 0
                            
                             Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                            'For i = 1 To .grdMain.Rows - 1
                            '    If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            '        If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                            '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                            '        End If
                            '    End If
                            'Next i
                            'If Val(Swiss_Round) <> 0 Then
                            '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                            '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                            '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                            '            Sale_Total = Sale_Total - q
                            '        End If
                            '    End If
                            'End If
                            Sale_Total = TillData.SaleTotal
                            .lblTender.Caption = Format(Sale_Total, "0.00")
                            
                            .lblCash.Caption = "Subtotal"
                        End With
                    Case 1
                        frmSales1.grdMain.RemoveItem (frmSales1.grdMain.Rows - 1)
                        KeyRegister = "Corr - Service Charge"
                        KeyRegister = ""
                        With frmSales1
                            Sale_Total = 0
                            
                             Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                            'For i = 1 To .grdMain.Rows - 1
                            '    If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            '        If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                            '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                            '        End If
                            '    End If
                            'Next i
                            'If Val(Swiss_Round) <> 0 Then
                            '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                            '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                            '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                            '            Sale_Total = Sale_Total - q
                            '        End If
                            '    End If
                            'End If
                            Sale_Total = TillData.SaleTotal
                            .lblTender.Caption = Format(Sale_Total, "0.00")
                            
                            .lblCash.Caption = "Subtotal"
                        End With
                    Case 2
                        frmBar.grdMain.RemoveItem (frmBar.grdMain.Rows - 1)
                        KeyRegister = "Corr - Service Charge"
                        KeyRegister = ""
                        With frmBar
                            Sale_Total = 0
                             Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                            'For i = 1 To .grdMain.Rows - 1
                            '    If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            '        If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                            '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                            '        End If
                            '    End If
                            'Next i
                            'If Val(Swiss_Round) <> 0 Then
                            '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                            '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                            '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                            '            Sale_Total = Sale_Total - q
                            '        End If
                            '    End If
                            'End If
                            Sale_Total = TillData.SaleTotal
                            .lblTender.Caption = Format(Sale_Total, "0.00")
                            
                            .lblCash.Caption = "Subtotal"
                        End With
                End Select
                Screen.MousePointer = 1
                Exit Sub
            End If
            GlobalMode = TillMode.Inputmode
            KeyRegister = KeyRegister & " (Corr) "
            ExecFunction Keystring
            UpdateDisplay KeyType.FunctionKey
            KeyRegister = ""
        Case "Search"
            Select Case GlobalMode
                Case TillMode.StartMode, TillMode.FinMode
                    GlobalMode = TillMode.Inputmode
                    KeyRegister = ""
                    KeyRegister = KeyRegister & Keystring
            End Select
            If Panel_no = 0 Then frmSales.KeyPreview = False
            KeyRegister = " (Search) "
            UpdateDisplay KeyType.FunctionKey
        Case "Void"
            GlobalMode = TillMode.Inputmode
            KeyRegister = KeyRegister & " (Void) "
            UpdateDisplay KeyType.FunctionKey
        Case "Return Item"
            GlobalMode = TillMode.Inputmode
            KeyRegister = KeyRegister & " (Return Item) "
            UpdateDisplay KeyType.FunctionKey
        Case "Wastage"
            GlobalMode = TillMode.Inputmode
            KeyRegister = KeyRegister & " (Wastage) "
            UpdateDisplay KeyType.FunctionKey
        Case "Price O/V"
            GlobalMode = TillMode.Inputmode
            KeyRegister = KeyRegister & " (Price O/V) "
            UpdateDisplay KeyType.FunctionKey
        Case "No Sale"
        If UserRecord.No_Sales = False Then
        Screen.MousePointer = 0
        Exit Sub
        End If
            TillData.ExtraFunc = ""
            GlobalMode = TillMode.FinMode
            KeyRegister = "<No Sale>"
            TillCounters Keystring
            UpdateDisplay KeyType.FinalizationKey
            DrawerKick Keystring
            WriteJournals Keystring
            PrintSlip Keystring
            ExecFunction Keystring
            KeyRegister = ""
        Case "Cash", "Voucher", "Card", "Charge", "Loyalty", "Quote"
        If Finalizing = True Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        Finalizing = True
        

        
        
        
' Charge sale
If Keystring = "Charge" Then
            Screen.MousePointer = 1
            
            'Credit for member
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            If TillData.Account_No <> "" Then
            If Gotcredit = 0 Then
            Finalizing = False
                    Exit Sub
                    End If
            End If

           If TillData.Account_No = "" Or TillData.Account_No = "0" Then
                Load frmCharge
                frmCharge.Tag = ""
                frmCharge.Show vbModal
                If frmCharge.Tag = "Cancel" Or frmCharge.Tag = "" Then
                    Finalizing = False
                    Exit Sub
                    Unload frmCharge
                Else
                    Select Case Trim(Mid(frmCharge.Tag, InStr(frmCharge.Tag, ">") + 1))
                        Case "Select a Room to Charge to."
                            TillData.Room_No = Mid(frmCharge.Tag, 1, InStr(frmCharge.Tag, ">") - 1)
                            ActiveReadServer "Select Res_No from Reservations where Room_No = " & TillData.Room_No & " and Res_Type = 2"
                            TillData.Res_No = rs.Fields("Res_No")
                            rs.Close
                        Case "Select a Debtor to Charge to."
                            TillData.Account_No = Trim(Mid(frmCharge.Tag, 1, InStr(frmCharge.Tag, ">") - 1))
                        Case "Select a Staff Account to Charge to."
                            TillData.Account_No = Trim(Mid(frmCharge.Tag, 1, InStr(frmCharge.Tag, ">") - 1))
                        Case "Select a Management Account to Charge to."
                            TillData.Account_No = Trim(Mid(frmCharge.Tag, 1, InStr(frmCharge.Tag, ">") - 1))
                        Case "Select a Travel Agent to Charge to."
                            TillData.Account_No = Trim(Mid(frmCharge.Tag, 1, InStr(frmCharge.Tag, ">") - 1))
                        Case "Select a Member to Charge to."
                            TillData.Account_No = Trim(Mid(frmCharge.Tag, 1, InStr(frmCharge.Tag, ">") - 1))
                    End Select
                    Unload frmCharge
                End If
                If TillData.Account_No <> "" Then
                    ActiveReadServer "Select * from Debtor_Discounts where Debtor_No = '" & TillData.Account_No & "'"
                    While Not rs.EOF
                        Select Case Panel_no
                            Case 0
                                With frmSales
                                   For i = 1 To .grdMain.Rows - 1
                                        If .grdMain.ValueMatrix(i, 0) <> 0 And .grdMain.TextMatrix(i, 8) = "" Or .grdMain.TextMatrix(i, 8) = "Return Item" Then
                                            If rs.Fields("Department_No") = .grdMain.TextMatrix(i, 10) Then
                                                If Val(rs.Fields("Selling_Price") & "") <> 0 Then
                                                    Discount1 = 0
                                                    ActiveReadServer2 "Select Price" & Val(rs.Fields("Selling_Price") & "") & " as Price from Product_Prices where Product_Code = '" & .grdMain.TextMatrix(i, 9) & "'"
                                                    If rs2.RecordCount > 0 Then
                                                        Price = rs2.Fields("Price")
                                                    End If
                                                    rs2.Close
                                                    If Val(Price) <> 0 Then
                                                        TillData.DiscountVal = .grdMain.ValueMatrix(i, 2) - (Price * .grdMain.ValueMatrix(i, 0))
                                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 19) = TillData.DiscountVal
                                                        TillData.TotDiscountCount = TillData.TotDiscountCount + 1
                                                        Discount1 = TillData.DiscountVal
                                                        TaxPortion = TillData.DiscountVal - (TillData.DiscountVal / ((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                                        .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) - TillData.DiscountVal, "0.00")
                                                        TillData.TotDiscount = TillData.TotDiscount + Discount1
                                                        TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                                                        If TillData.TaxRate <> 0 Then
                                                            TillData.TaxableSales = TillData.TaxableSales - Discount1
                                                            TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                                                        Else
                                                            TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                                        End If
                                                        Sale_Total = 0
                                                         Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                                                        
                                                        'For ib = 1 To .grdMain.Rows - 1
                                                        '    If .grdMain.TextMatrix(ib, 8) <> "Corr" Then
                                                        '        If .grdMain.TextMatrix(ib, 3) <> "Subtotal" Then
                                                        '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(ib, 2))
                                                        '        End If
                                                        '    End If
                                                        'Next ib
                                                        'If Val(Swiss_Round) <> 0 Then
                                                        '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                                                        '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                                                        '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                                                        '            Sale_Total = Sale_Total - q
                                                        '        End If
                                                        '    End If
                                                        'End If
                                                        Sale_Total = TillData.SaleTotal
                                                        .lblTender.Caption = Format(Sale_Total, "0.00")
                                                        
                                                    End If
                                                End If
                                                If Val(rs.Fields("Cost_Disc") & "") <> 0 Then
                                                    Discount1 = 0
                                                    TillData.DiscountVal = .grdMain.ValueMatrix(i, 2) - ((.grdMain.ValueMatrix(i, 4) * .grdMain.ValueMatrix(i, 0) * ((100 + .grdMain.ValueMatrix(i, 5)) / 100)) * ((100 + Val(rs.Fields("Cost_Disc") & "")) / 100))
                                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 19) = TillData.DiscountVal
                                                    TillData.TotDiscountCount = TillData.TotDiscountCount + 1
                                                    Discount1 = TillData.DiscountVal
                                                    TaxPortion = TillData.DiscountVal - (TillData.DiscountVal / ((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                                    .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) - TillData.DiscountVal, "0.00")
                                                    TillData.TotDiscount = TillData.TotDiscount + Discount1
                                                    TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                                                    If TillData.TaxRate <> 0 Then
                                                        TillData.TaxableSales = TillData.TaxableSales - Discount1
                                                        TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                                                    Else
                                                        TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                                    End If
                                                    Sale_Total = 0
                                                     Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                                                   ' For ib = 1 To .grdMain.Rows - 1
                                                   '     If .grdMain.TextMatrix(ib, 8) <> "Corr" Then
                                                   '         If .grdMain.TextMatrix(ib, 3) <> "Subtotal" Then
                                                   '             Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(ib, 2))
                                                   '         End If
                                                   '     End If
                                                   ' Next ib
                                                   ' If Val(Swiss_Round) <> 0 Then
                                                   '     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                                                   '         If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                                                   '             q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                                                   '             Sale_Total = Sale_Total - q
                                                   '         End If
                                                   '     End If
                                                   ' End If
                                                   Sale_Total = TillData.SaleTotal
                                                    .lblTender.Caption = Format(Sale_Total, "0.00")
                                                    
                                                End If
                                                If Val(rs.Fields("Sell_Disc") & "") <> 0 Then
                                                    Discount1 = 0
                                                    TillData.DiscountVal = .grdMain.ValueMatrix(i, 2) - (.grdMain.ValueMatrix(i, 2) * ((100 - Val(rs.Fields("Sell_Disc") & "")) / 100))
                                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 19) = TillData.DiscountVal
                                                    TillData.TotDiscountCount = TillData.TotDiscountCount + 1
                                                    Discount1 = TillData.DiscountVal
                                                    TaxPortion = TillData.DiscountVal - (TillData.DiscountVal / ((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                                    .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) - TillData.DiscountVal, "0.00")
                                                    TillData.TotDiscount = TillData.TotDiscount + Discount1
                                                    TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                                                    If TillData.TaxRate <> 0 Then
                                                        TillData.TaxableSales = TillData.TaxableSales - Discount1
                                                        TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                                                    Else
                                                        TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                                    End If
                                                    Sale_Total = 0
                                                     Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                                                    
                                                    'For ib = 1 To .grdMain.Rows - 1
                                                    '    If .grdMain.TextMatrix(ib, 8) <> "Corr" Then
                                                    '        If .grdMain.TextMatrix(ib, 3) <> "Subtotal" Then
                                                    '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(ib, 2))
                                                    '        End If
                                                    '    End If
                                                    'Next ib
                                                    'If Val(Swiss_Round) <> 0 Then
                                                    '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                                                    '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                                                    '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                                                    '            Sale_Total = Sale_Total - q
                                                    '        End If
                                                    '    End If
                                                    'End If
                                                    Sale_Total = TillData.SaleTotal
                                                    .lblTender.Caption = Format(Sale_Total, "0.00")
                                                    
                                                End If
                                                If rs.Fields("Sell_Cost") <> "No" Then
                                                    Discount1 = 0
                                                    TillData.DiscountVal = .grdMain.ValueMatrix(i, 2) - ((.grdMain.ValueMatrix(i, 4) * ((100 + .grdMain.ValueMatrix(i, 5)) / 100)) * .grdMain.ValueMatrix(i, 0))
                                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 19) = TillData.DiscountVal
                                                    TillData.TotDiscountCount = TillData.TotDiscountCount + 1
                                                    Discount1 = TillData.DiscountVal
                                                    TaxPortion = TillData.DiscountVal - (TillData.DiscountVal / ((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                                    .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) - TillData.DiscountVal, "0.00")
                                                    TillData.TotDiscount = TillData.TotDiscount + Discount1
                                                    TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                                                    If TillData.TaxRate <> 0 Then
                                                        TillData.TaxableSales = TillData.TaxableSales - Discount1
                                                        TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                                                    Else
                                                        TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                                    End If
                                                    Sale_Total = 0
                                                    
                                                     Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                                                    'For ib = 1 To .grdMain.Rows - 1
                                                    '    If .grdMain.TextMatrix(ib, 8) <> "Corr" Then
                                                    '        If .grdMain.TextMatrix(ib, 3) <> "Subtotal" Then
                                                    '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(ib, 2))
                                                    '        End If
                                                    '    End If
                                                    'Next ib
                                                    'If Val(Swiss_Round) <> 0 Then
                                                    '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                                                    '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                                                    '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                                                    '            Sale_Total = Sale_Total - q
                                                    '        End If
                                                    '    End If
                                                    'End If
                                                    Sale_Total = TillData.SaleTotal
                                                    .lblTender.Caption = Format(Sale_Total, "0.00")
                                                    
                                                End If
                                            End If
                                        End If
                                    Next i
                                End With
                            Case 2
                                With frmBar
                                    For i = 1 To .grdMain.Rows - 1
                                        If .grdMain.ValueMatrix(i, 0) <> 0 Then
                                            If rs.Fields("Department_No") = .grdMain.TextMatrix(i, 10) Then
                                                If Val(rs.Fields("Selling_Price") & "") <> 0 Then
                                                    Discount1 = 0
                                                    ActiveReadServer2 "Select Price" & Val(rs.Fields("Selling_Price") & "") & " as Price from Product_Prices where Product_Code = '" & .grdMain.TextMatrix(i, 9) & "'"
                                                    If rs2.RecordCount > 0 Then
                                                        Price = rs2.Fields("Price")
                                                    End If
                                                    rs2.Close
                                                    If Val(Price) <> 0 Then
                                                         TillData.DiscountVal = .grdMain.ValueMatrix(i, 2) - (Price * .grdMain.ValueMatrix(i, 0))
                                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 19) = TillData.DiscountVal
                                                        TillData.TotDiscountCount = TillData.TotDiscountCount + 1
                                                        Discount1 = TillData.DiscountVal
                                                        TaxPortion = TillData.DiscountVal - (TillData.DiscountVal / ((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                                        .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) - TillData.DiscountVal, "0.00")
                                                        TillData.TotDiscount = TillData.TotDiscount + Discount1
                                                        TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                                                        If TillData.TaxRate <> 0 Then
                                                            TillData.TaxableSales = TillData.TaxableSales - Discount1
                                                            TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                                                        Else
                                                            TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                                        End If
                                                        Sale_Total = 0
                                                         Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                                                        'For ib = 1 To .grdMain.Rows - 1
                                                        '    If .grdMain.TextMatrix(ib, 8) <> "Corr" Then
                                                        '        If .grdMain.TextMatrix(ib, 3) <> "Subtotal" Then
                                                        '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(ib, 2))
                                                        '        End If
                                                        '    End If
                                                        'Next ib
                                                        'If Val(Swiss_Round) <> 0 Then
                                                        '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                                                        '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                                                        '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                                                        '            Sale_Total = Sale_Total - q
                                                        '        End If
                                                        '    End If
                                                        'End If
                                                        Sale_Total = TillData.SaleTotal
                                                        .lblTender.Caption = Format(Sale_Total, "0.00")
                                                        
                                                    End If
                                                End If
                                                If Val(rs.Fields("Cost_Disc") & "") <> 0 Then
                                                    Discount1 = 0
                                                    TillData.DiscountVal = .grdMain.ValueMatrix(i, 2) - ((.grdMain.ValueMatrix(i, 4) * .grdMain.ValueMatrix(i, 0) * ((100 + .grdMain.ValueMatrix(i, 5)) / 100)) * ((100 + Val(rs.Fields("Cost_Disc") & "")) / 100))
                                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 19) = TillData.DiscountVal
                                                    TillData.TotDiscountCount = TillData.TotDiscountCount + 1
                                                    Discount1 = TillData.DiscountVal
                                                    TaxPortion = TillData.DiscountVal - (TillData.DiscountVal / ((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                                    .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) - TillData.DiscountVal, "0.00")
                                                    TillData.TotDiscount = TillData.TotDiscount + Discount1
                                                    TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                                                    If TillData.TaxRate <> 0 Then
                                                        TillData.TaxableSales = TillData.TaxableSales - Discount1
                                                        TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                                                    Else
                                                        TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                                    End If
                                                    Sale_Total = 0
                                                     Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                                                   ' For ib = 1 To .grdMain.Rows - 1
                                                   '     If .grdMain.TextMatrix(ib, 8) <> "Corr" Then
                                                   '         If .grdMain.TextMatrix(ib, 3) <> "Subtotal" Then
                                                   '             Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(ib, 2))
                                                   '         End If
                                                   '     End If
                                                   ' Next ib
                                                   ' If Val(Swiss_Round) <> 0 Then
                                                   '     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                                                   '         If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                                                   '             q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                                                   '             Sale_Total = Sale_Total - q
                                                   '         End If
                                                   '     End If
                                                   ' End If
                                                    Sale_Total = TillData.SaleTotal
                                                    .lblTender.Caption = Format(Sale_Total, "0.00")
                                                    
                                                End If
                                                If Val(rs.Fields("Sell_Disc") & "") <> 0 Then
                                                    Discount1 = 0
                                                    TillData.DiscountVal = .grdMain.ValueMatrix(i, 2) - (.grdMain.ValueMatrix(i, 2) * ((100 - Val(rs.Fields("Sell_Disc") & "")) / 100))
                                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 19) = TillData.DiscountVal
                                                    TillData.TotDiscountCount = TillData.TotDiscountCount + 1
                                                    Discount1 = TillData.DiscountVal
                                                    TaxPortion = TillData.DiscountVal - (TillData.DiscountVal / ((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                                    .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) - TillData.DiscountVal, "0.00")
                                                    TillData.TotDiscount = TillData.TotDiscount + Discount1
                                                    TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                                                    If TillData.TaxRate <> 0 Then
                                                        TillData.TaxableSales = TillData.TaxableSales - Discount1
                                                        TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                                                       
                                                    Else
                                                        TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                                    End If
                                                    Sale_Total = 0
                                                     Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                                                   ' For ib = 1 To .grdMain.Rows - 1
                                                   '     If .grdMain.TextMatrix(ib, 8) <> "Corr" Then
                                                   '         If .grdMain.TextMatrix(ib, 3) <> "Subtotal" Then
                                                   '             Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(ib, 2))
                                                   '         End If
                                                   '     End If
                                                   ' Next ib
                                                   ' If Val(Swiss_Round) <> 0 Then
                                                   '     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                                                   '        q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                                                   '         Sale_Total = Sale_Total - q
                                                   '     End If
                                                   ' End If
                                                   Sale_Total = TillData.SaleTotal
                                                    .lblTender.Caption = Format(Sale_Total, "0.00")
                                                    
                                                End If
                                                If rs.Fields("Sell_Cost") <> "No" Then
                                                    Discount1 = 0
                                                    TillData.DiscountVal = .grdMain.ValueMatrix(i, 2) - ((.grdMain.ValueMatrix(i, 4) * ((100 + .grdMain.ValueMatrix(i, 5)) / 100)) * .grdMain.ValueMatrix(i, 0))
                                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 19) = TillData.DiscountVal
                                                    TillData.TotDiscountCount = TillData.TotDiscountCount + 1
                                                    Discount1 = TillData.DiscountVal
                                                    TaxPortion = TillData.DiscountVal - (TillData.DiscountVal / ((100 + Val(.grdMain.TextMatrix(i, 5))) / 100))
                                                    .grdMain.TextMatrix(i, 2) = Format(.grdMain.TextMatrix(i, 2) - TillData.DiscountVal, "0.00")
                                                    TillData.TotDiscount = TillData.TotDiscount + Discount1
                                                    TillData.TaxTotal = TillData.TaxTotal - TaxPortion
                                                    If TillData.TaxRate <> 0 Then
                                                        TillData.TaxableSales = TillData.TaxableSales - Discount1
                                                        TillData.CollectedTax = TillData.CollectedTax - TaxPortion
                                                    Else
                                                        TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                                    End If
                                                    Sale_Total = 0
                                                     Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                                                   ' For ib = 1 To .grdMain.Rows - 1
                                                   '     If .grdMain.TextMatrix(ib, 8) <> "Corr" Then
                                                   '         If .grdMain.TextMatrix(ib, 3) <> "Subtotal" Then
                                                   '             Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(ib, 2))
                                                   '         End If
                                                   '     End If
                                                   ' Next ib
                                                   ' If Val(Swiss_Round) <> 0 Then
                                                   '     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                                                   '        q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                                                   '         Sale_Total = Sale_Total - q
                                                   '     End If
                                                    'End If
                                                    Sale_Total = TillData.SaleTotal
                                                    .lblTender.Caption = Format(Sale_Total, "0.00")
                                                    
                                                End If
                                            End If
                                        End If
                                    Next i
                                End With
                        End Select
                        rs.MoveNext
                    Wend
                    rs.Close
                End If
            End If
        End If
Cash:
          ' Kotie 11/04/2013
          '  If TillData.Change = 0 Or TillData.Change > 0 Then
          '      Select Case Keystring
          '          Case "Cash"
          '              If UserRecord.Draw_Cash = True Then DrawerKick Keystring
          '          Case "Voucher"
          '              If UserRecord.Draw_Cheque = True Then DrawerKick Keystring
          '          Case "Card"
          '              If UserRecord.Draw_Card = True Then DrawerKick Keystring
          '          Case "Charge"
          '              If UserRecord.Draw_Charge = True Then DrawerKick Keystring
          '          Case "Loyalty"
          '              If UserRecord.Draw_Loyalty = True Then DrawerKick Keystring
          '      End Select
          '  End If
            
            TillData.ExtraFunc = ""
            GlobalMode = TillMode.FinMode
            KeyRegister = KeyRegister & "<" & Keystring & ">"
            CalculateChange Keystring
            
            ' Kotie 11/04/2013
            If TillData.Change = 0 Or TillData.Change > 0 Then
                Select Case Keystring
                    Case "Cash"
                        If UserRecord.Draw_Cash = True Then DrawerKick Keystring
                    Case "Voucher"
                        If UserRecord.Draw_Cheque = True Then DrawerKick Keystring
                    Case "Card"
                        If UserRecord.Draw_Card = True Then DrawerKick Keystring
                    Case "Charge"
                        If UserRecord.Draw_Charge = True Then DrawerKick Keystring
                    Case "Loyalty"
                        If UserRecord.Draw_Loyalty = True Then DrawerKick Keystring
                End Select
            End If
            
            TillCounters Keystring
            UpdateDisplay KeyType.FinalizationKey
            If AskLog = 1 Then
                Load frmAskLoc
                Select Case Panel_no
                    Case 0
                        If frmAskLoc.Height > 5400 Then
                            frmAskLoc.Height = 7180
                            frmAskLoc.Width = 6040
                        End If
                        frmAskLoc.top = frmSales.Picture1.top
                        frmAskLoc.Left = frmSales.Picture1.Left
                    Case 2
                        If frmAskLoc.Height > 5400 Then
                            frmAskLoc.Height = 6810
                            frmAskLoc.Width = 5880
                        End If
                        frmAskLoc.top = frmBar.cmdPlu(0).top - 30
                        frmAskLoc.Left = frmBar.cmdPlu(0).Left - 30
                End Select
                frmAskLoc.Show vbModal
            End If
            If Round(TillData.Change, 2) = 0 Or Round(TillData.Change, 2) > 0 Then
                KitchenPrint
                WriteJournals Keystring
                LastTab = TillData.TabNo
                LastTable = TillData.TableNo
                Last_KeyString = Keystring
                PrintSlip Keystring
            End If
            ExecFunction Keystring
            If AskLog = 1 Then
                Unload frmAskLoc
            End If
            DoEvents
            
            'Logout users that must not remain logged in after a sale
            '**************************************************
            'Waiter
            If UserRecord.uType = 3 And Panel_no = 0 And GlobalMode = TillMode.FinMode Then
                Select Case UserRecord.Logged_in
                    Case False
                        frmSales.KeyPreview = False
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
                        frmSales.grdMain.Rows = 1
                        frmSplash.Show
                        KeyCode = 0
                        KeyRegister = ""
                        frmSales.picHoldFocus.Tag = "1"
                        DoEvents
                        frmSales.Hide
                        Finalizing = False
                        Screen.MousePointer = 1
                        Exit Sub
                    Case True
                        frmSales.picHoldFocus.Tag = "1"
                        Screen.MousePointer = 11
                        frmSales1.grdMain.Rows = 1
                        frmSales1.lblTender.Caption = "0.00"
                        frmSales1.lblCash.Caption = ""
                        frmSales1.lblKeyRegister.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
                        Finalizing = False
                        frmSales1.Show
                        DoEvents
                        frmSales.Hide
                        Finalizing = False
                        Screen.MousePointer = 0
                End Select
            End If
            'Barman
            If UserRecord.uType = 4 And Panel_no = 2 And GlobalMode = TillMode.FinMode Then
                Select Case UserRecord.Logged_in
                    Case False
                        frmBar.KeyPreview = False
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
                        frmSales.grdMain.Rows = 1
                        frmSplash.cmdChange.Visible = True
                        frmSplash.cmdChange.Caption = "Change " & Format(TillData.Change, "0.00")
'                        frmSplash.cmdChange.TextDescrRM.Text = "Total " & Format(TillData.SaleTotal, "0.00")
                        frmSplash.Show
                        KeyCode = 0
                        KeyRegister = ""
                        frmBar.picHoldFocus.Tag = "1"
                        DoEvents
                        frmBar.Hide
                        Finalizing = False
                        Screen.MousePointer = 1
                        Exit Sub
                    Case True
                        Screen.MousePointer = 11
                        frmBar.Show
                        frmBar.picHoldFocus.Tag = "1"
                        DoEvents
                        Screen.MousePointer = 0
                End Select
            End If
            
          'Cashier
            If UserRecord.uType = 8 And Panel_no = 0 And GlobalMode = TillMode.FinMode Then
                Select Case UserRecord.Logged_in
                    Case False
                        frmSales.KeyPreview = False
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
                        frmSales.grdMain.Rows = 1
                        frmSplash.cmdChange.Visible = True
                        frmSplash.cmdChange.Caption = "Change " & Format(TillData.Change, "0.00")
'                        frmSplash.cmdChange.TextDescrRM.Text = "Total " & Format(TillData.SaleTotal, "0.00")
                        frmSplash.Show
                        KeyCode = 0
                        KeyRegister = ""
                        frmBar.picHoldFocus.Tag = "1"
                        DoEvents
                        frmSales.Hide
                        Finalizing = False
                        Screen.MousePointer = 1
                        Exit Sub
                    Case True
                        Screen.MousePointer = 11
                        frmSales.Show
                        frmSales.picHoldFocus.Tag = "1"
                        DoEvents
                        Screen.MousePointer = 0
                End Select
            End If
       '  *************************************************

            KeyRegister = ""
            If TillData.Tendered > TillData.SaleTotal Or TillData.Tendered = TillData.SaleTotal Then
                AskMember
            End If
            Finalizing = False
        Case "Subtotal"
            GlobalMode = TillMode.Inputmode
            KeyRegister = KeyRegister & " (Subtotal) "
            ExecFunction Keystring
            UpdateDisplay KeyType.FunctionKey
        Case "R200-00", "R100-00", "R50-00", "R20-00", "R10-00"
            GlobalMode = TillMode.Inputmode
            ExecFunction Keystring
            UpdateDisplay KeyType.FunctionKey
            Keystring = "Cash"
            GoTo Cash
    End Select
 
    
    Screen.MousePointer = 1
End Sub




Private Sub kittMess()
    Screen.MousePointer = 11
    Select Case Panel_no
        Case 0
     
        Case 1
            frmSales1.Enabled = False
        Case 2
            frmBar.Enabled = False
            'frmBar.lblPayReason = ""
            
    End Select

    For Each Form In Forms
        If Form.Name = "frmKeyBoard" Then
            GoTo far
        End If
    Next
    Load frmKeyBoard
far:
    DoEvents
    frmKeyBoard.Tag = "Message"
    DoEvents
    Screen.MousePointer = 0
    
    Select Case Panel_no  ' Kotie 14-03-2013 09:43
        Case 0

        Case 1
        frmSales1.Tag = "1"
        frmSales1.Enabled = True
        frmKeyBoard.Show vbModal
        If frmSales1.lblPayReason <> "" Then
            frmSales1.grdMain.Rows = frmSales1.grdMain.Rows + 1
            frmSales1.grdMain.Row = frmSales1.grdMain.Rows - 1
            frmSales1.grdMain.TextMatrix(frmSales1.grdMain.Row, 1) = "    >" & frmSales1.lblPayReason
            frmSales1.grdMain.TextMatrix(frmSales1.grdMain.Rows - 1, 14) = "P"
            frmSales1.grdMain.TextMatrix(frmSales1.grdMain.Rows - 1, 11) = TillData.Kitchen1
            frmSales1.grdMain.TextMatrix(frmSales1.grdMain.Rows - 1, 12) = TillData.Kitchen2
            frmSales1.grdMain.Cell(flexcpBackColor, frmSales1.grdMain.Rows - 1, 14, frmSales1.grdMain.Rows - 1, 14) = &HC0FFFF
            frmSales1.grdMain.Cell(flexcpForeColor, frmSales1.grdMain.Rows - 1, 0, frmSales1.grdMain.Rows - 1, 2) = &HC00000
            frmSales1.grdMain.Cell(flexcpFontBold, frmSales1.grdMain.Rows - 1, 0, frmSales1.grdMain.Rows - 1, 2) = True
        End If
        
        Case 2
        frmBar.Tag = "1"
        frmBar.Enabled = True
        frmKeyBoard.Show vbModal
        If frmSales1.lblPayReason <> "" Then
            frmBar.grdMain.Rows = frmBar.grdMain.Rows + 1
            frmBar.grdMain.Row = frmBar.grdMain.Rows - 1
            frmBar.grdMain.TextMatrix(frmBar.grdMain.Row, 1) = "    >" & frmSales1.lblPayReason
            frmBar.grdMain.TextMatrix(frmBar.grdMain.Rows - 1, 14) = "P"
            frmBar.grdMain.TextMatrix(frmBar.grdMain.Rows - 1, 11) = TillData.Kitchen1
            frmBar.grdMain.TextMatrix(frmBar.grdMain.Rows - 1, 12) = TillData.Kitchen2
            frmBar.grdMain.Cell(flexcpBackColor, frmBar.grdMain.Rows - 1, 14, frmBar.grdMain.Rows - 1, 14) = &HC0FFFF
            frmBar.grdMain.Cell(flexcpForeColor, frmBar.grdMain.Rows - 1, 0, frmBar.grdMain.Rows - 1, 2) = &HC00000
            frmBar.grdMain.Cell(flexcpFontBold, frmBar.grdMain.Rows - 1, 0, frmBar.grdMain.Rows - 1, 2) = True
        End If
    End Select

End Sub
Public Sub PrintSlip(Keystring$)
Screen.MousePointer = 11
ImPrinting = True
PrintErr = 0
Slip_Port = ""
If Trim(Slip_Printer) = "" Or Slip_Printer = "<None>" Then
    ImPrinting = False
    Screen.MousePointer = 0
    Exit Sub
End If
Select Case Panel_no
    Case 0
        frmSales.Tag = "1"
    Case 2
        frmBar.Tag = "1"
End Select
If Trim(Slip_Printer) = "<A4 Wide Invoice>" Then
    rptInvoice1.Show vbModal
    ImPrinting = False
    Exit Sub
End If
If Trim(Slip_Printer) = "<Choose Printer>" Then
    Load frmPrintChoose
    frmPrintChoose.top = 3300
    frmPrintChoose.Left = 4650
    If frmPrintChoose.Height > 5400 Then
        frmPrintChoose.Height = 7180
        frmPrintChoose.Width = 6040
    End If
    DoEvents
    Screen.MousePointer = 0
    frmPrintChoose.Show vbModal
    Select Case frmPrintChoose.Tag
        Case "<A4 Wide Invoice>"
            Unload frmPrintChoose
            rptInvoice1.Show vbModal
            ImPrinting = False
            Exit Sub
        Case Else
            Slip_Printer = frmPrintChoose.Tag
            Unload frmPrintChoose
    End Select
    Screen.MousePointer = 0
    ImPrinting = False
    Slip_Printer = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Slip_Printer")
    Slip_Printer_Type = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Slip_Printer_Type", Default:=0)
    Slip_PrinterPort = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Slip_PrinterPort", Default:=0)
    
    
    Exit Sub
End If
Screen.MousePointer = 0
On Error GoTo trap
If Keystring = "Print Bill Table" Then Keystring = "Print Bill"
If Keystring = "Print Bill Tab" Then Keystring = "Print Bill"
Load frmPrint
frmPrint.LoadLines1

nosale = False
Select Case Keystring
    Case "Cash", "Voucher", "Charge", "Card", "Loyalty", "No Sale", "Quote"
        Select Case Panel_no
            Case 0
                If frmSales.cmdSlip.Caption = "Slip off" Then
                    If TillData.TotDiscountVal + TillData.TotDiscount <> 0 Then
                        Thediscounttotal = TillData.TotDiscount
                        Reprintdiscount = Reprintdiscount + 2
                    End If
                    
                    
                    
                    
                    If Keystring = "Charge" Then
                        frmSales.Tag = "1"
                        Load frmQuestion
                        frmQuestion.Tag = "Fin"
                        frmQuestion.lblCap = "Do you want to Print a Slip?"
                        frmQuestion.Show vbModal
                        Select Case GQAnswer
                            Case "Yes"
                                GoTo far
                            Case Else
                                On Error GoTo 0
                                ImPrinting = False
                                Exit Sub
                        End Select
                    Else
                        On Error GoTo 0
                        ImPrinting = False
                        Exit Sub
                    End If
                End If
            Case 2
                If frmBar.cmdSlip1.Caption = "Slip off" Then
                    If Keystring = "Charge" Then
                        frmBar.Tag = "1"
                        Load frmQuestion
                        frmQuestion.lblCap = "Do you want to Print a Slip?"
                        frmQuestion.Tag = "Fin"
                        frmQuestion.Show vbModal
                        Select Case GQAnswer
                            Case "Yes"
                                GQAnswer = ""
                                GoTo far
                            Case Else
                                GQAnswer = ""
                                On Error GoTo 0
                                ImPrinting = False
                                Exit Sub
                        End Select
                    Else
                        On Error GoTo 0
                        ImPrinting = False
                        LastTab = TillData.TabNo
                        Last_KeyString = Keystring
                        Exit Sub
                    End If
                End If
        End Select
far:
End Select
Last_KeyString = ""
Reprint:
'Currentwork
Reprintplease = False
'Currentwork
With frmSlipDetails
    filenum = FreeFile
    Close #filenum
    
    'temp-----------
  '  Open "c:\print.prn" For Output As #filenum
    '--------------
    If Slip_PrinterPort = 0 Then ' Kotie 17-03-2013
        If InStr(Trim(Slip_Printer), "\\") = 0 Then
            If Slip_Port = "" Then
                Open "\\" & Comp_Name & "\" & Slip_Printer For Output As #filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        Else
            If Slip_Port = "" Then
                Open Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
        If Slip_Port <> "" Then
            If UCase(Left(Slip_Port, 2)) = "NE" Then
                Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
        Print #filenum, Chr(27) & Chr(64);
        If Slip_Printer_Type = 0 Then
            Print #filenum, Chr(27) & Chr(69) & Chr(1);
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
    
    If Keystring = "Print Bill" Then
        If TillData.Print_Count > 0 Then
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
            Print #filenum, "* * * REPRINT * * *"
        End If
    End If
    
    For i = 1 To .grdHead.Rows - 1
        If Trim(.grdHead.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdHead.TextMatrix(i, 1))
                Case "Line Feeds"
                    If Slip_Printer_Type = 0 Then
                        Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdHead.TextMatrix(i, 0)));
                    End If
                Case Else
                    Select Case .grdHead.TextMatrix(i, 2)
                        Case "Left"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdHead.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    If frmPrint.grdPrint.Rows > 0 Then
        If frmPrint.grdPrint.TextMatrix(0, 2) <> "No Sale" Then
            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
            Print #filenum, Chr(27) & Chr(33) & Chr(16);
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
            If Keystring = "Print Bill" Then
                Print #filenum, "PRO FORMA INVOICE"
            Else
                If Keystring = "Reprint" Then
                    Print #filenum, "TAX INVOICE"
                Else
                    If Keystring = "Quote" Then
                        Print #filenum, "QUOTATION"
                    Else
                        Print #filenum, "TAX INVOICE"
                    End If
                End If
            End If
            If Keystring = "Print Bill" Then
                If TillData.Print_Count > 0 Then
                    Print #filenum, "* * * REPRINT * * *"
                End If
            End If
            Print #filenum, Chr(27) & Chr(33) & Chr(0);
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
            If Slip_Printer_Type = 0 Then
                Print #filenum, Chr(27) & Chr(77) & Chr(49);
                Print #filenum, String(40, "=")
            Else
                Print #filenum, String(33, "=")
            End If
        End If
    End If
    If TillData.SaleTotal = 0 And frmPrint.grdPrint.Rows - 1 = 0 Then 'No Sale
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
        Print #filenum, Chr(27) & Chr(33) & Chr(16);
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
        If frmPrint.grdPrint.TextMatrix(0, 2) = "No Sale" Then
            Print #filenum, "No Sale"
            nosale = True
        Else
            Print #filenum, "Sale Cancelled"
        End If
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    End If
    For i = 0 To frmPrint.grdPrint.Rows - 1
        If frmPrint.grdPrint.TextMatrix(i, 8) <> "" And frmPrint.grdPrint.TextMatrix(i, 1) <> "" Then
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
            Print #filenum, Chr(27) & Chr(33) & Chr(16);
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
            ActiveReadServer "Select Dept_Name from Departments where Department_No = '" & frmPrint.grdPrint.TextMatrix(i, 8) & "'"
                Print #filenum, rs.Fields("Dept_Name")
            rs.Close
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
            Print #filenum, Chr(27) & Chr(33) & Chr(0);
            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
        End If
        If PrintZeroItems = 0 And frmPrint.grdPrint.ValueMatrix(i, 3) = 0 Then
            If TillData.SaleTotal <> 0 Then GoTo skip
        End If
        If PrintVoids = 0 And Left(frmPrint.grdPrint.TextMatrix(i, 4), 4) = "Void" Then
            If TillData.SaleTotal <> 0 Then GoTo skip
        End If
        If frmPrint.grdPrint.TextMatrix(i, 2) = "Service Charge" Then
            TillData.Tipp = frmPrint.grdPrint.ValueMatrix(i, 3)
            If TillData.SaleTotal <> 0 Then GoTo skip
        End If
        If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(33) & Chr(0);
        Select Case frmPrint.grdPrint.TextMatrix(i, 1)
            Case ""
                If (Left(frmPrint.grdPrint.TextMatrix(i, 2), 5) = "    >") And (frmPrint.grdPrint.ValueMatrix(i, 3) <> 0) Then
                    If 40 - Len(frmPrint.grdPrint.TextMatrix(i, 1) & " x " & frmPrint.grdPrint.TextMatrix(i, 2) & " @ R" & Price & frmPrint.grdPrint.TextMatrix(i, 3)) < 1 Then
                       Print #filenum, frmPrint.grdPrint.TextMatrix(i, 1) & "x" & Left(frmPrint.grdPrint.TextMatrix(i, 2), 26) & _
                       String(40 - Len(frmPrint.grdPrint.TextMatrix(i, 1) & "x" & Left(frmPrint.grdPrint.TextMatrix(i, 2), 26) & frmPrint.grdPrint.TextMatrix(i, 3)), " ") & _
                       frmPrint.grdPrint.TextMatrix(i, 3)
                    Else
                       Print #filenum, frmPrint.grdPrint.TextMatrix(i, 1) & frmPrint.grdPrint.TextMatrix(i, 2) & _
                       String(40 - Len(frmPrint.grdPrint.TextMatrix(i, 1) & frmPrint.grdPrint.TextMatrix(i, 2) & frmPrint.grdPrint.TextMatrix(i, 3)), " ") & _
                       frmPrint.grdPrint.TextMatrix(i, 3)
                    End If
                End If
            Case -1
                If 40 - Len(frmPrint.grdPrint.TextMatrix(i, 1) & " x " & frmPrint.grdPrint.TextMatrix(i, 2) & " @ R" & Price & frmPrint.grdPrint.TextMatrix(i, 3)) < 1 Then
                    Print #filenum, frmPrint.grdPrint.TextMatrix(i, 1) & "x" & Left(frmPrint.grdPrint.TextMatrix(i, 2), 26) & _
                    String(40 - Len(frmPrint.grdPrint.TextMatrix(i, 1) & "x" & Left(frmPrint.grdPrint.TextMatrix(i, 2), 26) & frmPrint.grdPrint.TextMatrix(i, 3)), " ") & _
                    frmPrint.grdPrint.TextMatrix(i, 3)
                Else
                    Print #filenum, frmPrint.grdPrint.TextMatrix(i, 1) & "x" & frmPrint.grdPrint.TextMatrix(i, 2) & _
                    String(40 - Len(frmPrint.grdPrint.TextMatrix(i, 1) & "x" & frmPrint.grdPrint.TextMatrix(i, 2) & frmPrint.grdPrint.TextMatrix(i, 3)), " ") & _
                    frmPrint.grdPrint.TextMatrix(i, 3)
                End If
            Case 1
                If Slip_Printer_Type = 0 Then
                    Print #filenum, frmPrint.grdPrint.TextMatrix(i, 1) & "x" & frmPrint.grdPrint.TextMatrix(i, 2) & _
                    String(40 - Len(frmPrint.grdPrint.TextMatrix(i, 1) & "x" & frmPrint.grdPrint.TextMatrix(i, 2) & frmPrint.grdPrint.TextMatrix(i, 3)), " ") & _
                    frmPrint.grdPrint.TextMatrix(i, 3)
                Else
                    Print #filenum, Mid(frmPrint.grdPrint.TextMatrix(i, 1) & "x" & frmPrint.grdPrint.TextMatrix(i, 2), 1, 33) & _
                    String(33 - Len(Mid(frmPrint.grdPrint.TextMatrix(i, 1) & "x" & frmPrint.grdPrint.TextMatrix(i, 2) & frmPrint.grdPrint.TextMatrix(i, 3), 1, 33)), " ") & _
                    frmPrint.grdPrint.TextMatrix(i, 3)
                End If
            Case Else
                Price = Format(Str(Val(frmPrint.grdPrint.TextMatrix(i, 3)) / Val(frmPrint.grdPrint.TextMatrix(i, 1))), "0.00")
                If 40 - Len(frmPrint.grdPrint.TextMatrix(i, 1) & " x " & frmPrint.grdPrint.TextMatrix(i, 2) & " @ R" & Price & frmPrint.grdPrint.TextMatrix(i, 3)) < 1 Then
                    Print #filenum, frmPrint.grdPrint.TextMatrix(i, 1) & "x" & Mid(frmPrint.grdPrint.TextMatrix(i, 2), 1, 20) & " @ R" & Price & _
                    String(40 - Len(frmPrint.grdPrint.TextMatrix(i, 1) & "x" & Mid(frmPrint.grdPrint.TextMatrix(i, 2), 1, 20) & " @ R" & Price & frmPrint.grdPrint.TextMatrix(i, 3)), " ") & _
                    frmPrint.grdPrint.TextMatrix(i, 3)
                Else
                    Print #filenum, frmPrint.grdPrint.TextMatrix(i, 1) & "x" & frmPrint.grdPrint.TextMatrix(i, 2) & " @ R" & Price & _
                    String(40 - Len(frmPrint.grdPrint.TextMatrix(i, 1) & "x" & frmPrint.grdPrint.TextMatrix(i, 2) & " @ R" & Price & frmPrint.grdPrint.TextMatrix(i, 3)), " ") & _
                    frmPrint.grdPrint.TextMatrix(i, 3)
                End If
        End Select
skip:
    Next i
    If nosale = False Then
        If Slip_Printer_Type = 0 Then
            Print #filenum, String(40, "-")
        Else
            Print #filenum, String(33, "-")
        End If
    End If
    If Keystring = "Print Bill" Then
        If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
        Print #filenum, Chr(27) & Chr(33) & Chr(16);
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
        If TillData.Tipp <> 0 Then
            If TillData.TableNo <> 9999 Then
            Print #filenum, "Sub Total: " & String(18 - Len(Format(TillData.SaleTotal - TillData.Tipp, "0.00")), " ") & Format(TillData.SaleTotal - TillData.Tipp, "0.00")
            End If
            If TillData.TableNo = 9999 Then
            Print #filenum, "Training Table *DO NOT PAY*"
            End If
            Print #filenum, " "
            Print #filenum, " Gratuity: " & String(18 - Len(Format(TillData.Tipp, "0.00")), " ") & Format(TillData.Tipp, "0.00")
            Print #filenum, " "
            Print #filenum, "    Total: " & String(18, "_")
            Select Case Branch_Type
                Case 1, 2, 3, 4, 6
                    Print #filenum, " "
                    Print #filenum, "  Room No: " & String(18, "_")
                    Print #filenum, " "
                    Print #filenum, " "
                    Print #filenum, " "
                    Print #filenum, "  Signature " & String(30, "-")
            End Select
        Else
            
            If TillData.TableNo <> 9999 Then
            Print #filenum, "Sub Total: " & String(18 - Len(Format(TillData.SaleTotal, "0.00")), " ") & Format(TillData.SaleTotal, "0.00")
            End If
            If TillData.TableNo = 9999 Then
            Print #filenum, "Training Table *DO NOT PAY*"
            End If
            Print #filenum, " "
            Print #filenum, " Gratuity: " & String(18, "_")
            Print #filenum, " "
            Print #filenum, "    Total: " & String(18, "_")
            Select Case Branch_Type
                Case 1, 2, 3, 4, 6
                    Print #filenum, " "
                    Print #filenum, "  Room No: " & String(18, "_")
                    Print #filenum, " "
                    Print #filenum, " "
                    Print #filenum, " "
                    Print #filenum, "  Signature " & String(30, "-")
            End Select
        End If
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    Else
        If TillData.SaleTotal <> 0 Then 'No Sale
            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
            Print #filenum, Chr(27) & Chr(33) & Chr(16);
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
            InsaleTipp = False
            Select Case Panel_no
                Case -1
                    If TillData.Tipp <> 0 Then
                        InsaleTipp = True
                        If Reprint = 0 Then TillData.SaleTotal = TillData.SaleTotal - TillData.Tipp
                    End If
                Case 0
                    If frmSales.grdMain.TextMatrix(frmSales.grdMain.Rows - 2, 3) = "Service Charge" Then
                        InsaleTipp = True
                        If Reprint = 0 Then TillData.SaleTotal = TillData.SaleTotal - TillData.Tipp
                    End If
                Case 1
                    If frmSales1.grdMain.TextMatrix(frmSales1.grdMain.Rows - 2, 3) = "Service Charge" Then
                        InsaleTipp = True
                        If Reprint = 0 Then TillData.SaleTotal = TillData.SaleTotal - TillData.Tipp
                    End If
                Case 2
                    If frmBar.grdMain.TextMatrix(frmBar.grdMain.Rows - 2, 3) = "Service Charge" Then
                        InsaleTipp = True
                        If Reprint = 0 Then TillData.SaleTotal = TillData.SaleTotal - TillData.Tipp
                    End If
            End Select
            'Print #filenum, "    Sub Total:" & String(12 - Len(Format(TillData.SaleTotal - TillData.TaxTotal, "0.00")), " ") & Format(TillData.SaleTotal - TillData.TaxTotal, "0.00")
            Print #filenum, "          Vat:" & String(12 - Len(Format(TillData.TaxTotal, "0.00")), " ") & Format(TillData.TaxTotal, "0.00")
            
            
            
            
            
            If TillData.Change > 0 Then
                Print #filenum, "Total:" & String(12 - Len(Format(TillData.SaleTotal, "0.00")), " ") & Format(TillData.SaleTotal, "0.00")
            Else
                If Keystring = "Reprint" Then
                    Print #filenum, "Tendered:" & String(12 - Len(Format(TillData.SaleTotal, "0.00")), " ") & Format(TillData.SaleTotal, "0.00")
                Else
                    If TillData.ShortTender = True Then
                        Print #filenum, " Tendered:" & String(12 - Len(Format(TillData.SaleTotal + TillData.Change, "0.00")), " ") & Format(TillData.SaleTotal + TillData.Change, "0.00")
                    Else
                        If InsaleTipp = True Then
                            Print #filenum, "Total:" & String(12 - Len(Format(TillData.SaleTotal, "0.00")), " ") & Format(TillData.SaleTotal, "0.00")
                            Print #filenum, "Gratuity:" & String(12 - Len(Format(TillData.Tipp, "0.00")), " ") & Format(TillData.Tipp, "0.00")
                            Print #filenum, Keystring & " Tendered:" & String(12 - Len(Format(TillData.SaleTotal + TillData.Tipp, "0.00")), " ") & Format(TillData.SaleTotal + TillData.Tipp, "0.00")
                        Else
                            If Keystring = "Quote" Then
                                Print #filenum, "Quotation Total:" & String(12 - Len(Format(TillData.SaleTotal, "0.00")), " ") & Format(TillData.SaleTotal, "0.00")
                            Else
                                If Panel_no = -1 Then
                                    ActiveReadServer "Select Line_Total,Function_Key from Sales_Journal where Function_Key in (9,10,11,12,13) and  Invoice_No = " & TillData.DocNo
                                    While Not rs.EOF
                                        Select Case rs.Fields("Function_Key")
                                            Case 9: Print #filenum, "Cash Tendered:" & String(12 - Len(Format(rs.Fields("Line_Total"), "0.00")), " ") & Format(rs.Fields("Line_Total"), "0.00")
                                            Case 10: Print #filenum, "Card Tendered:" & String(12 - Len(Format(rs.Fields("Line_Total"), "0.00")), " ") & Format(rs.Fields("Line_Total"), "0.00")
                                            Case 11: Print #filenum, "Voucher Tendered:" & String(12 - Len(Format(rs.Fields("Line_Total"), "0.00")), " ") & Format(rs.Fields("Line_Total"), "0.00")
                                            Case 12: Print #filenum, "Charge Tendered:" & String(12 - Len(Format(rs.Fields("Line_Total"), "0.00")), " ") & Format(rs.Fields("Line_Total"), "0.00")
                                            Case 13: Print #filenum, "Voucher Tendered:" & String(12 - Len(Format(rs.Fields("Line_Total"), "0.00")), " ") & Format(rs.Fields("Line_Total"), "0.00")
                                        End Select
                                        rs.MoveNext
                                    Wend
                                    If rs.RecordCount > 1 Then
                                        Print #filenum, "Total Tendered:" & String(12 - Len(Format(TillData.SaleTotal, "0.00")), " ") & Format(TillData.SaleTotal, "0.00")
                                    End If
                                    rs.Close
                                Else
                                    Print #filenum, Keystring & " Tendered:" & String(12 - Len(Format(TillData.SaleTotal, "0.00")), " ") & Format(TillData.SaleTotal, "0.00")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If TillData.Change > 0 Then
                If Keystring = "Reprint" Then
                    Print #filenum, "Tendered:" & String(12 - Len(Format(TillData.Tendered, "0.00")), " ") & Format(TillData.Tendered, "0.00")
                Else
                    If TillData.ShortTender = True Then
                        Print #filenum, " Tendered:" & String(12 - Len(Format(TillData.SaleTotal + TillData.Change, "0.00")), " ") & Format(TillData.SaleTotal + TillData.Change, "0.00")
                    Else
                        Print #filenum, Keystring & " Tendered:" & String(12 - Len(Format(TillData.Tendered, "0.00")), " ") & Format(TillData.Tendered, "0.00")
                    End If
                End If
                If TillData.Tipp <> 0 Then
                    If TillData.Tipp = TillData.Change Then
                        Print #filenum, " Gratuity:" & String(12 - Len(Format(TillData.Change, "0.00")), " ") & Format(TillData.Change, "0.00")
                    Else
                        Print #filenum, " Gratuity:" & String(12 - Len(Format(TillData.Tipp, "0.00")), " ") & Format(TillData.Tipp, "0.00")
                        Print #filenum, "   Change:" & String(12 - Len(Format(TillData.Change, "0.00")), " ") & Format(TillData.Change, "0.00")
                    End If
                Else
                    Print #filenum, "   Change:" & String(12 - Len(Format(TillData.Change, "0.00")), " ") & Format(TillData.Change, "0.00")
                End If
            End If
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
            Print #filenum, Chr(27) & Chr(33) & Chr(0);
            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
        Else
            If nosale = False Then
                If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                Print #filenum, Chr(27) & Chr(33) & Chr(16);
                If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                Print #filenum, "    Total:" & String(12 - Len(Format(TillData.SaleTotal, "0.00")), " ") & Format(TillData.SaleTotal, "0.00")
                If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                Print #filenum, Chr(27) & Chr(33) & Chr(0);
                If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
            End If
        End If
    End If
    'Kotie 28-03-2013
    '********* Print Convertion estimate on slip
    If Conversion_Rate <> 0 Then
        If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
        Print #filenum, Chr(27) & Chr(33) & Chr(16);
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
        Print #filenum, Conversion_description & " Total:" & String(12 - Len(Format(TillData.SaleTotal * Conversion_Rate, "0.00")), " ") & Format(TillData.SaleTotal * Conversion_Rate, "0.00")
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    End If
    '*********
    If Keystring <> "Print Bill" Then
        Print #filenum, String(33, "-")
    End If
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
    
    needline = False
    If TillData.ShortTender = True Then
        needline = True
        Select Case Panel_no
            Case 0
                For i = frmSales.grdMain.Rows - 1 To 1 Step -1
                    If frmSales.grdMain.TextMatrix(i, 0) <> "" Then Exit For
                    Print #filenum, frmSales.grdMain.TextMatrix(i, 1) & ":" & String(12 - Len(Format(frmSales.grdMain.TextMatrix(i, 2), "0.00")), " ") & frmSales.grdMain.TextMatrix(i, 2)
                Next i
            Case 2
                For i = frmBar.grdMain.Rows - 1 To 1 Step -1
                    If frmBar.grdMain.TextMatrix(i, 0) <> "" Then Exit For
                    Print #filenum, frmBar.grdMain.TextMatrix(i, 1) & ":" & String(12 - Len(Format(frmBar.grdMain.TextMatrix(i, 2), "0.00")), " ") & frmBar.grdMain.TextMatrix(i, 2)
                Next i
        End Select
        
    End If
    If needline = True Then
        Print #filenum, String(33, "-")
    End If
    If Keystring = "Print Bill" Then
        If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
        Print #filenum, String(33, "_")
        If TillData.TableNo = 0 Then
            Print #filenum, "Barman: " & UserRecord.User_Number; " - " & Trim(UserRecord.Name)
            Print #filenum, "Tab: " & TillData.TabName
        Else
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
            Print #filenum, "Waitron: " & UserRecord.User_Number; " - " & Trim(UserRecord.Name)
            Print #filenum, "Table No: " & TillData.TableNo
        End If
    Else
        If TillData.TabNo <> 0 Then
            If Panel_no = -1 Then
                ActiveReadServer "Select User_No,(Select User_Name from Users where Users.User_No = Sales_Journal.User_No) as User_Name from Sales_Journal where Date_Time is not null and Invoice_No = " & TillData.DocNo
                Print #filenum, "Barman: " & rs.Fields("User_No"); " - " & Trim(rs.Fields("User_Name"))
                rs.Close
            Else
                Print #filenum, "Barman: " & UserRecord.User_Number; " - " & Trim(UserRecord.Name)
            End If
            Print #filenum, "Tab: " & TillData.TabName
        End If
        If TillData.TableNo <> 0 Then
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
            If Panel_no = -1 Then
                ActiveReadServer "Select User_No,(Select User_Name from Users where Users.User_No = Sales_Journal.User_No) as User_Name from Sales_Journal where User_No is not Null and Date_Time is not null and Invoice_No = " & TillData.DocNo
                Print #filenum, "Waitron: " & rs.Fields("User_No"); " - " & Trim(rs.Fields("User_Name"))
                rs.Close
            Else
                Print #filenum, "Waitron: " & UserRecord.User_Number; " - " & Trim(UserRecord.Name)
            End If
            Print #filenum, "Table No: " & TillData.TableNo
            If Keystring = "Charge" Then
                Print #filenum, " "
                Print #filenum, "  Room No: " & String(18, "_")
                Print #filenum, " "
                Print #filenum, " "
                Print #filenum, " "
                Print #filenum, "  Signature " & String(30, "-")
            End If
        End If
        If TillData.TableNo = 0 And TillData.TableNo = 0 Then
            Print #filenum, "Cashier: " & UserRecord.User_Number; " - " & Trim(UserRecord.Name)
        End If
    End If
    If Panel_no = -1 Then
        ActiveReadServer "Select Date_Time from Sales_Journal where Date_Time is not null and Invoice_No = " & TillData.DocNo
        Print #filenum, "Date: " & Format(rs.Fields("Date_Time"), "dd MMMM yyyy DDD") & " " & Format(rs.Fields("Date_Time"), "HH:MM:SS")
        rs.Close
    Else
        Print #filenum, "Date: " & Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
    End If
    If Keystring = "Quote" Then
        Print #filenum, "Quotation No: " & Format(TillData.DocNo, "000000")
    Else
        
        'Makebold
        Print #filenum, "Invoice No: " & Format(TillData.DocNo, "000000")
    End If
    
    'Discount Value
    '__Currentwork____________________________________________________________________________
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If TillData.TableNo = 9999 Then Print #filenum, "TRAINING-DO NOT PAY"
    If TillData.TotDiscountVal + TillData.TotDiscount <> 0 Then
        Print #filenum, "Discount Value: " & Format(TillData.TotDiscountVal + TillData.TotDiscount, "0.00")
        Thediscounttotal = TillData.TotDiscount
        Reprintdiscount = Reprintdiscount + 1
    End If
    
    If Reprintdiscount > 1 Then
    If Thediscounttotal <> 0 Then
    Print #filenum, "Discount Value: " & Format(Thediscounttotal, "0.00")
        Reprintdiscount = Reprintdiscount + 1

    End If
    End If
    Reprintdiscount = Reprintdiscount + 1
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    '______________________________________________________________________________
    If TillData.Room_No <> 0 Then
        Print #filenum, "Charged to Room: " & Format(TillData.Room_No, "000000")
    End If
    If TillData.Account_No <> "" Then
        ActiveReadServer "Select * from Debtors where Debtor_No = '" & TillData.Account_No & "'"
        If rs.RecordCount > 0 Then
            Print #filenum, "Charged to: " & Format(TillData.Account_No, "000000") & " - " & rs.Fields("Debtor_Name")
            Print #filenum, "Old Balance: " & Format(rs.Fields("Balance") - TillData.Charge, "0.00")
            Print #filenum, "  This Sale: " & Format(TillData.Charge, "0.00")
            Print #filenum, "New Balance: " & Format(rs.Fields("Balance"), "0.00")
        End If
        rs.Close
    End If
    For i = 1 To .grdFoot.Rows - 1
        If Trim(.grdFoot.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdFoot.TextMatrix(i, 1))
                Case "Line Feeds"
                    Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdFoot.TextMatrix(i, 0)));
                Case Else
                    Select Case .grdFoot.TextMatrix(i, 2)
                        Case "Left": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdFoot.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, Chr(29) & "V" & Chr(49);
    
    send_data_steam_keylog (filenum)
    Close #filenum
    'temp----
   ' Open "c:\print.prn" For Input As #filenum
   ' filenum1 = FreeFile
   ' Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum1
   ' Print #filenum1, filenum
   ' Close #filenum1
    '---------
End With
If Keystring = "Quote" Then
    If TillData.Room_No = 0 And Reprint = 0 Then
        Slip_Port = ""
        PrintErr = 0
        Reprint = 1
        GoTo Reprint
    End If
End If
If Keystring = "Charge" Then
    If TillData.Room_No = 0 And Reprint = 0 Then
        Slip_Port = ""
        PrintErr = 0
        Reprint = 1
        If ChargeSlip = 1 Then GoTo Reprint
    End If
End If

On Error GoTo 0
ImPrinting = False


Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        Dim x As Printer
        For Each x In Printers
            If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = x.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = err.Description
    send_data_steam_keylog ("Print error -" & err.Description)
    DoEvents
    frmError.Show vbModal
    Close #filenum
    ImPrinting = False
    On Error GoTo 0
End Sub
Private Sub GetDocNumbers(Keystring$)
    
'    If TillData.DocNo = 0 Then
'        ActiveReadServer "Select (Select isnull(Max(convert(int,Trans_No)),0)+1 from Sales_Journal) as Trans_No,(Select isnull(Max(convert(int,Invoice_No)),0)+1 from Sales_Journal) as Invoice_No"
'            If rs.RecordCount > 0 Then
'            TillData.DocNo = rs.Fields("Invoice_No")
'            TillData.TransNo = rs.Fields("Trans_No")
'            End If
'        rs.Close
'        DoEvents
'        ActiveUpdateServer "Insert into Sales_Journal (Date_Time,Invoice_No,Trans_No,Table_No,Tab_No,Function_Key,User_Overide)  values (Getdate()," & TillData.DocNo & "," & TillData.TransNo & ",'" & TillData.TableNo & "','" & TillData.TabNo & "',14," & UserRecord.User_Number & ")"
'    End If


Redo:
    If TillData.DocNo > 0 Then GoTo out
    If TillData.DocNo = 0 Then
    ActiveReadServer "Select (Select isnull(Max(convert(int,Trans_No)),0)+1 from Sales_Journal) as Trans_No,(Select isnull(Max(convert(int,Invoice_No)),0)+1 from Sales_Journal) as Invoice_No"
 
    If rs.RecordCount > 0 Then
            
            If Val(rs.Fields("Invoice_No")) <> 0 Then
            If Val(rs.Fields("Trans_No")) <> 0 Then
            
            TillData.DocNo = rs.Fields("Invoice_No")
            TillData.TransNo = rs.Fields("Trans_No")
            
            Else
            
    MsgBox " Cannot connect to SQLSERVER!, you may have a possible Network/Server error, will retry as soon as you click ok! ", vbDefaultButton1
    filenum = FreeFile
    Close filenum
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    If LOF(filenum) > 144000 Then
        Close filenum
        Kill Trim(App.Path) & "\Logs\ServerErrors.log"
        
    End If
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, "Connection error to get transaction no or invoice no"
    Print #filenum, "Select (Select isnull(Max(convert(int,Trans_No)),0)+1 from Sales_Journal) as Trans_No,(Select isnull(Max(convert(int,Invoice_No)),0)+1 from Sales_Journal) as Invoice_No"
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
    'MsgBox Err.Description, vbInformation
   ' Resume Next
            GoTo Redo
            End If
            
            Else
            
    MsgBox " Cannot connect to SQLSERVER!, you may have a possible Network/Server error, will retry as soon as you click ok! ", vbDefaultButton1
            filenum = FreeFile
    Close filenum
    Open Trim(App.Path) & "\Logs\ServerErrors.log" For Append As filenum
    If LOF(filenum) > 144000 Then
        Close filenum
        Kill Trim(App.Path) & "\Logs\ServerErrors.log"
        
    End If
    Print #filenum, "User Logged On"
    Print #filenum, "User Number: " & UserRecord.User_Number
    Print #filenum, "User Name: " & UserRecord.Name
    Print #filenum, "Date & Time: " & Format(Date, "YYYY-MM-DD (DDD)") & " " & Time
    Print #filenum, "Connection error to get transaction no or invoice no"
    Print #filenum, "Select (Select isnull(Max(convert(int,Trans_No)),0)+1 from Sales_Journal) as Trans_No,(Select isnull(Max(convert(int,Invoice_No)),0)+1 from Sales_Journal) as Invoice_No"
    Print #filenum, "*********************************************************"
    ' Close before reopening in another mode.
    Close filenum
    'MsgBox Err.Description, vbInformation
   ' Resume Next
            GoTo Redo
            End If
            
            End If

End If

        
        DoEvents
        
        If Val(TillData.DocNo) <> 0 Then
        If Val(TillData.TransNo) <> 0 Then
        ActiveUpdateServer "Insert into Sales_Journal (Date_Time,Invoice_No,Trans_No,Table_No,Tab_No,Function_Key,User_Overide)  values (Getdate()," & TillData.DocNo & "," & TillData.TransNo & ",'" & TillData.TableNo & "','" & TillData.TabNo & "',14," & UserRecord.User_Number & ")"
End If
End If

out:
End Sub
Private Sub WriteJournals(Keystring$)

If TillData.TableNo = 9999 Then Exit Sub
GetDocNumbers Keystring
    On Error Resume Next
    SaveLocation = Location_No
    If AskLog = 1 Then Location_No = Val(Mid(frmAskLoc.Tag, 1, InStr(frmAskLoc.Tag, "-") - 1))
    Select Case Keystring
        Case "No Sale"
            ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No) values " & _
            "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & ",6," & Location_No & "," & Branch_No & "," & TillData.Cashup_No & ")"
        Case "Cash", "Voucher", "Card", "Charge", "Loyalty"
            Select Case Panel_no
                Case 0
                    With frmSales
                        For i = 1 To .grdMain.Rows - 1
                            Select Case .grdMain.TextMatrix(i, 3)
                                Case "Plu"
                                    Function_Key = 7
                                Case "Dept"
                                    Function_Key = 8
                                Case "Cash"
                                    Function_Key = 9
                                Case "Card"
                                    Function_Key = 10
                                Case "Voucher"
                                    Function_Key = 11
                                Case "Charge"
                                    Function_Key = 12
                                Case "Loyalty"
                                    Function_Key = 13
                                Case "Service Charge"
                                    Function_Key = 16
                                Case "Short"
                                    Select Case .grdMain.TextMatrix(i, 1)
                                        Case "Cash Tendered"
                                            Function_Key = 9
                                        Case "Card Tendered"
                                            Function_Key = 10
                                        Case "Charge Tendered"
                                            Function_Key = 12
                                        Case "Voucher Tendered"
                                            Function_Key = 11
                                        Case "Loyalty Tendered"
                                            Function_Key = 13
                                    End Select
                            End Select
                            If .grdMain.TextMatrix(i, 3) = "Short" And i = .grdMain.Rows - 1 Then
                                linetotal = .grdMain.TextMatrix(i, 2) - TillData.Change
                            Else
                                linetotal = .grdMain.TextMatrix(i, 2)
                                If .grdMain.TextMatrix(i - 1, 3) = "Service Charge" Then
                                    linetotal = Val(.grdMain.TextMatrix(i, 2)) - TillData.Tipp
                                End If
                            End If
                            
                            ParentProd = ""
                            If Mid(.grdMain.TextMatrix(i, 1), 5, 1) = ">" Then
                                .grdMain.TextMatrix(i, 8) = Mid(.grdMain.TextMatrix(i, 1), 6)
                                For b = i To 1 Step -1
                                    If .grdMain.TextMatrix(b, 0) <> "" Then
                                        ParentProd = .grdMain.TextMatrix(b, 9)
                                        Exit For
                                    End If
                                Next b
                            End If
                            Qty = Val(.grdMain.TextMatrix(i, 0))
                            If Qty = 0 Then Qty = 1
                            
                            'NewLocation = Location_No
                            
                            If NewLocation = 0 Then NewLocation = Location_No
                            
                            If NewLocation = 0 Then
                            
                            ActiveReadServer1 "Select * from Printer_Links where Printer = '" & Trim(.grdMain.TextMatrix(i, 11)) & "'"
                            If rs1.RecordCount > 0 Then
                                If Val(rs1.Fields("Sales_Location_No") & "") <> 0 Then
                                    NewLocation = Val(rs1.Fields("Sales_Location_No") & "")
                                Else
                                    NewLocation = Location_No
                                End If
                            End If
                            rs1.Close
                            If NewLocation = 0 Then NewLocation = Location_No
                            End If
                            
                            
                            SaveUser = 0
                            If TillData.TableNo <> 0 Then
                                ActiveReadServer "Select * from Table_Listing_View where Table_No = " & TillData.TableNo
                                If rs.RecordCount > 0 Then
                                    If UserRecord.User_Number <> rs.Fields("User_No") Then
                                        If UserRecord.uType = 3 Or UserRecord.uType = 4 Then
                                            If UserRecord.Bar_Cash = 1 Then
                                                SaveUser = UserRecord.User_Number
                                                UserRecord.User_Number = rs.Fields("User_No")
                                            End If
                                        End If
                                    End If
                                End If
                                rs.Close
                            End If
                            
                            If TillData.TableNo <> 9999 Then
                            
                            '*********** counters panel1
                            If Function_Key = 7 Then
                            ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No,Product_Code,Department_No,Qty,Ave_Cost,Sales_Tax,Tax_Type,Line_Total,Extra,Table_No,Tab_No,Covers,User_Overide,Account_No,Room_No,Res_No,Discount_Amt,Dicount_Value, Conversion_Rate) values " & _
                            "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & "," & Val(Function_Key) & "," & NewLocation & "," & Branch_No & "," & _
                            TillData.Cashup_No & ",'" & .grdMain.TextMatrix(i, 9) & "','" & .grdMain.TextMatrix(i, 10) & "','" & Qty & "','" & .grdMain.TextMatrix(i, 4) & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "'," & Val(linetotal) & ",'" & .grdMain.TextMatrix(i, 8) & "','" & TillData.TableNo & "','" & TillData.TabNo & "','" & TillData.Covers & "','" & .grdMain.ValueMatrix(i, 17) & "','" & Trim(TillData.Account_No) & "'," & TillData.Room_No & "," & TillData.Res_No & "," & .grdMain.ValueMatrix(i, 18) & "," & TillData.TotDiscount & "," & Conversion_Rate & ")"
                            
                            Else
                             If Val(Swiss_Round) <> 0 Then
                                If Str(linetotal / Swiss_Round) <> Str(Round(linetotal / Swiss_Round, 0)) Then
                                   q = Round((linetotal / Swiss_Round - Int(linetotal / Swiss_Round)) * Swiss_Round, 2)
                                    linetotal = linetotal - q
                                End If
                            End If
                              'Tax_Total = TillData.TaxTotal * (linetotal - TillData.Change) / TillData.SaleTotal
                            Tax_Total = TillData.TaxTotal
                            ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No,Product_Code,Department_No,Qty,Ave_Cost,Sales_Tax,Tax_Type,Line_Total,Extra,Table_No,Tab_No,Covers,User_Overide,Account_No,Room_No,Res_No,Discount_Amt,Dicount_Value, Conversion_Rate) values " & _
                              "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & "," & Val(Function_Key) & "," & NewLocation & "," & Branch_No & "," & _
                              TillData.Cashup_No & ",'" & .grdMain.TextMatrix(i, 9) & "','" & .grdMain.TextMatrix(i, 10) & "','" & Qty & "','" & .grdMain.TextMatrix(i, 4) & "','" & Tax_Total & "','" & .grdMain.TextMatrix(i, 6) & "'," & Val(linetotal) & ",'" & .grdMain.TextMatrix(i, 8) & "','" & TillData.TableNo & "','" & TillData.TabNo & "','" & TillData.Covers & "','" & .grdMain.ValueMatrix(i, 17) & "','" & Trim(TillData.Account_No) & "'," & TillData.Room_No & "," & TillData.Res_No & "," & .grdMain.ValueMatrix(i, 18) & "," & TillData.TotDiscount & "," & Conversion_Rate & ")"
                            
                             
                             End If
                            '*********************************
                            If SaveUser <> 0 Then UserRecord.User_Number = SaveUser
                            DoEvents
                            If Function_Key = 7 Then
                                If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                                    If ParentProd <> "" And Mid(.grdMain.TextMatrix(i, 1), 5, 1) = ">" Then
                                        ActiveReadServer "Select Line_Type,Unit_of_Measure as Recipe_Unit, Line_Code,Qty_Used,(Select Unit_Size from Products where Products.Product_code = Recipes.Line_Code) as Unit_Size, (Select Unit_of_Measure from Products where Products.Product_code = Recipes.Line_Code) as Unit_of_Measure,(Select Ave_Cost from Products where Products.Product_code = Recipes.Line_Code) as Ave_Cost from Recipes where Product_Code= '" & ParentProd & "' and line_Code='" & .grdMain.TextMatrix(i, 9) & "'"
                                        If rs.RecordCount > 0 Then
                                            Select Case rs.Fields("Qty_Used")
                                                Case "1 x 25ml"
                                                    Qty = Round(25 / rs.Fields("Unit_Size"), 4)
                                                Case "2 x 25ml"
                                                    Qty = Round(50 / rs.Fields("Unit_Size"), 4)
                                                Case Else
                                                    UnitSize = rs.Fields("Unit_Size")
                                                    If rs.Fields("Unit_of_Measure") <> rs.Fields("Recipe_Unit") Then
                                                        Select Case UCase(rs.Fields("Unit_of_Measure") & " to " & rs.Fields("Recipe_Unit"))
                                                            Case "ML TO LT"
                                                                UnitSize = rs.Fields("Unit_Size") / 1000
                                                            Case "LT TO ML"
                                                                UnitSize = rs.Fields("Unit_Size") * 1000
                                                            Case "G TO KG"
                                                                UnitSize = rs.Fields("Unit_Size") / 1000
                                                            Case "KG TO G"
                                                                UnitSize = rs.Fields("Unit_Size") * 1000
                                                            Case Else
                                                                UnitSize = rs.Fields("Unit_Size")
                                                        End Select
                                                    End If
                                                    If Val(UnitSize & "") <> 0 Then
                                                        Qty = Round(rs.Fields("Qty_Used") / UnitSize, 4)
                                                    Else
                                                        Qty = rs.Fields("Qty_Used")
                                                    End If
                                            End Select
                                        End If
                                        rs.Close
                                        UpdateQuantities .grdMain.TextMatrix(i, 9), Qty, .grdMain.TextMatrix(i, 11), TillData.DocNo, .grdMain.TextMatrix(i, 4)
                                    Else
                                        UpdateQuantities .grdMain.TextMatrix(i, 9), .grdMain.TextMatrix(i, 0), .grdMain.TextMatrix(i, 11), TillData.DocNo, .grdMain.TextMatrix(i, 4)
                                    End If
                                End If
                            End If
                            End If
                            If Function_Key = 10 Then
                                If TillData.ShortTender = True Then
                                    ActiveUpdateServer "Update Counters set " & _
                                    "CardsinDrawer_Value=isnull(CardsinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 2) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                Else
                                    ActiveUpdateServer "Update Counters set " & _
                                    "CardsinDrawer_Value=isnull(CardsinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 4) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                End If
                            End If
                            If Function_Key = 11 Then
                                If TillData.ShortTender = True Then
                                    ActiveUpdateServer "Update Counters set " & _
                                    "ChequeinDrawer_Value=isnull(ChequeinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 2) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                Else
                                    ActiveUpdateServer "Update Counters set " & _
                                    "ChequeinDrawer_Value=isnull(ChequeinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 4) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                End If
                            End If
                            If Function_Key = 12 Then
                                If TillData.ShortTender = True Then
                                    ActiveUpdateServer "Update Counters set " & _
                                    "ChequeinDrawer_Value=isnull(ChequeinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 2) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                Else
                                    ActiveUpdateServer "Update Counters set " & _
                                    "ChargeinDrawer_Value=isnull(ChargeinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 4) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                End If
                            End If
                        Next i
                    End With
                Case 2
                    With frmBar
                        For i = 1 To .grdMain.Rows - 1
                            Select Case .grdMain.TextMatrix(i, 3)
                                Case "Plu"
                                    Function_Key = 7
                                Case "Dept"
                                    Function_Key = 8
                                Case "Cash"
                                    Function_Key = 9
                                Case "Card"
                                    Function_Key = 10
                                Case "Voucher"
                                    Function_Key = 11
                                Case "Charge"
                                    Function_Key = 12
                                Case "Loyalty"
                                    Function_Key = 13
                                Case "Short"
                                    Select Case .grdMain.TextMatrix(i, 1)
                                        Case "Cash Tendered"
                                            Function_Key = 9
                                        Case "Card Tendered"
                                            Function_Key = 10
                                        Case "Charge Tendered"
                                            Function_Key = 12
                                        Case "Voucher Tendered"
                                            Function_Key = 11
                                        Case "Loyalty Tendered"
                                            Function_Key = 13
                                    End Select
                            End Select
                            If .grdMain.TextMatrix(i, 3) = "Short" And i = .grdMain.Rows - 1 Then
                                linetotal = .grdMain.TextMatrix(i, 2) - TillData.Change
                            Else
                                If .grdMain.TextMatrix(i, 3) = "Short" Then .grdMain.TextMatrix(i, 8) = "Short"
                                linetotal = .grdMain.TextMatrix(i, 2)
                            End If
                            If Mid(.grdMain.TextMatrix(i, 1), 5, 1) = ">" Then
                                .grdMain.TextMatrix(i, 8) = Mid(.grdMain.TextMatrix(i, 1), 6)
                                For b = i To 1 Step -1
                                    If .grdMain.TextMatrix(b, 0) <> "" Then
                                        ParentProd = .grdMain.TextMatrix(b, 9)
                                        Exit For
                                    End If
                                Next b
                            End If
                            Qty = Val(.grdMain.TextMatrix(i, 0))
                            If Qty = 0 Then Qty = 1
                            
                            NewLocation = Location_No
                            ActiveReadServer1 "Select * from Printer_Links where Printer = '" & Trim(.grdMain.TextMatrix(i, 11)) & "'"
                            If rs1.RecordCount > 0 Then
                                If Val(rs1.Fields("Sales_Location_No") & "") <> 0 Then
                                    NewLocation = Val(rs1.Fields("Sales_Location_No") & "")
                                Else
                                    NewLocation = Location_No
                                End If
                            End If
                            rs1.Close
                            If NewLocation = 0 Then NewLocation = Location_No
                            SaveUser = 0
                            If TillData.TableNo <> 0 Then
                                ActiveReadServer "Select * from Table_Listing_View where Table_No = " & TillData.TableNo
                                If rs.RecordCount > 0 Then
                                    If UserRecord.User_Number <> rs.Fields("User_No") Then
                                        If UserRecord.uType = 3 Or UserRecord.uType = 4 Then
                                            If UserRecord.Bar_Cash = 1 Then
                                                SaveUser = UserRecord.User_Number
                                                UserRecord.User_Number = rs.Fields("User_No")
                                            End If
                                        End If
                                    End If
                                End If
                                rs.Close
                            End If
                            
                            
                            
'                            ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No,Product_Code,Department_No,Qty,Ave_Cost,Sales_Tax,Tax_Type,Line_Total,Extra,Table_No,Tab_No,Covers, Account_No,Room_No,Res_No,Discount_Amt,Dicount_Value) values " & _
'                            "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & "," & Val(Function_key) & "," & NewLocation & "," & Branch_No & "," & _
'                            TillData.Cashup_No & ",'" & .grdMain.TextMatrix(i, 9) & "','" & .grdMain.TextMatrix(i, 10) & "','" & Qty & "','" & .grdMain.TextMatrix(i, 4) & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "'," & Val(linetotal) & ",'" & .grdMain.TextMatrix(i, 8) & "','" & TillData.TableNo & "','" & TillData.TabNo & "','" & TillData.Covers & "','" & Trim(TillData.Account_No) & "'," & TillData.Room_No & "," & TillData.Res_No & "," & .grdMain.ValueMatrix(i, 18) & "," & .grdMain.ValueMatrix(i, 19) & ")"
                             
                             '*********** counters panel2
                            If Function_Key = 7 Then
                            ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No,Product_Code,Department_No,Qty,Ave_Cost,Sales_Tax,Tax_Type,Line_Total,Extra,Table_No,Tab_No,Covers,User_Overide,Account_No,Room_No,Res_No,Discount_Amt,Dicount_Value, Conversion_Rate) values " & _
                            "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & "," & Val(Function_Key) & "," & NewLocation & "," & Branch_No & "," & _
                            TillData.Cashup_No & ",'" & .grdMain.TextMatrix(i, 9) & "','" & .grdMain.TextMatrix(i, 10) & "','" & Qty & "','" & .grdMain.TextMatrix(i, 4) & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "'," & Val(linetotal) & ",'" & .grdMain.TextMatrix(i, 8) & "','" & TillData.TableNo & "','" & TillData.TabNo & "','" & TillData.Covers & "','" & .grdMain.ValueMatrix(i, 17) & "','" & Trim(TillData.Account_No) & "'," & TillData.Room_No & "," & TillData.Res_No & "," & .grdMain.ValueMatrix(i, 18) & "," & TillData.TotDiscount & "," & Conversion_Rate & ")"
                            
                            Else
                             If Val(Swiss_Round) <> 0 Then
                                If Str(linetotal / Swiss_Round) <> Str(Round(linetotal / Swiss_Round, 0)) Then
                                   q = Round((linetotal / Swiss_Round - Int(linetotal / Swiss_Round)) * Swiss_Round, 2)
                                    linetotal = linetotal - q
                                End If
                                End If
                                
                             ' Tax_Total = TillData.TaxTotal * (linetotal - TillData.Change) / TillData.SaleTotal
                              Tax_Total = TillData.TaxTotal
                              ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No,Product_Code,Department_No,Qty,Ave_Cost,Sales_Tax,Tax_Type,Line_Total,Extra,Table_No,Tab_No,Covers,User_Overide,Account_No,Room_No,Res_No,Discount_Amt,Dicount_Value, Conversion_Rate) values " & _
                            "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & "," & Val(Function_Key) & "," & NewLocation & "," & Branch_No & "," & _
                            TillData.Cashup_No & ",'" & .grdMain.TextMatrix(i, 9) & "','" & .grdMain.TextMatrix(i, 10) & "','" & Qty & "','" & .grdMain.TextMatrix(i, 4) & "','" & Tax_Total & "','" & .grdMain.TextMatrix(i, 6) & "'," & Val(linetotal) & ",'" & .grdMain.TextMatrix(i, 8) & "','" & TillData.TableNo & "','" & TillData.TabNo & "','" & TillData.Covers & "','" & .grdMain.ValueMatrix(i, 17) & "','" & Trim(TillData.Account_No) & "'," & TillData.Room_No & "," & TillData.Res_No & "," & .grdMain.ValueMatrix(i, 18) & "," & TillData.TotDiscount & "," & Conversion_Rate & ")"
                            
                             
                             End If
                            '*********************************
                            
                            
                            
                            If SaveUser <> 0 Then UserRecord.User_Number = SaveUser
                            DoEvents
                            If Function_Key = 7 Then
                                If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                                    If ParentProd <> "" And Mid(.grdMain.TextMatrix(i, 1), 5, 1) = ">" Then
                                        ActiveReadServer "Select Line_Type,Unit_of_Measure as Recipe_Unit, Line_Code,Qty_Used,(Select Unit_Size from Products where Products.Product_code = Recipes.Line_Code) as Unit_Size, (Select Unit_of_Measure from Products where Products.Product_code = Recipes.Line_Code) as Unit_of_Measure,(Select Ave_Cost from Products where Products.Product_code = Recipes.Line_Code) as Ave_Cost from Recipes where Product_Code= '" & ParentProd & "' and line_Code='" & .grdMain.TextMatrix(i, 9) & "'"
                                        If rs.RecordCount > 0 Then
                                            Select Case rs.Fields("Qty_Used")
                                                Case "1 x 25ml"
                                                    Qty = Round(25 / rs.Fields("Unit_Size"), 4)
                                                Case "2 x 25ml"
                                                    Qty = Round(50 / rs.Fields("Unit_Size"), 4)
                                                Case Else
                                                    UnitSize = rs.Fields("Unit_Size")
                                                    If rs.Fields("Unit_of_Measure") <> rs.Fields("Recipe_Unit") Then
                                                        Select Case UCase(rs.Fields("Unit_of_Measure") & " to " & rs.Fields("Recipe_Unit"))
                                                            Case "ML TO LT"
                                                                UnitSize = rs.Fields("Unit_Size") / 1000
                                                            Case "LT TO ML"
                                                                UnitSize = rs.Fields("Unit_Size") * 1000
                                                            Case "G TO KG"
                                                                UnitSize = rs.Fields("Unit_Size") / 1000
                                                            Case "KG TO G"
                                                                UnitSize = rs.Fields("Unit_Size") * 1000
                                                            Case Else
                                                                UnitSize = rs.Fields("Unit_Size")
                                                        End Select
                                                    End If
                                                    If Val(UnitSize & "") <> 0 Then
                                                        Qty = Round(rs.Fields("Qty_Used") / UnitSize, 4)
                                                    Else
                                                        Qty = rs.Fields("Qty_Used")
                                                    End If
                                            End Select
                                        End If
                                        rs.Close
                                        UpdateQuantities .grdMain.TextMatrix(i, 9), Qty, .grdMain.TextMatrix(i, 11), TillData.DocNo, .grdMain.TextMatrix(i, 4)
                                    Else
                                        UpdateQuantities .grdMain.TextMatrix(i, 9), .grdMain.TextMatrix(i, 0), .grdMain.TextMatrix(i, 11), TillData.DocNo, .grdMain.TextMatrix(i, 4)
                                    End If
                                End If
                            End If
                            If Function_Key = 10 Then
                                If TillData.ShortTender = True Then
                                    ActiveUpdateServer "Update Counters set " & _
                                    "CardsinDrawer_Value=isnull(CardsinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 2) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                Else
                                    ActiveUpdateServer "Update Counters set " & _
                                    "CardsinDrawer_Value=isnull(CardsinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 4) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                End If
                            End If
                            If Function_Key = 11 Then
                                If TillData.ShortTender = True Then
                                    ActiveUpdateServer "Update Counters set " & _
                                    "ChequeinDrawer_Value=isnull(ChequeinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 2) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                Else
                                    ActiveUpdateServer "Update Counters set " & _
                                    "ChequeinDrawer_Value=isnull(ChequeinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 4) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                End If
                            End If
                            If Function_Key = 12 Then
                                If TillData.ShortTender = True Then
                                    ActiveUpdateServer "Update Counters set " & _
                                    "ChargeinDrawer_Value=isnull(ChargeinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 2) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                Else
                                    ActiveUpdateServer "Update Counters set " & _
                                    "ChargeinDrawer_Value=isnull(ChargeinDrawer_Value,0) + " & .grdMain.ValueMatrix(i, 4) & _
                                    " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                                End If
                            End If
                        Next i
                    End With
            End Select
    End Select
    If TillData.Account_No <> "" Then
        ActiveUpdateServer "Update Debtors set Balance=Balance + " & TillData.Charge & " where Debtor_No = '" & Trim(TillData.Account_No) & "'"
        ActiveReadServer "Select isnull(Balance,0) as Balance from Debtors where Debtor_No = '" & TillData.Account_No & "'"
        If rs.RecordCount > 0 Then
            NewBalance = NewBalance + rs.Fields("Balance")
        Else
            NewBalance = 0
        End If
        rs.Close
        ActiveUpdateServer "INSERT INTO [Debtor_Accounts]([Transaction_Type],[Date_Time], [Invoice_No], [Account_No], [Debit], [Credit], [Balance])" & _
        "VALUES('Invoice',Getdate()," & TillData.DocNo & ",'" & TillData.Account_No & "'," & TillData.Charge & ",0," & NewBalance & ")"
    End If
    If TillData.Room_No <> 0 Then
        Balance = 0
        ActiveReadServer "Select Balance from Room_Accounts where Res_No = " & TillData.Res_No & " order by Line_No"
        If rs.RecordCount > 0 Then
            rs.MoveLast
            Balance = rs.Fields("Balance")
        End If
        rs.Close
        
        ActiveUpdateServer "INSERT INTO [Room_Accounts]([Transaction_Type],[Date_Time], [Invoice_No], [Account_No], [Debit], [Credit], [Balance],[Res_No])" & _
        "VALUES('Invoice',Getdate()," & TillData.DocNo & ",'" & TillData.Account_No & "'," & TillData.SaleTotal & ",0," & Balance + TillData.SaleTotal & "," & TillData.Res_No & ")"
    End If
    If AskLog = 1 Then Location_No = SaveLocation
    On Error GoTo 0
End Sub
Private Sub UpdateQuantities(Product_Code, Qty, Kitchen_Printer, Invoice_No, Ave_Cost)
    On Error GoTo trap
   
    If Product_Code = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    If Val(Qty) = 0 Then
        Qty = 1
    End If
    NewLocation = Location_No
    If Product_Code <> "" Then
        ActiveReadServer "Select Recipe_Item from Products where Recipe_Item = 1 and Product_Code = '" & Product_Code & "'"
        Select Case rs.RecordCount
            Case 0 'No Recipe
                Select Case Trim(Kitchen_Printer & "")
                    Case "<None>", ""
                        ActiveReadServer2 "Select Link_Code,(Select Pack_Size from Products where Pack_Links.Product_Code=Products.Product_Code) as Pack_Size from Pack_Links where Product_Code = '" & Product_Code & "'"
                        If rs2.RecordCount > 0 Then
                            
                            ActiveReadServer1 "Select * from Quantities where Product_Code = '" & rs2.Fields("Link_Code") & "' and Location_No = " & NewLocation
                    
                            If rs1.RecordCount > 0 Then
                                
                                ActiveUpdateServer "Update Quantities set Stock_on_Hand = Stock_on_Hand - " & Qty * Val(rs2.Fields("Pack_Size") & "") & " where Product_Code = '" & rs2.Fields("Link_Code") & "' and Location_No = " & NewLocation
                        
                            Else
                                
                                ActiveUpdateServer "Insert into Quantities (Product_Code,Stock_on_Hand,Location_No) values ('" & rs2.Fields("Link_Code") & "'," & Qty * Val(rs2.Fields("Pack_Size") & "") * -1 & "," & NewLocation & ")"
                            
                            End If
                            rs1.Close
                            rs2.Close
                        Else
                            rs2.Close
                            ActiveReadServer1 "Select * from Quantities where Product_Code = '" & Product_Code & "' and Location_No = " & NewLocation
                            If rs1.RecordCount > 0 Then
                                
                                ActiveUpdateServer "Update Quantities set Stock_on_Hand = Stock_on_Hand - " & Qty & " where Product_Code = '" & Product_Code & "' and Location_No = " & NewLocation
                               
                            Else
                                
                                ActiveUpdateServer "Insert into Quantities (Product_Code,Stock_on_Hand,Location_No) values ('" & Product_Code & "'," & Qty * -1 & "," & NewLocation & ")"
                             
                            End If
                            rs1.Close
                        End If
                    Case Else 'Location from Kitchen Printer
                        ActiveReadServer1 "Select * from Printer_Links where Printer = '" & Trim(Kitchen_Printer) & "'"
                        If rs1.RecordCount > 0 Then
                            If Val(rs1.Fields("Location_No") & "") <> 0 Then
                                NewLocation = Val(rs1.Fields("Location_No") & "")
                            Else
                                NewLocation = Location_No
                            End If
                        End If
                        rs1.Close
                        
                        ActiveReadServer1 "Select * from Quantities where Product_Code = '" & Product_Code & "' and Location_No = " & NewLocation
                        If rs1.RecordCount > 0 Then
                     
                            ActiveUpdateServer "Update Quantities set Stock_on_Hand = Stock_on_Hand - " & Qty & " where Product_Code = '" & Product_Code & "' and Location_No = " & NewLocation
                           
                        Else
                           
                            ActiveUpdateServer "Insert into Quantities (Product_Code,Stock_on_Hand,Location_No) values ('" & Product_Code & "'," & Qty * -1 & "," & NewLocation & ")"
                       
                        End If
                        rs1.Close
                End Select
               
                ActiveUpdateServer "Insert into Consumption_Journal (Product_Code,Location_No,Ave_Cost,Qty_Consumed,Date_Time,Invoice_No) values ('" & Product_Code & "'," & NewLocation & "," & Ave_Cost * Qty & "," & Qty & ",getdate()," & Invoice_No & ")"
              
            Case 1 'Using a Recipe
                Select Case Kitchen_Printer
                    Case "<None>", ""
                        SaveQty = Qty
                        ActiveReadServer1 "Select Line_Type,Unit_of_Measure as Recipe_Unit, Line_Code,Qty_Used,(Select Unit_Size from Products where Products.Product_code = Recipes.Line_Code) as Unit_Size, (Select Unit_of_Measure from Products where Products.Product_code = Recipes.Line_Code) as Unit_of_Measure,(Select Ave_Cost from Products where Products.Product_code = Recipes.Line_Code) as Ave_Cost from Recipes where Product_Code = '" & Product_Code & "' and Line_Type in (3,4)"
                        While Not rs1.EOF
                            Qty = SaveQty
                            Select Case rs1.Fields("Qty_Used")
                                Case "1 x 25ml"
                                    Qty = Round(25 / rs1.Fields("Unit_Size"), 4) * Qty
                                Case "2 x 25ml"
                                    Qty = Round(50 / rs1.Fields("Unit_Size"), 4) * Qty
                                Case Else
                                    UnitSize = rs1.Fields("Unit_Size")
                                    If rs1.Fields("Unit_of_Measure") <> rs1.Fields("Recipe_Unit") Then
                                        Select Case UCase(rs1.Fields("Unit_of_Measure") & " to " & rs1.Fields("Recipe_Unit"))
                                            Case "ML TO LT"
                                                UnitSize = rs1.Fields("Unit_Size") / 1000
                                            Case "LT TO ML"
                                                If rs1.Fields("Unit_Size") = 0 Then
                                                    UnitSize = 1000
                                                Else
                                                    UnitSize = rs1.Fields("Unit_Size") * 1000
                                                End If
                                            Case "G TO KG"
                                                UnitSize = rs1.Fields("Unit_Size") / 1000
                                            Case "KG TO G"
                                                If rs1.Fields("Unit_Size") = 0 Then
                                                    UnitSize = 1000
                                                Else
                                                    UnitSize = rs1.Fields("Unit_Size") * 1000
                                                End If
                                            Case Else
                                                UnitSize = rs1.Fields("Unit_Size")
                                        End Select
                                    End If
                                    If Val(UnitSize & "") <> 0 Then
                                        Qty = Round(rs1.Fields("Qty_Used") / UnitSize, 4) * Qty
                                    Else
                                        Qty = rs1.Fields("Qty_Used") * Qty
                                    End If
                            End Select
                            Ave_Cost = rs1.Fields("Ave_Cost") * Qty
                            ActiveReadServer2 "Select * from Quantities where Product_Code = '" & rs1.Fields("Line_Code") & "' and Location_No = " & NewLocation
                            If rs2.RecordCount > 0 Then
                                
                                ActiveUpdateServer "Update Quantities set Stock_on_Hand = Stock_on_Hand - " & Qty & " where Product_Code = '" & rs1.Fields("Line_Code") & "' and Location_No = " & NewLocation
                               
                            Else
                                
                                ActiveUpdateServer "Insert into Quantities (Product_Code,Stock_on_Hand,Location_No) values ('" & rs1.Fields("Line_Code") & "'," & Qty * -1 & "," & Location_No & ")"
                              
                            End If
                        
                            ActiveUpdateServer "Insert into Consumption_Journal (Product_Code,Location_No,Ave_Cost,Qty_Consumed,Date_Time,Invoice_No) values ('" & rs1.Fields("Line_Code") & "'," & NewLocation & "," & Ave_Cost & "," & Qty & ",getdate()," & Invoice_No & ")"
                            rs2.Close
                            rs1.MoveNext
                            
                        Wend
                        rs1.Close
                        
                    Case Else 'Location from Kitchen Printer
                        ActiveReadServer1 "Select * from Printer_Links where Printer = '" & Trim(Kitchen_Printer) & "'"
                        If rs1.RecordCount > 0 Then
                            If Val(rs1.Fields("Location_No") & "") <> 0 Then
                                NewLocation = Val(rs1.Fields("Location_No") & "")
                            Else
                                NewLocation = Location_No
                            End If
                        End If
                        rs1.Close
                        ActiveReadServer1 "Select Unit_of_Measure as Recipe_Unit, Line_Code,Qty_Used,isnull((Select Unit_Size from Products where Products.Product_code = Recipes.Line_Code),0) as Unit_Size, (Select Unit_of_Measure from Products where Products.Product_code = Recipes.Line_Code) as Unit_of_Measure,(Select Ave_Cost from Products where Products.Product_code = Recipes.Line_Code) as Ave_Cost from Recipes where Product_Code = '" & Product_Code & "' and Line_Type in (3,4)"
                        SaveQty = Qty
                        While Not rs1.EOF
                            Qty = SaveQty
                            Select Case rs1.Fields("Qty_Used")
                                Case "1 x 25ml"
                                    If Val(rs1.Fields("Unit_Size") & "") <> 0 Then
                                        Qty = Round(25 / rs1.Fields("Unit_Size"), 4) * Qty
                                    End If
                                Case "2 x 25ml"
                                    If Val(rs1.Fields("Unit_Size") & "") <> 0 Then
                                        Qty = Round(50 / rs1.Fields("Unit_Size"), 4) * Qty
                                    End If
                                Case Else
                                    UnitSize = rs1.Fields("Unit_Size")
                                    If rs1.Fields("Unit_of_Measure") <> rs1.Fields("Recipe_Unit") Then
                                        Select Case UCase(rs1.Fields("Unit_of_Measure") & " to " & rs1.Fields("Recipe_Unit"))
                                            Case "ML TO LT"
                                                UnitSize = rs1.Fields("Unit_Size") / 1000
                                            Case "LT TO ML"
                                                If rs1.Fields("Unit_Size") = 0 Then
                                                    UnitSize = 1000
                                                Else
                                                    UnitSize = rs1.Fields("Unit_Size") * 1000
                                                End If
                                            Case "G TO KG"
                                                UnitSize = rs1.Fields("Unit_Size") / 1000
                                            Case "KG TO G"
                                                If rs1.Fields("Unit_Size") = 0 Then
                                                    UnitSize = 1000
                                                Else
                                                    UnitSize = rs1.Fields("Unit_Size") * 1000
                                                End If
                                            Case Else
                                                UnitSize = rs1.Fields("Unit_Size")
                                        End Select
                                    End If
                                    If Val(UnitSize & "") <> 0 Then
                                        Qty = Round(rs1.Fields("Qty_Used") / UnitSize, 4) * Qty
                                    Else
                                        Qty = rs1.Fields("Qty_Used") * Qty
                                    End If
                            End Select
                            Ave_Cost = rs1.Fields("Ave_Cost") * Qty
                            ActiveReadServer2 "Select * from Quantities where Product_Code = '" & rs1.Fields("Line_Code") & "' and Location_No = " & NewLocation
                            If rs2.RecordCount > 0 Then
                                
                                ActiveUpdateServer "Update Quantities set Stock_on_Hand = Stock_on_Hand - " & Qty & " where Product_Code = '" & rs1.Fields("Line_Code") & "' and Location_No = " & NewLocation
                          
                            Else
                                
                                ActiveUpdateServer "Insert into Quantities (Product_Code,Stock_on_Hand,Location_No) values ('" & rs1.Fields("Line_Code") & "'," & Qty * -1 & "," & NewLocation & ")"
                              
                            End If
                          
                            ActiveUpdateServer "Insert into Consumption_Journal (Product_Code,Location_No,Ave_Cost,Qty_Consumed,Date_Time,Invoice_No) values ('" & rs1.Fields("Line_Code") & "'," & NewLocation & "," & Ave_Cost & "," & Qty & ",getdate()," & Invoice_No & ")"
                            rs2.Close
                            rs1.MoveNext
                      
                        Wend
                        rs1.Close
                        
                End Select
        End Select
        rs.Close
    End If
    On Error GoTo 0
    Exit Sub
trap:
    MsgBox err.Description, vbCritical, "HeroPOS"
    Resume Next
End Sub
Private Sub UpdateDisplay(KeyFunction)
    On Error Resume Next
    Select Case Panel_no
        Case 0
            With frmSales
                .grdMain.HighLight = flexHighlightAlways
                Select Case KeyFunction
                    Case KeyType.ClearKey
                        .picDigit.Visible = False
                        .lblKeyRegister.TextAlign = fmTextAlignLeft
                        .grdMain.Rows = 1
                        .lblTender.Caption = "0.00"
                        .lblCash.Caption = ""
                        If KeyRegister = "*" Then
                            .lblKeyRegister.Caption = KeyRegister
                        End If
                    Case KeyType.InputKey
                        If Asc(Right(KeyRegister, 1)) > 47 And Asc(Right(KeyRegister, 1)) < 58 Then
                            If InStr(KeyRegister, Chr(215)) = Len(KeyRegister) - 2 Then
                                .picDigit.Visible = False
                            Else
                                If Left(.lblKeyRegister.Caption, 1) <> "*" Then
                                    .picDigit.Visible = True
                                    .lblDigit = Right(KeyRegister, 2)
                                End If
                            End If
                        Else
                            .picDigit.Visible = False
                        End If
                        .lblKeyRegister.TextAlign = fmTextAlignRight
                        If .lblKeyRegister.Caption = "*" Then
                            If Left(.lblKeyRegister.Caption, 1) = "*" Then
                                .lblKeyRegister.Caption = .lblKeyRegister.Caption & KeyRegister
                            End If
                            KeyRegister = .lblKeyRegister.Caption
                        Else
                            .lblKeyRegister.Caption = KeyRegister
                        End If
                        If Left(.lblKeyRegister.Caption, 1) = "*" Then
                            .lblKeyRegister.TextAlign = fmTextAlignLeft
                            KeyRegister = Replace(KeyRegister, "**", "*")
                            .lblKeyRegister.Caption = Replace(.lblKeyRegister.Caption, "**", "*")
                        End If
                    Case KeyType.FunctionKey
                        .picDigit.Visible = False
                        If KeyRegister = " (Load Member Number) " Then
                            frmSales.Tag = "1"
                            frmMember.Show vbModal
                            Select Case KeyRegister
                                Case ""
                                    frmSales.lblKeyRegister = ""
                                    frmSales.lblDebtor = ""
                                    TillData.Account_No = ""
                                Case Else
                                    ActiveReadServer "Select * from Debtors where Debtor_No ='" & KeyRegister & "'"
                                    If rs.RecordCount > 0 Then
                                        frmSales.lblKeyRegister = "Member - " & rs.Fields("Debtor_Name") & " (" & KeyRegister & ")"
                                        frmSales.lblDebtor = "Member - " & rs.Fields("Debtor_Name") & " (" & KeyRegister & ")"
                                        TillData.Account_No = KeyRegister
                                        TillData.Creditbalance = rs.Fields("Balance")
                                        TillData.Creditlimit = rs.Fields("Credit_Limit")
                                    End If
                                    rs.Close
                                    KeyRegister = ""
                            End Select
                        End If
                        If KeyRegister = " (Search) " Then
                            If InStr(.lblKeyRegister, "(") = 0 Then
                                If InStr(.lblKeyRegister, Chr(215)) = 0 Then
                                    .lblKeyRegister = ""
                                End If
                            End If
                            .adoData.ConnectionString = cnnMain.ConnectionString
                            .adoData.CursorLocation = adUseServer
                            .adoData.CursorType = adOpenStatic
                            .adoData.LockType = adLockReadOnly
                            .adoData.RecordSource = "Select Product_Code,Description,Department,SOH,Landed_Cost,Tax_Rate,Selling_Price from Product_List where Sales_Item=1 and Touch_Item=1 order by Description"
                            .adoData.Refresh
                            .grdFind.ColHidden(4) = True
                            .cmdArr(2).Tag = ""
                            .grdFind.RowHeight(0) = 740
                            .grdFind.TextMatrix(0, 0) = " Product Code"
                            .grdFind.TextMatrix(0, 1) = "Description"
                            .grdFind.TextMatrix(0, 2) = "Department"
                            .grdFind.TextMatrix(0, 3) = "Stock on Hand "
                            .grdFind.TextMatrix(0, 4) = "Landed Cost "
                            .grdFind.TextMatrix(0, 5) = "Tax Rate "
                            .grdFind.TextMatrix(0, 6) = "Price (Incl) "
                            .grdFind.ColAlignment(0) = flexAlignLeftCenter
                            .grdFind.ColAlignment(1) = flexAlignLeftCenter
                            .grdFind.ColAlignment(2) = flexAlignLeftCenter
                            .grdFind.ColAlignment(3) = flexAlignRightCenter
                            .grdFind.ColAlignment(4) = flexAlignRightCenter
                            .grdFind.ColAlignment(5) = flexAlignRightCenter
                            .grdFind.ColAlignment(6) = flexAlignRightCenter
                            .grdFind.ColWidth(0) = .grdFind.Width * 0.12
                            .grdFind.ColWidth(1) = .grdFind.Width * 0.33
                            .grdFind.ColWidth(2) = .grdFind.Width * 0.25
                            .grdFind.ColWidth(3) = .grdFind.Width * 0.1
                            .grdFind.ColWidth(4) = .grdFind.Width * 0.1
                            .grdFind.ColWidth(5) = .grdFind.Width * 0.1
                            .grdFind.ColWidth(6) = .grdFind.Width * 0.15
                            .grdFind.Col = 1
                            .grdFind.Row = 1
                            .grdDept.Rows = 0
                            For i = 0 To 6
                                .cmdDeptStrip(i).Value = 0
                            Next i
                            .cmdDeptStrip(0).Caption = ""
                            .cmdDeptStrip(0).Picture = ""
                            DoEvents
                            ActiveReadServer "Select * from Departments_Panel1"
                            i = -1
                            b = 0
                            While Not rs.EOF
                                i = i + 1
                                .grdDept.Rows = .grdDept.Rows + 1
                                If i < 7 And Not rs.EOF Then
                                    .cmdDeptStrip(i).Caption = Replace(rs.Fields("Dept_Name"), "&", "&&")
                                    .cmdDeptStrip(i).Tag = rs.Fields("Department_no")
                                    If .cmdDeptStrip(i).Visible = False Then .cmdDeptStrip(i).Visible = True
                                    .grdDept.Row = .grdDept.Rows - 1
                                    .grdDept.TextMatrix(.grdDept.Rows - 1, 0) = Replace(rs.Fields("Dept_Name"), "&", "&&")
                                    .grdDept.TextMatrix(.grdDept.Rows - 1, 1) = rs.Fields("Department_No")
                                Else
                                    .grdDept.TextMatrix(.grdDept.Rows - 1, 0) = Replace(rs.Fields("Dept_Name"), "&", "&&")
                                    .grdDept.TextMatrix(.grdDept.Rows - 1, 1) = rs.Fields("Department_No")
                                End If
                                rs.MoveNext
                            Wend
                            rs.Close
                            For b = i + 1 To .cmdDeptStrip.Count - 1
                               .cmdDeptStrip(b).Caption = "1"
                               .cmdDeptStrip(b).Tag = ""
                               .cmdDeptStrip(b).Visible = False
                            Next b
                            .picSearch.Height = 10935
                            .grdFind.ColHidden(8) = True
                            DoEvents
                            .picSearch.Visible = True
                            DoEvents
                            .grdFind.SetFocus
                            On Error GoTo 0
                            Exit Sub
                        End If
                        If KeyRegister = " (Corr) " Then
                            For i = .grdMain.Rows - 1 To 1 Step -1
                                If .grdMain.TextMatrix(i, 0) <> "" Then
                                    .grdMain.Rows = i + 1
                                    Exit For
                                End If
                            Next i
                            .grdMain.Row = .grdMain.Rows - 1
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = TillData.ExtraFunc
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 17) = TillData.UserOveride
                            TillData.UserOveride = 0
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            .grdMain.Cell(flexcpFontStrikethru, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = True
                            .lblKeyRegister.Caption = "Item Correct - " & TillData.Description & " @ R" & Format(TillData.Price, "0.00")
                            If .grdMain.TextMatrix(.grdMain.Rows - 1, 8) <> "Wastage" Then  'Kotie
                                TillData.TaxTotal = Round(TillData.TaxTotal, 3) - Round(((TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100))), 3)
                                If TillData.TaxRate <> 0 Then
                                    TillData.TaxableSales = TillData.TaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                    TillData.CollectedTax = TillData.CollectedTax - ((TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100)))
                                Else
                                    TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                End If
                            End If
                            Sale_Total = 0
                             Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                           ' For i = 1 To .grdMain.Rows - 1
                           '     If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                           '         If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                           '             Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                           '         End If
                           '     End If
                           ' Next i
                           ' If Val(Swiss_Round) <> 0 Then
                           '     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                           '        q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                           '         Sale_Total = Sale_Total - q
                           '     End If
                           ' End If
                           Sale_Total = TillData.SaleTotal
                            .lblTender.Caption = Format(Sale_Total, "0.00")
                            
                            .lblCash.Caption = "Subtotal"
                            If .grdMain.Rows > 11 Then .grdMain.TopRow = .grdMain.Row - 11
                            TillData.Corrects = TillData.Corrects + (TillData.Price * TillData.Qty)
                            TillData.CorrectCount = TillData.CorrectCount + 1
                            On Error GoTo 0
                            Exit Sub
                        End If
                        If KeyRegister = " (Subtotal) " Then
                            Sale_Total = 0
                             Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                            'For i = 1 To .grdMain.Rows - 1
                            '    If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            '        If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                            '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                            '        End If
                            '    End If
                            'Next i
                            'If Val(Swiss_Round) <> 0 Then
                            '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                            '       q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                            '        Sale_Total = Sale_Total - q
                            '    End If
                            'End If
                            
                            Sale_Total = TillData.SaleTotal
                            .grdMain.Rows = .grdMain.Rows + 1
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 0) = ""
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Subtotal"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Sale_Total, "0.00")
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = TillData.Keystring
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = ""
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = ""
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = KeyRegister
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = ""
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0FFC0
                            .lblKeyRegister.TextAlign = fmTextAlignLeft
                            .lblKeyRegister.Caption = "Subtotal"
                            .lblTender.Caption = Format(Sale_Total, "0.00")
                            .lblCash.Caption = "Subtotal"
                            KeyRegister = ""
                            .grdMain.Row = .grdMain.Rows - 1
                            If .grdMain.Rows > 11 Then .grdMain.TopRow = .grdMain.Row - 11
                        End If
                        .lblKeyRegister.TextAlign = fmTextAlignRight
                        .lblKeyRegister.Caption = KeyRegister
                    Case KeyType.ItemizerKey
                        MarkforPrinting = 0
                        .picDigit.Visible = False
                        If TillData.Keystring = "Plu" Or TillData.Keystring = "Dept" Then .cmdInput(14).Caption = "Corr"
                        .lblKeyRegister.TextAlign = fmTextAlignLeft
                        If TillData.Keystring = "Dept" Then
                            .lblKeyRegister.Caption = "Department - " & TillData.Qty & " x " & TillData.Description & " @ R" & Format(TillData.Price, "0.00")
                        Else
                            .lblKeyRegister.Caption = TillData.Qty & " x " & TillData.Description & " @ R" & Format(TillData.Price, "0.00")
                        End If
                        .grdMain.Rows = .grdMain.Rows + 1
                        If TillData.Kitchen1 <> "" And TillData.Kitchen1 <> "<None>" Then
                            .grdMain.ColHidden(14) = False
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 14, .grdMain.Rows - 1, 14) = &HC0FFFF
                            MarkforPrinting = .grdMain.Rows - 1
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 14) = "P"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 11) = TillData.Kitchen1
                        End If
                        If TillData.Kitchen2 <> "" And TillData.Kitchen2 <> "<None>" Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 12) = TillData.Kitchen2
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 0) = TillData.Qty
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = TillData.ShortDesc
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Val(TillData.Price) * Val(TillData.Qty), "0.00")
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = TillData.Keystring
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = TillData.Cost
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = TillData.TaxRate
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 6) = TillData.TaxType
                      '  If .grdMain.TextMatrix(.grdMain.Rows - 1, 8) <> "Wastage" Then  'Kotie
                      '      If TillData.TaxRate <> 0 Then
                      '          TillData.TaxableSales = TillData.TaxableSales + Val(TillData.Price) * Val(TillData.Qty)
                      '          TillData.CollectedTax = TillData.CollectedTax + (TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100))
                      '      Else
                      '          TillData.NonTaxableSales = TillData.NonTaxableSales + Val(TillData.Price) * Val(TillData.Qty)
                      '      End If
                      '  End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = KeyRegister
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 9) = TillData.ProductCode
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 10) = TillData.DeptNo
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 13) = TillData.PriceOveride
                        If InStr(KeyRegister, "Void") <> 0 Then  'Kotie 19-03-2013  06:45  Made void to be on any line via Void all
                            voidline = 0
                            'LineToVoid = .grdMain.Row
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Void - " & .grdMain.TextMatrix(.grdMain.Rows - 1, 1)
                            .lblKeyRegister.Caption = "(Void) " & .lblKeyRegister.Caption
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            TillData.VoidTotal = TillData.VoidTotal + (Val(TillData.Price) * Val(TillData.Qty))
                            TillData.VoidCount = TillData.VoidCount + 1
                            TillData.ExtraFunc = "Void"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 17) = TillData.UserOveride
                            TillData.UserOveride = 0
                            .grdMain.Cell(flexcpForeColor, LineToVoid, 0, LineToVoid, 2) = vbRed
                            .grdMain.TextMatrix(LineToVoid, 8) = "Voided on - " & .grdMain.Rows - 1
                            voidline = LineToVoid
                        End If
                        If InStr(KeyRegister, "Return Item") <> 0 Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Return Item - " & .grdMain.TextMatrix(.grdMain.Rows - 1, 1)
                            .lblKeyRegister.Caption = "(Return Item) " & .lblKeyRegister.Caption
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            TillData.ReturnTotal = TillData.ReturnTotal + Val(TillData.Price) * Val(TillData.Qty)
                            TillData.ReturnCount = TillData.ReturnCount + 1
                            TillData.ExtraFunc = "Return Item"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 17) = TillData.UserOveride
                            TillData.UserOveride = 0
                        End If
                        If InStr(KeyRegister, "Wastage") <> 0 Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Wastage - " & .grdMain.TextMatrix(.grdMain.Rows - 1, 1)
                            .lblKeyRegister.Caption = "(Wastage) " & .lblKeyRegister.Caption
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            TillData.UllageTotal = TillData.UllageTotal + Val(TillData.Price) * Val(TillData.Qty)
                            TillData.UllageCount = TillData.UllageCount + 1
                            TillData.ExtraFunc = "Wastage"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 17) = TillData.UserOveride
                            TillData.UserOveride = 0
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = TillData.ExtraFunc
                        If .grdMain.TextMatrix(.grdMain.Rows - 1, 8) <> "Wastage" Then  'Kotie
                            If TillData.TaxRate <> 0 Then
                                TillData.TaxableSales = TillData.TaxableSales + Val(TillData.Price) * Val(TillData.Qty)
                                TillData.CollectedTax = TillData.CollectedTax + (TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100))
                            Else
                                TillData.NonTaxableSales = TillData.NonTaxableSales + Val(TillData.Price) * Val(TillData.Qty)
                            End If
                        End If
                        
                        .grdMain.Row = .grdMain.Rows - 1
                        If TillData.Recipe = 1 And TillData.ExtraFunc <> "Void" And .grdMenu.Rows > 0 Then
                            For i = 0 To frmRecipe.grdMess.Rows - 1
                                .grdMain.Rows = .grdMain.Rows + 1
                                .grdMain.Row = .grdMain.Rows - 1
                                If MarkforPrinting > 0 Then
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 14) = "P"
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 11) = TillData.Kitchen1
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 12) = TillData.Kitchen2
                                    .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 14, .grdMain.Rows - 1, 14) = &HC0FFFF
                                End If
                                .grdMain.TextMatrix(.grdMain.Row, 1) = "    >" & frmRecipe.grdMess.TextMatrix(i, 1)
                                .grdMain.Cell(flexcpForeColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC00000
                                .grdMain.Cell(flexcpFontBold, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = True
                                'Product Code
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 9) = frmRecipe.grdMess.TextMatrix(i, 0)
                                If Val(frmRecipe.grdMess.TextMatrix(i, 0)) <> 0 Then
                                    ActiveReadServer "Select Department_No,Sales_Tax,Tax_Type,Ave_Cost, Selling_Price  from Products where Product_Code = '" & frmRecipe.grdMess.TextMatrix(i, 0) & "'"
                                    If rs.RecordCount > 0 Then
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = rs.Fields("Ave_Cost")
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = rs.Fields("Sales_Tax")
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 6) = rs.Fields("Tax_Type")
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = frmRecipe.grdMess.TextMatrix(i, 1)
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 10) = rs.Fields("Department_No")
                                        
                                        If rs.Fields("Receipe_Charge_Item") = 1 Then
                                            .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(rs.Fields("selling_price"), "0.00")
                                            If TillData.TaxRate <> 0 Then
                                                TillData.TaxableSales = TillData.TaxableSales + rs.Fields("selling_price")
                                                TillData.CollectedTax = TillData.CollectedTax + (rs.Fields("selling_price")) - ((rs.Fields("selling_price")) / ((100 + rs.Fields("Sales_Tax")) / 100))
                                            Else
                                                TillData.NonTaxableSales = TillData.NonTaxableSales + Val(rs.Fields("selling_price"))
                                            End If
                                        End If
                                            If InStr(KeyRegister, "Return Item") <> 0 Then
                                                .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(.grdMain.TextMatrix(.grdMain.Rows - 1, 2) * -1, "0.00")
                                            End If
                                       
                                    End If
                                    rs.Close
                                Else
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = frmRecipe.grdMess.TextMatrix(i, 1)
                                End If
                            Next i
                        Else
                            While Trim(.grdMain.TextMatrix(voidline + 1, 0)) = ""
                                .grdMain.Rows = .grdMain.Rows + 1
                                .grdMain.Row = .grdMain.Rows - 1
                                For i = 0 To .grdMain.Cols - 1
                                    If i = 1 Then
                                        .grdMain.Cell(flexcpFontBold, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = True
                                        .grdMain.Cell(flexcpForeColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = vbRed
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, i) = "Void " & Trim(.grdMain.TextMatrix(voidline + 1, i))
                                    Else
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, i) = .grdMain.TextMatrix(voidline + 1, i)
                                        .grdMain.Cell(flexcpForeColor, voidline + 1, 0, voidline + 1, 2) = vbRed
                                        If i = 2 Then
                                            If Val(.grdMain.TextMatrix(.grdMain.Rows - 1, i)) <> 0 Then
                                            .grdMain.TextMatrix(.grdMain.Rows - 1, i) = Format(Val(.grdMain.TextMatrix(.grdMain.Rows - 1, i)) * -1, "0.00")
                                            End If
                                        End If
                                        If InStr(0, KeyRegister, "Return Item") = True Then
                                            .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(.grdMain.TextMatrix(.grdMain.Rows - 1, 2) * -1, "0.00")
                                        End If
                                    End If
                                Next i
                                voidline = voidline + 1
                            Wend
                        End If
                        KeyRegister = ""
                        Sale_Total = 0
                       ' For i = 1 To .grdMain.Rows - 1
                       '     If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                       '         If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                       '             Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                       '         End If
                       '     End If
                       ' Next i
                       ' If Val(Swiss_Round) <> 0 Then
                       '     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                       '         If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                       '             q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                       '             Sale_Total = Sale_Total - q
                       '         End If
                       '     End If
                       ' End If
                        Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                        Sale_Total = TillData.SaleTotal
                        .lblTender.Caption = Format(Sale_Total, "0.00")
                        .lblCash.Caption = "Subtotal"
                        If .grdMain.Rows > 11 Then .grdMain.TopRow = .grdMain.Row - 11
                        If .grdMain.Rows = 1 Then
                            .cmdFancy(4).Caption = "Member No"
                        Else
                            .cmdFancy(4).Caption = "Discount"
                        End If
                    Case KeyType.FinalizationKey
                        .picDigit.Visible = False
                        Sale_Total = 0
                         Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                        'For i = 1 To .grdMain.Rows - 1
                        '    If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                        '        If .grdMain.TextMatrix(i, 8) <> "Wastage" Then
                        '            If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                        '                If .grdMain.TextMatrix(i, 3) <> "Short" Then
                        '                    Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                        '                End If
                        '            End If
                        '        End If
                        '    End If
                        'Next i
                        Sale_Total = TillData.SaleTotal
                        If TillData.Change < 0 Then
                            .lblTender.Caption = Format(TillData.Change * -1, "0.00")
                            .lblCash.Caption = "Short Tendered"
                        Else
                            If TillData.Change = 0 Then
                                If .lblCash = "Short Tendered" Then
                                    .lblCash.Caption = "Change"
                                    .lblTender.Caption = Format(TillData.Change, "0.00")
                                Else
                                    .lblCash.Caption = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "")
                                End If
                            Else
                                .lblCash.Caption = "Change"
                                .lblTender.Caption = Format(TillData.Change, "0.00")
                            End If
                        End If
                        .grdMain.Rows = .grdMain.Rows + 1
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 0) = ""
                        If KeyRegister = "<No Sale>" Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "No Sale"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = "No Sale"
                        Else
                            If TillData.Change < 0 Then
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "") & " Tendered"
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = "Short"
                            Else
                                If .grdMain.TextMatrix(.grdMain.Rows - 2, 3) = "Short" Then
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "") & " Tendered"
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = "Short"
                                Else
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "") & " Sale"
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "")
                                End If
                            End If
                        End If
                        If TillData.Change < 0 Then
                            KeyRegister = Replace(KeyRegister, "-", "")
                            If Left(KeyRegister, 1) = "R" Then
                                KeyRegister = Mid(KeyRegister, 2)
                            End If
                             .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Val(Mid(KeyRegister, 1, InStr(KeyRegister, "<") - 1)) / 100, "0.00")
                        Else
                            If .grdMain.TextMatrix(.grdMain.Rows - 2, 3) = "Short" Then
                                KeyRegister = Replace(KeyRegister, "-", "")
                                If Left(KeyRegister, 1) = "R" Then
                                    KeyRegister = Mid(KeyRegister, 2)
                                End If
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Val(Mid(KeyRegister, 1, InStr(KeyRegister, "<") - 1)) / 100, "0.00")
                            Else
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Sale_Total, "0.00")
                            End If
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = TillData.Tendered
                        If Right(.grdMain.TextMatrix(.grdMain.Rows - 1, 1), 8) = "Tendered" Then
                            If TillData.Change < 0 Then
                                Tax_Total = TillData.TaxTotal * (.grdMain.TextMatrix(.grdMain.Rows - 1, 2) / TillData.SaleTotal)
                            Else
                                Tax_Total = TillData.TaxTotal * ((.grdMain.TextMatrix(.grdMain.Rows - 1, 2) - TillData.Change) / TillData.SaleTotal)
                            End If
                        Else
                            Tax_Total = TillData.TaxTotal
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = Tax_Total
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 13) = TillData.Change
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = KeyRegister
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = TillData.ExtraFunc
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 17) = TillData.UserOveride
                        TillData.UserOveride = 0

                        .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0FFFF
                        .lblKeyRegister.TextAlign = fmTextAlignLeft
                        If KeyRegister = "<No Sale>" Then
                            .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - No Sale"
                        Else
                             If TillData.Change < 0 Then
                                .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - R" & Format(TillData.Tendered, "0.00") & " Tendered of R" & Format(Sale_Total, "0.00")
                            Else
                                If .grdMain.TextMatrix(.grdMain.Rows - 2, 3) = "Short" Then
                                    .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - Sale Total: R" & Format(Sale_Total, "0.00")
                                Else
                                    .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - R" & Format(TillData.Tendered, "0.00") & " Tendered"
                                End If
                            End If
                        End If
                        .cmdInput(14).Caption = "No Sale"
                        .grdMain.Row = .grdMain.Rows - 1
                        If .grdMain.Rows > 11 Then .grdMain.TopRow = .grdMain.Row - 11
                        .grdMain.HighLight = flexHighlightWithFocus
                        .picHoldFocus.SetFocus
                End Select
                
                
                If Devices.DisplayModel <> "" Then
                
                Select Case Devices.DisplayModel
                
                Case "CD7220", "Posiflex"
                
                'If Devices.DisplayModel = "CD7220" Then
                    filenum = FreeFile
                    Open Devices.DisplayPort For Output As filenum
                    Print #filenum, Chr(27) & "@"
                    Select Case KeyFunction
                        Case KeyType.InputKey
                            Print #filenum, Replace(String(20 - Len(Trim(Left(.lblKeyRegister.Caption, 20))), " ") & Trim(Left(.lblKeyRegister.Caption, 20)) & "SUBTOTAL" & UCase(String(20 - Len("SUBTOTAL") - Len(.lblTender.Caption), " ") & .lblTender.Caption), Chr(215), "X")
                        Case KeyType.ItemizerKey
                            DisplayLine = Replace(UCase(Trim(Mid(.lblKeyRegister.Caption, 1, InStr(.lblKeyRegister.Caption, "@") - 1))), "1 X ", "")
                            Print #filenum, Trim(Left(DisplayLine, 20)) & String(20 - Len(Trim(Left(DisplayLine, 20))), " ");
                            Print #filenum, UCase(Trim(.lblCash.Caption & String(20 - Len(.lblCash.Caption) - Len(.lblTender.Caption), " ") & .lblTender.Caption))
                        Case KeyType.FinalizationKey
                            DisplayLine = Replace(UCase(Trim(Mid(.lblKeyRegister.Caption, InStr(.lblKeyRegister.Caption, "-") + 1))), "1 X ", "")
                            Print #filenum, Trim(Left(DisplayLine, 20)) & String(20 - Len(Trim(Left(DisplayLine, 20))), " ");
                            Print #filenum, UCase(Trim(.lblCash.Caption & String(20 - Len(.lblCash.Caption) - Len(.lblTender.Caption), " ") & .lblTender.Caption))
                    End Select
                    Close #filenum
                    End Select
                End If
                
            End With
        
        
        
        
        
        Case 1
            With frmSales1
                .grdMain.HighLight = flexHighlightAlways
                Select Case KeyFunction
                    Case KeyType.ClearKey
                        .picDigit.Visible = False
                        .lblKeyRegister.TextAlign = fmTextAlignLeft
                        .grdMain.Rows = 1
                        .lblTender.Caption = "0.00"
                        .lblCash.Caption = ""
                    Case KeyType.InputKey
                        If Asc(Right(KeyRegister, 1)) > 47 And Asc(Right(KeyRegister, 1)) < 58 Then
                            .picDigit.Visible = True
                            .lblDigit = Right(KeyRegister, 2)
                        Else
                            .picDigit.Visible = False
                        End If
                        .lblKeyRegister.TextAlign = fmTextAlignRight
                        .lblKeyRegister.Caption = KeyRegister
                    Case KeyType.FunctionKey
                        .picDigit.Visible = False
                        If KeyRegister = " (Corr) " Then
                            For i = .grdMain.Rows - 1 To 1 Step -1
                                If .grdMain.TextMatrix(i, 0) <> "" Then
                                    .grdMain.Rows = i + 1
                                    Exit For
                                End If
                            Next i
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = TillData.ExtraFunc
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 17) = TillData.UserOveride
                            TillData.UserOveride = 0
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            .grdMain.Cell(flexcpFontStrikethru, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = True
                            .lblKeyRegister.Caption = "Item Correct - " & TillData.Description & " @ R" & Format(TillData.Price, "0.00")
                            TillData.TaxTotal = Round(TillData.TaxTotal, 3) - Round(((TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100))), 3)
                            If TillData.TaxRate <> 0 Then
                                TillData.TaxableSales = TillData.TaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                TillData.CollectedTax = TillData.CollectedTax - ((TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100)))
                            Else
                                TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                            End If
                            Sale_Total = 0
                             Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                           ' For i = 1 To .grdMain.Rows - 1
                           '     If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                           '         If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                           '             Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                           '         End If
                           '     End If
                           ' Next i
                           ' If Val(Swiss_Round) <> 0 Then
                           '     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                           '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                           '             q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                           '             Sale_Total = Sale_Total - q
                           '         End If
                           '     End If
                           ' End If
                            Sale_Total = TillData.SaleTotal
                            .lblTender.Caption = Format(Sale_Total, "0.00")
                            
                            .lblCash.Caption = "Subtotal"
                            TillData.Corrects = TillData.Corrects + (TillData.Price * TillData.Qty)
                            TillData.CorrectCount = TillData.CorrectCount + 1
                            If .grdMain.Rows > 11 Then .grdMain.TopRow = .grdMain.Row - 11
                            On Error GoTo 0
                            Exit Sub
                        End If
                        If KeyRegister = " (Subtotal) " Then
                            Sale_Total = 0
                             Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                           ' For i = 1 To .grdMain.Rows - 1
                           '     If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                           '         If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                           '             Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                           '         End If
                           '     End If
                           ' Next i
                           ' If Val(Swiss_Round) <> 0 Then
                           '     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                           '         If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                           '             q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                           '             Sale_Total = Sale_Total - q
                           '         End If
                           '     End If
                           ' End If
                           
                            Sale_Total = TillData.SaleTotal
                            .grdMain.Rows = .grdMain.Rows + 1
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 0) = ""
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Subtotal"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Sale_Total, "0.00")
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = TillData.Keystring
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = ""
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = ""
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = KeyRegister
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = ""
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0FFC0
                            .lblKeyRegister.TextAlign = fmTextAlignLeft
                            .lblKeyRegister.Caption = "Subtotal"
                            .lblTender.Caption = Format(Sale_Total, "0.00")
                            .lblCash.Caption = "Subtotal"
                            KeyRegister = ""
                            .grdMain.Row = .grdMain.Rows - 1
                            .grdMain.ShowCell .grdMain.Rows - 1, 0
                        End If
                        .lblKeyRegister.TextAlign = fmTextAlignRight
                        .lblKeyRegister.Caption = KeyRegister
                    Case KeyType.ItemizerKey
                        MarkforPrinting = 0
                        .picDigit.Visible = False
                        .lblKeyRegister.TextAlign = fmTextAlignLeft
                        If TillData.Keystring = "Plu" Or TillData.Keystring = "Dept" Then .cmdDept(6).Caption = "Corr"
                        If TillData.Keystring = "Dept" Then
                            .lblKeyRegister.Caption = "Department - " & TillData.Qty & " x " & TillData.Description & " @ R" & Format(TillData.Price, "0.00")
                        Else
                            If TillData.Keystring = "Service Charge" Then
                                TillData.Recipe = 0
                                .lblKeyRegister.Caption = "Service Charge: R" & Format(TillData.Tipp, "0.00")
                                TillData.ShortDesc = "Service Charge"
                                TillData.Price = TillData.Tipp
                                TillData.ProductCode = ""
                                TillData.Weight = 0
                                TillData.Qty = ""
                            Else
                                .lblKeyRegister.Caption = TillData.Qty & " x " & TillData.Description & " @ R" & Format(TillData.Price, "0.00")
                            End If
                        End If
                        .grdMain.Rows = .grdMain.Rows + 1
                        If TillData.Keystring = "Service Charge" Then
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HFBCFBF
                            .grdMain.Row = .grdMain.Rows - 1
                        End If
                        If TillData.Kitchen1 <> "" And TillData.Kitchen1 <> "<None>" Then
                            .grdMain.ColHidden(14) = False
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 14, .grdMain.Rows - 1, 14) = &HC0FFFF
                            MarkforPrinting = .grdMain.Rows - 1
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 14) = "P"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 11) = TillData.Kitchen1
                        End If
                        If TillData.Kitchen2 <> "" And TillData.Kitchen2 <> "<None>" Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 12) = TillData.Kitchen2
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 0) = TillData.Qty
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = TillData.ShortDesc
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Val(TillData.Price) * Val(TillData.Qty), "0.00")
                        If TillData.Keystring = "Service Charge" Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Val(TillData.Price), "0.00")
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = TillData.Keystring
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = TillData.Cost
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = TillData.TaxRate
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 6) = TillData.TaxType

                        .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = KeyRegister
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 9) = TillData.ProductCode
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 10) = TillData.DeptNo
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 13) = TillData.PriceOveride
                        If InStr(KeyRegister, "Void") <> 0 Then
                            voidline = 0
                            LineToVoid = .grdMain.Row
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Void - " & .grdMain.TextMatrix(.grdMain.Rows - 1, 1)
                            .lblKeyRegister.Caption = "(Void) " & .lblKeyRegister.Caption
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            TillData.VoidTotal = TillData.VoidTotal + (Val(TillData.Price) * Val(TillData.Qty))
                            TillData.VoidCount = TillData.VoidCount + 1
                            TillData.ExtraFunc = "Void"
                            .grdMain.Cell(flexcpForeColor, .grdMain.Row, 0, .grdMain.Row, 2) = vbRed
                            .grdMain.TextMatrix(.grdMain.Row, 8) = "Voided on - " & .grdMain.Rows - 1
                        End If
                        If InStr(KeyRegister, "Return Item") <> 0 Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Return Item - " & .grdMain.TextMatrix(.grdMain.Rows - 1, 1)
                            .lblKeyRegister.Caption = "(Return Item) " & .lblKeyRegister.Caption
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            TillData.ReturnTotal = TillData.ReturnTotal + Val(TillData.Price) * Val(TillData.Qty)
                            TillData.ReturnCount = TillData.ReturnCount + 1
                            TillData.ExtraFunc = "Return Item"
                        End If
                        If InStr(KeyRegister, "Wastage") <> 0 Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Wastage - " & .grdMain.TextMatrix(.grdMain.Rows - 1, 1)
                            .lblKeyRegister.Caption = "(Wastage) " & .lblKeyRegister.Caption
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            TillData.UllageTotal = TillData.UllageTotal + Val(TillData.Price) * Val(TillData.Qty)
                            TillData.UllageCount = TillData.UllageCount + 1
                            TillData.ExtraFunc = "Wastage"
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = TillData.ExtraFunc
                        If TillData.Keystring <> "Service Charge" Then
                            If .grdMain.TextMatrix(.grdMain.Rows - 1, 8) <> "Wastage" Then  'Kotie
                                If TillData.TaxRate <> 0 Then
                                    TillData.TaxableSales = TillData.TaxableSales + Val(TillData.Price) * Val(TillData.Qty)
                                    TillData.CollectedTax = TillData.CollectedTax + (TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100))
                                Else
                                    TillData.NonTaxableSales = TillData.NonTaxableSales + Val(TillData.Price) * Val(TillData.Qty)
                                End If
                            End If
                        End If
                         Update_Sale_Total (Panel_no) 'Kotie 10/04/2013

                       ' Sale_Total = 0
                       ' For i = 1 To .grdMain.Rows - 1
                       '     If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                       '         If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                       '             Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                       '         End If
                       '     End If
                       ' Next i
                       ' If Val(Swiss_Round) <> 0 Then
                       '     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                       '         If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                       '             q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                       '             Sale_Total = Sale_Total - q
                       '         End If
                       '     End If
                       ' End If
                        Sale_Total = TillData.SaleTotal
                        .lblTender.Caption = Format(Sale_Total, "0.00")
                        .lblCash.Caption = "Subtotal"
                        
                        .grdMain.Row = .grdMain.Rows - 1
                        If TillData.Recipe = 1 And .grdMenu.Rows > 0 Then
                            For i = 0 To frmRecipe.grdMess.Rows - 1
                                .grdMain.Rows = .grdMain.Rows + 1
                                .grdMain.Row = .grdMain.Rows - 1
                                If MarkforPrinting > 0 Then
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 14) = "P"
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 11) = TillData.Kitchen1
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 12) = TillData.Kitchen2
                                    .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 14, .grdMain.Rows - 1, 14) = &HC0FFFF
                                End If
                                .grdMain.TextMatrix(.grdMain.Row, 1) = "    >" & frmRecipe.grdMess.TextMatrix(i, 1)
                                .grdMain.Cell(flexcpForeColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC00000
                                .grdMain.Cell(flexcpFontBold, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = True
                                'Product Code
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 9) = frmRecipe.grdMess.TextMatrix(i, 0)
                                If Val(frmRecipe.grdMess.TextMatrix(i, 0)) <> 0 Then
                                    ActiveReadServer "Select Department_No,Sales_Tax,Tax_Type,Ave_Cost, Selling_price,ISNULL(Receipe_Charge_Item,0) as Receipe_Charge_Item  from Products where Product_Code = '" & frmRecipe.grdMess.TextMatrix(i, 0) & "'"
                                    If rs.RecordCount > 0 Then
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = rs.Fields("Ave_Cost")
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = rs.Fields("Sales_Tax")
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 6) = rs.Fields("Tax_Type")
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = frmRecipe.grdMess.TextMatrix(i, 1)
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 10) = rs.Fields("Department_No")
                                        
                                        If rs.Fields("Receipe_Charge_Item") = 1 Then
                                            .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(rs.Fields("selling_price"), "0.00")
                                            If TillData.TaxRate <> 0 Then
                                                TillData.TaxableSales = TillData.TaxableSales + rs.Fields("selling_price")
                                                TillData.CollectedTax = TillData.CollectedTax + (rs.Fields("selling_price")) - ((rs.Fields("selling_price")) / ((100 + rs.Fields("Sales_Tax")) / 100))
                                            Else
                                                TillData.NonTaxableSales = TillData.NonTaxableSales + Val(rs.Fields("selling_price"))
                                            End If
                                        End If
                                            If InStr(KeyRegister, "Return Item") <> 0 Then
                                                .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(.grdMain.TextMatrix(.grdMain.Rows - 1, 2) * -1, "0.00")
                                            End If
                                        
                                    End If
                                    rs.Close
                                Else
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = frmRecipe.grdMess.TextMatrix(i, 1)
                                End If
                            Next i
                        End If
                        KeyRegister = ""
                        ' Kotie 26-03-2013
                        Sale_Total = 0
                         Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                        'For i = 1 To .grdMain.Rows - 1
                        '    If (.grdMain.TextMatrix(i, 8) <> "Corr") Then
                        '        If .grdMain.TextMatrix(i, 8) <> "Wastage" Then
                        '            If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                        '                Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                        '            End If
                        '        End If
                        '    End If
                        'Next i
                        'If Val(Swiss_Round) <> 0 Then
                        '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                        '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                        '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                        '            Sale_Total = Sale_Total - q
                        '        End If
                        '    End If
                        'End If
                        Sale_Total = TillData.SaleTotal
                        .lblTender.Caption = Format(Sale_Total, "0.00")
                        .lblCash.Caption = "Subtotal"
                        If .grdMain.Rows > 11 Then .grdMain.TopRow = .grdMain.Row - 11
                    Case KeyType.FinalizationKey
                        .picDigit.Visible = False
                        Sale_Total = 0
                        
                         Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                        'For i = 1 To .grdMain.Rows - 1
                        '    If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                        '        If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                        '           If .grdMain.TextMatrix(i, 3) <> "Short" Then
                        '                Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                        '            End If
                        '        End If
                        '    End If
                        'Next i
                        Sale_Total = TillData.SaleTotal
                        If TillData.Change < 0 Then
                            .lblTender.Caption = Format(TillData.Change * -1, "0.00")
                            .lblCash.Caption = "Short Tendered"
                        Else
                            If TillData.Change = 0 Then
                                If .lblCash = "Short Tendered" Then
                                    .lblCash.Caption = "Change"
                                    .lblTender.Caption = Format(TillData.Change, "0.00")
                                Else
                                    .lblCash.Caption = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "")
                                End If
                            Else
                                .lblCash.Caption = "Change"
                                .lblTender.Caption = Format(TillData.Change, "0.00")
                            End If
                        End If
                        .grdMain.Rows = .grdMain.Rows + 1
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 0) = ""
                        If KeyRegister = "<No Sale>" Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "No Sale"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = "No Sale"
                        Else
                            If TillData.Change < 0 Then
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "") & " Tendered"
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = "Short"
                            Else
                                If .grdMain.TextMatrix(.grdMain.Rows - 2, 3) = "Short" Then
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "") & " Tendered"
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = "Short"
                                Else
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "") & " Sale"
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "")
                                End If
                            End If
                        End If
                        If TillData.Change < 0 Then
                            KeyRegister = Replace(KeyRegister, "-", "")
                            If Left(KeyRegister, 1) = "R" Then
                                KeyRegister = Mid(KeyRegister, 2)
                            End If
                             .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Val(Mid(KeyRegister, 1, InStr(KeyRegister, "<") - 1)) / 100, "0.00")
                        Else
                            If .grdMain.TextMatrix(.grdMain.Rows - 2, 3) = "Short" Then
                                KeyRegister = Replace(KeyRegister, "-", "")
                                If Left(KeyRegister, 1) = "R" Then
                                    KeyRegister = Mid(KeyRegister, 2)
                                End If
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Val(Mid(KeyRegister, 1, InStr(KeyRegister, "<") - 1)) / 100, "0.00")
                            Else
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Sale_Total, "0.00")
                            End If
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = TillData.Tendered
                        If Right(.grdMain.TextMatrix(.grdMain.Rows - 1, 1), 8) = "Tendered" Then
                            If TillData.Change < 0 Then
                                Tax_Total = TillData.TaxTotal * (.grdMain.TextMatrix(.grdMain.Rows - 1, 2) / TillData.SaleTotal)
                            Else
                                Tax_Total = TillData.TaxTotal * ((.grdMain.TextMatrix(.grdMain.Rows - 1, 2) - TillData.Change) / TillData.SaleTotal)
                            End If
                        Else
                            Tax_Total = TillData.TaxTotal
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = Tax_Total
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 13) = TillData.Change
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = KeyRegister
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = TillData.ExtraFunc
                        .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0FFFF
                        .lblKeyRegister.TextAlign = fmTextAlignLeft
                        If KeyRegister = "<No Sale>" Then
                            .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - No Sale"
                        Else
                             If TillData.Change < 0 Then
                                .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - R" & Format(TillData.Tendered, "0.00") & " Tendered of R" & Format(Sale_Total, "0.00")
                            Else
                                If .grdMain.TextMatrix(.grdMain.Rows - 2, 3) = "Short" Then
                                    .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - Sale Total: R" & Format(Sale_Total, "0.00")
                                Else
                                    .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - R" & Format(TillData.Tendered, "0.00") & " Tendered"
                                End If
                            End If
                        End If
                        .cmdDept(6).Caption = "No Sale"
                        .grdMain.Row = .grdMain.Rows - 1
                        .grdMain.ShowCell .grdMain.Rows - 1, 0
                        .grdMain.HighLight = flexHighlightWithFocus
                        .picHoldFocus.SetFocus
                End Select
                
                
                
                If Devices.DisplayModel <> "" Then
                
                Select Case Devices.DisplayModel
                
                Case "CD7220", "Posiflex"
                
                'If Devices.DisplayModel = "CD7220" Then
                    filenum = FreeFile
                    Open Devices.DisplayPort For Output As filenum
                    Print #filenum, Chr(27) & "@"
                    Select Case KeyFunction
                        Case KeyType.InputKey
                            Print #filenum, Replace(String(20 - Len(Trim(Left(.lblKeyRegister.Caption, 20))), " ") & Trim(Left(.lblKeyRegister.Caption, 20)) & "SUBTOTAL" & UCase(String(20 - Len("SUBTOTAL") - Len(.lblTender.Caption), " ") & .lblTender.Caption), Chr(215), "X")
                        Case KeyType.ItemizerKey
                            DisplayLine = Replace(UCase(Trim(Mid(.lblKeyRegister.Caption, 1, InStr(.lblKeyRegister.Caption, "@") - 1))), "1 X ", "")
                            Print #filenum, Trim(Left(DisplayLine, 20)) & String(20 - Len(Trim(Left(DisplayLine, 20))), " ");
                            Print #filenum, UCase(Trim(.lblCash.Caption & String(20 - Len(.lblCash.Caption) - Len(.lblTender.Caption), " ") & .lblTender.Caption))
                        Case KeyType.FinalizationKey
                            DisplayLine = Replace(UCase(Trim(Mid(.lblKeyRegister.Caption, InStr(.lblKeyRegister.Caption, "-") + 1))), "1 X ", "")
                            Print #filenum, Trim(Left(DisplayLine, 20)) & String(20 - Len(Trim(Left(DisplayLine, 20))), " ");
                            Print #filenum, UCase(Trim(.lblCash.Caption & String(20 - Len(.lblCash.Caption) - Len(.lblTender.Caption), " ") & .lblTender.Caption))
                    End Select
                    Close #filenum
                    End Select
                End If
            
            
            
            
            
            End With
        Case 2
            With frmBar
                .grdMain.HighLight = flexHighlightAlways
                Select Case KeyFunction
                    Case KeyType.ClearKey
                        .picDigit.Visible = False
                        .lblKeyRegister.TextAlign = fmTextAlignLeft
                        .grdMain.Rows = 1
                        .lblTender.Caption = "0.00"
                        .lblCash.Caption = ""
                    Case KeyType.InputKey
                        If Asc(Right(KeyRegister, 1)) > 47 And Asc(Right(KeyRegister, 1)) < 58 Then
                            .picDigit.Visible = True
                            .lblDigit = Right(KeyRegister, 2)
                        Else
                            .picDigit.Visible = False
                        End If
                        .lblKeyRegister.TextAlign = fmTextAlignRight
                        .lblKeyRegister.Caption = KeyRegister
                    Case KeyType.FunctionKey
                        .picDigit.Visible = False
                        If KeyRegister = " (Corr) " Then
                            For i = .grdMain.Rows - 1 To 1 Step -1
                                If .grdMain.TextMatrix(i, 0) <> "" Then
                                    .grdMain.Rows = i + 1
                                    Exit For
                                End If
                            Next i
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = TillData.ExtraFunc
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            .grdMain.Cell(flexcpFontStrikethru, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = True
                            .lblKeyRegister.Caption = "Item Correct - " & TillData.Description & " @ R" & Format(TillData.Price, "0.00")
                            If .grdMain.TextMatrix(.grdMain.Rows - 1, 8) <> "Wastage" Then  'Kotie
                                TillData.TaxTotal = Round(TillData.TaxTotal, 3) - Round(((TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100))), 3)
                                If TillData.TaxRate <> 0 Then
                                    TillData.TaxableSales = TillData.TaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                    TillData.CollectedTax = TillData.CollectedTax - ((TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100)))
                                Else
                                    TillData.NonTaxableSales = TillData.NonTaxableSales - Val(TillData.Price) * Val(TillData.Qty)
                                End If
                            End If
                            Sale_Total = 0
                             Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                            'For i = 1 To .grdMain.Rows - 1
                            '    If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            '        If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                            '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                            '        End If
                            '    End If
                            'Next i
                            'If Val(Swiss_Round) <> 0 Then
                            '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                            '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                            '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                            '            Sale_Total = Sale_Total - q
                            '        End If
                            '    End If
                            'End If
                            Sale_Total = TillData.SaleTotal
                            .lblTender.Caption = Format(Sale_Total, "0.00")
                            .lblCash.Caption = "Subtotal"
                            If .grdMain.Rows > 10 Then .grdMain.TopRow = .grdMain.Row - 10
                            On Error GoTo 0
                            Exit Sub
                        End If
                        If KeyRegister = " (Subtotal) " Then
                            Sale_Total = 0
                            
                             Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                           ' For i = 1 To .grdMain.Rows - 1
                           '     If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                           '         If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                           '             Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                           '         End If
                           '     End If
                           ' Next i
                           ' If Val(Swiss_Round) <> 0 Then
                           '     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                           '         If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                           '             q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                           '             Sale_Total = Sale_Total - q
                           '             End If
                           '     End If
                           ' End If
                            Sale_Total = TillData.SaleTotal
                            .grdMain.Rows = .grdMain.Rows + 1
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 0) = ""
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Subtotal"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Sale_Total, "0.00")
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = TillData.Keystring
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = ""
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = ""
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = KeyRegister
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = ""
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0FFC0
                            .lblKeyRegister.TextAlign = fmTextAlignLeft
                            .lblKeyRegister.Caption = "Subtotal"
                            .lblTender.Caption = Format(Sale_Total, "0.00")
                            .lblCash.Caption = "Subtotal"
                            KeyRegister = ""
                            .grdMain.Row = .grdMain.Rows - 1
                            If .grdMain.Rows > 10 Then .grdMain.TopRow = .grdMain.Row - 10
                        End If
                        .lblKeyRegister.TextAlign = fmTextAlignRight
                        .lblKeyRegister.Caption = KeyRegister
                    Case KeyType.ItemizerKey
                        MarkforPrinting = 0
                        .picDigit.Visible = False
                        .lblKeyRegister.TextAlign = fmTextAlignLeft
                        If TillData.Keystring = "Plu" Or TillData.Keystring = "Dept" Then .cmdKey(4).Caption = "Corr"
                        If TillData.Keystring = "Dept" Then
                            .lblKeyRegister.Caption = "Department - " & TillData.Qty & " x " & TillData.Description & " @ R" & Format(TillData.Price, "0.00")
                        Else
                            .lblKeyRegister.Caption = TillData.Qty & " x " & TillData.Description & " @ R" & Format(TillData.Price, "0.00")
                        End If
                        .grdMain.Rows = .grdMain.Rows + 1
                        If TillData.Kitchen1 <> "" And TillData.Kitchen1 <> "<None>" Then
                            .grdMain.ColHidden(14) = False
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 14, .grdMain.Rows - 1, 14) = &HC0FFFF
                            MarkforPrinting = .grdMain.Rows - 1
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 14) = "P"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 11) = TillData.Kitchen1
                        End If
                        If TillData.Kitchen2 <> "" And TillData.Kitchen2 <> "<None>" Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 12) = TillData.Kitchen2
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 0) = TillData.Qty
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = TillData.ShortDesc
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Val(TillData.Price) * Val(TillData.Qty), "0.00")
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = TillData.Keystring
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = TillData.Cost
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = TillData.TaxRate
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 6) = TillData.TaxType
                      '  If TillData.TaxRate <> 0 Then
                      '      TillData.TaxableSales = TillData.TaxableSales + Val(TillData.Price) * Val(TillData.Qty)
                      '      TillData.CollectedTax = TillData.CollectedTax + (TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100))
                      '  Else
                      '      TillData.NonTaxableSales = TillData.NonTaxableSales + Val(TillData.Price) * Val(TillData.Qty)
                      '  End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = KeyRegister
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 9) = TillData.ProductCode
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 10) = TillData.DeptNo
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 13) = TillData.PriceOveride
                        If InStr(KeyRegister, "Void") <> 0 Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Void - " & .grdMain.TextMatrix(.grdMain.Rows - 1, 1)
                            .lblKeyRegister.Caption = "(Void) " & .lblKeyRegister.Caption
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            TillData.VoidTotal = TillData.VoidTotal + (Val(TillData.Price) * Val(TillData.Qty))
                            TillData.VoidCount = TillData.VoidCount + 1
                            TillData.ExtraFunc = "Void"
                            .grdMain.Cell(flexcpForeColor, .grdMain.Row, 0, .grdMain.Row, 2) = vbRed
                            .grdMain.TextMatrix(.grdMain.Row, 8) = "Voided on - " & .grdMain.Rows - 1
                        End If
                        If InStr(KeyRegister, "Return Item") <> 0 Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Return Item - " & .grdMain.TextMatrix(.grdMain.Rows - 1, 1)
                            .lblKeyRegister.Caption = "(Return Item) " & .lblKeyRegister.Caption
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            TillData.ReturnTotal = TillData.ReturnTotal + Val(TillData.Price) * Val(TillData.Qty)
                            TillData.ReturnCount = TillData.ReturnCount + 1
                            TillData.ExtraFunc = "Return Item"
                        End If
                        If InStr(KeyRegister, "Wastage") <> 0 Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "Wastage - " & .grdMain.TextMatrix(.grdMain.Rows - 1, 1)
                            .lblKeyRegister.Caption = "(Wastage) " & .lblKeyRegister.Caption
                            .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0C0FF
                            TillData.UllageTotal = TillData.UllageTotal + Val(TillData.Price) * Val(TillData.Qty)
                            TillData.UllageCount = TillData.UllageCount + 1
                            TillData.ExtraFunc = "Wastage"
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = TillData.ExtraFunc
                        
                        If TillData.Keystring <> "Service Charge" Then
                            If .grdMain.TextMatrix(.grdMain.Rows - 1, 8) <> "Wastage" Then  'Kotie
                                If TillData.TaxRate <> 0 Then
                                    TillData.TaxableSales = TillData.TaxableSales + Val(TillData.Price) * Val(TillData.Qty)
                                    TillData.CollectedTax = TillData.CollectedTax + (TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100))
                                Else
                                    TillData.NonTaxableSales = TillData.NonTaxableSales + Val(TillData.Price) * Val(TillData.Qty)
                                End If
                            End If
                        End If
                        '****************************************************
                        'Moved to after repcipe check
                        'Sale_Total = 0
                        'For i = 1 To .grdMain.Rows - 1
                        '    If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                        '        If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                        '            Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                        '        End If
                        '    End If
                        'Next i
                        'If Val(Swiss_Round) <> 0 Then
                        '    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                        '        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                        '            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                        '            Sale_Total = Sale_Total - q
                        '        End If
                        '    End If
                        'End If
                        'TillData.SaleTotal = Sale_Total
                        '.lblTender.Caption = Format(Sale_Total, "0.00")
                        '.lblCash.Caption = "Subtotal"
                        '****************************************************
                        
                        .grdMain.Row = .grdMain.Rows - 1
                        If TillData.Recipe = 1 And .grdMenu.Rows > 0 Then
                            For i = 0 To frmRecipe.grdMess.Rows - 1
                                .grdMain.Rows = .grdMain.Rows + 1
                                .grdMain.Row = .grdMain.Rows - 1
                                If MarkforPrinting > 0 Then
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 14) = "P"
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 11) = TillData.Kitchen1
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 12) = TillData.Kitchen2
                                    .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 14, .grdMain.Rows - 1, 14) = &HC0FFFF
                                End If
                                .grdMain.TextMatrix(.grdMain.Row, 1) = "    >" & frmRecipe.grdMess.TextMatrix(i, 1)
                                .grdMain.Cell(flexcpForeColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC00000
                                .grdMain.Cell(flexcpFontBold, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = True
                                'Product Code
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 9) = frmRecipe.grdMess.TextMatrix(i, 0)
                                If Val(frmRecipe.grdMess.TextMatrix(i, 0)) <> 0 Then
                                    ActiveReadServer "Select Department_No,Sales_Tax,Tax_Type,Ave_Cost, selling_price, ISNULL(Receipe_Charge_Item,0) as Receipe_Charge_Item  from Products where Product_Code = '" & frmRecipe.grdMess.TextMatrix(i, 0) & "'"
                                    If rs.RecordCount > 0 Then

                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = rs.Fields("Ave_Cost")
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = rs.Fields("Sales_Tax")
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 6) = rs.Fields("Tax_Type")
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = frmRecipe.grdMess.TextMatrix(i, 1)
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 10) = rs.Fields("Department_No")
                                        
                                        
                                            If rs.Fields("Receipe_Charge_Item") = 1 Then
                                                .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(rs.Fields("selling_price"), "0.00")
                                                If TillData.TaxRate <> 0 Then
                                                    TillData.TaxableSales = TillData.TaxableSales + rs.Fields("selling_price")
                                                    TillData.CollectedTax = TillData.CollectedTax + (rs.Fields("selling_price")) - ((rs.Fields("selling_price")) / ((100 + rs.Fields("Sales_Tax")) / 100))
                                                Else
                                                    TillData.NonTaxableSales = TillData.NonTaxableSales + Val(rs.Fields("selling_price"))
                                                End If
                                            End If
                                            If InStr(KeyRegister, "Return Item") <> 0 Then
                                                .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(.grdMain.TextMatrix(.grdMain.Rows - 1, 2) * -1, "0.00")
                                                .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 14) = &HC0C0FF
                                            End If
                                      
                                        
                                    End If
                                    rs.Close
                                Else
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = frmRecipe.grdMess.TextMatrix(i, 1)
                                End If
                            Next i
                        End If
                        KeyRegister = ""
                        'Kotie 25-03-2013  Moved from top to here
                        Sale_Total = 0
                        'For i = 1 To .grdMain.Rows - 1
                        '    If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                        '        If .grdMain.TextMatrix(i, 8) <> "Wastage" Then  'Kotie
                        '            If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                        '                sale_total = Val(sale_total) + Val(.grdMain.TextMatrix(i, 2))
                        '            End If
                        '        End If
                        '    End If
                        'Next i
                        'If Val(Swiss_Round) <> 0 Then
                        '    If Str(sale_total / Swiss_Round) <> Str(Round(sale_total / Swiss_Round, 0)) Then
                        '        If Right(Swiss_Round, 1) <> Right(sale_total, 1) Then
                        '            q = Round((sale_total / Swiss_Round - Int(sale_total / Swiss_Round)) * Swiss_Round, 2)
                        '            sale_total = sale_total - q
                        '        End If
                        '    End If
                        'End If
                        On Error GoTo 0
                         Update_Sale_Total (Panel_no) 'Kotie 10/04/2013

                        If .grdMain.Rows > 10 Then .grdMain.TopRow = .grdMain.Row - 10
                    Case KeyType.FinalizationKey
                        .picSlip.Visible = True
                        .picDigit.Visible = False
                        Sale_Total = 0
                       ' For i = 1 To .grdMain.Rows - 1
                       '     If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                       '         If .grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                       '             If .grdMain.TextMatrix(i, 3) <> "Short" Then
                       '                 sale_total = Val(sale_total) + Val(.grdMain.TextMatrix(i, 2))
                       '             End If
                       '         End If
                       '     End If
                       ' Next i
                       ' TillData.SaleTotal = sale_total
                         Update_Sale_Total (Panel_no) 'Kotie 10/04/2013
                        Sale_Total = TillData.SaleTotal
                        If TillData.Change < 0 Then
                            .lblTender.Caption = Format(TillData.Change * -1, "0.00")
                            .lblCash.Caption = "Short Tendered"
                        Else
                            If TillData.Change = 0 Then
                                If .lblCash = "Short Tendered" Then
                                    .lblCash.Caption = "Change"
                                    .lblTender.Caption = Format(TillData.Change, "0.00")
                                Else
                                    .lblCash.Caption = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "")
                                End If
                            Else
                                .lblCash.Caption = "Change"
                                .lblTender.Caption = Format(TillData.Change, "0.00")
                            End If
                        End If
                        .grdMain.Rows = .grdMain.Rows + 1
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 0) = ""
                        If KeyRegister = "<No Sale>" Then
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = "No Sale"
                            .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = "No Sale"
                        Else
                            If TillData.Change < 0 Then
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "") & " Tendered"
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = "Short"
                            Else
                                If .grdMain.TextMatrix(.grdMain.Rows - 2, 3) = "Short" Then
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "") & " Tendered"
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = "Short"
                                Else
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 1) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "") & " Sale"
                                    .grdMain.TextMatrix(.grdMain.Rows - 1, 3) = Replace(Mid(KeyRegister, InStr(KeyRegister, "<") + 1), ">", "")
                                End If
                            End If
                        End If
                        If TillData.Change < 0 Then
                            KeyRegister = Replace(KeyRegister, "-", "")
                            If Left(KeyRegister, 1) = "R" Then
                                KeyRegister = Mid(KeyRegister, 2)
                            End If
                             .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Val(Mid(KeyRegister, 1, InStr(KeyRegister, "<") - 1)) / 100, "0.00")
                        Else
                            If .grdMain.TextMatrix(.grdMain.Rows - 2, 3) = "Short" Then
                                KeyRegister = Replace(KeyRegister, "-", "")
                                If Left(KeyRegister, 1) = "R" Then
                                    KeyRegister = Mid(KeyRegister, 2)
                                End If
                                Select Case Mid(KeyRegister, InStr(KeyRegister, "<"))
                                    Case "<Cash>"
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(TillData.Cash, "0.00")
                                    Case "<Card>"
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(TillData.Card, "0.00")
                                    Case "<Voucher>"
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(TillData.Cheque, "0.00")
                                    Case "<Charge>"
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(TillData.Charge, "0.00")
                                    Case "<Loyalty>"
                                        .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(TillData.Loyalty, "0.00")
                                End Select
                            Else
                                .grdMain.TextMatrix(.grdMain.Rows - 1, 2) = Format(Sale_Total, "0.00")
                            End If
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 4) = TillData.Tendered
                        If Right(.grdMain.TextMatrix(.grdMain.Rows - 1, 1), 8) = "Tendered" Then
                            If TillData.Change < 0 Then
                                Tax_Total = TillData.TaxTotal * (.grdMain.TextMatrix(.grdMain.Rows - 1, 2) / TillData.SaleTotal)
                            Else
                                Tax_Total = TillData.TaxTotal * ((.grdMain.TextMatrix(.grdMain.Rows - 1, 2) - TillData.Change) / TillData.SaleTotal)
                            End If
                        Else
                            Tax_Total = TillData.TaxTotal
                        End If
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 5) = Tax_Total
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 13) = TillData.Change
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 7) = KeyRegister
                        .grdMain.TextMatrix(.grdMain.Rows - 1, 8) = TillData.ExtraFunc
                        .grdMain.Cell(flexcpBackColor, .grdMain.Rows - 1, 0, .grdMain.Rows - 1, 2) = &HC0FFFF
                        .lblKeyRegister.TextAlign = fmTextAlignLeft
                        If KeyRegister = "<No Sale>" Then
                            .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - No Sale"
                        Else
                             If TillData.Change < 0 Then
                                .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - R" & Format(TillData.Tendered, "0.00") & " Tendered of R" & Format(Sale_Total, "0.00")
                            Else
                                If .grdMain.TextMatrix(.grdMain.Rows - 2, 3) = "Short" Then
                                    .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - Sale Total: R" & Format(Sale_Total, "0.00")
                                Else
                                    .lblKeyRegister.Caption = "Invoice No: " & TillData.DocNo & " - R" & Format(TillData.Tendered, "0.00") & " Tendered"
                                End If
                            End If
                        End If
                        .cmdKey(4).Caption = "No Sale"
                        .grdMain.Row = .grdMain.Rows - 1
                        .grdMain.ShowCell .grdMain.Rows - 1, 0
                        .grdMain.HighLight = flexHighlightWithFocus
                        .cmdSlip.Caption = "Close Slip"
                        .cmdFancy(3).Caption = "Create Tab"
                        .picHoldFocus.SetFocus
                End Select
                
                
                
                
                If Devices.DisplayModel <> "" Then
                
                Select Case Devices.DisplayModel
                
                Case "CD7220", "Posiflex"
                
                'If Devices.DisplayModel = "CD7220" Then
                    filenum = FreeFile
                    Open Devices.DisplayPort For Output As filenum
                    Print #filenum, Chr(27) & "@"
                    Select Case KeyFunction
                        Case KeyType.InputKey
                            Print #filenum, Replace(String(20 - Len(Trim(Left(.lblKeyRegister.Caption, 20))), " ") & Trim(Left(.lblKeyRegister.Caption, 20)) & "SUBTOTAL" & UCase(String(20 - Len("SUBTOTAL") - Len(.lblTender.Caption), " ") & .lblTender.Caption), Chr(215), "X")
                        Case KeyType.ItemizerKey
                            DisplayLine = Replace(UCase(Trim(Mid(.lblKeyRegister.Caption, 1, InStr(.lblKeyRegister.Caption, "@") - 1))), "1 X ", "")
                            Print #filenum, Trim(Left(DisplayLine, 20)) & String(20 - Len(Trim(Left(DisplayLine, 20))), " ");
                            Print #filenum, UCase(Trim(.lblCash.Caption & String(20 - Len(.lblCash.Caption) - Len(.lblTender.Caption), " ") & .lblTender.Caption))
                        Case KeyType.FinalizationKey
                            DisplayLine = Replace(UCase(Trim(Mid(.lblKeyRegister.Caption, InStr(.lblKeyRegister.Caption, "-") + 1))), "1 X ", "")
                            Print #filenum, Trim(Left(DisplayLine, 20)) & String(20 - Len(Trim(Left(DisplayLine, 20))), " ");
                            Print #filenum, UCase(Trim(.lblCash.Caption & String(20 - Len(.lblCash.Caption) - Len(.lblTender.Caption), " ") & .lblTender.Caption))
                    End Select
                    Close #filenum
                    End Select
                End If
            End With
    End Select
    On Error GoTo 0
End Sub

Private Function Update_Sale_Total(Parent_form) As Double
    Sale_Total = 0
    
    Select Case Parent_form
        Case 0
            With frmSales
                 For i = 1 To .grdMain.Rows - 1
                     If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                        If .grdMain.TextMatrix(i, 8) <> "Wastage" Then  'Kotie
                            If (.grdMain.TextMatrix(i, 3) <> "Subtotal") And (.grdMain.TextMatrix(i, 3) <> "Short") Then
                                Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                            End If
                         End If
                     End If
                 Next i
                 If Val(Swiss_Round) <> 0 Then
                     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                         If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                             q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                             Sale_Total = Sale_Total - q
                         End If
                     End If
                 End If
            End With
            
        Case 1
            With frmSales1
                 For i = 1 To .grdMain.Rows - 1
                     If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                        If .grdMain.TextMatrix(i, 8) <> "Wastage" Then  'Kotie
                            If (.grdMain.TextMatrix(i, 3) <> "Subtotal") And (.grdMain.TextMatrix(i, 3) <> "Short") Then
                                Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                            End If
                         End If
                     End If
                 Next i
                 If Val(Swiss_Round) <> 0 Then
                     If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                         If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                             q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                             Sale_Total = Sale_Total - q
                         End If
                     End If
                 End If
            End With
            
        Case 2
            With frmBar
                For i = 1 To .grdMain.Rows - 1
                    If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                        If .grdMain.TextMatrix(i, 8) <> "Wastage" Then  'Kotie
                            If (.grdMain.TextMatrix(i, 3) <> "Subtotal") And (.grdMain.TextMatrix(i, 3) <> "Short") Then
                                Sale_Total = Val(Sale_Total) + Val(.grdMain.TextMatrix(i, 2))
                            End If
                        End If
                    End If
                Next i
                If Val(Swiss_Round) <> 0 Then
                    If Str(Sale_Total / Swiss_Round) <> Str(Round(Sale_Total / Swiss_Round, 0)) Then
                        If Right(Swiss_Round, 1) <> Right(Sale_Total, 1) Then
                            q = Round((Sale_Total / Swiss_Round - Int(Sale_Total / Swiss_Round)) * Swiss_Round, 2)
                            Sale_Total = Sale_Total - q
                        End If
                    End If
                End If
                .lblTender.Caption = Format(Sale_Total, "0.00")
                .lblCash.Caption = "Subtotal"
            End With
    End Select
    TillData.SaleTotal = Sale_Total
End Function
Private Function Validatator(Keystring$)
    Validatator = True
    Select Case Keystring
        Case "Kitchen Message"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If KeyRegister <> "" Then Validatator = False
            If InStr(KeyRegister, "Void") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Return Item") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Wastage") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Corr") <> 0 Then Validatator = False
            If TillData.ExtraFunc = "Corr" Then Validatator = False
            If TillData.ExtraFunc = "Void" Then Validatator = False
            If TillData.ExtraFunc = "Return Item" Then Validatator = False
            If TillData.ExtraFunc = "Wastage" Then Validatator = False
            If TillData.DocNo = 0 Then Validatator = False
        Case "P/R"
            If KeyRegister <> "" Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If InStr(KeyRegister, "Void") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Return Item") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Wastage") <> 0 Then Validatator = False
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
        Case "R/A"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If TillData.DocNo <> 0 Then Validatator = False
        Case "Pay Out"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If TillData.DocNo <> 0 Then Validatator = False
            If Val(frmSales.lblKeyRegister) = 0 Then Validatator = False
            If UserRecord.Payouts = False Then
                TillData.UserOveride = 0
                Load frmValidate
                frmValidate.Tag = Keystring
                frmValidate.Show vbModal
                If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                    Validatator = False
                Else
                    frmValidate.Tag = ""
                End If
            End If
        Case "Discount"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If TillData.DocNo = 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If Right(KeyRegister, 2) = ") " Then Validatator = False
            If InStr(KeyRegister, Chr(215)) <> 0 Then Validatator = False
            If InStr(KeyRegister, "Price O/V") <> 0 Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
            If TillData.SaleTotal = 0 Then Validatator = False
            If InStr(KeyRegister, "Void") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Return Item") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Wastage") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Corr") <> 0 Then Validatator = False
            'If TillData.ExtraFunc = "Corr" Then Validatator = False
            If TillData.ExtraFunc = "Wastage" Then Validatator = False
            If UserRecord.Disc_Amt = False Or UserRecord.Disc_Perc = False Then
                TillData.UserOveride = 0
                Load frmValidate
                frmValidate.Tag = Keystring
                frmValidate.Show vbModal
                If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                    Validatator = False
                Else
                    frmValidate.Tag = ""
                End If
            End If
        Case "X 2", "X 3", "X 4", "X 5", "X 6", "X 10", "X 12", "X 20", "X 30"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If KeyRegister <> "" Then Validatator = False
        Case "Print Bill"
            If TillData.ShortTender = True Then Validatator = False
            If TillData.TableNo = 0 And TillData.TabNo = 0 Then Validatator = False
            If TillData.DocNo = 0 Then Validatator = False
                '****************** Kotie 20-03-2013
                If TillData.Print_Count > 0 Then
                    If UserRecord.Reprint = False Then
                        TillData.UserOveride = 0
                        Load frmValidate
                        frmValidate.Tag = Keystring
                        frmValidate.Show vbModal
                        If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                            Validatator = False
                        Else
                            frmValidate.Tag = ""
                        End If
                    End If
                End If
                '******************
            Select Case Panel_no

                Case 0
                    If TillData.TableNo = 0 Then Validatator = False
                    If frmSales.grdMain.Rows = 1 Then Validatator = False
                Case 1
                    If TillData.TableNo = 0 Then Validatator = False
                    If frmSales1.grdMain.Rows = 1 Then Validatator = False
                Case 2
                    If TillData.TabNo = 0 Then Validatator = False
                    If frmBar.grdMain.Rows = 1 Then Validatator = False
            End Select
        Case "Reprint"
            'If LastTab = 0 Then Validatator = False
            'If TillData.TableNo = 0 And TillData.TabNo = 0 Then Validatator = False
            If TillData.DocNo <> 0 Then Validatator = False
            If TillData.Prev_Doc_No = 0 Then Validatator = False
                '****************** Kotie 20-03-2013
                If Validatator = True Then
                    If UserRecord.Reprint = False Then
                        TillData.UserOveride = 0
                        Load frmValidate
                        frmValidate.Tag = Keystring
                        frmValidate.Show vbModal
                        If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                            Validatator = False
                        Else
                            frmValidate.Tag = ""
                        End If
                    End If
                End If
                '******************
            Select Case Panel_no
                Case 0
                    If frmSales.grdMain.Rows = 1 Then Validatator = False
                Case 2
                    If frmBar.grdMain.Rows = 1 Then Validatator = False
            End Select
        Case "Close Tab"
            If Panel_no = 2 Then
                If TillData.TabNo = 0 Then Validatator = False
            End If
        Case "Pickup Tab"
            If TillData.TabNo <> 0 Then Validatator = False
            If TillData.DocNo <> 0 Then Validatator = False
        Case "Create Tab"
            If TillData.TabNo = 0 Then
                If frmBar.grdMain.Rows = 1 Then Validatator = False
                If GlobalMode = TillMode.FinMode Then Validatator = False
                If TillData.ShortTender = True Then Validatator = False
            End If
            If TillData.ShortTender = True Then Validatator = False
            Select Case UserRecord.uType
                Case 4, 8
                Case Else
                    Validatator = False
            End Select
        Case "Service Charge"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If System_Service = 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            Select Case Panel_no
                Case 1
                    If TillData.TableNo = 0 Then Validatator = False
                    If frmSales1.grdMain.Rows = 1 Then Validatator = False
                Case 2
                    If TillData.TabNo = 0 Then Validatator = False
                    If frmBar.grdMain.Rows = 1 Then Validatator = False
            End Select
            ActiveReadServer "Select User_No from Users where User_No= '" & UserRecord.User_Number & "' and Service_Charge =" & "'1'"
                    If rs.RecordCount > 0 Then
                    UserRecord.Service_Charge = True
                    Else
                    UserRecord.Service_Charge = False
                    End If
                    
            If UserRecord.Service_Charge = False Then
                TillData.UserOveride = 0
                Load frmValidate
                frmValidate.Tag = Keystring
                frmValidate.Show vbModal
                If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                    Validatator = False
                Else
                    frmValidate.Tag = ""
                    Validatator = True
                End If
            End If
        
        
        If UserRecord.Service_Charge = True Then
                TillData.UserOveride = 1
                Load frmTipp
                frmValidate.Tag = Keystring
                
                    Validatator = True
            
            End If
        
        
        
        
        
        
        Case "Transfer Items"
            If TillData.Tipp <> 0 Then Validatator = False
            Select Case Panel_no
                Case 1
                    If TillData.TableNo = 0 Then Validatator = False
                    If frmSales1.grdMain.Rows = 1 Then Validatator = False
                Case 2
                    If TillData.TabNo = 0 Then Validatator = False
                    If frmBar.grdMain.Rows = 1 Then Validatator = False
            End Select
            If UserRecord.Transfers = False Then
                TillData.UserOveride = 0
                Load frmValidate
                frmValidate.Tag = Keystring
                frmValidate.Show vbModal
                If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                    Validatator = False
                Else
                    frmValidate.Tag = ""
                End If
            End If
        Case "Print Bill"
            If TillData.TableNo = 0 And TillData.TabNo = 0 Then Validatator = False
            Select Case Panel_no
                Case 1
                    If frmSales1.grdMain.Rows = 1 Then Validatator = False
                Case 2
                    If frmBar.grdMain.Rows = 1 Then Validatator = False
            End Select
        Case "Split Bill"
            If TillData.ShortTender = True Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            Select Case Panel_no
                Case 1
                    If TillData.TableNo <> Int(TillData.TableNo) Then Validatator = False
                    If TillData.TableNo = 0 Then Validatator = False
                    If frmSales1.grdMain.Rows = 1 Then Validatator = False
                Case 2
                    If TillData.TabNo <> Int(TillData.TabNo) Then Validatator = False
                    If TillData.TabNo = 0 Then Validatator = False
                    If frmBar.grdMain.Rows = 1 Then Validatator = False
            End Select
        Case "Transfer Table"
            If TillData.Tipp <> 0 Then Validatator = False
            If TillData.TableNo = 0 Then Validatator = False
            If frmSales1.grdMain.Rows = 1 Then Validatator = False
        Case "Transfer Tab"
            If TillData.Tipp <> 0 Then Validatator = False
            If TillData.TabNo = 0 Then Validatator = False
            If frmBar.grdMain.Rows = 1 Then Validatator = False
        Case "View Tables"
            If TillData.TableNo > 0 Then Validatator = False
            If TillData.DocNo <> 0 Then Validatator = False
        Case "Place Order"
            If frmSales1.grdMain.Rows = 1 Then Validatator = False
            If TillData.TableNo = 0 Then Validatator = False
            If TillData.ShortTender = True Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
        Case "Send to Tab"
            If TillData.Tipp <> 0 Then Validatator = False
            If frmBar.grdMain.Rows = 1 Then Validatator = False
            If TillData.TabNo = 0 Then Validatator = False
            If TillData.ShortTender = True Then Validatator = False
        Case "New Tab"
            If TillData.Tipp <> 0 Then Validatator = False
            If TillData.TabNo <> 0 Then Validatator = False
            If TillData.ShortTender = True Then Validatator = False
        Case "New Table"
            If TillData.TableNo <> 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If TillData.ShortTender = True Then Validatator = False
        Case "x"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If KeyRegister <> "" Then
                If InStr(KeyRegister, "*") <> 0 Then Validatator = False
            End If
            If KeyRegister = "" Or KeyRegister = " (Return Item) " Then
                If KeyRegister <> " (Return Item) " Then
                    KeyRegister = "*"
                End If
                Keystring = "*"
            End If
            If Right(KeyRegister, 2) = ") " Then
                If KeyRegister <> " (Return Item) " Then
                    Validatator = False
                End If
            End If
            If InStr(KeyRegister, Chr(215)) <> 0 Then
                If Right(KeyRegister, 2) <> Chr(215) & " " Then
                    Validatator = False
                End If
            End If
            If InStr(KeyRegister, "Price O/V") <> 0 Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
        Case "00"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If KeyRegister = "" Then Validatator = False
            If Right(KeyRegister, 1) = " " Then Validatator = False
            If Right(KeyRegister, 2) = ") " Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
        Case "0" To "9"
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
        Case " "
            If InStr(KeyRegister, "*") = 0 Then Validatator = False
        Case "."
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If InStr(KeyRegister, ".") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Price O/V") <> 0 Then Validatator = False
            If InStr(KeyRegister, Chr(215)) <> 0 Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
        Case "Plu"
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If Right(KeyRegister, 2) = ") " Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
        Case "Dept"
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If Right(KeyRegister, 2) = ") " Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
        Case "Corr"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If KeyRegister <> "" Then Validatator = False
            If InStr(KeyRegister, "Void") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Return Item") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Wastage") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Corr") <> 0 Then Validatator = False
            If TillData.ExtraFunc = "Corr" Then Validatator = False
            If TillData.ExtraFunc = "Void" Then Validatator = False
            If TillData.ExtraFunc = "Return Item" Then Validatator = False
            If TillData.ExtraFunc = "Wastage" Then Validatator = False
            If TillData.Keystring <> "Plu" And TillData.Keystring <> "Dept" Then
                If TillData.Keystring <> "Service Charge" Then
                    Validatator = False
                End If
            End If
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
            If TillData.DocNo = 0 Then Validatator = False
            If UserRecord.Item_Corrects = False Then
                TillData.UserOveride = 0
                Load frmValidate
                frmValidate.Tag = Keystring
                frmValidate.Show vbModal
                If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                    Validatator = False
                Else
                    frmValidate.Tag = ""
                End If
            End If
            Select Case Panel_no
                Case 0
                    If frmSales.grdMain.TextMatrix(frmSales.grdMain.Rows - 1, 14) <> "" Then
                        If Asc(frmSales.grdMain.TextMatrix(frmSales.grdMain.Rows - 1, 14)) = 187 Then Validatator = False
                    End If
                Case 1
                    If frmSales1.grdMain.TextMatrix(frmSales1.grdMain.Rows - 1, 14) <> "" Then
                        If Asc(frmSales1.grdMain.TextMatrix(frmSales1.grdMain.Rows - 1, 14)) = 187 Then Validatator = False
                    End If
                Case 2
                    If frmBar.grdMain.TextMatrix(frmBar.grdMain.Rows - 1, 14) <> "" Then
                        If Asc(frmBar.grdMain.TextMatrix(frmBar.grdMain.Rows - 1, 14)) = 187 Then Validatator = False
                    End If
            End Select
        Case "Void"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If InStr(KeyRegister, "Void") <> 0 Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Return Item") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Wastage") <> 0 Then Validatator = False
            If TillData.SaleTotal = 0 Then Validatator = False
            If TillData.DocNo = 0 Then Validatator = False
            If InStr(KeyRegister, Chr(215)) <> 0 Then Validatator = False
            If UserRecord.Voids = False Then
                TillData.UserOveride = 0
                Load frmValidate
                frmValidate.Tag = Keystring
                frmValidate.Show vbModal
                If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                    Validatator = False
                Else
                    frmValidate.Tag = ""
                End If
            End If
        Case "Return Item"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If InStr(KeyRegister, "Void") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Return Item") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Wastage") <> 0 Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
            If UserRecord.Returns = False Then
                TillData.UserOveride = 0
                Load frmValidate
                frmValidate.Tag = Keystring
                frmValidate.Show vbModal
                If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                    Validatator = False
                Else
                    frmValidate.Tag = ""
                End If
            End If
        Case "Wastage"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If InStr(KeyRegister, "Void") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Return Item") <> 0 Then Validatator = False
            If InStr(KeyRegister, "Wastage") <> 0 Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
            If UserRecord.Ullages = False Then
                TillData.UserOveride = 0
                Load frmValidate
                frmValidate.Tag = Keystring
                frmValidate.Show vbModal
                If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                    Validatator = False
                Else
                    frmValidate.Tag = ""
                End If
            End If
        Case "Price O/V"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If KeyRegister = "" Then Validatator = False
            If InStr(KeyRegister, "Price O/V") <> 0 Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
            If Right(KeyRegister, 2) = ") " Then Validatator = False
            If UserRecord.Overides = False Then
                TillData.UserOveride = 0
                Load frmValidate
                frmValidate.Tag = Keystring
                frmValidate.Show vbModal
                If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                    Validatator = False
                Else
                    frmValidate.Tag = ""
                End If
            End If
        Case "Subtotal"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If KeyRegister <> "" Then Validatator = False
            If TillData.DocNo = 0 Then Validatator = False
            If InStr(KeyRegister, "-") <> 0 Then Validatator = False
            Select Case Panel_no
                Case 0: If frmSales.grdMain.TextMatrix(frmSales.grdMain.Rows - 1, 1) = "Subtotal" Then Validatator = False
                Case 1: If frmSales1.grdMain.TextMatrix(frmSales1.grdMain.Rows - 1, 1) = "Subtotal" Then Validatator = False
                Case 2: If frmBar.grdMain.TextMatrix(frmBar.grdMain.Rows - 1, 1) = "Subtotal" Then Validatator = False
            End Select
        Case "Cash", "Card", "Voucher", "Charge", "Quote"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If Keystring = "Quote" Then
'                If UserRecord.uType = 3 Or UserRecord.uType = 4 Then
'                    Validatator = False
'                End If
            End If
            If TillData.DocNo = 0 Then Validatator = False
            If Right(KeyRegister, 2) = ") " Then Validatator = False
            If InStr(KeyRegister, ".") <> 0 Then Validatator = False
            If InStr(KeyRegister, Chr(215)) <> 0 Then Validatator = False
            If TillData.TableNo <> 0 Then
                ActiveReadServer "Select * from Table_Listing_View where Table_No = " & TillData.TableNo
                If rs.RecordCount > 0 Then
                    If UserRecord.User_Number <> rs.Fields("User_No") Then
                        If UserRecord.uType = 3 Or UserRecord.uType = 4 Then
                            If UserRecord.Bar_Cash = 0 Then
                                Validatator = False
                            End If
                        End If
                    End If
                End If
                rs.Close
            End If
            If UserRecord.Cash_Sales = False And Keystring = "Cash" Then
                Validatator = False
            End If
            If UserRecord.Card_Sales = False And Keystring = "Card" Then
                Validatator = False
            End If
            If UserRecord.Cheque_Sales = False And Keystring = "Voucher" Then
                Validatator = False
            End If
            If UserRecord.Charge_Sales = False And Keystring = "Charge" Then
                If TillData.Account_No = "" Then
                    Validatator = False
                End If
            End If
            If UserRecord.Loyalty_Sales = False And Keystring = "Loyalty" Then
                Validatator = False
            End If
            If Keystring = "Charge" Then
                If Val(KeyRegister) <> 0 Then
                    If Val(KeyRegister) / 100 < TillData.SaleTotal Then
                        Validatator = False
                    End If
                End If
            End If
        Case "R200-00", "R100-00", "R50-00", "R20-00", "R10-00"
            If InStr(KeyRegister, "P/R") <> 0 Then Validatator = False
            If TillData.TableNo <> 0 Then
                ActiveReadServer "Select * from Table_Listing_View where Table_No = " & TillData.TableNo
                If rs.RecordCount > 0 Then
                    If UserRecord.User_Number <> rs.Fields("User_No") Then
                        If UserRecord.uType = 3 Or UserRecord.uType = 4 Then
                            If UserRecord.Bar_Cash = 0 Then
                                Validatator = False
                            End If
                        End If
                    End If
                End If
                rs.Close
            End If
            If UserRecord.Cash_Sales = False Then
                Validatator = False
            End If
            If TillData.DocNo = 0 Then Validatator = False
            If Right(KeyRegister, 2) = ") " Then Validatator = False
            If InStr(KeyRegister, ".") <> 0 Then
                If InStr(KeyRegister, "-") = 0 Then Validatator = False
            End If
            If InStr(KeyRegister, Chr(215)) <> 0 Then Validatator = False
        Case "No Sale"
            If TillData.Tipp <> 0 Then Validatator = False
            If GlobalMode = TillMode.TenderMode Then Validatator = False
            If TillData.DocNo <> 0 Then Validatator = False
    End Select
    If Validatator = False Then
        Keystring = ""
    End If
End Function
Public Sub DisplayErr(Title$)
    Select Case Panel_no
        Case 0
            frmSales.cmdErr.Caption = Title
            frmSales.cmdErr.Visible = True
            frmSales.errTimer.Enabled = True
        Case 1:
            frmSales1.cmdErr.Caption = Title
            frmSales1.cmdErr.Visible = True
            frmSales1.errTimer.Enabled = True
        Case 2:
            frmBar.cmdErr.Caption = Title
            frmBar.cmdErr.Visible = True
            frmBar.errTimer.Enabled = True
    End Select
End Sub
Private Sub TillCounters(Keystring$)

    'Quote-work and write to Quote_journal
    If Keystring = "Quote" Then Exit Sub
'    ActiveReadServer1 " SELECT     ISNULL(MAX(Invoice_No), 0) + 1 AS Invoice_No FROM Quote_Journal "
'    If rs1.RecordCount > 0 Then
'    Dim NextQuote As Integer
'    NextQuote = Val(rs1.Fields("Invoice_No"))
'    rs1.Close
'    ActiveUpdateServer "Insert into Quote_Journal"
'    End If
'    Exit Sub
'    End If
    
    If TillData.TableNo = 9999 Then Exit Sub
    On Error Resume Next
    TillData.Cashup_No = 0
    ActiveReadServer "Select * from Counters where User_no= " & UserRecord.User_Number & " and Finalized= 0"
    If rs.RecordCount > 0 Then
        TillData.Cashup_No = rs.Fields("Cashup_No")
    Else
        ActiveReadServer1 "Select isnull(max(Cashup_No),0)+1 as Cashup_No from Counters"
        TillData.Cashup_No = rs1.Fields("Cashup_No")
        rs1.Close
        ActiveReadServer1 "Select Function_Key,Date_Time from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
        "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
        If rs1.RecordCount > 0 Then
            If rs1.Fields("Function_Key") = 3 Then
                ClockinTime = rs1.Fields("Date_Time") & ""
            End If
        End If
        rs1.Close
        If Keystring = "No Sale" And UserRecord.uType = 0 Then
            rs.Close
            On Error GoTo 0
            Exit Sub
        Else
            ActiveUpdateServer "Insert into Counters (User_No,Cashup_No,Finalized,Counted,Shift_Start) values (" & UserRecord.User_Number & ", " & TillData.Cashup_No & ",0,0,'" & ClockinTime & "')"
        End If
    End If
    rs.Close
    DeductTipp = 0
    Select Case Panel_no
        Case 0
            If frmSales.grdMain.TextMatrix(frmSales.grdMain.Rows - 1, 3) = "Service Charge" Then
                DeductTipp = TillData.Tipp
            End If
            If TillData.ShortTender = True Then
                DeductTipp = TillData.Tipp
            End If
        Case 1
            If frmSales1.grdMain.TextMatrix(frmSales1.grdMain.Rows - 1, 3) = "Service Charge" Then
                DeductTipp = TillData.Tipp
            End If
            If TillData.ShortTender = True Then
                DeductTipp = TillData.Tipp
            End If
        Case 2
            If frmBar.grdMain.TextMatrix(frmBar.grdMain.Rows - 1, 3) = "Service Charge" Then
                DeductTipp = TillData.Tipp
            End If
            If TillData.ShortTender = True Then
                DeductTipp = TillData.Tipp
            End If
    End Select
    
    
    If TillData.Tendered > TillData.SaleTotal Or TillData.Tendered = TillData.SaleTotal Then
        If TillData.TotDiscountCount <> 0 Then
            
            
            
            ActiveUpdateServer "Update Counters set " & _
            "Discount_Perc_Value = isnull(Discount_Perc_Value,0) + " & (TillData.TotDiscount) & _
            ",Discount_Perc_Qty=isnull(Discount_Perc_Qty,0) +" & TillData.TotDiscountCount & _
            " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
        End If
        If TillData.TotDiscountValCount <> 0 Then
            ActiveUpdateServer "Update Counters set " & _
            "Discount_Amt_Value = isnull(Discount_Amt_Value,0) + " & (TillData.TotDiscount) & _
            ",Discount_Amt_Qty=isnull(Discount_Amt_Qty,0) +" & TillData.TotDiscountValCount & _
            " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
        End If
        If TillData.CorrectCount <> 0 Then
            ActiveUpdateServer "Update Counters set " & _
            "Item_Corrects_Value = isnull(Item_Corrects_Value,0) + " & TillData.Corrects & _
            ",Item_Corrects_Qty=isnull(Item_Corrects_Qty,0) + 1 " & _
            " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
        End If
        If TillData.VoidCount <> 0 Then
            ActiveUpdateServer "Update Counters set " & _
            "Voids_Value = isnull(Voids_Value,0) + " & (TillData.VoidTotal * -1) & _
            ",Voids_Qty=isnull(Voids_Qty,0) + 1 " & _
            " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
        End If
        If TillData.ReturnCount <> 0 Then
            ActiveUpdateServer "Update Counters set " & _
            "RTMD_Value = isnull(RTMD_Value,0) + " & (TillData.ReturnTotal * -1) & _
            ",RTMD_Qty=isnull(RTMD_Qty,0) + 1 " & _
            " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
        End If
        If TillData.UllageCount <> 0 Then
            ActiveUpdateServer "Update Counters set " & _
            "Ullage_Value = isnull(Ullage_Value,0) + " & (TillData.UllageTotal * -1) & _
            ",Ullage_Qty=isnull(Ullage_Qty,0) + 1 " & _
            " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
        End If
        If Val(TillData.Covers) = 0 Then
            ActiveUpdateServer "Update Counters set " & _
            "Customer_Count = isnull(Customer_Count,0) +  1 " & _
            " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
        Else
            ActiveUpdateServer "Update Counters set " & _
            "Customer_Count = isnull(Customer_Count,0) + " & TillData.Covers & _
            " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
        End If
        If TillData.Tipp <> 0 Then
            ActiveUpdateServer "Update Counters set " & _
            "Tipp = isnull(Tipp,0) + " & Abs(TillData.Tipp) & _
            ",Tipp_Count=isnull(Tipp_Count,0) + 1 " & _
            " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
        End If
    End If
    
    Select Case Keystring
        Case "No Sale"
            ActiveUpdateServer "Update Counters set No_Sales=isnull(No_Sales,0) + 1 where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
        Case "Cash"
            If TillData.Tendered > TillData.SaleTotal Or TillData.Tendered = TillData.SaleTotal Then
                If TillData.SaleTotal < 0 Then
                
                
                ActiveUpdateServer "Update Counters set " & _
                "Cash_Sales_Value=isnull(Cash_Sales_Value,0) + " & TillData.Tendered - TillData.Change - DeductTipp & _
                ",Cash_Sales_Qty=isnull(Cash_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                End If
                
                If TillData.SaleTotal > 0 Then
            x = Format(TillData.Cash, "00.00")
                ActiveUpdateServer "Update Counters set " & _
                "Cash_Sales_Value=isnull(Cash_Sales_Value,0) + " & TillData.Cash - TillData.Change - DeductTipp & _
                ",Cash_Sales_Qty=isnull(Cash_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                End If
            Else
                TillData.ShortTender = True
                ActiveUpdateServer "Update Counters set " & _
                "Cash_Sales_Value=isnull(Cash_Sales_Value,0) + " & TillData.Cash & _
                ",Cash_Sales_Qty=isnull(Cash_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                TillData.Cash = 0
            End If
            If TillData.Change = 0 Or TillData.Change > 0 Then
                ActiveUpdateServer "Update Counters set " & _
                "TaxableSales_Value = isnull(TaxableSales_Value,0) + " & TillData.TaxableSales & _
                ",TotalExcemptSales = isnull(TotalExcemptSales,0) + " & TillData.NonTaxableSales & _
                ",TotalCalculatedTax_Value = isnull(TotalCalculatedTax_Value,0) + " & TillData.CalculatedTax & _
                ",TotalCollectedTax_Value = isnull(TotalCollectedTax_Value,0) + " & TillData.CollectedTax & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
            End If
        Case "Voucher"
            If TillData.Tendered > TillData.SaleTotal Or TillData.Tendered = TillData.SaleTotal Then
                ActiveUpdateServer "Update Counters set " & _
                "Cheque_Sales_Value=isnull(Cheque_Sales_Value,0) + " & TillData.Cheque - TillData.Change - DeductTipp & _
                ",Cheque_Sales_Qty=isnull(Cheque_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
            Else
                TillData.ShortTender = True
                ActiveUpdateServer "Update Counters set " & _
                "Cheque_Sales_Value=isnull(Cheque_Sales_Value,0) + " & TillData.Cheque & _
                ",Cheque_Sales_Qty=isnull(Cheque_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                TillData.Cheque = 0
            End If
            If TillData.Change = 0 Or TillData.Change > 0 Then
                ActiveUpdateServer "Update Counters set " & _
                "TaxableSales_Value = isnull(TaxableSales_Value,0) + " & TillData.TaxableSales & _
                ",TotalExcemptSales = isnull(TotalExcemptSales,0) + " & TillData.NonTaxableSales & _
                ",TotalCalculatedTax_Value = isnull(TotalCalculatedTax_Value,0) + " & TillData.CalculatedTax & _
                ",TotalCollectedTax_Value = isnull(TotalCollectedTax_Value,0) + " & TillData.CollectedTax & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
            End If
        Case "Card"
            If TillData.Tendered > TillData.SaleTotal Or TillData.Tendered = TillData.SaleTotal Then
                ActiveUpdateServer "Update Counters set " & _
                "Card_Sales_Value=isnull(Card_Sales_Value,0) + " & TillData.Card - TillData.Change - DeductTipp & _
                ",Card_Sales_Qty=isnull(Card_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                DoEvents
                If System_Service = 1 And TillData.SaleTotal > 0 Then
                    If TillData.Tendered > TillData.SaleTotal And TillData.Tipp = 0 Then
                        Select Case Panel_no
                            Case 0
                                frmSales.Tag = "1"
                            Case 2
                                frmBar.Tag = "1"
                        End Select
                        Load frmTipp
                        frmTipp.Tag = ""
                        frmTipp.Show vbModal
                        If TillData.Tipp <> 0 Then
                            DoEvents
                            ActiveUpdateServer "Update Counters set " & _
                            "Tipp=isnull(Tipp,0) + " & TillData.Tipp & _
                            ",Tipp_Count=isnull(Tipp_Count,0) + 1 " & _
                            " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                        End If
                    End If
                End If
            Else
                DoEvents
                TillData.ShortTender = True
                ActiveUpdateServer "Update Counters set " & _
                "Card_Sales_Value=isnull(Card_Sales_Value,0) + " & TillData.Card & _
                ",Card_Sales_Qty=isnull(Card_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                TillData.Card = 0
            End If
            If TillData.Change = 0 Or TillData.Change > 0 Then
                DoEvents
                ActiveUpdateServer "Update Counters set " & _
                "TaxableSales_Value = isnull(TaxableSales_Value,0) + " & TillData.TaxableSales & _
                ",TotalExcemptSales = isnull(TotalExcemptSales,0) + " & TillData.NonTaxableSales & _
                ",TotalCalculatedTax_Value = isnull(TotalCalculatedTax_Value,0) + " & TillData.CalculatedTax & _
                ",TotalCollectedTax_Value = isnull(TotalCollectedTax_Value,0) + " & TillData.CollectedTax & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
            End If
        Case "Charge"
            If TillData.Tendered > TillData.SaleTotal Or TillData.Tendered = TillData.SaleTotal Then
                DoEvents
                ActiveUpdateServer "Update Counters set " & _
                "Charge_Sales_Value=isnull(Charge_Sales_Value,0) + " & TillData.Charge - TillData.Change - DeductTipp & _
                ",Charge_Sales_Qty=isnull(Charge_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
            Else
                TillData.ShortTender = True
                DoEvents
                ActiveUpdateServer "Update Counters set " & _
                "Charge_Sales_Value=isnull(Charge_Sales_Value,0) + " & TillData.Charge & _
                ",Charge_Sales_Qty=isnull(Charge_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                TillData.Charge = 0
            End If
            If TillData.Change = 0 Or TillData.Change > 0 Then
                DoEvents
                ActiveUpdateServer "Update Counters set " & _
                "TaxableSales_Value = isnull(TaxableSales_Value,0) + " & TillData.TaxableSales & _
                ",TotalExcemptSales = isnull(TotalExcemptSales,0) + " & TillData.NonTaxableSales & _
                ",TotalCalculatedTax_Value = isnull(TotalCalculatedTax_Value,0) + " & TillData.CalculatedTax & _
                ",TotalCollectedTax_Value = isnull(TotalCollectedTax_Value,0) + " & TillData.CollectedTax & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
            End If
         Case "Loyalty"
            If TillData.Tendered > TillData.SaleTotal Or TillData.Tendered = TillData.SaleTotal Then
                DoEvents
                ActiveUpdateServer "Update Counters set " & _
                "Loyalty_Sales_Value=isnull(Loyalty_Sales_Value,0) + " & TillData.Loyalty - TillData.Change - DeductTipp & _
                ",Loyalty_Sales_Qty=isnull(Loyalty_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
            Else
                TillData.ShortTender = True
                DoEvents
                ActiveUpdateServer "Update Counters set " & _
                "Loyalty_Sales_Value=isnull(Loyalty_Sales_Value,0) + " & TillData.Loyalty & _
                ",Loyalty_Sales_Qty=isnull(Loyalty_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                TillData.Loyalty = 0
            End If
            If TillData.Change = 0 Or TillData.Change > 0 Then
                DoEvents
                ActiveUpdateServer "Update Counters set " & _
                "TaxableSales_Value = isnull(TaxableSales_Value,0) + " & TillData.TaxableSales & _
                ",TotalExcemptSales = isnull(TotalExcemptSales,0) + " & TillData.NonTaxableSales & _
                ",TotalCalculatedTax_Value = isnull(TotalCalculatedTax_Value,0) + " & TillData.CalculatedTax & _
                ",TotalCollectedTax_Value = isnull(TotalCollectedTax_Value,0) + " & TillData.CollectedTax & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
            End If
    End Select
    On Error GoTo 0
End Sub
Private Sub GetData(Keystring$)
    
    
    TillData.Price = ""
    TillData.Qty = "1"
    TillData.ProductCode = ""
    TillData.Weight = 0
    Product_Code = ""
    TillData.DeptNo = ""
    TillData.Keystring = ""
    TillData.Cost = 0
    TillData.TaxRate = 0
    TillData.TaxType = 0
    TillData.Description = ""
    TillData.ShortDesc = ""
    TillData.UnitSize = ""
    TillData.Kitchen1 = ""
    TillData.Kitchen2 = ""
    TillData.PriceOveride = 0
    TillData.Recipe = 0
    TillData.Discount = 0
    TillData.DiscountVal = 0
    TillData.Deposit = 0
    DoEvents
    Select Case Keystring
        Case "Plu"
            If TillData.ExtraFunc = "Void" And KeyRegister = " <Plu>" Then
                DisplayErr "No Product Code Supplied"
                Exit Sub
            End If
            If TillData.ExtraFunc = "Return Item" And KeyRegister = " <Plu>" Then
                DisplayErr "No Product Code Supplied"
                Exit Sub
            End If
            If TillData.ExtraFunc = "Wastage" And KeyRegister = " <Plu>" Then
                DisplayErr "No Product Code Supplied"
                Exit Sub
            End If
            If TillData.ExtraFunc = "Corr" And KeyRegister = " <Plu>" Then
                DisplayErr "No Product Code Supplied"
                Exit Sub
            End If
            TillData.ExtraFunc = ""
            If KeyRegister = " <Plu>" And TillData.DocNo <> 0 Then
                If InStr(TillData.KeyReg, "Plu") <> 0 Then
                    KeyRegister = TillData.KeyReg
                End If
            End If
            
            For i = InStr(KeyRegister, "<") - 2 To 1 Step -1
                If Chr(Asc(Mid(KeyRegister, i, 1))) = " " Or Chr(Asc(Mid(KeyRegister, i, 1))) = "<" Then Exit For
                Product_Code = Mid(KeyRegister, i, 1) & Product_Code
            Next i
            If InStr(KeyRegister, Chr(215)) <> 0 Then
                TillData.Qty = ""
                For i = InStr(KeyRegister, Chr(215)) - 2 To 1 Step -1
               
                
                    If Asc(Mid(KeyRegister, i, 1)) < 46 Or Asc(Mid(KeyRegister, i, 1)) > 57 Then Exit For
                    TillData.Qty = Mid(KeyRegister, i, 1) & TillData.Qty
                Next i
            End If
            If InStr(KeyRegister, "Return Item") <> 0 Then
                TillData.Qty = Val(TillData.Qty) * -1
            End If
            If InStr(KeyRegister, "Wastage") <> 0 Then
                'TillData.Qty = Val(TillData.Qty) * -1  Kotie 17-03-2013
            End If
            If InStr(KeyRegister, "Void") <> 0 Then
                TillData.Qty = Val(TillData.Qty) * -1
            End If
            If InStr(KeyRegister, "Price O/V") <> 0 Then
                For i = InStr(KeyRegister, "(Price O/V") - 2 To 1 Step -1
                    If Asc(Mid(KeyRegister, i, 1)) < 46 Or Asc(Mid(KeyRegister, i, 1)) > 57 Then Exit For
                    TillData.Price = Mid(KeyRegister, i, 1) & TillData.Price
                Next i
            End If
            If Product_Code = "" Then
                DisplayErr "No Product Code Supplied"
                Exit Sub
            End If
            TillData.ProductCode = Product_Code
            ActiveReadServer "Select * from Products where Sales_Item = 1 and Product_Code='" & TillData.ProductCode & "'"
            If rs.RecordCount > 0 Then
                TillData.Deposit = Val(rs.Fields("Returnable_Item"))
                TillData.DeptNo = rs.Fields("Department_No")
                If rs.Fields("Unit_of_Measure") = "each" Then
                    TillData.Description = rs.Fields("Description")
                Else
                    TillData.Description = rs.Fields("Description") & " " & rs.Fields("Unit_Size") & rs.Fields("Unit_of_Measure")
                End If
                If IsNull(rs.Fields("Short_Description")) Then
                    TillData.ShortDesc = Mid(rs.Fields("Description"), 1, 25)
                Else
                    TillData.ShortDesc = rs.Fields("Short_Description")
                End If
                 If TillData.Weight <> 0 Then
                    TillData.Description = rs.Fields("Description")
                End If
                TillData.TaxRate = rs.Fields("Sales_Tax")
                TillData.TaxType = rs.Fields("Tax_Type")
                If TillData.Price = "" Then
                    TillData.Price = rs.Fields("Selling_Price")
                    If TillData.Account_No <> "" Then
                        ActiveReadServer2 "Select * from Debtor_Discounts where Debtor_No = '" & TillData.Account_No & "' and Department_No = '" & TillData.DeptNo & "'"
                        If rs2.RecordCount > 0 Then
                            If Val(rs2.Fields("Selling_Price") & "") <> 0 Then
                                Discount1 = 0
                                ActiveReadServer1 "Select Price" & Val(rs2.Fields("Selling_Price") & "") & " as Price from Product_Prices where Product_Code = '" & TillData.ProductCode & "'"
                                If rs1.RecordCount > 0 Then
                                    TillData.Price = rs1.Fields("Price")
                                End If
                                rs1.Close
                            End If
                            If Val(rs2.Fields("Cost_Disc") & "") <> 0 Then
                                TillData.Price = (rs.Fields("Ave_Cost") * ((100 + Val(rs2.Fields("Cost_Disc") & "")) / 100)) * ((100 + TillData.TaxRate) / 100)
                            End If
                            If Val(rs2.Fields("Sell_Disc") & "") <> 0 Then
                                TillData.Price = TillData.Price * ((100 - Val(rs2.Fields("Sell_Disc") & "")) / 100)
                            End If
                            If rs2.Fields("Sell_Cost") <> "No" Then
                                TillData.Price = rs.Fields("Ave_Cost")
                            End If
                        End If
                        rs2.Close
                    End If
                    If HappyHour = 1 Then
                        Select Case HappyHourPrice
                            Case 2
                                ActiveReadServer2 "Select Price2 from Product_Prices where Product_Code = '" & TillData.ProductCode & "'"
                                If rs2.RecordCount > 0 Then
                                    If Round(rs2.Fields("Price2"), 2) > 0 Then
                                        TillData.Price = rs2.Fields("Price2")
                                    End If
                                End If
                                rs2.Close
                            Case 3
                                ActiveReadServer2 "Select Price3 from Product_Prices where Product_Code = '" & TillData.ProductCode & "'"
                                If rs2.RecordCount > 0 Then
                                    If Round(rs2.Fields("Price3"), 2) > 0 Then
                                        TillData.Price = rs2.Fields("Price3")
                                    End If
                                End If
                                rs2.Close
                            Case 4
                                ActiveReadServer2 "Select Price4 from Product_Prices where Product_Code = '" & TillData.ProductCode & "'"
                                If rs2.RecordCount > 0 Then
                                    If Round(rs2.Fields("Price4"), 2) > 0 Then
                                        TillData.Price = rs2.Fields("Price4")
                                    End If
                                End If
                                rs2.Close
                            Case 5
                                ActiveReadServer2 "Select Price5 from Product_Prices where Product_Code = '" & TillData.ProductCode & "'"
                                If rs2.RecordCount > 0 Then
                                    If Round(rs2.Fields("Price5"), 2) > 0 Then
                                        TillData.Price = rs2.Fields("Price5")
                                    End If
                                End If
                                rs2.Close
                            Case 6
                                ActiveReadServer2 "Select Price6 from Product_Prices where Product_Code = '" & TillData.ProductCode & "'"
                                If rs2.RecordCount > 0 Then
                                    If Round(rs2.Fields("Price6"), 2) > 0 Then
                                        TillData.Price = rs2.Fields("Price6")
                                    End If
                                End If
                                rs2.Close
                        End Select
                    End If
                    If HappyHour1 = 1 Then
                        Select Case HappyHourPrice1
                            Case 2
                                ActiveReadServer2 "Select Price2 from Product_Prices where Product_Code = '" & TillData.ProductCode & "'"
                                If rs2.RecordCount > 0 Then
                                    If Round(rs2.Fields("Price2"), 2) > 0 Then
                                        TillData.Price = rs2.Fields("Price2")
                                    End If
                                End If
                                rs2.Close
                            Case 3
                                ActiveReadServer2 "Select Price3 from Product_Prices where Product_Code = '" & TillData.ProductCode & "'"
                                If rs2.RecordCount > 0 Then
                                    If Round(rs2.Fields("Price3"), 2) > 0 Then
                                        TillData.Price = rs2.Fields("Price3")
                                    End If
                                End If
                                rs2.Close
                            Case 4
                                ActiveReadServer2 "Select Price4 from Product_Prices where Product_Code = '" & TillData.ProductCode & "'"
                                If rs2.RecordCount > 0 Then
                                    If Round(rs2.Fields("Price4"), 2) > 0 Then
                                        TillData.Price = rs2.Fields("Price4")
                                    End If
                                End If
                                rs2.Close
                            Case 5
                                ActiveReadServer2 "Select Price5 from Product_Prices where Product_Code = '" & TillData.ProductCode & "'"
                                If rs2.RecordCount > 0 Then
                                    If Round(rs2.Fields("Price5"), 2) > 0 Then
                                        TillData.Price = rs2.Fields("Price5")
                                    End If
                                End If
                                rs2.Close
                            Case 6
                                ActiveReadServer2 "Select Price6 from Product_Prices where Product_Code = '" & TillData.ProductCode & "'"
                                If rs2.RecordCount > 0 Then
                                    If Round(rs2.Fields("Price6"), 2) > 0 Then
                                        TillData.Price = rs2.Fields("Price6")
                                    End If
                                End If
                                rs2.Close
                        End Select
                    End If
                    TillData.PriceOveride = 0
                Else
                    TillData.Price = Val(TillData.Price) / 100
                    TillData.PriceOveride = 1
                End If
                If TillData.ExtraFunc <> "Corr" Then
                    If TillData.ExtraFunc <> "Void" Then
                        If InStr(KeyRegister, Chr(215)) = 0 Then
                            
                           
                            If rs.Fields("Scale_Item") = 1 And rs.Fields("Scale_Prefix") = "<None>" Then
'                                If Devices.ScalePort <> "<Not Installed>" Then
                                If Devices.ScalePort <> "<Not Set>" Then
                                    frmScale.lblWeight.Visible = True
                                    frmScale.lblWeightType.Enabled = False
                                    frmScale.lblWeightType.Visible = False
                                    ReadScale
                                    TillData.Qty = TillData.Weight
                                    If TillData.Qty = 0 Then
                                        frmScale.Show vbModal
                                        If TillData.Weight <> 0 Then
                                            TillData.Qty = TillData.Weight
                                        Else
                                            DisplayErr "No Selling Quantity Supplied"
                                            TillData.ProductCode = ""
                                            TillData.Weight = 0
                                            rs.Close
                                            Exit Sub
                                        End If
                                    End If
                                Else
                                    frmScale.lblWeight.Visible = False
                                    frmScale.lblWeightType.Enabled = True
                                    frmScale.lblWeightType.Visible = True
                                    TillData.Qty = 0
                                    If TillData.Qty = 0 Then
                                        frmScale.Show vbModal
                                        If TillData.Weight <> 0 Then
                                            TillData.Qty = TillData.Weight
                                        Else
                                            DisplayErr "No Selling Quantity Supplied"
                                            TillData.ProductCode = ""
                                            TillData.Weight = 0
                                            rs.Close
                                            Exit Sub
                                        End If
                                    End If
'                                    DisplayErr "Communication failure with the Checkout Scale."
'                                    TillData.ProductCode = ""
'                                    TillData.Weight = 0
'                                    rs.Close
'                                    Exit Sub
                                End If
                            End If
                            
                            
                            
                            
                        End If
                    End If
                End If
                If InStr(KeyRegister, "Wastage") = 0 Then
                    TillData.TaxTotal = TillData.TaxTotal + (TillData.Price * TillData.Qty) - ((TillData.Price * TillData.Qty) / ((100 + TillData.TaxRate) / 100))
                End If
                TillData.Keystring = Keystring
                TillData.Cost = rs.Fields("Ave_Cost")
                Select Case Kitchen_Printer_No
                    Case 0
                        TillData.Kitchen1 = rs.Fields("Kitchen1") & ""
                    Case 1
                        TillData.Kitchen1 = rs.Fields("Kitchen2") & ""
                    Case 2
                        TillData.Kitchen1 = rs.Fields("Kitchen1") & ""
                        TillData.Kitchen2 = rs.Fields("Kitchen2") & ""
                End Select
                TillData.Kitchen2 = rs.Fields("Kitchen2") & ""
                
                TillData.Recipe = rs.Fields("Recipe_Item")
            Else
                DisplayErr "Unknown Product Code - " & TillData.ProductCode
                TillData.ProductCode = ""
                TillData.Weight = 0
                rs.Close
                Exit Sub
            End If
            rs.Close
            If Val(TillData.Price) = 0 And Product_Code <> "" Then
                ActiveReadServer "Select Zero_Price from Departments where Department_No = '" & TillData.DeptNo & "'"
                If rs.RecordCount > 0 Then
                    If Val(rs.Fields("Zero_Price") & "") = 0 Then
                        DisplayErr "No Selling Price Supplied"
                        TillData.ProductCode = ""
                        TillData.Weight = 0
                    End If
                    rs.Close
                End If
                Exit Sub
            End If
            ActiveReadServer "Select * from Specials where Product_Code = '" & TillData.ProductCode & "' and Active = 1"
            If rs.RecordCount > 0 Then
                If Val(rs.Fields("Price") & "") <> 0 Then
                    TillData.Price = Val(rs.Fields("Price") & "")
                End If
                rs.Close
            End If
            TillData.KeyReg = KeyRegister
            If TillData.Deposit = 1 Then
                TillData.Qty = TillData.Qty * -1
            End If
        Case "Dept"
            If TillData.ExtraFunc = "Void" And KeyRegister = " <Dept>" Then
                DisplayErr "No Department Number Supplied"
                Exit Sub
            End If
            If TillData.ExtraFunc = "Return Item" And KeyRegister = " <Dept>" Then
                DisplayErr "No Department Number Supplied"
                Exit Sub
            End If
            If TillData.ExtraFunc = "Wastage" And KeyRegister = " <Dept>" Then
                DisplayErr "No Department Number Supplied"
                Exit Sub
            End If
            If TillData.ExtraFunc = "Corr" And KeyRegister = " <Dept>" Then
                DisplayErr "No Department Number Supplied"
                Exit Sub
            End If
            TillData.ExtraFunc = ""
            If KeyRegister = " <Dept>" And TillData.DocNo <> 0 Then
                If InStr(TillData.KeyReg, "Dept") <> 0 Then
                    KeyRegister = TillData.KeyReg
                End If
            End If
            For i = InStr(KeyRegister, "<") - 2 To 1 Step -1
                If Asc(Mid(KeyRegister, i, 1)) < 48 Or Asc(Mid(KeyRegister, i, 1)) > 57 Then Exit For
                TillData.DeptNo = Mid(KeyRegister, i, 1) & TillData.DeptNo
            Next i
            If InStr(KeyRegister, Chr(215)) <> 0 Then
                TillData.Qty = ""
                For i = InStr(KeyRegister, Chr(215)) - 2 To 1 Step -1
                    If Asc(Mid(KeyRegister, i, 1)) < 46 Or Asc(Mid(KeyRegister, i, 1)) > 57 Then Exit For
                    TillData.Qty = Mid(KeyRegister, i, 1) & TillData.Qty
                Next i
            End If
            If InStr(KeyRegister, "Return Item") <> 0 Then
                TillData.Qty = Val(TillData.Qty) * -1
            End If
            If InStr(KeyRegister, "Wastage") <> 0 Then
                TillData.Qty = Val(TillData.Qty) * -1
            End If
            If InStr(KeyRegister, "Price O/V") <> 0 Then
                For i = InStr(KeyRegister, "(Price O/V") - 2 To 1 Step -1
                    If Asc(Mid(KeyRegister, i, 1)) < 46 Or Asc(Mid(KeyRegister, i, 1)) > 57 Then Exit For
                    TillData.Price = Mid(KeyRegister, i, 1) & TillData.Price
                Next i
            End If
            If TillData.DeptNo = "" Then
                DisplayErr "No Department Number Supplied"
                Exit Sub
            End If
            If TillData.Price = "" Then
                DisplayErr "No Selling Price Supplied"
                TillData.DeptNo = ""
                Exit Sub
            End If
            ActiveReadServer "SELECT Departments.Department_No, Departments.Dept_Name,Short_Name, Departments.Sales_Tax, Departments.Tax_Type " & _
            "FROM Departments INNER JOIN Department_Links ON Departments.Department_No = Department_Links.Dept_No INNER JOIN " & _
            "Locations ON Department_Links.Location_No = Locations.Location_No " & _
            "WHERE (Locations.Loc_Type = 0) AND (Departments.Dept_Type = 0) AND (Departments.Department_No = '" & TillData.DeptNo & "') " & _
            "GROUP BY Departments.Department_No, Departments.Dept_Name, Short_Name,Departments.Sales_Tax, Departments.Tax_Type"
            If rs.RecordCount > 0 Then
                TillData.Description = rs.Fields("Dept_Name")
                If IsNull(rs.Fields("Short_Name")) Then
                    TillData.ShortDesc = Mid(rs.Fields("Dept_Name"), 1, 25)
                Else
                    TillData.ShortDesc = rs.Fields("Short_Name")
                End If
                TillData.TaxRate = rs.Fields("Sales_Tax")
                TillData.TaxType = rs.Fields("Tax_Type")
                TillData.Price = Val(TillData.Price) / 100
                TillData.PriceOveride = 1
                TillData.Keystring = Keystring
                TillData.Cost = 0
            Else
                DisplayErr "Unknown Department Number - " & TillData.DeptNo
                TillData.ProductCode = ""
                TillData.Weight = 0
            End If
            rs.Close
            TillData.KeyReg = KeyRegister
    End Select
End Sub
Private Sub ExecFunction(Keystring$)
    On Error Resume Next
    Select Case Keystring
        Case "Corr"
            TillData.ExtraFunc = "Corr"
        Case "Cash", "Card", "Voucher", "Charge", "No Sale", "Quote"
            ActiveUpdateServer "Delete from Table_Listing where Table_No = " & TillData.TableNo
            ActiveUpdateServer "Delete from Tab_Listing where Tab_No = " & TillData.TabNo
            If GlobalMode <> TillMode.TenderMode Then
                TillData.Prev_Doc_No = TillData.DocNo
                TillData.DocNo = 0
                TillData.TransNo = 0
                TillData.TableNo = 0
                TillData.Table_Name = ""
                TillData.Covers = 0
                TillData.TabName = ""
                TillData.TabNo = 0
                TillData.Tipp = 0
                TillData.TotDiscount = 0
                TillData.TotDiscountVal = 0
                TillData.TotDiscountCount = 0
                TillData.TotDiscountValCount = 0
                TillData.Room_No = 0
                TillData.Res_No = 0
                TillData.Account_No = ""
                If Panel_no = 0 Then frmSales.cmdFancy(4).Caption = "Member No"
                If Panel_no = 0 Then frmSales.lblDebtor.Caption = ""
            End If
            Finalizing = False
        Case "Subtotal"
            TillData.Keystring = "Subtotal"
        Case "R200-00", "R100-00", "R50-00", "R20-00", "R10-00"
            If InStr(Keystring, "-") <> 0 Then
                KeyRegister = "R" & Replace(Format(Val(Replace(Mid(KeyRegister, 2), "-", ".")) + Val(Replace(Mid(Keystring, 2), "-", ".")), "0.00"), ".", "-")
            End If
    End Select
    On Error GoTo 0
    Screen.MousePointer = 0
End Sub
Private Sub CalculateChange(Keystring$)
    On Error Resume Next
    TillData.Change = 0
    tender = 0
    Select Case Panel_no
        Case 0
            
            If InStr(KeyRegister, "-") <> 0 Then
                TillData.Tendered = TillData.Tendered + Val(Replace(Mid(frmSales.lblKeyRegister.Caption, 2), "-", "."))
                tender = Val(Replace(Mid(frmSales.lblKeyRegister.Caption, 2), "-", "."))
            Else
                If InStr(KeyRegister, "<") = 1 Then
                    If TillData.SaleTotal < 0 Then
                        TillData.Tendered = TillData.SaleTotal
                        tender = TillData.SaleTotal
            
                    Else
                        KeyRegister = (TillData.SaleTotal - TillData.Tendered) * 100 & KeyRegister
                        
                        
                        If TillData.ShortTender = False Then
                            tender = TillData.SaleTotal
                           frmSales.cmdKey(7).Enabled = True
                   
                       
                        Else
                            tender = TillData.SaleTotal - TillData.Tendered
                            frmSales.cmdKey(7).Enabled = True
                        End If
                        TillData.Tendered = TillData.SaleTotal
'                        If Keystring$ = "Cash" Or Keystring$ = "Card" Or Keystring$ = "Voucher" And TillData.SaleTotal > TillData.Tendered Then frmSales.cmdKey(7).Enabled = False
'            If Keystring$ = "Cash" Or Keystring$ = "Card" Or Keystring$ = "Voucher" And TillData.SaleTotal < TillData.Tendered Then frmSales.cmdKey(7).Enabled = True
'
                    End If
                Else
                    TillData.Tendered = TillData.Tendered + Val(frmSales.lblKeyRegister) / 100
                    tender = Val(frmSales.lblKeyRegister) / 100
                    If TillData.Tendered < TillData.SaleTotal Then frmSales.cmdKey(7).Enabled = False
                   If TillData.ShortTender = True And TillData.Tendered = TillData.SaleTotal Or TillData.Tendered > TillData.SaleTotal Then frmSales.cmdKey(7).Enabled = True
                   If TillData.Tendered = TillData.SaleTotal Or TillData.Tendered > TillData.SaleTotal Then frmSales.cmdKey(7).Enabled = True
                   
                End If
            End If
        Case 1
            
            If InStr(KeyRegister, "-") <> 0 Then
                TillData.Tendered = TillData.Tendered + Val(Replace(Mid(frmSales1.lblKeyRegister.Caption, 2), "-", "."))
                tender = Val(Replace(Mid(frmSales1.lblKeyRegister.Caption, 2), "-", "."))
            Else
                If InStr(KeyRegister, "<") = 1 Then
                    If TillData.SaleTotal < 0 Then
                        TillData.Tendered = 0
                        tender = 0
                    Else
                        KeyRegister = (TillData.SaleTotal - TillData.Tendered) * 100 & KeyRegister
                        If TillData.ShortTender = False Then
                            tender = TillData.SaleTotal
                        Else
                            tender = TillData.SaleTotal - TillData.Tendered
                        End If
                        TillData.Tendered = TillData.SaleTotal
                    End If
                Else
                    TillData.Tendered = TillData.Tendered + Val(frmSales1.lblKeyRegister) / 100
                    tender = Val(frmSales1.lblKeyRegister) / 100
                    
                End If
            End If
        Case 2
             
            If InStr(KeyRegister, "-") <> 0 Then
                TillData.Tendered = TillData.Tendered + Val(Replace(Mid(frmBar.lblKeyRegister.Caption, 2), "-", "."))
                tender = Val(Replace(Mid(frmBar.lblKeyRegister.Caption, 2), "-", "."))
            Else
                If InStr(KeyRegister, "<") = 1 Then
                    If TillData.SaleTotal < 0 Then
                        TillData.Tendered = 0
                        tender = TillData.SaleTotal
                    Else
                        If TillData.ShortTender = False Then
                            tender = TillData.SaleTotal
                        frmBar.cmdInput(19).Enabled = True
                        Else
                            tender = TillData.SaleTotal - TillData.Tendered
                            
                        End If
                        TillData.Tendered = TillData.SaleTotal
                        frmBar.cmdInput(19).Enabled = True
                    End If
                Else
                    TillData.Tendered = TillData.Tendered + Val(frmBar.lblKeyRegister) / 100
                    tender = Val(frmBar.lblKeyRegister) / 100
                   If TillData.Tendered < TillData.SaleTotal Then frmBar.cmdInput(19).Enabled = False
                   If TillData.ShortTender = True And TillData.Tendered = TillData.SaleTotal Or TillData.Tendered > TillData.SaleTotal Then frmBar.cmdInput(19).Enabled = True
                   If TillData.Tendered = TillData.SaleTotal Or TillData.Tendered > TillData.SaleTotal Then frmBar.cmdInput(19).Enabled = True
                    
                  
                End If
            End If
    End Select
    Select Case Keystring
        Case "Cash": TillData.Cash = TillData.Cash + tender
        Case "Card": TillData.Card = TillData.Card + tender
        Case "Voucher": TillData.Cheque = TillData.Cheque + tender
        Case "Charge": TillData.Charge = TillData.Charge + tender
        Case "Loyalty": TillData.Loyalty = TillData.Loyalty + tender
    End Select
    If TillData.Tendered > 0 Or TillData.Tendered = 0 Then
        TillData.Change = TillData.Tendered - Val(TillData.SaleTotal)
    Else
        TillData.Change = 0
        If TillData.SaleTotal < 0 Then
            TillData.Change = TillData.Tendered - Val(TillData.SaleTotal)
        Else
            TillData.Tendered = TillData.SaleTotal
        End If
    End If
    TillData.Change = Round(TillData.Change, 2)
    If TillData.Change < 0 Then
        GlobalMode = TillMode.TenderMode
    Else
        GlobalMode = TillMode.FinMode
    End If
    On Error GoTo 0
End Sub
Private Sub AskMember()
    If Member_No = 1 Then
        Select Case Panel_no
            Case 0
                frmMember.Show vbModal
                Select Case KeyRegister
                    Case ""
                        frmSales.lblKeyRegister = ""
                        frmSales.lblDebtor = ""
                        TillData.Account_No = ""
                    Case Else
                        ActiveReadServer "Select * from Debtors where Debtor_No ='" & KeyRegister & "'"
                        If rs.RecordCount > 0 Then
                            frmSales.lblKeyRegister = "Member - " & rs.Fields("Debtor_Name") & " (" & KeyRegister & ")"
                            frmSales.lblDebtor = "Member - " & rs.Fields("Debtor_Name") & " (" & KeyRegister & ")"
                            TillData.Account_No = KeyRegister
                            TillData.Creditbalance = rs.Fields("Balance")
                            TillData.Creditlimit = rs.Fields("Credit_Limit")
                        End If
                        rs.Close
                        KeyRegister = ""
                End Select
            Case 2
                frmMember.Show vbModal
                Select Case KeyRegister
                    Case ""
                        frmBar.lblKeyRegister = ""
                        TillData.Account_No = ""
                    Case Else
                        ActiveReadServer "Select * from Debtors where Debtor_No ='" & KeyRegister & "'"
                        If rs.RecordCount > 0 Then
                            frmBar.lblKeyRegister = "Member - " & rs.Fields("Debtor_Name") & " (" & KeyRegister & ")"
                            TillData.Account_No = KeyRegister
                            TillData.Creditbalance = rs.Fields("Balance")
                            TillData.Creditlimit = rs.Fields("Credit_Limit")
                        End If
                        rs.Close
                        KeyRegister = ""
                End Select
        End Select
    End If
End Sub
Private Sub Payouts()
    frmSales.Tag = "1"
    frmSales.lblPayReason = ""
    DoEvents
    Load frmKeyBoard
    frmKeyBoard.Tag = "Payout"
    frmKeyBoard.Show vbModal
    DoEvents
    Load frmPayout
    frmSales.Tag = "1"
    frmPayout.Show vbModal
    Select Case frmPayout.Tag
        Case ""
        Case Else
            TillData.Account_No = frmPayout.Tag
            ActiveReadServer1 "Select isnull(max(Payment_No),0)+1 as Payment_No from Supplier_Accounts where Transaction_Type = 'Payment'"
            Payment_No = rs1.Fields("Payment_No")
            rs1.Close
            Pay1_Value = 0
            Balance = 0
            ActiveReadServer "Select Balance from Suppliers where Supplier_No = '" & TillData.Account_No & "'"
            If rs.RecordCount > 0 Then
                Balance = rs.Fields("Balance")
            End If
            rs.Close
            TillData.Change = 0
            If InStr(KeyRegister, ".") = 0 Then
                NewBalance = Balance - Val(KeyRegister) / 100
                Pay1_Value = Val(KeyRegister) / 100
            Else
                NewBalance = Balance - Val(KeyRegister) / 100
                Pay1_Value = Val(KeyRegister)
            End If
            ActiveUpdateServer "INSERT INTO [Supplier_Accounts]([User_No],[Date_Time],[Transaction_Type], [Payment_No], [Account_No], [Debit], [Credit], [Balance],[Tender_Type],[Ref_No])" & _
            "VALUES(" & UserRecord.User_Number & ",Getdate(),'Payment'," & Payment_No & ",'" & frmPayout.Tag & "'," & Pay1_Value & ",0," & Balance - ((Pay1_Value) * -1) & ",'Cash','" & frmSales.lblPayReason & "')"
            DoEvents
            
            ActiveUpdateServer "Update Suppliers set Balance=Balance - " & (Pay1_Value) & " where Supplier_No='" & frmPayout.Tag & "'"
            
            TillData.Cashup_No = 0
            ActiveReadServer "Select * from Counters where User_no= " & UserRecord.User_Number & " and Finalized= 0"
            If rs.RecordCount > 0 Then
                TillData.Cashup_No = rs.Fields("Cashup_No")
            Else
                ActiveReadServer1 "Select isnull(max(Cashup_No),0)+1 as Cashup_No from Counters"
                TillData.Cashup_No = rs1.Fields("Cashup_No")
                rs1.Close
                ActiveReadServer1 "Select Function_Key,Date_Time from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                If rs1.RecordCount > 0 Then
                    If rs1.Fields("Function_Key") = 3 Then
                        ClockinTime = rs1.Fields("Date_Time") & ""
                    End If
                End If
                rs1.Close
                ActiveUpdateServer "Insert into Counters (User_No,Cashup_No,Finalized,Counted,Shift_Start) values (" & UserRecord.User_Number & ", " & TillData.Cashup_No & ",0,0,'" & ClockinTime & "')"
            End If
            rs.Close
    
            ActiveUpdateServer "Update Counters set " & _
            "PayOuts_Value = isnull(PayOuts_Value,0) +" & (Pay1_Value) & _
            ",PayOuts_Qty=isnull(PayOuts_Qty,0) +" & 1 & _
            " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
            DoEvents
            
            With frmSales
                ActiveReadServer "Select Supplier_Name from Suppliers where Supplier_No = '" & frmPayout.Tag & "'"
                If rs.RecordCount > 0 Then
                    Debtor_Name = rs.Fields("Supplier_Name")
                End If
                rs.Close
                .grdMain.Rows = 4
                .grdMain.TextMatrix(1, 0) = ""
                .grdMain.TextMatrix(1, 1) = "Cash Payout"
                .grdMain.TextMatrix(1, 2) = Format(Pay1_Value, "0.00")
                .grdMain.TextMatrix(2, 1) = "Balance Account No - " & frmPayout.Tag
                .grdMain.TextMatrix(2, 2) = Format(Balance - Pay1_Value, "0.00")
                .grdMain.Cell(flexcpBackColor, 2, 0, 2, 2) = &HC0FFC0
                .grdMain.TextMatrix(3, 1) = "Cash"
                .grdMain.TextMatrix(3, 2) = Format(Pay1_Value, "0.00")
                .grdMain.Cell(flexcpBackColor, 3, 0, 3, 2) = &HC0FFFF
                .lblTender.Caption = Format(Pay1_Value, "0.00")
                .lblKeyRegister = "Payment to: " & frmPayout.Tag & " - " & Debtor_Name
                .lblCash = "Cash"
                .lblDigit.Caption = ""
            End With
            Print_Payout 0, 0, frmPayout.Tag, frmSales.lblTender.Caption, 1, Payment_No, "Cash"
            DrawerKick "Cash"
            DoEvents
            GlobalMode = TillMode.FinMode
    End Select
    Unload frmPayout
    KeyRegister = ""
    frmSales.lblPayReason = ""
End Sub
Public Function Gotcredit()

If TillData.Creditbalance + TillData.SaleTotal > TillData.Creditlimit Then
                    Finalizing = False
                    Select Case Panel_no
                    Case 0
                            frmSales.cmdErr.Caption = "Credit Limit Exceeded for this Account"
                            frmSales.cmdErr.Visible = True
                            frmSales.errTimer.Enabled = True
                            Gotcredit = False
                            Exit Function
                            
                     Case 2
                            frmBar.cmdErr.Caption = "Credit Limit Exceeded for this Account"
                            frmBar.cmdErr.Visible = True
                            frmBar.errTimer.Enabled = True
                            Gotcredit = False
                            Exit Function
                            
                    End Select
                    
                    
                    End If
                    Gotcredit = True
End Function

'Kotie 17-03-2013  20:57
Public Function Clear_TillData()
    Creditlimit = 0
    Creditbalance = 0
    ProductCode = ""
    DeptNo = ""
    Description = ""
    ShortDesc = ""
    Qty = ""
    TaxRate = 0
    TaxType = 0
    UnitSize = ""
    Cost = 0
    Price = ""
    Keystring = ""
    ExtraFunc = ""
    Change = 0
    Tendered = 0
    DocNo = 0
    TransNo = 0
    KeyReg = ""
    SaleTotal = 0
    VoidTotal = 0
    VoidCount = 0
    ReturnTotal = 0
    ReturnCount = 0
    UllageTotal = 0
    UllageCount = 0
    PriceOveride = 0
    Kitchen1 = ""
    Kitchen2 = ""
    Recipe = 0
    Cashup_No = 0
    Cash = 0
    Card = 0
    Cheque = 0
    CashCol = 0
    CardCol = 0
    ChequeCol = 0
    Charge = 0
    Loyalty = 0
    TaxTotal = 0
    TaxableSales = 0
    NonTaxableSales = 0
    CollectedTax = 0
    CalculatedTax = 0
    Corrects = 0
    CorrectCount = 0
    TableNo = 0
    TabName = ""
    TabNo = 0
    Covers = 0
    Tipp = 0
    TippCount = 0
    ShortTender = False
    UserOveride = 0
    Discount = 0
    DiscountVal = 0
    TotDiscount = 0
    TotDiscountVal = 0
    TotDiscountCount = 0
    TotDiscountValCount = 0
    Room_No = 0
    Account_No = ""
    Res_No = 0
    Weight = 0
    Deposit = 0
    TillData.Change = 0
    TillData.ReturnTotal = 0
    TillData.UllageTotal = 0
    TillData.VoidTotal = 0
    TillData.Tendered = 0
    TillData.Cash = 0
    TillData.Card = 0
    TillData.Cheque = 0
    TillData.Charge = 0
    TillData.Loyalty = 0
    TillData.TaxTotal = 0
    TillData.TaxableSales = 0
    TillData.NonTaxableSales = 0
    TillData.CollectedTax = 0
    TillData.CalculatedTax = 0
    TillData.Corrects = 0
    TillData.TabNo = 0
    TillData.TabName = ""
    TillData.CorrectCount = 0
    TillData.VoidCount = 0
    TillData.ReturnCount = 0
    TillData.UllageCount = 0
    TillData.Tipp = 0
    TillData.TippCount = 0
    TillData.ShortTender = False
    TillData.UserOveride = 0
    TillData.Discount = 0
    TillData.DiscountVal = 0
    TillData.TotDiscount = 0
    TillData.TotDiscountVal = 0
    TillData.TotDiscountCount = 0
    TillData.TotDiscountValCount = 0
    TillData.Account_No = ""
    TillData.Room_No = 0
    TillData.Res_No = 0
    TillData.Print_Count = 0
    TillData.Table_Name = ""
    'TillData.DocNo = 0
End Function

Public Function Delay(Optional i As Integer = 1)
    T1 = Now
    While (DateDiff("s", T1, Now) < i)
    
    Wend
    
End Function


