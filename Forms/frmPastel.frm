VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmPastel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pastel Partner Interface"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   Icon            =   "frmPastel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   90
      TabIndex        =   9
      Top             =   3600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.PictureBox picDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   5250
      ScaleHeight     =   2925
      ScaleWidth      =   5355
      TabIndex        =   6
      Top             =   510
      Visible         =   0   'False
      Width           =   5355
      Begin btButtonEx.ButtonEx cmdOk 
         Height          =   315
         Left            =   4110
         TabIndex        =   7
         Top             =   2460
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Appearance      =   3
         Caption         =   "Ok"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.MonthView mthViewStart 
         Height          =   2310
         Left            =   90
         TabIndex        =   8
         Top             =   90
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16239822
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxSelCount     =   365
         MonthBackColor  =   16777215
         MultiSelect     =   -1  'True
         ShowToday       =   0   'False
         StartOfWeek     =   66125826
         TitleBackColor  =   16761281
         TrailingForeColor=   -2147483639
         CurrentDate     =   38701
      End
      Begin MSComCtl2.MonthView mthViewEnd 
         Height          =   2310
         Left            =   2610
         TabIndex        =   11
         Top             =   90
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16239822
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxSelCount     =   365
         MonthBackColor  =   16777215
         MultiSelect     =   -1  'True
         ShowToday       =   0   'False
         StartOfWeek     =   66125826
         TitleBackColor  =   16761281
         TrailingForeColor=   -2147483639
         CurrentDate     =   38701
      End
      Begin MSForms.Image Image5 
         Height          =   2925
         Left            =   0
         Top             =   0
         Width           =   5355
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "9446;5159"
      End
      Begin MSForms.Image Image6 
         Height          =   2805
         Left            =   60
         Top             =   60
         Width           =   5235
         BackColor       =   16777215
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "9234;4948"
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grdPas 
      Height          =   2895
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   10560
      _cx             =   18627
      _cy             =   5106
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   15329975
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   700
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPastel.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   375
      Left            =   9420
      TabIndex        =   2
      ToolTipText     =   " Click to Search.... "
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdOpen 
      Height          =   375
      Left            =   7860
      TabIndex        =   3
      ToolTipText     =   " Click to Search.... "
      Top             =   3600
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Generate Batch"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   375
      Left            =   10140
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   105
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   8421504
      Caption         =   "¦"
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx ButtonEx2 
      Height          =   375
      Left            =   6300
      TabIndex        =   10
      ToolTipText     =   " Click to Search.... "
      Top             =   3600
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin MSForms.Label lblDate 
      Height          =   315
      Left            =   7800
      TabIndex        =   5
      Top             =   210
      Width           =   2235
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "1 Feb 2006 to 13 Feb 2006"
      Size            =   "3942;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblMaj 
      BackStyle       =   0  'Transparent
      Caption         =   "Export Settings"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Width           =   2715
   End
   Begin MSForms.Image Image2 
      Height          =   375
      Left            =   7740
      Top             =   105
      Width           =   2385
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "4207;661"
   End
   Begin MSForms.Image picMaj 
      Height          =   495
      Left            =   60
      Top             =   60
      Width           =   10560
      BackColor       =   14737632
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "18627;873"
   End
End
Attribute VB_Name = "frmPastel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonEx1_Click()
    Select Case ButtonEx1.Value
        Case 0
            picDate.Visible = True
        Case 1
            picDate.Visible = False
            If picDate.Visible = False Then Selection_Change
    End Select
End Sub

Private Sub ButtonEx2_Click()
    MsgBox "Setting Saved", vbInformation, "HeroPOS"
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOpen_Click()
     ActiveReadServer "Select * from Branch_Details"
    If rs.RecordCount > 0 Then
        If rs.Fields("Fin_Year") & "" = "" Then
            DTPeriod = Format(mthViewStart.Value, "M") - 1
        Else
            DTPeriod = Format(mthViewStart.Value, "M") - Format(rs.Fields("Fin_Year"), "M") - 1
        End If
    End If
    rs.Close
    If grdPas.ValueMatrix(1, 1) = -1 Then
        
    End If
    If grdPas.ValueMatrix(2, 1) = -1 Then
        If Right(Str(Time_Stop), 2) = "AM" And mthViewStart.Value = mthViewEnd.Value Then
            Selender = DateAdd("d", 1, mthViewEnd.Value)
            lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
        Else
            Selender = mthViewEnd.Value
        End If
        FunctionKey = 16
        ActiveReadServer "Select * from Purchase_Journal_View where Supplier_No is Not Null and (Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "') order by Date_Time"
        filenum = FreeFile
        Dim fso As New FileSystemObject
        If fso.FolderExists(App.Path & "\Export") = False Then fso.CreateFolder (App.Path & "\Export")
        Open Trim(App.Path) & "\Export\Purchase.Dat" For Output As filenum
        ProgressBar1.Value = 0
        ProgressBar1.Max = rs.RecordCount + 1
        While Not rs.EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            Print #filenum, DTPeriod & "," & Chr(34) & Format(rs.Fields("Date_Time"), "DD/MM/YYYY") & Chr(34) & _
            "," & Chr(34) & "G" & Chr(34) & _
            "," & Chr(34) & rs.Fields("Supplier_No") & Chr(34) & _
            "," & Chr(34) & "PURCHASE" & Chr(34) & _
            "," & Chr(34) & "HeroPOS IMPORT" & Chr(34) & _
            "," & rs.Fields("Line_Total") * -1 & _
            "," & "1" & _
            "," & rs.Fields("Vat_Rate") * -1 & _
            "," & Chr(34) & " " & Chr(34) & _
            "," & Chr(34) & "     " & Chr(34) & _
            "," & Chr(34) & GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Purchases4", Default:="") & Chr(34) & _
            ",1,1,0,0,0,0"
            rs.MoveNext
        Wend
        rs.Close
        Close #filenum
    End If
    If grdPas.ValueMatrix(3, 1) = -1 Then
        
    End If
    If grdPas.ValueMatrix(4, 1) = -1 Then
        
    End If
    If grdPas.ValueMatrix(5, 1) = -1 Then
        
    End If
    If grdPas.ValueMatrix(6, 1) = -1 Then
        
    End If
    Unload Me
End Sub
Private Sub Form_Load()
    grdPas.Rows = 7
    mthViewStart.Value = Date
    mthViewEnd.Value = Date
    lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(mthViewEnd.Value, "DD MMM YYYY")
    grdPas.ColWidth(0) = grdPas.Width * 0.2
    grdPas.ColWidth(1) = grdPas.Width * 0.1
    grdPas.ColWidth(2) = grdPas.Width * 0.3
    grdPas.ColWidth(3) = grdPas.Width * 0.2
    grdPas.ColWidth(4) = grdPas.Width * 0.2
    grdPas.ColAlignment(3) = flexAlignLeftCenter
    grdPas.ColAlignment(4) = flexAlignLeftCenter
    grdPas.TextMatrix(0, 0) = "Export Type"
    grdPas.TextMatrix(0, 1) = "Active"
    grdPas.TextMatrix(0, 2) = "Action"
    grdPas.TextMatrix(0, 3) = "Debit Account"
    grdPas.TextMatrix(0, 4) = "Credit Account"
    grdPas.TextMatrix(1, 0) = "Sales"
    grdPas.TextMatrix(2, 0) = "Purchases"
    grdPas.TextMatrix(3, 0) = "Stock"
    grdPas.TextMatrix(4, 0) = "Expences"
    grdPas.TextMatrix(5, 0) = "Debtors"
    grdPas.TextMatrix(6, 0) = "Creditors"
    grdPas.TextMatrix(1, 2) = "Sales by Location"
    grdPas.TextMatrix(2, 2) = "Line Invoices and Payments"
    grdPas.TextMatrix(3, 2) = "Stock by Location"
    grdPas.TextMatrix(4, 2) = "Expences by Location"
    grdPas.TextMatrix(5, 2) = "Line Invoices and Payments"
    grdPas.TextMatrix(6, 2) = "Creditors Balances"
    grdPas.TextMatrix(1, 3) = "Use Location GL"
    grdPas.TextMatrix(2, 3) = "<Please Supply>"
    grdPas.TextMatrix(3, 3) = "Use Department GL"
    grdPas.TextMatrix(4, 3) = "Use Location GL"
    grdPas.TextMatrix(5, 3) = "Use Location GL"
    grdPas.TextMatrix(6, 3) = "<Please Supply>"
    grdPas.TextMatrix(1, 4) = "<Please Supply>"
    grdPas.TextMatrix(2, 4) = "Use Supplier GL"
    grdPas.TextMatrix(3, 4) = "<Please Supply>"
    grdPas.TextMatrix(4, 4) = "<Please Supply>"
    grdPas.TextMatrix(5, 4) = "<Please Supply>"
    grdPas.TextMatrix(6, 4) = "Use Supplier GL"
    grdPas.ColDataType(1) = flexDTBoolean
    grdPas.ColAlignment(1) = flexAlignCenterCenter
    grdPas.TextMatrix(1, 1) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Sales", Default:=0)
    grdPas.TextMatrix(2, 1) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Purchases", Default:=0)
    grdPas.TextMatrix(3, 1) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Stock", Default:=0)
    grdPas.TextMatrix(4, 1) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Expences", Default:=0)
    grdPas.TextMatrix(5, 1) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Debtors", Default:=0)
    grdPas.TextMatrix(6, 1) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Creditors", Default:=0)
    
    grdPas.TextMatrix(1, 3) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Sales3", Default:="Use Location GL")
    grdPas.TextMatrix(2, 3) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Purchases3", Default:="<Please Supply>")
    grdPas.TextMatrix(3, 3) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Stock3", Default:="Use Department GL")
    grdPas.TextMatrix(4, 3) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Expences3", Default:="Use Location GL")
    grdPas.TextMatrix(5, 3) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Debtors3", Default:="Use Location GL")
    grdPas.TextMatrix(6, 3) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Creditors3", Default:="<Please Supply>")

    grdPas.TextMatrix(1, 4) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Sales4", Default:="<Please Supply>")
    grdPas.TextMatrix(2, 4) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Purchases4", Default:="Use Supplier GL")
    grdPas.TextMatrix(3, 4) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Stock4", Default:="<Please Supply>")
    grdPas.TextMatrix(4, 4) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Expences4", Default:="<Please Supply>")
    grdPas.TextMatrix(5, 4) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Debtors4", Default:="<Please Supply>")
    grdPas.TextMatrix(6, 4) = GetSetting(appname:=Trim(gblApp_Name), Section:="Pastel", key:="Creditors4", Default:="Use Supplier GL")

End Sub

Private Sub grdPas_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
        Case 1
            Select Case Row
                Case 1
                    SaveSetting Trim(gblApp_Name), "Pastel", "Sales", grdPas.TextMatrix(1, 1)
                Case 2
                    SaveSetting Trim(gblApp_Name), "Pastel", "Purchases", grdPas.TextMatrix(2, 1)
                Case 3
                    SaveSetting Trim(gblApp_Name), "Pastel", "Stock", grdPas.TextMatrix(3, 1)
                Case 4
                    SaveSetting Trim(gblApp_Name), "Pastel", "Expences", grdPas.TextMatrix(4, 1)
                Case 5
                    SaveSetting Trim(gblApp_Name), "Pastel", "Debtors", grdPas.TextMatrix(5, 1)
                Case 6
                    SaveSetting Trim(gblApp_Name), "Pastel", "Creditors", grdPas.TextMatrix(6, 1)
            End Select
        Case 3
            Select Case Row
                Case 1
                    SaveSetting Trim(gblApp_Name), "Pastel", "Sales3", grdPas.TextMatrix(1, 3)
                Case 2
                    SaveSetting Trim(gblApp_Name), "Pastel", "Purchases3", grdPas.TextMatrix(2, 3)
                Case 3
                    SaveSetting Trim(gblApp_Name), "Pastel", "Stock3", grdPas.TextMatrix(3, 3)
                Case 4
                    SaveSetting Trim(gblApp_Name), "Pastel", "Expences3", grdPas.TextMatrix(4, 3)
                Case 5
                    SaveSetting Trim(gblApp_Name), "Pastel", "Debtors3", grdPas.TextMatrix(5, 3)
                Case 6
                    SaveSetting Trim(gblApp_Name), "Pastel", "Creditors3", grdPas.TextMatrix(6, 3)
            End Select
        Case 4
            Select Case Row
                Case 1
                    SaveSetting Trim(gblApp_Name), "Pastel", "Sales4", grdPas.TextMatrix(1, 4)
                Case 2
                    SaveSetting Trim(gblApp_Name), "Pastel", "Purchases4", grdPas.TextMatrix(2, 4)
                Case 3
                    SaveSetting Trim(gblApp_Name), "Pastel", "Stock4", grdPas.TextMatrix(3, 4)
                Case 4
                    SaveSetting Trim(gblApp_Name), "Pastel", "Expences4", grdPas.TextMatrix(4, 4)
                Case 5
                    SaveSetting Trim(gblApp_Name), "Pastel", "Debtors4", grdPas.TextMatrix(5, 4)
                Case 6
                    SaveSetting Trim(gblApp_Name), "Pastel", "Creditors4", grdPas.TextMatrix(6, 4)
            End Select
    End Select
End Sub

Private Sub grdPas_EnterCell()
    Select Case grdPas.Col
        Case 1
            grdPas.Editable = flexEDKbdMouse
        Case 2
            grdPas.ColComboList(2) = ""
            Select Case grdPas.TextMatrix(grdPas.Row, 0)
                Case "Sales"
                    grdPas.ColComboList(2) = "Sales by Location|Sales by Department|Detailed Sales Invoices"
                Case "Purchases"
                    grdPas.ColComboList(2) = "Line Invoices and Payments|Detailed Line Invoices and Payments"
                Case "Stock"
                
                Case "Expences"
                
                Case "Debtors"
                
                Case "Creditors"
                
            End Select
        Case 3, 4
            If InStr(grdPas.TextMatrix(grdPas.Row, grdPas.Col), "<") <> 0 Then
                grdPas.Editable = flexEDKbdMouse
            Else
                grdPas.Editable = flexEDNone
            End If
    End Select
End Sub
Private Sub mthView_LostFocus()
    DoEvents
    If picDate.Visible = False Then Selection_Change
End Sub
Private Sub mthView_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(mthViewEnd.Value, "DD MMM YYYY")
End Sub
Private Sub Selection_Change()

End Sub
Private Sub cmdOk_Click()
    picDate.Visible = False
    If picDate.Visible = False Then Selection_Change
End Sub

