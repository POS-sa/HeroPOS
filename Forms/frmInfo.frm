VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Occupation"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   2550
      ScaleHeight     =   2925
      ScaleWidth      =   5355
      TabIndex        =   18
      Top             =   570
      Visible         =   0   'False
      Width           =   5355
      Begin MSComCtl2.MonthView mthView 
         Height          =   2310
         Left            =   90
         TabIndex        =   20
         Top             =   90
         Width           =   5160
         _ExtentX        =   9102
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
         MonthColumns    =   2
         MonthBackColor  =   16777215
         MultiSelect     =   -1  'True
         ShowToday       =   0   'False
         StartOfWeek     =   25755650
         TitleBackColor  =   16761281
         TrailingForeColor=   -2147483639
         CurrentDate     =   38701
      End
      Begin btButtonEx.ButtonEx cmdOk 
         Height          =   315
         Left            =   4110
         TabIndex        =   19
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
      Begin MSForms.Image Image12 
         Height          =   2925
         Left            =   0
         Top             =   0
         Width           =   5355
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "9446;5159"
      End
      Begin MSForms.Image Image11 
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
   Begin btButtonEx.ButtonEx cmdSupplier 
      Height          =   735
      Left            =   6750
      TabIndex        =   8
      ToolTipText     =   " Click to Search.... "
      Top             =   2730
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1296
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
      Style           =   1
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   735
      Left            =   5520
      TabIndex        =   9
      ToolTipText     =   " Click to Search.... "
      Top             =   2730
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1296
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
      Style           =   1
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx ButtonEx2 
      Height          =   375
      Left            =   7410
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   180
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
   Begin BTNENHLib4.BtnEnh fmType 
      Height          =   1275
      Left            =   4230
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   900
      Width           =   2955
      _Version        =   524298
      _ExtentX        =   5212
      _ExtentY        =   2249
      _StockProps     =   66
      Caption         =   "0 %"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   1
      CornerFactor    =   10
      Surface         =   3
      PictureTranspColor=   192
      BackColorContainer=   14215660
      ShadowColor     =   16777215
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      ForeColorDisabled=   12640511
      UserData        =   0.1
      textCaption     =   "frmInfo.frx":000C
      textLT          =   "frmInfo.frx":0072
      textCT          =   "frmInfo.frx":008A
      textRT          =   "frmInfo.frx":00A2
      textLM          =   "frmInfo.frx":00BA
      textRM          =   "frmInfo.frx":00D2
      textLB          =   "frmInfo.frx":00EA
      textCB          =   "frmInfo.frx":0102
      textRB          =   "frmInfo.frx":011A
      colorBack       =   "frmInfo.frx":0132
      colorIntern     =   "frmInfo.frx":015C
      colorMO         =   "frmInfo.frx":0186
      colorFocus      =   "frmInfo.frx":01B0
      colorDisabled   =   "frmInfo.frx":01DA
      colorPressed    =   "frmInfo.frx":0204
      HollowFrame     =   -1  'True
      LightDirection  =   8
   End
   Begin VB.Label lblProv 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   555
      Width           =   1530
   End
   Begin MSForms.Label lblDate 
      Height          =   315
      Left            =   3960
      TabIndex        =   17
      Top             =   270
      Width           =   3555
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "1 Feb 2006 to 13 Feb 2006"
      Size            =   "6271;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblAv 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   2190
      Width           =   1530
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rooms Available:"
      Height          =   195
      Left            =   135
      TabIndex        =   14
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label lblSold 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Top             =   1875
      Width           =   1530
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rooms Sold:"
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   1890
      Width           =   1695
   End
   Begin VB.Label lblOut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   1545
      Width           =   1530
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Checked Out:"
      Height          =   195
      Left            =   135
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1530
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Checked In:"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   1230
      Width           =   1695
   End
   Begin VB.Label lblCon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   870
      Width           =   1530
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmed:"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label lblRooms 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   210
      Width           =   1530
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Room Quantity:"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Provitional Bookings:"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   570
      Width           =   1695
   End
   Begin MSForms.Image Image7 
      Height          =   285
      Left            =   1890
      Top             =   180
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   13827793
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image8 
      Height          =   285
      Left            =   1890
      Top             =   1830
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   13827793
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Left            =   1890
      Top             =   1500
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image3 
      Height          =   285
      Left            =   1890
      Top             =   1170
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image2 
      Height          =   285
      Left            =   1890
      Top             =   840
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image9 
      Height          =   285
      Left            =   1890
      Top             =   2160
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   12648447
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image4 
      Height          =   285
      Left            =   1890
      Top             =   510
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image10 
      Height          =   375
      Left            =   3660
      Top             =   180
      Width           =   3705
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6535;661"
   End
   Begin MSForms.Image Image1 
      Height          =   2445
      Index           =   0
      Left            =   60
      Top             =   90
      Width           =   7875
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "13891;4313"
   End
   Begin MSForms.Image Image6 
      Height          =   2625
      Left            =   30
      Top             =   0
      Width           =   7980
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "14076;4630"
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonEx1_Click()
    Unload Me
End Sub
Private Sub ButtonEx2_Click()
    Select Case ButtonEx2.Value
        Case 0
            picDate.Visible = True
        Case 1
            picDate.Visible = False
            If picDate.Visible = False Then Selection_Change
    End Select
End Sub
Private Sub cmdOk_Click()
    picDate.Visible = False
    If picDate.Visible = False Then Selection_Change
End Sub
Private Sub cmdSupplier_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Selection_Change
End Sub

Private Sub Form_Load()
    mthView.SelStart = frmRes.grdRes.TextMatrix(2, 2)
    mthView.SelEnd = frmRes.grdRes.TextMatrix(2, 32)
    lblDate.Caption = Format(mthView.SelStart, "DD MMM YYYY") & " to " & Format(mthView.SelEnd, "DD MMM YYYY")
    DoEvents
    
End Sub
Private Sub mthView_LostFocus()
    DoEvents
    If picDate.Visible = False Then Selection_Change
End Sub
Private Sub mthView_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    lblDate.Caption = Format(mthView.SelStart, "DD MMM YYYY") & " to " & Format(mthView.SelEnd, "DD MMM YYYY")
End Sub
Private Sub Selection_Change()
    On Error Resume Next
    ActiveReadServer "Select Count(Room_No) as RoomsQty from Rooms"
    If rs.RecordCount > 0 Then
        Rooms = Val(rs.Fields("RoomsQty") & "")
    End If
    rs.Close
    DoEvents
    ActiveReadServer "Select Arrive_Date,Depart_Date,Res_No,Room_No,Res_Type,Title + ' ' + First_Name + ' ' + Last_name as Guest_Name,Res_Type from Reservations " & _
    " where (Arrive_Date > '" & DateAdd("D", -1, mthView.SelStart) & "'and Arrive_Date < '" & DateAdd("D", 1, mthView.SelEnd) & "')" & _
    " or (Depart_Date > '" & DateAdd("D", -1, mthView.SelStart) & "'and Depart_Date < '" & DateAdd("D", 1, mthView.SelEnd) & "')" & _
    " order by Room_No"
    lblProv.Caption = "0"
    lblCon.Caption = "0"
    lblCheck.Caption = "0"
    lblOut.Caption = "0"
    Days = (DateDiff("d", mthView.SelStart, mthView.SelEnd)) + 1
    lblRooms.Caption = Days * Rooms
    While Not rs.EOF
        Select Case rs.Fields("Res_Type")
            Case 0: lblProv.Caption = Val(lblProv.Caption) + DateDiff("d", rs.Fields("Arrive_Date"), rs.Fields("Depart_Date"))
            Case 1: lblCon.Caption = Val(lblCon.Caption) + DateDiff("d", rs.Fields("Arrive_Date"), rs.Fields("Depart_Date"))
            Case 2: lblCheck.Caption = Val(lblCheck.Caption) + DateDiff("d", rs.Fields("Arrive_Date"), rs.Fields("Depart_Date"))
            Case 3: lblOut.Caption = Val(lblOut.Caption) + DateDiff("d", rs.Fields("Arrive_Date"), rs.Fields("Depart_Date"))
        End Select
        rs.MoveNext
    Wend
    
    lblSold.Caption = Val(lblProv.Caption) + Val(lblCon.Caption) + Val(lblCheck.Caption) + Val(lblOut.Caption)
    lblAv.Caption = Val(lblRooms.Caption) - Val(lblSold.Caption)
    
    rs.Close
    fmType.Caption = Round((Val(lblSold.Caption) / Val(lblRooms.Caption)) * 100, 2) & " %"
    On Error GoTo 0
End Sub

