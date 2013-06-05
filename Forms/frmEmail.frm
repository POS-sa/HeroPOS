VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmail 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Email run..."
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9264.74
   ScaleMode       =   0  'User
   ScaleWidth      =   12468.22
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex8Ctl.VSFlexGrid grdList 
      Height          =   7650
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   12180
      _cx             =   21484
      _cy             =   13494
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
      BackColorSel    =   15391677
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmEmail.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   5
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   7830
      Visible         =   0   'False
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   375
      Left            =   10920
      TabIndex        =   2
      ToolTipText     =   " Click to Search.... "
      Top             =   8280
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
      Left            =   9630
      TabIndex        =   3
      ToolTipText     =   " Click to Search.... "
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Start"
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
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOpen_Click()
    Screen.MousePointer = 11
    Call Setpdfprinter
    ProgressBar1.Max = grdList.Rows - 1
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    For i = 1 To grdList.Rows - 1
        grdList.Row = i
        ProgressBar1.Value = ProgressBar1.Value + 1
        If grdList.ValueMatrix(i, 4) <> 0 Then
         '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
          Dim Accnum As String, Accnames As String, Accemail As String
         Accnum = grdList.TextMatrix(i, 0)
         Accnames = grdList.TextMatrix(i, 1)
         Accemail = grdList.TextMatrix(i, 3)
         
         Emailattachment Accnum, Accnames, Accemail
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        End If
        DoEvents
    Next i
    Screen.MousePointer = 0
    
    
    MsgBox "Email Run Completed", vbInformation, "HeroPOS"
    Unload Me
End Sub



Private Sub Form_Load()
Loadall
End Sub

Private Sub grdList_EnterCell()
    If grdList.Col = 4 Then
   
        grdList.Editable = flexEDKbdMouse
    Else
        grdList.Editable = flexEDNone
    End If
End Sub

Private Sub Loadall()
    grdList.Cols = 5
    grdList.TextMatrix(0, 0) = " Debtor No"
    grdList.TextMatrix(0, 1) = " Debtor Name"
    grdList.TextMatrix(0, 2) = " Balance"
    grdList.TextMatrix(0, 3) = " Email adress"
    grdList.TextMatrix(0, 4) = " Select"
    
    grdList.ColAlignment(0) = flexAlignLeftCenter
    grdList.ColAlignment(1) = flexAlignLeftCenter
    grdList.ColAlignment(2) = flexAlignRightCenter
    grdList.ColAlignment(3) = flexAlignLeftCenter
    grdList.ColAlignment(4) = flexAlignCenterCenter
    grdList.ColDataType(4) = flexDTBoolean
    grdList.ColWidth(0) = grdList.Width * 0.15
    grdList.ColWidth(1) = grdList.Width * 0.3
    grdList.ColWidth(2) = grdList.Width * 0.2
    grdList.ColWidth(3) = grdList.Width * 0.25
    grdList.ColWidth(4) = grdList.Width * 0.15
    grdList.Rows = 1
    ActiveReadServer "Select * from Debtors order by Debtor_Name"
    While Not rs.EOF
        grdList.Rows = grdList.Rows + 1
        grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Debtor_No")
        grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("Debtor_Name")
        grdList.TextMatrix(grdList.Rows - 1, 2) = Format(rs.Fields("Balance"), "0.00")
        grdList.TextMatrix(grdList.Rows - 1, 3) = rs.Fields("E_Mail")
        grdList.TextMatrix(grdList.Rows - 1, 4) = ""
        
        rs.MoveNext
    Wend
    rs.Close
End Sub
