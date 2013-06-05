VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmTelTrace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TelTrace SOHO Exchange"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   Icon            =   "frmTelTrace.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkActivate 
      Caption         =   "Activated"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   8100
      Width           =   2145
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1020
      TabIndex        =   4
      Text            =   "C:\"
      Top             =   180
      Width           =   4965
   End
   Begin VSFlex8Ctl.VSFlexGrid grdRoom 
      Height          =   7320
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   5985
      _cx             =   10557
      _cy             =   12912
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorSel    =   16639711
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTelTrace.frx":000C
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
      ExplorerBar     =   0
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
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9540
         TabIndex        =   1
         Top             =   11700
         Width           =   1005
      End
   End
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   345
      Left            =   4860
      TabIndex        =   2
      Top             =   8130
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Close"
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
   Begin btButtonEx.ButtonEx cmdOk 
      Height          =   345
      Left            =   3630
      TabIndex        =   3
      Top             =   8130
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   180
      Width           =   1185
   End
   Begin MSForms.Image Image1 
      Height          =   405
      Index           =   2
      Left            =   960
      Top             =   90
      Width           =   5085
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "8969;714"
   End
   Begin MSForms.Image Image6 
      Height          =   7995
      Left            =   0
      Top             =   0
      Width           =   6150
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "10848;14102"
   End
End
Attribute VB_Name = "frmTelTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    ActiveUpdateServer "Delete from Tel_Ex"
    For i = 1 To grdRoom.Rows - 1
        ActiveUpdateServer "Insert Into Tel_Ex (Room_No,Tel_Ex) values ('" & grdRoom.TextMatrix(i, 0) & "','" & grdRoom.TextMatrix(i, 2) & "')"
    Next i
    SaveSetting Trim(gblApp_Name), "Workstation", "Exchange", txtPath.Text
    SaveSetting Trim(gblApp_Name), "Workstation", "Ex_Activated", chkActivate.Value
    Unload Me
End Sub

Private Sub Form_Load()
    grdRoom.Col = 3
    grdRoom.Rows = 1
    grdRoom.TextMatrix(0, 0) = "Room No."
    grdRoom.TextMatrix(0, 1) = "Description"
    grdRoom.TextMatrix(0, 2) = "Teltrace Ex Number"
    grdRoom.ColAlignment(0) = flexAlignLeftCenter
    grdRoom.ColAlignment(1) = flexAlignLeftCenter
    grdRoom.ColAlignment(2) = flexAlignLeftCenter
    grdRoom.ColWidth(0) = grdRoom.Width * 0.2
    grdRoom.ColWidth(1) = grdRoom.Width * 0.5
    grdRoom.ColWidth(2) = grdRoom.Width * 0.3
    ActiveReadServer "Select * from Room_Ex_View order by convert(int,Room_No)"
    While Not rs.EOF
        grdRoom.Rows = grdRoom.Rows + 1
        grdRoom.TextMatrix(grdRoom.Rows - 1, 0) = rs.Fields("Room_No")
        grdRoom.TextMatrix(grdRoom.Rows - 1, 1) = rs.Fields("Description")
        grdRoom.TextMatrix(grdRoom.Rows - 1, 2) = rs.Fields("Tel_Ex") & ""
        rs.MoveNext
    Wend
    rs.Close
    txtPath.Text = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", Key:="Exchange", Default:="C:\")
    chkActivate = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", Key:="Ex_Activated", Default:=0)
    Tel_Ex = chkActivate
End Sub
Private Sub grdRoom_EnterCell()
    If grdRoom.Col = 2 Then
        grdRoom.Editable = flexEDKbdMouse
    Else
        grdRoom.Editable = flexEDNone
    End If
End Sub
