VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInquiry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Inquiry"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9660
   Icon            =   "frmInquiry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1620
      TabIndex        =   1
      Top             =   690
      Width           =   7815
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4080
      TabIndex        =   18
      Top             =   300
      Width           =   4395
   End
   Begin VB.TextBox txtProdCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1620
      TabIndex        =   0
      Top             =   300
      Width           =   2265
   End
   Begin VB.PictureBox picDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   60
      ScaleHeight     =   915
      ScaleWidth      =   5265
      TabIndex        =   14
      Top             =   4500
      Visible         =   0   'False
      Width           =   5295
      Begin btButtonEx.ButtonEx cmdOk 
         Height          =   315
         Left            =   3960
         TabIndex        =   15
         Top             =   510
         Width           =   1155
         _ExtentX        =   2037
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
      Begin MSComCtl2.DTPicker mthViewStart 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy (ddd)"
         Format          =   68288515
         CurrentDate     =   39414
      End
      Begin MSComCtl2.DTPicker mthViewEnd 
         Height          =   315
         Left            =   2670
         TabIndex        =   20
         Top             =   120
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy (ddd)"
         Format          =   68288515
         CurrentDate     =   39414
      End
      Begin MSForms.Image Image10 
         Height          =   915
         Left            =   0
         Top             =   0
         Width           =   5265
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "9287;1614"
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grdProd 
      Height          =   90
      Left            =   60
      TabIndex        =   11
      Top             =   1050
      Visible         =   0   'False
      Width           =   9525
      _cx             =   16801
      _cy             =   159
      Appearance      =   2
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmInquiry.frx":000C
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
   Begin VSFlex8Ctl.VSFlexGrid grdStock 
      Height          =   3660
      Left            =   90
      TabIndex        =   2
      Top             =   1590
      Width           =   5895
      _cx             =   10398
      _cy             =   6456
      Appearance      =   2
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
      BackColorSel    =   12582912
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   11
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmInquiry.frx":0084
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   End
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   345
      Left            =   8400
      TabIndex        =   4
      Top             =   5520
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
   Begin VSFlex8Ctl.VSFlexGrid grdRev 
      Height          =   1350
      Left            =   6090
      TabIndex        =   7
      Top             =   1590
      Width           =   3465
      _cx             =   6112
      _cy             =   2381
      Appearance      =   2
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
      BackColorSel    =   12582912
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   1800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmInquiry.frx":00FC
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   End
   Begin VSFlex8Ctl.VSFlexGrid grdRev1 
      Height          =   1350
      Left            =   6090
      TabIndex        =   8
      Top             =   3540
      Width           =   3465
      _cx             =   6112
      _cy             =   2381
      Appearance      =   2
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
      BackColorSel    =   12582912
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   1800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmInquiry.frx":0174
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   End
   Begin btButtonEx.ButtonEx cmdSearch 
      Height          =   345
      Left            =   180
      TabIndex        =   10
      Top             =   210
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Product Code..."
      CaptionOffsetY  =   1
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
      Height          =   375
      Left            =   2460
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5490
      Width           =   435
      _ExtentX        =   767
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
   Begin MSForms.Image Image13 
      Height          =   345
      Left            =   3990
      Top             =   210
      Width           =   4545
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "8017;609"
   End
   Begin MSForms.Image Image12 
      Height          =   345
      Left            =   1530
      Top             =   600
      Width           =   7965
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "14049;609"
   End
   Begin MSForms.Image Image11 
      Height          =   345
      Left            =   1530
      Top             =   210
      Width           =   2385
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "4207;609"
   End
   Begin MSForms.Label lblInfo 
      Height          =   255
      Left            =   6150
      TabIndex        =   17
      Top             =   5040
      Width           =   3405
      ForeColor       =   4210752
      Caption         =   "Ave Cost: 0.00    Selling: 0.00"
      Size            =   "6006;450"
      FontName        =   "Arial Narrow"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox cmbLocs 
      Height          =   375
      Left            =   2970
      TabIndex        =   16
      Tag             =   "Up"
      Top             =   5490
      Width           =   5355
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "9446;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      BorderColor     =   12632256
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblDate 
      Height          =   285
      Left            =   180
      TabIndex        =   13
      Top             =   5550
      Width           =   2235
      ForeColor       =   4210752
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "1 Feb 2006 to 13 Feb 2006"
      Size            =   "3942;503"
      FontName        =   "Arial Narrow"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   3060
      Width           =   3405
      ForeColor       =   8421504
      Caption         =   "Actual GP"
      Size            =   "6006;661"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Image Image6 
      Height          =   1425
      Left            =   6060
      Top             =   3510
      Width           =   3525
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6218;2514"
   End
   Begin MSForms.Image Image5 
      Height          =   1425
      Left            =   6060
      Top             =   1560
      Width           =   3525
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6218;2514"
   End
   Begin MSForms.Label Label5 
      Height          =   345
      Left            =   6120
      TabIndex        =   6
      Top             =   1140
      Width           =   3405
      ForeColor       =   8421504
      Caption         =   "Theoretical GP"
      Size            =   "6006;609"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   345
      Left            =   1260
      TabIndex        =   5
      Top             =   1140
      Width           =   3675
      ForeColor       =   8421504
      Caption         =   "Stock Movement"
      Size            =   "6482;609"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description: "
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   630
      Width           =   1335
   End
   Begin MSForms.Image Image4 
      Height          =   915
      Left            =   60
      Top             =   120
      Width           =   9525
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "16801;1614"
   End
   Begin MSForms.Image Image3 
      Height          =   3730
      Left            =   60
      Top             =   1560
      Width           =   5955
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "10504;6597"
   End
   Begin MSForms.Image Image2 
      Height          =   465
      Left            =   6060
      Top             =   1050
      Width           =   3525
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6218;820"
   End
   Begin MSForms.Image Image1 
      Height          =   465
      Left            =   60
      Top             =   1050
      Width           =   5955
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "10504;820"
   End
   Begin MSForms.Image Image7 
      Height          =   465
      Left            =   6060
      Top             =   3000
      Width           =   3525
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6218;820"
   End
   Begin MSForms.Image Frame1 
      Height          =   6285
      Left            =   0
      Top             =   -900
      Width           =   9675
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "17066;11086"
   End
   Begin MSForms.Image Image8 
      Height          =   375
      Left            =   60
      Top             =   5490
      Width           =   2385
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "4207;661"
   End
End
Attribute VB_Name = "frmInquiry"
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
Private Sub Selection_Change()
    lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(mthViewEnd.Value, "DD MMM YYYY")
    LoadProduct
End Sub
Private Sub LoadProduct()
    If txtProdCode.Text <> "" Then
        If mthViewEnd.Value < Date Then
            grdStock.TextMatrix(10, 1) = "0"
            grdStock.TextMatrix(10, 2) = "0.00"
            lblInfo.Caption = "Ave Cost: 0.00    Selling: 0.00"
            Select Case cmbLocs.Text
                Case "<All Locations>"
                    LocString = "%"
                Case Else
                    LocString = Trim(Mid(cmbLocs.Text, 1, InStr(cmbLocs.Text, "-") - 1))
            End Select
        Else
            'Stock on Hand
            Select Case cmbLocs.Text
                Case "<All Locations>"
                    LocString = "%"
                    ActiveReadServer "SELECT Products.Product_Code, SUM(Quantities.Stock_on_Hand) AS Stock_on_Hand, Products.Ave_Cost, Products.Landed_Cost, " & _
                    " Products.Selling_Price" & _
                    " FROM Products LEFT OUTER JOIN" & _
                    " Quantities ON Products.Product_Code = Quantities.Product_Code where Products.Product_Code ='" & txtProdCode.Text & "'" & _
                    " GROUP BY Products.Product_Code, Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price"
                Case Else
                    LocString = Trim(Mid(cmbLocs.Text, 1, InStr(cmbLocs.Text, "-") - 1))
                    ActiveReadServer "SELECT Products.Product_Code, SUM(Quantities.Stock_on_Hand) AS Stock_on_Hand, Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price , Quantities.Location_No" & _
                    " FROM Products INNER JOIN" & _
                    " Quantities ON Products.Product_Code = Quantities.Product_Code" & _
                    " GROUP BY Products.Product_Code, Quantities.Location_No, Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price" & _
                    " HAVING (Quantities.Location_No = " & Mid(cmbLocs.Text, 1, InStr(cmbLocs.Text, "-") - 1) & ") and (Products.Product_Code = '" & txtProdCode.Text & "')"
            End Select
            If rs.RecordCount > 0 Then
                 grdStock.TextMatrix(10, 1) = Round(Val(rs.Fields("Stock_on_hand") & ""), 3)
                 grdStock.TextMatrix(10, 2) = Format(Val(rs.Fields("Stock_on_hand") & "") * rs.Fields("Ave_Cost"), "0.00")
                 lblInfo.Caption = "Ave Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & "    Selling: " & Format(rs.Fields("Selling_Price"), "0.00")
            Else
                grdStock.TextMatrix(10, 1) = "0"
                grdStock.TextMatrix(10, 2) = "0.00"
                lblInfo.Caption = "Ave Cost: 0.00    Selling: 0.00"
            End If
            rs.Close
        End If
        'Goods Received
        ActiveReadServer "Select sum(Qty_Delivered) as Qty_Received,sum(Qty_Invoiced) as Qty_Invoiced,sum(Line_Total) as Line_Total from Purchase_Journal where Product_Code='" & txtProdCode.Text & "' and Location_No like '" & LocString & "' and " & _
        "Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & mthViewEnd.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' group by Product_Code"
        If rs.RecordCount > 0 Then
            Qty_Received = 0
            If Round(Val(rs.Fields("Qty_Received") & ""), 3) = 0 Then
                Qty_Received = Round(Val(rs.Fields("Qty_Invoiced") & ""), 3)
            Else
                Qty_Received = Round(Val(rs.Fields("Qty_Received") & ""), 3)
            End If
            grdStock.TextMatrix(1, 1) = Qty_Received
            grdStock.TextMatrix(1, 2) = Format(rs.Fields("Line_Total"), "0.00")
        Else
            grdStock.TextMatrix(1, 1) = "0"
            grdStock.TextMatrix(1, 2) = "0.00"
        End If
        rs.Close
        'Stock Takes
        ActiveReadServer "Select sum(Qty_Counted-Qty_on_Hand ) as Variance,Sum((Qty_Counted-Qty_on_Hand )*Ave_Cost) as Line_Total from Stock_Take_Journal where Product_Code='" & txtProdCode.Text & "' and Location_No like '" & LocString & "' and " & _
        "Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & mthViewEnd.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' group by Product_Code"
        If rs.RecordCount > 0 Then
            grdStock.TextMatrix(9, 1) = Round(Val(rs.Fields("Variance") & ""), 3)
            grdStock.TextMatrix(9, 2) = Format(rs.Fields("Line_Total"), "0.00")
        Else
            grdStock.TextMatrix(9, 1) = "0"
            grdStock.TextMatrix(9, 2) = "0.00"
        End If
        rs.Close
        'Stock Consumption
        ActiveReadServer "Select sum(Qty_Consumed) as Qty_Consumed,Sum(Ave_Cost) as Line_Total from Consumption_Journal where Product_Code='" & txtProdCode.Text & "' and Location_No like '" & LocString & "' and " & _
        "Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & mthViewEnd.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' group by Product_Code"
        If rs.RecordCount > 0 Then
            grdStock.TextMatrix(7, 1) = Round(Val(rs.Fields("Qty_Consumed") & ""), 3)
            grdStock.TextMatrix(7, 2) = Format(rs.Fields("Line_Total"), "0.00")
        Else
            grdStock.TextMatrix(7, 1) = "0"
            grdStock.TextMatrix(7, 2) = "0.00"
        End If
        rs.Close
        'Tranfers In
        ActiveReadServer "Select sum(Qty) as Qty_Transfered,Sum(Ave_Cost) as Line_Total from Transfer_Journal where Product_Code='" & txtProdCode.Text & "' and Rec_Location_No like '" & LocString & "' and " & _
        "Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & mthViewEnd.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' group by Product_Code"
        If rs.RecordCount > 0 Then
            grdStock.TextMatrix(5, 1) = Round(Val(rs.Fields("Qty_Transfered") & ""), 3)
            grdStock.TextMatrix(5, 2) = Format(rs.Fields("Line_Total"), "0.00")
        Else
            grdStock.TextMatrix(5, 1) = "0"
            grdStock.TextMatrix(5, 2) = "0.00"
        End If
        rs.Close
        'Transfers Out
        ActiveReadServer "Select sum(Qty) as Qty_Transfered,Sum(Ave_Cost) as Line_Total from Transfer_Journal where Product_Code='" & txtProdCode.Text & "' and Trans_Location_No like '" & LocString & "' and " & _
        "Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & mthViewEnd.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' group by Product_Code"
        If rs.RecordCount > 0 Then
            grdStock.TextMatrix(4, 1) = Round(Val(rs.Fields("Qty_Transfered") & ""), 3)
            grdStock.TextMatrix(4, 2) = Format(rs.Fields("Line_Total"), "0.00")
        Else
            grdStock.TextMatrix(4, 1) = "0"
            grdStock.TextMatrix(4, 2) = "0.00"
        End If
        rs.Close
        'Calculate Opening
        grdStock.TextMatrix(8, 1) = Round(grdStock.ValueMatrix(10, 1) - grdStock.ValueMatrix(9, 1), 3)
        grdStock.TextMatrix(0, 1) = Round(grdStock.ValueMatrix(8, 1) + grdStock.ValueMatrix(7, 1) - grdStock.ValueMatrix(6, 1) - grdStock.ValueMatrix(5, 1) + grdStock.ValueMatrix(4, 1) + grdStock.ValueMatrix(3, 1) + grdStock.ValueMatrix(2, 1) - grdStock.ValueMatrix(1, 1), 3)
        grdStock.TextMatrix(8, 2) = Format(grdStock.ValueMatrix(10, 2) - grdStock.ValueMatrix(9, 2), "0.00")
        grdStock.TextMatrix(0, 2) = Format(grdStock.ValueMatrix(8, 2) + grdStock.ValueMatrix(7, 2) - grdStock.ValueMatrix(6, 2) - grdStock.ValueMatrix(5, 2) + grdStock.ValueMatrix(4, 2) + grdStock.ValueMatrix(3, 2) + grdStock.ValueMatrix(2, 2) - grdStock.ValueMatrix(1, 2), "0.00")
    End If
    'txtProdCode.SetFocus
    txtProdCode.SelStart = Len(txtProdCode.Text)
End Sub

Private Sub cmbLocs_Change()
    LoadProduct
End Sub
Private Sub cmdClose_Click()
    Select Case cmdclose.Caption
        Case "Ok"
            cmdSearch.Value = Up
            cmdSearch_Click
            KeyCode = 0
            txtProdCode.Text = grdProd.TextMatrix(grdProd.Row, 0)
            txtDescription.Text = grdProd.TextMatrix(grdProd.Row, 1)
            txtType.Text = grdProd.TextMatrix(grdProd.Row, 2)
            LoadProduct
            cmdclose.Caption = "Close"
        Case "Close"
            Unload Me
    End Select
End Sub
Private Sub cmdOk_Click()
    picDate.Visible = False
    If picDate.Visible = False Then Selection_Change
End Sub
Private Sub Form_Activate()
    If frmInquiry.Tag = "Not Now" Then
        frmInquiry.Tag = ""
        Exit Sub
    End If
    If frmInquiry.Tag <> "" Then
        ActiveReadServer "SELECT Products.Product_Code," & _
            " CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' +" & _
            " CONVERT(nvarchar(20), Unit_Size)" & _
            " + Unit_of_Measure END AS Description," & _
            " Sales_Item,Stock_Item," & _
            " Products.Selling_Price" & _
            " FROM Products " & _
            " Where Product_Code = '" & frmInquiry.Tag & "'"
        If rs.RecordCount > 0 Then
            txtProdCode.Text = frmInquiry.Tag
            txtDescription.Text = rs.Fields("Description")
            If rs.Fields("Sales_Item") = 1 Then
                 txtType.Text = "Sales Item"
            End If
            If rs.Fields("Stock_Item") = 1 Then
                 txtType.Text = "Stock Item"
            End If
            If rs.Fields("Sales_Item") = 1 And rs.Fields("Stock_Item") = 1 Then
                 txtType.Text = "Sales and Stock Item"
            End If
        End If
        rs.Close
        frmInquiry.Tag = "'"
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(mthViewEnd.Value, "DD MMM YYYY")
        LoadProduct
    Else
        mthViewStart.Value = Date
        mthViewEnd.Value = Date
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(mthViewEnd.Value, "DD MMM YYYY")
    End If
End Sub

Private Sub grdProd_DblClick()
    cmdclose.Caption = "Close"
    cmdSearch.Value = Up
    cmdSearch_Click
    KeyCode = 0
    txtProdCode.Text = grdProd.TextMatrix(grdProd.Row, 0)
    txtDescription.Text = grdProd.TextMatrix(grdProd.Row, 1)
    txtType.Text = grdProd.TextMatrix(grdProd.Row, 2)
    LoadProduct
End Sub

Private Sub grdStock_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Select Case Row
        Case 7
            frmInquiry.Tag = "Not Now"
            frmSaleInfo.Show vbModal
    End Select
End Sub

Private Sub grdStock_EnterCell()
    Select Case grdStock.Col
        Case 1, 2
            Select Case grdStock.Row
                Case 7
                    grdStock.Editable = flexEDKbdMouse
                    grdStock.ColComboList(grdStock.Col) = "..."
                Case Else
                    grdStock.Editable = flexEDNone
                    grdStock.ColComboList(grdStock.Col) = ""
            End Select
    End Select
End Sub

Private Sub mthView_LostFocus()
    DoEvents
    If picDate.Visible = False Then Selection_Change
End Sub
Private Sub mthView_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(mthViewEnd.Value, "DD MMM YYYY")
End Sub
Private Sub cmdSearch_Click()
    Select Case cmdSearch.Value
        Case 0
            cmdclose.Caption = "Ok"
            grdProd.Visible = True
            DoEvents
            Screen.MousePointer = 11
            txtProdCode.Enabled = False
            txtDescription.Enabled = False
            txtType.Enabled = False
            grdProd.Rows = 1
            ActiveReadServer "SELECT Products.Product_Code," & _
            " CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' +" & _
            " CONVERT(nvarchar(20), Unit_Size)" & _
            " + Unit_of_Measure END AS Description," & _
            " Sales_Item,Stock_Item," & _
            " Products.Selling_Price," & _
            " SUM(IsNull(Quantities.Stock_on_Hand, 0)) As SOH" & _
            " FROM Products LEFT OUTER JOIN" & _
            " Quantities ON Products.Product_Code = Quantities.Product_Code" & _
            " Where Sales_Item = 1 Or Stock_Item = 1" & _
            " GROUP BY Products.Product_Code, Products.Description,Unit_of_Measure,Unit_Size," & _
            " Products.Selling_Price , Sales_Item, Stock_Item order by Description"
            While Not rs.EOF
                grdProd.Rows = grdProd.Rows + 1
                grdProd.TextMatrix(grdProd.Rows - 1, 0) = rs.Fields("Product_Code")
                grdProd.TextMatrix(grdProd.Rows - 1, 1) = rs.Fields("Description")
                If rs.Fields("Sales_Item") = 1 Then
                    grdProd.TextMatrix(grdProd.Rows - 1, 2) = "Sales Item"
                End If
                If rs.Fields("Stock_Item") = 1 Then
                    grdProd.TextMatrix(grdProd.Rows - 1, 2) = "Stock Item"
                End If
                If rs.Fields("Sales_Item") = 1 And rs.Fields("Stock_Item") = 1 Then
                    grdProd.TextMatrix(grdProd.Rows - 1, 2) = "Sales and Stock Item"
                End If
                grdProd.TextMatrix(grdProd.Rows - 1, 3) = Round(rs.Fields("SOH"), 3)
                grdProd.TextMatrix(grdProd.Rows - 1, 4) = Format(rs.Fields("Selling_Price"), "0.00")
                rs.MoveNext
            Wend
            rs.Close
            If grdProd.Rows > 0 Then grdProd.Row = 1
            grdProd.Col = 1
            On Error Resume Next
            grdProd.SetFocus
            On Error GoTo 0
            Screen.MousePointer = 0
        Case 1
            grdProd.Visible = False
            txtProdCode.Enabled = True
            txtDescription.Enabled = True
            txtType.Enabled = True
    End Select
End Sub

Private Sub Form_Load()
    picDate.Height = 945
    grdProd.ColWidth(0) = grdProd.Width * 0.15
    grdProd.ColWidth(1) = grdProd.Width * 0.3
    grdProd.ColWidth(2) = grdProd.Width * 0.2
    grdProd.ColWidth(3) = grdProd.Width * 0.2
    grdProd.ColWidth(4) = grdProd.Width * 0.14
    grdProd.TextMatrix(0, 0) = "Product Code"
    grdProd.TextMatrix(0, 1) = "Description"
    grdProd.TextMatrix(0, 2) = "Product Type"
    grdProd.TextMatrix(0, 3) = "Stock on Hand"
    grdProd.TextMatrix(0, 4) = "Selling Price"
    grdProd.ColAlignment(0) = flexAlignLeftCenter
    grdProd.ColAlignment(1) = flexAlignLeftCenter
    grdProd.ColAlignment(2) = flexAlignLeftCenter
    grdProd.ColAlignment(3) = flexAlignRightCenter
    grdProd.ColAlignment(4) = flexAlignRightCenter
    grdStock.TextMatrix(0, 0) = "Opening Stock"
    grdStock.Cell(flexcpFontBold, 0, 0, 0, 2) = True
    grdStock.Cell(flexcpBackColor, 0, 1, 0, 2) = &HFDE2D9
    grdStock.TextMatrix(1, 0) = "+Goods Received"
    grdStock.TextMatrix(2, 0) = "-Goods Returned"
    grdStock.TextMatrix(3, 0) = "-Wastage"
    grdStock.TextMatrix(4, 0) = "-Out Going Transfers"
    grdStock.TextMatrix(5, 0) = "+In Comming Transfers"
    grdStock.TextMatrix(6, 0) = "+Production"
    grdStock.TextMatrix(7, 0) = "-Sales Consumption"
    grdStock.TextMatrix(8, 0) = "=Theoretical Closing Stock"
    grdStock.TextMatrix(0, 1) = "0"
    grdStock.TextMatrix(1, 1) = "0"
    grdStock.TextMatrix(2, 1) = "0"
    grdStock.TextMatrix(3, 1) = "0"
    grdStock.TextMatrix(4, 1) = "0"
    grdStock.TextMatrix(5, 1) = "0"
    grdStock.TextMatrix(6, 1) = "0"
    grdStock.TextMatrix(7, 1) = "0"
    grdStock.TextMatrix(8, 1) = "0"
    grdStock.TextMatrix(9, 1) = "0"
    grdStock.TextMatrix(10, 1) = "0"
    grdStock.TextMatrix(0, 2) = "0.00"
    grdStock.TextMatrix(1, 2) = "0.00"
    grdStock.TextMatrix(2, 2) = "0.00"
    grdStock.TextMatrix(3, 2) = "0.00"
    grdStock.TextMatrix(4, 2) = "0.00"
    grdStock.TextMatrix(5, 2) = "0.00"
    grdStock.TextMatrix(6, 2) = "0.00"
    grdStock.TextMatrix(7, 2) = "0.00"
    grdStock.TextMatrix(8, 2) = "0.00"
    grdStock.TextMatrix(9, 2) = "0.00"
    grdStock.TextMatrix(10, 2) = "0.00"
    grdStock.Cell(flexcpFontBold, 8, 0, 8, 2) = True
    grdStock.Cell(flexcpBackColor, 8, 1, 8, 2) = &HFDE2D9
    grdStock.TextMatrix(9, 0) = "Stock Take Variance"
    grdStock.TextMatrix(10, 0) = "=Closing Stock"
    grdStock.Cell(flexcpFontBold, 10, 0, 10, 2) = True
    grdStock.Cell(flexcpBackColor, 10, 1, 10, 2) = &HFDE2D9
    grdStock.ColWidth(0) = grdStock.Width * 0.5
    grdStock.ColWidth(1) = grdStock.Width * 0.25
    grdStock.ColWidth(2) = grdStock.Width * 0.25
    grdStock.ColAlignment(1) = flexAlignRightCenter
    grdStock.ColAlignment(2) = flexAlignRightCenter
    grdRev.TextMatrix(0, 0) = "Revenue"
    grdRev.TextMatrix(1, 0) = "-Cost of Sales"
    grdRev.TextMatrix(2, 0) = "=Gross Profit"
    grdRev.TextMatrix(3, 0) = "=Gross Profit%"
    grdRev.TextMatrix(0, 1) = "0.00"
    grdRev.TextMatrix(1, 1) = "0.00"
    grdRev.TextMatrix(2, 1) = "0.00"
    grdRev.TextMatrix(3, 1) = "0%"
    grdRev.Cell(flexcpFontBold, 2, 0, 2, 1) = True
    grdRev.Cell(flexcpFontBold, 3, 0, 3, 1) = True
    grdRev1.TextMatrix(0, 0) = "Revenue"
    grdRev1.TextMatrix(1, 0) = "-Cost of Sales"
    grdRev1.TextMatrix(2, 0) = "=Gross Profit"
    grdRev1.TextMatrix(3, 0) = "=Gross Profit%"
    grdRev1.TextMatrix(0, 1) = "0.00"
    grdRev1.TextMatrix(1, 1) = "0.00"
    grdRev1.TextMatrix(2, 1) = "0.00"
    grdRev1.TextMatrix(3, 1) = "0%"
    grdRev1.Cell(flexcpFontBold, 2, 0, 2, 1) = True
    grdRev1.Cell(flexcpFontBold, 3, 0, 3, 1) = True
    grdRev.Cell(flexcpBackColor, 2, 1, 3, 1) = &HFDE2D9
    grdRev1.Cell(flexcpBackColor, 2, 1, 3, 1) = &HFDE2D9
    cmbLocs.Clear
    ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
    cmbLocs.AddItem "<All Locations>"
    While Not rs.EOF
        cmbLocs.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmbLocs.Text = "<All Locations>"
    grdProd.Height = 4230
End Sub
Private Sub grdProd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            KeyCode = 0
            cmdSearch.Value = Up
            cmdSearch_Click
            KeyCode = 0
            txtProdCode.Text = grdProd.TextMatrix(grdProd.Row, 0)
            txtDescription.Text = grdProd.TextMatrix(grdProd.Row, 1)
            txtType.Text = grdProd.TextMatrix(grdProd.Row, 2)
            LoadProduct
            cmdclose.Caption = "Close"
        Case 27
            KeyCode = 0
            cmdSearch.Value = Up
            cmdSearch_Click
            KeyCode = 0
            txtProdCode.SetFocus
            cmdclose.Caption = "Close"
    End Select
End Sub

Private Sub txtProdCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            If txtProdCode.Text = "" Then
                KeyCode = 0
                cmdSearch.Value = Down
                cmdSearch_Click
                KeyCode = 0
            Else
                ActiveReadServer "SELECT Products.Product_Code," & _
                    " CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' +" & _
                    " CONVERT(nvarchar(20), Unit_Size)" & _
                    " + Unit_of_Measure END AS Description," & _
                    " Sales_Item,Stock_Item," & _
                    " Products.Selling_Price" & _
                    " FROM Products " & _
                    " Where Product_Code = '" & txtProdCode.Text & "'"
                If rs.RecordCount > 0 Then
                    txtProdCode.Text = txtProdCode.Text
                    txtDescription.Text = rs.Fields("Description")
                    If rs.Fields("Sales_Item") = 1 Then
                         txtType.Text = "Sales Item"
                    End If
                    If rs.Fields("Stock_Item") = 1 Then
                         txtType.Text = "Stock Item"
                    End If
                    If rs.Fields("Sales_Item") = 1 And rs.Fields("Stock_Item") = 1 Then
                         txtType.Text = "Sales and Stock Item"
                    End If
                    LoadProduct
                Else
                    rs.Close
                    MsgBox "Unknown Product Code. Please find your Product on the List Below", vbInformation, "HeroPOS"
                    txtProdCode.Text = ""
                    cmdSearch.Value = Down
                    cmdSearch_Click
                    KeyCode = 0
                    Exit Sub
                End If
                KeyCode = 0
            End If
        Case 38
            If txtProdCode.Text = "" Then
                KeyCode = 0
                cmdSearch.Value = Down
                cmdSearch_Click
                KeyCode = 0
            End If
        Case 40
            If txtProdCode.Text = "" Then
                KeyCode = 0
                cmdSearch.Value = Down
                cmdSearch_Click
                KeyCode = 0
            End If
    End Select
End Sub
Private Sub txtProdCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub
