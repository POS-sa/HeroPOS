VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmAdrbook 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8490
   ClientLeft      =   4545
   ClientTop       =   0
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11493.38
   ScaleMode       =   0  'User
   ScaleWidth      =   14202.63
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "List All"
      Height          =   375
      Left            =   2550
      TabIndex        =   43
      Top             =   6930
      Width           =   1875
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   180
      TabIndex        =   41
      Top             =   6930
      Width           =   2115
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear form"
      Height          =   765
      Left            =   1140
      Picture         =   "FrmAdrbook.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Delete an entry"
      Top             =   7440
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   825
      Left            =   10260
      Picture         =   "FrmAdrbook.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Close Addressbook"
      Top             =   7380
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   765
      Left            =   150
      Picture         =   "FrmAdrbook.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Delete an entry"
      Top             =   7440
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Search ID"
      Height          =   345
      Left            =   7200
      TabIndex        =   37
      Top             =   990
      Width           =   975
   End
   Begin VB.CommandButton Cmdupdate 
      Caption         =   "Update"
      Height          =   765
      Left            =   3510
      Picture         =   "FrmAdrbook.frx":1C9E
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Update an existing entry"
      Top             =   7440
      Width           =   915
   End
   Begin VB.CommandButton Cmdadd 
      Caption         =   "Add New"
      Height          =   765
      Left            =   2520
      Picture         =   "FrmAdrbook.frx":1FA8
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Add a new entry"
      Top             =   7440
      Width           =   915
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8610
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   990
      Width           =   825
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   1455
      Left            =   5010
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "FrmAdrbook.frx":22B2
      Top             =   6510
      Width           =   4875
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   5010
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   5730
      Width           =   2115
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   7320
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   4890
      Width           =   2535
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   5010
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   4890
      Width           =   2115
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   7320
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   5010
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   4080
      Width           =   2115
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   7320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3270
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   5010
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3270
      Width           =   2115
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   7320
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   5010
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2520
      Width           =   2115
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   7320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1770
      Width           =   2475
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   5010
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1770
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   5010
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   990
      Width           =   2115
   End
   Begin VSFlex8Ctl.VSFlexGrid Grdadress 
      Height          =   5295
      Left            =   270
      TabIndex        =   0
      Top             =   1140
      Width           =   4005
      _cx             =   7064
      _cy             =   9340
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   0
      BackColorSel    =   16642749
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   4194368
      GridColorFixed  =   -2147483630
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
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
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   2
      Left            =   210
      TabIndex        =   42
      Top             =   6690
      Width           =   2205
      ForeColor       =   0
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Search by Surname ..."
      Size            =   "3889;423"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Surname"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   2310
      TabIndex        =   33
      Top             =   810
      Width           =   1665
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   270
      TabIndex        =   32
      Top             =   810
      Width           =   1665
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8820
      TabIndex        =   31
      Top             =   240
      Width           =   1275
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   7320
      TabIndex        =   30
      Top             =   210
      Width           =   1440
      ForeColor       =   0
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Last Update:"
      Size            =   "2540;423"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5010
      TabIndex        =   23
      Top             =   6210
      Width           =   855
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5010
      TabIndex        =   22
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Province"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      TabIndex        =   21
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Postcode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5010
      TabIndex        =   20
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      TabIndex        =   19
      Top             =   3750
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5010
      TabIndex        =   18
      Top             =   3750
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cell"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   17
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5010
      TabIndex        =   16
      Top             =   2970
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   15
      Top             =   2250
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5010
      TabIndex        =   14
      Top             =   2220
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Surname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   13
      Top             =   1470
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5010
      TabIndex        =   12
      Top             =   1470
      Width           =   1125
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8610
      TabIndex        =   11
      Top             =   690
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5010
      TabIndex        =   10
      Top             =   690
      Width           =   1215
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7695
      Left            =   4800
      Top             =   510
      Width           =   5355
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   0
      Left            =   4590
      TabIndex        =   4
      Top             =   210
      Width           =   1440
      ForeColor       =   0
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Details ..."
      Size            =   "2540;423"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   465
      Left            =   330
      Top             =   240
      Width           =   3135
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   13
      Left            =   450
      TabIndex        =   3
      Top             =   390
      Width           =   2835
      ForeColor       =   0
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Accomodation Address Book"
      Size            =   "5001;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5865
      Left            =   150
      Top             =   750
      Width           =   4275
   End
End
Attribute VB_Name = "FrmAdrbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Populatebook()
Grdadress.Clear
Grdadress.Cols = 3
Grdadress.Rows = 0
Grdadress.ColWidth(0) = Grdadress.Width * 0.38
Grdadress.ColWidth(1) = Grdadress.Width * 0.38
Grdadress.ColHidden(2) = True

Dim i As Integer, y As Integer
If rs4.State = 1 Then
If rs4.RecordCount > 0 Then
y = rs4.RecordCount
End If
End If

If y > 0 Then
While Not rs4.EOF
Grdadress.Rows = Grdadress.Rows + 1
Grdadress.TextMatrix(Grdadress.Rows - 1, 0) = rs4.Fields("First_name")
Grdadress.TextMatrix(Grdadress.Rows - 1, 1) = rs4.Fields("Last_name")
Grdadress.TextMatrix(Grdadress.Rows - 1, 2) = rs4.Fields("Id_No")
rs4.MoveNext
Wend
rs4.Close
Grdadress.Row = 0
End If
End Sub



Private Sub cmdadd_Click()
If Text1.Text = "" Then
MsgBox "Please insert ID NO first before trying to Add to Addressbook"
Exit Sub
End If


On Error GoTo Errors
ActiveReadServer4 " Select * from Adressbook where Id_no = '" & Text1.Text & "'ORDER by Last_Name, First_Name "
If rs4.RecordCount = 0 Then
Text10.Text = UCase(Text10.Text)
Text11.Text = UCase(Text11.Text)
Text12.Text = UCase(Text12.Text)
Text13.Text = UCase(Text13.Text)
Text14.Text = UCase(Text14.Text)
Text2.Text = UCase(Text2.Text)
Text3.Text = UCase(Text3.Text)
Text4.Text = UCase(Text4.Text)
Text5.Text = UCase(Text5.Text)
Text6.Text = UCase(Text6.Text)
Text7.Text = UCase(Text7.Text)
Text8.Text = UCase(Text8.Text)
Text9.Text = UCase(Text9.Text)

ActiveUpdateServer " Insert into Adressbook (Id_No ,First_Name,Last_Name,Email,Tel_No," & _
"Fax_No,Cell_No,Address,City,Post_code,Province,Country,Remarks,Title,Last_Booking)" & _
"Values ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & _
Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & _
Text10.Text & "','" & Text11.Text & "','" & Text12.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Date & "')"
MsgBox "Information added to Addressbook!'"
Else
MsgBox "This ID NO already exist. Could not add to the Addressbook!"
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Label16.Caption = ""
ActiveReadServer4 " Select * from Adressbook ORDER by Last_Name, First_Name"

Populatebook
Grdadress_Click
Errors:
End Sub

Private Sub cmdUpdate_Click()


If Text1.Text = "" Then
MsgBox "Please insert ID NO first before trying to Update Addressbook"
Exit Sub
End If

Text10.Text = UCase(Text10.Text)
Text11.Text = UCase(Text11.Text)
Text12.Text = UCase(Text12.Text)
Text13.Text = UCase(Text13.Text)
Text14.Text = UCase(Text14.Text)
Text2.Text = UCase(Text2.Text)
Text3.Text = UCase(Text3.Text)
Text4.Text = UCase(Text4.Text)
Text5.Text = UCase(Text5.Text)
Text6.Text = UCase(Text6.Text)
Text7.Text = UCase(Text7.Text)
Text8.Text = UCase(Text8.Text)
Text9.Text = UCase(Text9.Text)

ActiveReadServer4 " Select * from Adressbook where Id_no = '" & Text1.Text & "'"
If rs4.RecordCount > 0 Then
ActiveUpdateServer "Update Adressbook set Id_No = '" & Text1.Text & "', First_Name = '" & Text2.Text & " '," & _
"Last_Name = '" & Text3.Text & "',Email =  '" & Text4.Text & "',Tel_No = '" & Text5.Text & "', " & _
"Fax_No = '" & Text6.Text & "' , Cell_No = '" & Text7.Text & "', Address = '" & Text8.Text & "', City = '" & Text9.Text & "', Post_code = '" & Text10.Text & "', Province = '" & Text11.Text & "',Country = '" & Text12.Text & "', Remarks = '" & Text13.Text & "', Title = '" & Text14.Text & "', Last_Booking = '" & Date & "'" & _
"Where Id_No = '" & Text1.Text & "'"
MsgBox "Addressbook Information updated!"
Else
MsgBox "Information could not be updated as this ID NO does not exist, Use ADDNEW to do this operation!"
End If
Grdadress.Clear
ActiveReadServer4 " Select * from Adressbook ORDER by Last_Name, First_Name"
Populatebook
Grdadress_Click
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please insert ID NO first before trying to search the Addressbook"
Exit Sub
End If
ActiveReadServer4 " Select * from Adressbook where Id_no = '" & Text1.Text & "' ORDER by Last_Name, First_Name"
If rs4.RecordCount > 0 Then
Text1.Text = rs4.Fields("Id_No")
Text2.Text = rs4.Fields("First_Name")
Text3.Text = rs4.Fields("Last_Name")
Text4.Text = rs4.Fields("Email")
Text5.Text = rs4.Fields("Tel_No")
Text6.Text = rs4.Fields("Fax_No")
Text7.Text = rs4.Fields("Cell_No")
Text8.Text = rs4.Fields("Address")
Text9.Text = rs4.Fields("City")
Text10.Text = rs4.Fields("Post_code")
Text11.Text = rs4.Fields("Province")
Text12.Text = rs4.Fields("Country")
Text13.Text = rs4.Fields("Remarks")
Text14.Text = rs4.Fields("Title")
Label16.Caption = Format(rs4.Fields("Last_Booking"), "dd/mm/yyyy")
End If

End Sub

Private Sub Command2_Click()
Dim Deletename As String, Answer As Integer
Deletename = Grdadress.TextMatrix(Grdadress.Row, 2)
If Deletename <> "" Then
Answer = MsgBox("Are you sure you want to delete " & Trim(Grdadress.TextMatrix(Grdadress.Row, 0)) & _
" " & Trim(Grdadress.TextMatrix(Grdadress.Row, 1)), vbOKCancel)
If Answer = 0 Then
Exit Sub
End If
If Answer = 1 Then
ActiveUpdateServer "delete from Adressbook where ID_NO = '" & Deletename & "' "
ActiveReadServer4 " Select * from Adressbook ORDER by Last_Name, First_Name"
Populatebook
Grdadress_Click
End If


End If
End Sub

Private Sub Command3_Click()
frmMain.Command1.Tag = ""
Unload FrmAdrbook
frmDetails.Show
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()
ActiveReadServer4 " Select * from Adressbook ORDER by Last_Name, First_Name"
Populatebook
Grdadress_Click
End Sub

Private Sub Grdadress_Click()

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Label16.Caption = ""
If Grdadress.Rows = 0 Then Exit Sub
Dim Selected As String
Selected = Grdadress.TextMatrix(Grdadress.Row, 2)
ActiveReadServer4 " Select * from Adressbook where Id_No = '" & Selected & "' ORDER by Last_Name, First_Name"
If rs4.RecordCount > 0 Then
Text1.Text = rs4.Fields("Id_No")
Text2.Text = rs4.Fields("First_Name")
Text3.Text = rs4.Fields("Last_Name")
Text4.Text = rs4.Fields("Email")
Text5.Text = rs4.Fields("Tel_No")
Text6.Text = rs4.Fields("Fax_No")
Text7.Text = rs4.Fields("Cell_No")
Text8.Text = rs4.Fields("Address")
Text9.Text = rs4.Fields("City")
Text10.Text = rs4.Fields("Post_code")
Text11.Text = rs4.Fields("Province")
Text12.Text = rs4.Fields("Country")
Text13.Text = rs4.Fields("Remarks")
Text14.Text = rs4.Fields("Title")
Label16.Caption = Format(rs4.Fields("Last_Booking"), "dd/mm/yyyy")

End If
If rs4.State = 1 Then rs4.Close


End Sub

Private Sub Form_Initialize()
ActiveReadServer4 " Select * from Adressbook ORDER by Last_Name, First_Name"

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Label16.Caption = ""
Grdadress.Clear
Grdadress.Cols = 2
Grdadress.Rows = 0
Grdadress.ColWidth(0) = Grdadress.Width * 0.38
Grdadress.ColWidth(1) = Grdadress.Width * 0.38
Populatebook
Grdadress_Click
End Sub



Private Sub txtSearch_Change()
ActiveReadServer4 " Select * from Adressbook where (Last_Name like '" & UCase(txtSearch.Text) & "%')  ORDER by Last_Name, First_Name"
Populatebook
Grdadress_Click
End Sub
