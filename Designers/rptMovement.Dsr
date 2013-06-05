VERSION 5.00
Begin {78E93846-85FD-11D0-8487-00A0C90DC8A9} rptMovement 
   Caption         =   "Stock Movement..."
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "rptMovement.dsx":0000
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19420
   _Version        =   393216
   _DesignerVersion=   100688210
   ReportWidth     =   7920
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   GridX           =   20
   GridY           =   20
   LeftMargin      =   500
   RightMargin     =   500
   TopMargin       =   500
   BottomMargin    =   500
   _Settings       =   31
   NumSections     =   5
   SectionCode0    =   1
   BeginProperty Section0 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section4"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode1    =   2
   BeginProperty Section1 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section2"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode2    =   4
   BeginProperty Section2 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section1"
      Object.Height          =   1440
      NumControls     =   0
   EndProperty
   SectionCode3    =   7
   BeginProperty Section3 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section3"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode4    =   8
   BeginProperty Section4 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section5"
      Object.Height          =   375
      NumControls     =   0
   EndProperty
End
Attribute VB_Name = "rptMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataReport_Initialize()
    Me.Orientation = rptOrientLandscape
    Dim MyConn As ADODB.Connection
    Dim MyRecSet As ADODB.Recordset
    Dim sSQL As String
    Set MyConn = New ADODB.Connection
    Set MyRecSet = New ADODB.Recordset
    MyConn.ConnectionString = cnnMain.ConnectionString
    MyConn.Open

    sSQL = "Select * from Pay_Temp"
    MyRecSet.Open sSQL, MyConn, adOpenStatic, adLockReadOnly

    Set rptMovement.DataSource = MyRecSet
    On Error Resume Next
    Set rptMovement.Sections(1).Controls(1).Picture = LoadPicture(Logo_File)
    rptMovement.Sections(1).Controls(3).Caption = Branch_Name
    rptMovement.Sections(1).Controls(4).Caption = Branch_Address
    rptMovement.Sections(2).Controls(9).Caption = "Reporting from " & frmReports.lblDate.Caption
    Set rptMovement.Sections(1).Controls(1).Picture = LoadPicture(Logo_File)
    On Error GoTo 0
'    On Error Resume Next
'    For i = 1 To rptMovement.Sections(2).Controls.Count
'        rptMovement.Sections(2).Controls(i).Caption = Str(i)
'    Next i
'    On Error GoTo 0
'    Exit Sub
End Sub
Private Sub DataReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    ActiveUpdateServer "Delete from Pay_Temp"
End Sub

