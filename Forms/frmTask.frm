VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmTask 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tasks"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   Icon            =   "frmTask.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2070
      Width           =   945
   End
   Begin VB.TextBox txtDepart 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2745
   End
   Begin VB.TextBox txtArrive 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   990
      Width           =   2745
   End
   Begin VB.TextBox txtRoom 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   1275
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Width           =   4185
   End
   Begin VB.TextBox txtLName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2820
      Width           =   4275
   End
   Begin VB.TextBox txtFName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2430
      Width           =   3165
   End
   Begin VB.TextBox txtNights 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1800
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1710
      Width           =   1035
   End
   Begin btButtonEx.ButtonEx cmdRates 
      Height          =   345
      Left            =   4920
      TabIndex        =   10
      Top             =   5580
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      Appearance      =   3
      AutoMask        =   0   'False
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
   End
   Begin btButtonEx.ButtonEx cmdOk 
      Height          =   345
      Left            =   3570
      TabIndex        =   11
      Top             =   5580
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      Appearance      =   3
      AutoMask        =   0   'False
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
   Begin MSComCtl2.DTPicker DTDepTime 
      Height          =   345
      Left            =   4920
      TabIndex        =   20
      Top             =   3330
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      Format          =   38141954
      CurrentDate     =   38862.4583333333
   End
   Begin MSComCtl2.DTPicker DTDate 
      Height          =   345
      Left            =   2130
      TabIndex        =   9
      Top             =   3330
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   609
      _Version        =   393216
      CalendarTitleBackColor=   16761024
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   38141955
      CurrentDate     =   38862
   End
   Begin RichTextLib.RichTextBox txtRemarks 
      Height          =   1425
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   " Check In Task "
      Top             =   3840
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2514
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmTask.frx":000C
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   -60
      TabIndex        =   21
      Top             =   3390
      Width           =   2115
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Task Date and Time:"
      Size            =   "3731;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   6
      Left            =   1650
      Top             =   2010
      Width           =   1215
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "2143;556"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   5
      Left            =   1650
      Top             =   1290
      Width           =   3015
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "5318;556"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   4
      Left            =   1650
      Top             =   930
      Width           =   3015
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "5318;556"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   3
      Left            =   1650
      Top             =   210
      Width           =   1545
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "2725;556"
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   19
      Top             =   1380
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Departure Date:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   18
      Top             =   990
      Width           =   1395
      BackColor       =   16777215
      Caption         =   "Arrival Date:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   315
      Index           =   12
      Left            =   210
      TabIndex        =   17
      Top             =   630
      Width           =   1395
      VariousPropertyBits=   8388627
      Caption         =   " Room Description:"
      Size            =   "2461;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   6
      Left            =   210
      TabIndex        =   16
      Top             =   270
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Room Number:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   2100
      Width           =   1395
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nights:"
      Height          =   255
      Index           =   0
      Left            =   630
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      Height          =   255
      Left            =   180
      TabIndex        =   13
      Top             =   2430
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   2790
      Width           =   1395
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   0
      Left            =   1650
      Top             =   570
      Width           =   4455
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "7858;556"
   End
   Begin MSForms.Image Image6 
      Height          =   315
      Left            =   1650
      Top             =   1650
      Width           =   1215
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "2143;556"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   1
      Left            =   1650
      Top             =   2370
      Width           =   3345
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "5900;609"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   2
      Left            =   1650
      Top             =   2760
      Width           =   4455
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "7858;609"
   End
   Begin MSForms.Image Image1 
      Height          =   3165
      Index           =   0
      Left            =   75
      Top             =   90
      Width           =   6135
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "10821;5583"
   End
   Begin MSForms.Image Image1 
      Height          =   1635
      Index           =   8
      Left            =   75
      Top             =   3750
      Width           =   6135
      BackColor       =   16777215
      BorderStyle     =   0
      MousePointer    =   8
      SpecialEffect   =   3
      Size            =   "10821;2884"
   End
   Begin MSForms.Image Image1 
      Height          =   6045
      Index           =   1
      Left            =   0
      Top             =   -570
      Width           =   6285
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "11086;10663"
   End
End
Attribute VB_Name = "frmTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    If Trim(txtRemarks.Text) <> "" Then
        ActiveReadServer1 "Select isnull(max(Task_No),0)+1 as Task_No from Room_Tasks"
        Task_No = rs1.Fields("Task_No")
        rs1.Close
        ActiveUpdateServer "INSERT INTO [Room_Tasks]([Task_No], [Res_No], [Description], [Date_Time], [Remarks])" & _
        " VALUES(" & Task_No & "," & TillData.Res_No & ",'Room Task','" & DTDate.Value & "','" & txtRemarks.Text & "')"
    End If
    Unload Me
End Sub
Private Sub cmdRates_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    DTDate.Value = Date
    ActiveReadServer "Select * from Reservations where Res_No = " & TillData.Res_No
    If rs.RecordCount > 0 Then
        txtArrive.Text = Format(rs.Fields("Arrive_Date"), "DDD DD MMM YYYY")
        txtDepart.Text = Format(rs.Fields("Depart_Date"), "DDD DD MMM YYYY")
        txtTitle.Text = rs.Fields("Title") & ""
        txtFName.Text = rs.Fields("First_Name") & ""
        txtLName.Text = rs.Fields("Last_Name") & ""
        txtNights.Text = Val(rs.Fields("Depart_Date") - rs.Fields("Arrive_Date"))
        txtRoom.Text = rs.Fields("Room_No")
        ActiveReadServer1 "Select Description from Rooms where Room_No = " & rs.Fields("Room_No")
        If rs1.RecordCount > 0 Then
            txtDescription.Text = rs1.Fields("Description")
        End If
        rs1.Close
    End If
    rs.Close
End Sub
