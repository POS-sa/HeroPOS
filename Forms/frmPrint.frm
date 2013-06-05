VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPrint 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Kitchen Printing"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11310
   FillColor       =   &H00C0FFC0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   405
      Left            =   5040
      TabIndex        =   1
      Top             =   8220
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
   Begin VSFlex8Ctl.VSFlexGrid grdPrint 
      Height          =   8025
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   11145
      _cx             =   19659
      _cy             =   14155
      Appearance      =   1
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
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   11
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub LoadLines1()
    On Error Resume Next
    grdPrint.Rows = 0
    Select Case Panel_no
        Case -1
            With frmTransView
                For i = 1 To .grdMain.Rows - 1
                    If .grdMain.Cell(flexcpForeColor, i, 1, i, 1) <> &HC00000 Then
                        If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            If Left(.grdMain.TextMatrix(i, 8), 4) <> "Void" Then
                                ' Kotie
                                'If .grdMain.TextMatrix(i, 0) <> "" Then
                                If (.grdMain.TextMatrix(i, 0) <> "") Or (.grdMain.TextMatrix(i, 2) <> 0) Then
                                    grdPrint.Rows = grdPrint.Rows + 1
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 2)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 8)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 5)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 7) = .grdMain.TextMatrix(i, 6)
                                    If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                Else
                                    If .grdMain.TextMatrix(i, 1) = "No Sale" Then
                                        grdPrint.Rows = grdPrint.Rows + 1
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 2)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 8)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 5)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 7) = .grdMain.TextMatrix(i, 6)
                                        If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
                TillData.SaleTotal = frmTransView.lblTotal
                TillData.TaxTotal = frmTransView.lblVat
            End With
        Case 0
            With frmSales
                For i = 1 To .grdMain.Rows - 1
                    If .grdMain.Cell(flexcpForeColor, i, 1, i, 1) <> &HC00000 Then
                        If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            If Left(.grdMain.TextMatrix(i, 8), 4) <> "Void" Then
                                'Kotie
                                'If .grdMain.TextMatrix(i, 0) <> "" Then
                                If (.grdMain.TextMatrix(i, 0) <> "") Or (.grdMain.TextMatrix(i, 2) <> 0) Then
                                    grdPrint.Rows = grdPrint.Rows + 1
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 2)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 8)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 5)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 7) = .grdMain.TextMatrix(i, 6)
                                    If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                Else
                                    If .grdMain.TextMatrix(i, 1) = "No Sale" Then
                                        grdPrint.Rows = grdPrint.Rows + 1
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 2)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 8)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 5)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 7) = .grdMain.TextMatrix(i, 6)
                                        If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End With
        Case 1
            With frmSales1
                For i = 1 To .grdMain.Rows - 1
                    If .grdMain.Cell(flexcpForeColor, i, 1, i, 1) <> &HC00000 Then
                        If Left(.grdMain.TextMatrix(i, 8), 4) <> "Void" Then
                            If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            'Kotie
                                'If .grdMain.TextMatrix(i, 0) <> "" Then
                                If (.grdMain.TextMatrix(i, 0) <> "") Or (.grdMain.TextMatrix(i, 2) <> 0) Then
                                    grdPrint.Rows = grdPrint.Rows + 1
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 2)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 8)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 5)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 7) = .grdMain.TextMatrix(i, 6)
                                    If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                Else
                                    If .grdMain.TextMatrix(i, 1) = "No Sale" Then
                                        grdPrint.Rows = grdPrint.Rows + 1
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 2)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 8)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 5)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 7) = .grdMain.TextMatrix(i, 6)
                                        If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End With
        Case 2
            With frmBar
                For i = 1 To .grdMain.Rows - 1
                    If .grdMain.Cell(flexcpForeColor, i, 1, i, 1) <> &HC00000 Then
                        If Left(.grdMain.TextMatrix(i, 8), 4) <> "Void" Then
                            If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            'Kotie
                                'If .grdMain.TextMatrix(i, 0) <> "" Then
                                If (.grdMain.TextMatrix(i, 0) <> "") Or (.grdMain.TextMatrix(i, 2) <> 0) Then
                                    grdPrint.Rows = grdPrint.Rows + 1
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 2)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 8)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 5)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 7) = .grdMain.TextMatrix(i, 6)
                                    If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                Else
                                    If .grdMain.TextMatrix(i, 1) = "No Sale" Then
                                        grdPrint.Rows = grdPrint.Rows + 1
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 2)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 8)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 5)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 7) = .grdMain.TextMatrix(i, 6)
                                        If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End With
    End Select
    If grdPrint.Rows = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    grdPrint.Select 0, 5, grdPrint.Rows - 1, 5
    grdPrint.Sort = flexSortStringAscending
top:
    For i = 0 To grdPrint.Rows - 1
        If i > 0 And grdPrint.TextMatrix(i, 1) <> "" Then
            If grdPrint.TextMatrix(i, 4) <> "Void" And grdPrint.TextMatrix(i - 1, 4) <> "Void" Then
            'Buffet item not to merge
             If InStr(grdPrint.TextMatrix(i, 2), "Buffet") = 0 Then
                If grdPrint.TextMatrix(i, 5) = grdPrint.TextMatrix(i - 1, 5) Then
                    If Val(grdPrint.ValueMatrix(i, 3) / grdPrint.ValueMatrix(i, 1)) = Val(grdPrint.ValueMatrix(i, 3) / grdPrint.ValueMatrix(i, 1)) Then
                        grdPrint.TextMatrix(i - 1, 1) = grdPrint.ValueMatrix(i - 1, 1) + grdPrint.ValueMatrix(i, 1)
                        grdPrint.TextMatrix(i - 1, 3) = Format(grdPrint.ValueMatrix(i - 1, 3) + grdPrint.ValueMatrix(i, 3), "0.00")
                        grdPrint.RemoveItem i
                        GoTo top
                    End If
                End If
                
                End If
            End If
        End If
    Next i
    grdPrint.Select 0, 1, grdPrint.Rows - 1, 1
    grdPrint.Sort = flexSortNumericAscending
    
    grdPrint.Select 0, 5, grdPrint.Rows - 1, 5
    grdPrint.Sort = flexSortNumericAscending
    
    grdPrint.Select 0, 8, grdPrint.Rows - 1, 8
    grdPrint.Sort = flexSortStringDescending
    For i = 0 To grdPrint.Rows - 1
        grdPrint.TextMatrix(i, 0) = i
        If i > 0 And grdPrint.TextMatrix(i, 1) <> "" Then
            If grdPrint.TextMatrix(i, 8) = grdPrint.TextMatrix(i - 1, 8) Then
                grdPrint.TextMatrix(i - 1, 8) = ""
            End If
        End If
    Next i
    
    grdPrint.Select 0, 0, grdPrint.Rows - 1, 0
    grdPrint.Sort = flexSortNumericDescending
    On Error GoTo 0
End Sub
Public Sub LoadLines()
    On Error Resume Next
    grdPrint.Rows = 0
    Select Case Panel_no
        Case 0
            With frmSales
                For i = 1 To .grdMain.Rows - 1
                    If .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = &HC0FFFF Then
                        If UCase(Trim(Slip_Printer)) <> UCase(Trim(.grdMain.TextMatrix(i, 11))) Then
                            If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                                grdPrint.Rows = grdPrint.Rows + 1
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 1)) = "" Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = grdPrint.TextMatrix(grdPrint.Rows - 2, 1)
                                End If
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 11)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 7)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 8)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                MasterCode = ""
                                For ib = InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), "<") - 2 To 1 Step -1
                                    If Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) < 48 Or Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) > 57 Then Exit For
                                    MasterCode = Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1) & MasterCode
                                Next ib
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = MasterCode
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 12)
                                If Left(Trim(.grdMain.TextMatrix(i, 1)), 1) = ">" Then
                                    If Right(.grdMain.TextMatrix(i, 1), 1) = Chr(160) Then
                                        .grdMain.TextMatrix(i, 1) = Mid(.grdMain.TextMatrix(i, 1), 1, Len(.grdMain.TextMatrix(i, 1)) - 1)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                        .grdMain.TextMatrix(i, 10) = .grdMain.TextMatrix(i - 1, 10)
                                    End If
                                End If
                                If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 8)) = "" Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                End If
                                If InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 2), ">") <> 0 Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                End If
                            End If
                        End If
                    End If
                Next i
            End With
        Case 1
            With frmSales1
                For i = 1 To .grdMain.Rows - 1
                    If .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = &HC0FFFF Then
                        If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            If UCase(Trim(Slip_Printer)) <> UCase(Trim(.grdMain.TextMatrix(i, 11))) Then
                                grdPrint.Rows = grdPrint.Rows + 1
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 1)) = "" Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = grdPrint.TextMatrix(grdPrint.Rows - 2, 1)
                                End If
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 11)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 7)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 8)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                MasterCode = ""
                                For ib = InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), "<") - 2 To 1 Step -1
                                    If Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) < 48 Or Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) > 57 Then Exit For
                                    MasterCode = Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1) & MasterCode
                                Next ib
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = MasterCode
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 12)
                                If Left(Trim(.grdMain.TextMatrix(i, 1)), 1) = ">" Then
                                    If Right(.grdMain.TextMatrix(i, 1), 1) = Chr(160) Then
                                        .grdMain.TextMatrix(i, 1) = Mid(.grdMain.TextMatrix(i, 1), 1, Len(.grdMain.TextMatrix(i, 1)) - 1)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                        .grdMain.TextMatrix(i, 10) = .grdMain.TextMatrix(i - 1, 10)
                                    End If
                                End If
                                If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 8)) = "" Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                End If
                                If InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 2), ">") <> 0 Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                End If
                            End If
                        End If
                    End If
                Next i
            End With
        Case 2
            With frmBar
                For i = 1 To .grdMain.Rows - 1
                    If .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = &HC0FFFF Then
                        If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            If UCase(Trim(Slip_Printer)) <> UCase(Trim(.grdMain.TextMatrix(i, 11))) Then
                                grdPrint.Rows = grdPrint.Rows + 1
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 1)) = "" Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = grdPrint.TextMatrix(grdPrint.Rows - 2, 1)
                                End If
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 11)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 7)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 8)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                MasterCode = ""
                                For ib = InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), "<") - 2 To 1 Step -1
                                    If Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) < 48 Or Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) > 57 Then Exit For
                                    MasterCode = Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1) & MasterCode
                                Next ib
                                If Trim(MasterCode) = "" Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = grdPrint.TextMatrix(grdPrint.Rows - 2, 4)
                                Else
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = MasterCode
                                End If
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 12)
                                If Left(Trim(.grdMain.TextMatrix(i, 1)), 1) = ">" Then
                                    If Right(.grdMain.TextMatrix(i, 1), 1) = Chr(160) Then
                                        .grdMain.TextMatrix(i, 1) = Mid(.grdMain.TextMatrix(i, 1), 1, Len(.grdMain.TextMatrix(i, 1)) - 1)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                        .grdMain.TextMatrix(i, 10) = .grdMain.TextMatrix(i - 1, 10)
                                    End If
                                End If
                                If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 8)) = "" Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                End If
                                If InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 2), ">") <> 0 Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                End If
                            End If
                        End If
                    End If
                Next i
            End With
    End Select
    
    grdPrint.Select 0, 0, grdPrint.Rows - 1, 0
    grdPrint.Sort = flexSortNumericAscending

    grdPrint.Select 0, 4, grdPrint.Rows - 1, 4
    grdPrint.Sort = flexSortNumericAscending

top:
    If Kitchen_Con = 1 Then
        For i = 0 To grdPrint.Rows - 1
            If i > 0 And grdPrint.TextMatrix(i, 1) <> "" Then
                If grdPrint.TextMatrix(i, 6) <> "Void" And grdPrint.TextMatrix(i - 1, 6) <> "Void" Then
                    If grdPrint.TextMatrix(i, 4) = grdPrint.TextMatrix(i - 1, 4) And grdPrint.TextMatrix(i, 4) <> "" Then
                        If grdPrint.TextMatrix(i, 2) = grdPrint.TextMatrix(i - 1, 2) Then
                            grdPrint.TextMatrix(i - 1, 1) = grdPrint.ValueMatrix(i - 1, 1) + grdPrint.ValueMatrix(i, 1)
                            grdPrint.RemoveItem i
                            GoTo top
                        End If
                    End If
                End If
            End If
        Next i
    End If
    grdPrint.Select 0, 0, grdPrint.Rows - 1, 0
    grdPrint.Sort = flexSortNumericDescending
    
    grdPrint.Select 0, 8, grdPrint.Rows - 1, 8
    grdPrint.Sort = flexSortStringAscending
    
    For i = 0 To grdPrint.Rows - 1
        grdPrint.TextMatrix(i, 9) = grdPrint.TextMatrix(i, 0)
        grdPrint.TextMatrix(i, 0) = i
        grdPrint.TextMatrix(i, 10) = grdPrint.TextMatrix(i, 8)
        If i > 0 And grdPrint.TextMatrix(i, 1) <> "" Then
            If grdPrint.TextMatrix(i, 8) = grdPrint.TextMatrix(i - 1, 8) Then
                grdPrint.TextMatrix(i - 1, 8) = ""
            End If
        End If
    Next i
    
    grdPrint.Select 0, 3, grdPrint.Rows - 1, 3
    grdPrint.Sort = flexSortStringAscending
    
    grdPrint.Select 0, 9, grdPrint.Rows - 1, 9
    grdPrint.Sort = flexSortNumericAscending
    
    grdPrint.Select 0, 10, grdPrint.Rows - 1, 10
    grdPrint.Sort = flexSortStringAscending
On Error GoTo 0
End Sub
Public Sub LoadLines2()
    On Error Resume Next
    grdPrint.Rows = 0
    Select Case Panel_no
        Case 0
            With frmSales
                For i = 1 To .grdMain.Rows - 1
                    If .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = &HC0FFFF Then
                        If UCase(Trim(Slip_Printer)) <> UCase(Trim(.grdMain.TextMatrix(i, 12))) Then
                            If UCase(Trim(.grdMain.TextMatrix(i, 12))) <> "" Then
                                If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                                    grdPrint.Rows = grdPrint.Rows + 1
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                    If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 1)) = "" Then
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = grdPrint.TextMatrix(grdPrint.Rows - 2, 1)
                                    End If
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 12)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 7)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 8)
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                    MasterCode = ""
                                    For ib = InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), "<") - 2 To 1 Step -1
                                        If Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) < 48 Or Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) > 57 Then Exit For
                                        MasterCode = Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1) & MasterCode
                                    Next ib
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = MasterCode
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 12)
                                    If Left(Trim(.grdMain.TextMatrix(i, 1)), 1) = ">" Then
                                        If Right(.grdMain.TextMatrix(i, 1), 1) = Chr(160) Then
                                            .grdMain.TextMatrix(i, 1) = Mid(.grdMain.TextMatrix(i, 1), 1, Len(.grdMain.TextMatrix(i, 1)) - 1)
                                            grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                            .grdMain.TextMatrix(i, 10) = .grdMain.TextMatrix(i - 1, 10)
                                        End If
                                    End If
                                    If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                    If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 8)) = "" Then
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                    End If
                                    If InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 2), ">") <> 0 Then
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End With
        Case 1
            With frmSales1
                For i = 1 To .grdMain.Rows - 1
                    If .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = &HC0FFFF Then
                        If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            If UCase(Trim(Slip_Printer)) <> UCase(Trim(.grdMain.TextMatrix(i, 12))) Then
                                If UCase(Trim(.grdMain.TextMatrix(i, 12))) <> "" Then
                                    If UCase(Trim(.grdMain.TextMatrix(i, 12))) <> "" Then
                                        grdPrint.Rows = grdPrint.Rows + 1
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                        If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 1)) = "" Then
                                            grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = grdPrint.TextMatrix(grdPrint.Rows - 2, 1)
                                        End If
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 12)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 7)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 8)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                        MasterCode = ""
                                        For ib = InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), "<") - 2 To 1 Step -1
                                            If Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) < 48 Or Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) > 57 Then Exit For
                                            MasterCode = Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1) & MasterCode
                                        Next ib
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = MasterCode
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 12)
                                        If Left(Trim(.grdMain.TextMatrix(i, 1)), 1) = ">" Then
                                            If Right(.grdMain.TextMatrix(i, 1), 1) = Chr(160) Then
                                                .grdMain.TextMatrix(i, 1) = Mid(.grdMain.TextMatrix(i, 1), 1, Len(.grdMain.TextMatrix(i, 1)) - 1)
                                                grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                                .grdMain.TextMatrix(i, 10) = .grdMain.TextMatrix(i - 1, 10)
                                            End If
                                        End If
                                        If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                        If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 8)) = "" Then
                                            grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                        End If
                                        If InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 2), ">") <> 0 Then
                                            grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End With
        Case 2
            With frmBar
                For i = 1 To .grdMain.Rows - 1
                    If .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = &HC0FFFF Then
                        If .grdMain.TextMatrix(i, 8) <> "Corr" Then
                            If UCase(Trim(Slip_Printer)) <> UCase(Trim(.grdMain.TextMatrix(i, 12))) Then
                                grdPrint.Rows = grdPrint.Rows + 1
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 0) = i
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = .grdMain.TextMatrix(i, 0)
                                If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 1)) = "" Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 1) = grdPrint.TextMatrix(grdPrint.Rows - 2, 1)
                                End If
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 3) = .grdMain.TextMatrix(i, 12)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = .grdMain.TextMatrix(i, 7)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 6) = .grdMain.TextMatrix(i, 8)
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 9)
                                MasterCode = ""
                                For ib = InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), "<") - 2 To 1 Step -1
                                    If Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) < 48 Or Asc(Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1)) > 57 Then Exit For
                                    MasterCode = Mid(grdPrint.TextMatrix(grdPrint.Rows - 1, 4), ib, 1) & MasterCode
                                Next ib
                                If Trim(MasterCode) = "" Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = grdPrint.TextMatrix(grdPrint.Rows - 2, 4)
                                Else
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 4) = MasterCode
                                End If
                                grdPrint.TextMatrix(grdPrint.Rows - 1, 5) = .grdMain.TextMatrix(i, 12)
                                If Left(Trim(.grdMain.TextMatrix(i, 1)), 1) = ">" Then
                                    If Right(.grdMain.TextMatrix(i, 1), 1) = Chr(160) Then
                                        .grdMain.TextMatrix(i, 1) = Mid(.grdMain.TextMatrix(i, 1), 1, Len(.grdMain.TextMatrix(i, 1)) - 1)
                                        grdPrint.TextMatrix(grdPrint.Rows - 1, 2) = .grdMain.TextMatrix(i, 1)
                                        .grdMain.TextMatrix(i, 10) = .grdMain.TextMatrix(i - 1, 10)
                                    End If
                                End If
                                If .grdMain.TextMatrix(i, 10) <> "" Then grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = Mid(.grdMain.TextMatrix(i, 10), 1, InStr(.grdMain.TextMatrix(i, 10), "-") - 1)
                                If Trim(grdPrint.TextMatrix(grdPrint.Rows - 1, 8)) = "" Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                End If
                                If InStr(grdPrint.TextMatrix(grdPrint.Rows - 1, 2), ">") <> 0 Then
                                    grdPrint.TextMatrix(grdPrint.Rows - 1, 8) = grdPrint.TextMatrix(grdPrint.Rows - 2, 8)
                                End If
                            End If
                        End If
                    End If
                Next i
            End With
    End Select
    
    grdPrint.Select 0, 0, grdPrint.Rows - 1, 0
    grdPrint.Sort = flexSortNumericAscending

    grdPrint.Select 0, 4, grdPrint.Rows - 1, 4
    grdPrint.Sort = flexSortNumericAscending

top:
    If Kitchen_Con = 1 Then
        For i = 0 To grdPrint.Rows - 1
            If i > 0 And grdPrint.TextMatrix(i, 1) <> "" Then
                If grdPrint.TextMatrix(i, 6) <> "Void" And grdPrint.TextMatrix(i - 1, 6) <> "Void" Then
                    If grdPrint.TextMatrix(i, 4) = grdPrint.TextMatrix(i - 1, 4) And grdPrint.TextMatrix(i, 4) <> "" Then
                        If grdPrint.TextMatrix(i, 2) = grdPrint.TextMatrix(i - 1, 2) Then
                            grdPrint.TextMatrix(i - 1, 1) = grdPrint.ValueMatrix(i - 1, 1) + grdPrint.ValueMatrix(i, 1)
                            grdPrint.RemoveItem i
                            GoTo top
                        End If
                    End If
                End If
            End If
        Next i
    End If
    grdPrint.Select 0, 0, grdPrint.Rows - 1, 0
    grdPrint.Sort = flexSortNumericDescending
    
    grdPrint.Select 0, 8, grdPrint.Rows - 1, 8
    grdPrint.Sort = flexSortStringAscending
    
    For i = 0 To grdPrint.Rows - 1
        grdPrint.TextMatrix(i, 9) = grdPrint.TextMatrix(i, 0)
        grdPrint.TextMatrix(i, 0) = i
        grdPrint.TextMatrix(i, 10) = grdPrint.TextMatrix(i, 8)
        If i > 0 And grdPrint.TextMatrix(i, 1) <> "" Then
            If grdPrint.TextMatrix(i, 8) = grdPrint.TextMatrix(i - 1, 8) Then
                grdPrint.TextMatrix(i - 1, 8) = ""
            End If
        End If
    Next i
    
    grdPrint.Select 0, 3, grdPrint.Rows - 1, 3
    grdPrint.Sort = flexSortStringAscending
    
    grdPrint.Select 0, 9, grdPrint.Rows - 1, 9
    grdPrint.Sort = flexSortNumericAscending
    
    grdPrint.Select 0, 10, grdPrint.Rows - 1, 10
    grdPrint.Sort = flexSortStringAscending
On Error GoTo 0
End Sub
Private Sub ButtonEx1_Click()
    End
    Me.Hide
End Sub

