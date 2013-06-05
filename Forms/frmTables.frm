VERSION 5.00
Begin VB.Form frmTables 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6780
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadTables()
    grdTable.Rows = 0
    cmdTable(0).Caption = ""
    cmdTable(0).Picture = ""
    DoEvents
    ActiveReadServer "Select * from Table_Listing_View where User_No = " & UserRecord.User_Number
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdTable.Rows = grdTable.Rows + 1
        If i < 14 And Not rs.EOF Then
            cmdTable(i).Caption = "Table No: " & rs.Fields("Table_No")
            cmdTable(i).Tag = rs.Fields("Table_No")
            If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
            grdTable.Row = grdTable.Rows - 1
            grdTable.TextMatrix(grdTable.Rows - 1, 0) = rs.Fields("Table_No")
        Else
            If b = 0 Then
                grdTable.TextMatrix(grdTable.Rows - 1, 0) = "ßArrow"
                grdTable.Rows = grdTable.Rows + 1
                If i = 14 Then
                    cmdTable(14).Caption = ""
                    cmdTable(14).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdTable(14).Visible = False Then cmdTable(14).Visible = True
                End If
            End If
            b = b + 1
            grdTable.TextMatrix(grdTable.Rows - 1, 0) = rs.Fields("Table_No")
            If b = 13 Then b = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    For b = i + 1 To cmdTable.Count - 1
       cmdTable(b).Caption = "1"
       cmdTable(b).Visible = False
    Next b
    If grdTable.Rows > 0 Then
         frmTables.Show 0, frmSales1
    End If
End Sub
Private Sub Form_Load()
    frmTables.top = frmSales1.cmdPlu(0).top - 30
    frmTables.Left = frmSales1.cmdPlu(0).Left - 30
    frmTables.Height = 6810
    frmTables.Width = 5880
    LoadTables
End Sub
