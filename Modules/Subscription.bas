Attribute VB_Name = "Subscription"
Public chkdateagain As String
Public authcodecurrent As String
Public checked As String

Public Function dateauthenticate2(dtvalue As String, txtauthtext As String) As Boolean
    Dim i As String
    Dim Thepword As String
    Thepword = "subscription"
    i = Dcode(txtauthtext, Thepword)
    i = Format(i, "DD-MMMM-YYYY")
    If i = chkdateagain Then
        dateauthenticate2 = True
    Else
        dateauthenticate2 = False
    End If
End Function

Public Function Checksubscriptiondb() As Boolean
    On Error GoTo working
    filenum = FreeFile
    Open "c:\Windows\Optimumw32.dll" For Input As filenum
    Close filenum
    Checksubscriptiondb = False
    Exit Function
working:
    ActiveReadServer6 " select * from Subscription"
    If rs6.RecordCount > 0 Then
        chkdateagain = Format(rs6.Fields("Subscriptiondate"), "DD-MMMM-YYYY")
        authcodecurrent = rs6.Fields("Authcode")
        x = dateauthenticate2(chkdateagain, authcodecurrent)
        If x = True Then Checksubscriptiondb = True
        If DateValue(Maindate) > DateValue(chkdateagain) Then Checksubscriptiondb = False
        If DateValue(Maindate) > DateValue(chkdateagain) Then x = False
        If x = False Then Checksubscriptiondb = False
    End If
    If x = False Then
        filenum = FreeFile
        Open "c:\Windows\Optimumw32.dll" For Append As filenum
        Print #filenum, Format(rs6.Fields("Subscriptiondate"), "DD-MMMM-YYYY")
    Close filenum
    End If
End Function
