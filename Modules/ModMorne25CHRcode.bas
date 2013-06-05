Attribute VB_Name = "ModMorne25CHRcode"
Public Morne As Boolean
Public Tstring As String
Public Sstring As String
Public Savedowner As Boolean
Public Hd As String

Private Function getPlusMinus(chrr) As Boolean ' <<< This function retunrs either true or false
chrr = UCase(chrr)                             '     depending on if a charachter is more than
                                               '     halfway through the alphabet or not...
If Asc(chrr) - 65 < 12 Then
    getPlusMinus = True
Else
    getPlusMinus = False
End If
End Function

Public Function genNumber(Sstring)


Dim stringVal As Long
Dim genVal As Long
Dim tmpVar As String
Dim i As Integer
Dim seedMod As Integer

For i = 1 To Len(Sstring) - 0
    stringVal = stringVal + Val(Asc(Mid$(Sstring, i, 1))) ' <<< Counts the value of each ascii chr
Next                                                '     in the app name
seedMod = Int((Day(Date) & Month(Date) & Year(Date) & Hour(Time) & Minute(Time) & Second(Time)) ^ 0.2)
For i = 0 To Int(seedMod + Minute(Time))  ' <<< Vb's random num generator is not
    Rnd                                                 '     very random so i will make it more
Next                                                    '     random

tmpVar = ""
For i = 1 To 20                                   ' <<< Randomly create the 1st 4 parts of the code
    If Rnd < 0.5 Then                             ' <<< 1 in two chance of a letter or a number
        tmpVar = tmpVar & Chr(Int(Rnd * 25) + 65)
    Else
        tmpVar = tmpVar & Int(Rnd * 9)
    End If
    
    If Int(i / 5) = i / 5 And i <> 25 Then    ' <<< Add a ' - ' every 5 charachters
        tmpVar = tmpVar & " - "
    End If
Next

For i = 1 To Len(tmpVar) - 0                              ' <<< Creates a number based on the
    If i < Len(Sstring) Then                              '     first sections. Adds or takes
        If getPlusMinus(Mid(Sstring, i, 1)) = False Then  '     depending on various things
            genVal = genVal + Val(Asc(Mid$(tmpVar, i, 1))) '    Makes it mathematicaly harder
        Else                                              '     to re-order the code.
            genVal = genVal - Val(Asc(Mid$(tmpVar, i, 1)))
        End If
    Else
        If Int(i / 2) = i / 2 Then
            genVal = genVal - Val(Asc(Mid$(tmpVar, i, 1)))
        Else
            genVal = genVal + Val(Asc(Mid$(tmpVar, i, 1)))
        End If
    End If
Next
If genVal < 0 Then genVal = 0 - genVal      ' <<< If the number is less than 0 then make it
                                            '     positive

tmpVar = tmpVar & Mid((genVal * stringVal) & "COMPU", 1, 5) ' <<< Last part of the code is the
                                                         '     'value' of the first part of
                                                         '     the code times the 'value'
                                                         '     of the program name, limited
                                                         '     to 5 charachters. "COMPU" is
                                                         '     to make sure the result is
                                                         '     atleast 5 chars.

genNumber = UCase(tmpVar)    ' <<< Returns the new key

End Function


Public Function authKey(key, Tstring) As Boolean
authKey = False
On Error GoTo err

Dim splt() As String
Dim stringVal As Long
Dim genVal As Long
Dim tempVar As String
Dim i As Integer
key = UCase(key)

For i = 1 To Len(Tstring) - 0
    stringVal = stringVal + Val(Asc(Mid$(Tstring, i, 1)))
Next

splt = Split(key, " - ")
splt(4) = ""

tempVar = Join(splt, " - ")

For i = 1 To Len(tempVar) - 0
    If i < Len(Tstring) Then
        If getPlusMinus(Mid(Tstring, i, 1)) = False Then
            genVal = genVal + Val(Asc(Mid$(tempVar, i, 1)))
        Else
            genVal = genVal - Val(Asc(Mid$(tempVar, i, 1)))
        End If
    Else
        If Int(i / 2) = i / 2 Then
            genVal = genVal - Val(Asc(Mid$(tempVar, i, 1)))
        Else
            genVal = genVal + Val(Asc(Mid$(tempVar, i, 1)))
        End If
    End If
Next
If genVal < 0 Then genVal = 0 - genVal

splt = Split(key, " - ")

If genVal = Val(splt(4)) / stringVal Then
    authKey = True
Else
    authKey = False
End If



If Mid((stringVal * genVal) & "COMPU", 1, 5) = splt(4) Then
    authKey = True
Else
    authKey = False
End If

err:

End Function

Public Function Detailssaved(full As Boolean)


End Function
Public Function SNumber(strDrv As String) As Long
Dim Snum As Long
Dim Result As Long
Dim Tempted1 As String
Dim Tempted2 As String
  
Tempted1 = String$(255, Chr$(0))
Tempted2 = String$(255, Chr$(0))
Result = GetVolumeInformation(strDrv, Tempted1, Len(Tempted1), Snum, 0, 0, _
Tempted2, Len(Tempted2))
  
' this will be the value returned by the function
SNumber = Snum

End Function

