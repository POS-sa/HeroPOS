Attribute VB_Name = "Modcryptinout"


Public Function EnCode(a$, pas$)
x = Len(a$) / Len(pas$)
x = Fix(x)
y = Len(a$) Mod Len(pas$)
a2$ = ""
pass$ = pas$
For i = 1 To x
a1$ = Mid$(a$, Len(pass$) * (i - 1) + 1, Len(pass$))
For j = 1 To Len(pass$)
rah$ = Chr$(Asc(Mid$(a1$, j, 1)) - 32)
a2$ = a2$ + Chr$(Asc(rah$) Xor Asc(Mid$(pass$, j, 1)))
Next
p1 = (Asc(Mid$(pass$, 2, 1)) + Asc(Right$(pass$, 1))) Xor 255
p1 = p1 Mod 159
pas1$ = Chr$(p1)
For j = 2 To Len(pass$) - 1
p1 = (Asc(Mid$(pass$, j - 1, 1)) + Asc(Mid$(pass$, j + 1, 1))) Xor 255
p1 = p1 Mod 159
pas1$ = pas1$ + Chr$(p1)
Next
p1 = (Asc(Mid$(pass$, 1, 1)) + Asc(Mid$(pass$, Len(pass$) - 1, 1))) Xor 255
p1 = p1 Mod 159
pas1$ = pas1$ + Chr$(p1)
pass$ = pas1$
Next
a1$ = Mid$(a$, Len(pas$) * x + 1, Len(pas$))
For j = 1 To y
a2$ = a2$ + Chr$(Asc(Mid$(a1$, j, 1)) Xor Asc(Mid$(pass$, j, 1)))
Next
a3$ = ""
For i = 1 To Len(a2$)
If Asc(Mid$(a2$, i, 1)) <> 26 And Asc(Mid$(a2$, i, 1)) <> 34 And Asc(Mid$(a2$, i, 1)) <> 0 Then a3$ = a3$ + Mid$(a2$, i, 1)
If Asc(Mid$(a2$, i, 1)) = 26 Then a3$ = a3$ + Chr$(254)
If Asc(Mid$(a2$, i, 1)) = 34 Then a3$ = a3$ + Chr$(255)
If Asc(Mid$(a2$, i, 1)) = 0 Then a3$ = a3$ + Chr$(253)
Next
a$ = a3$
EnCode = a$

End Function

Public Function Dcode(a$, pas$)
a2$ = a$
a3$ = ""
For i = 1 To Len(a2$)
If Asc(Mid$(a2$, i, 1)) <> 254 And Asc(Mid$(a2$, i, 1)) <> 255 And Asc(Mid$(a2$, i, 1)) <> 253 Then a3$ = a3$ + Mid$(a2$, i, 1)
If Asc(Mid$(a2$, i, 1)) = 254 Then a3$ = a3$ + Chr$(26)
If Asc(Mid$(a2$, i, 1)) = 255 Then a3$ = a3$ + Chr$(34)
If Asc(Mid$(a2$, i, 1)) = 253 Then a3$ = a3$ + Chr$(0)
Next
a$ = a3$
a3$ = ""
a2$ = ""
x = Len(a$) / Len(pas$)
x = Fix(x)
y = Len(a$) Mod Len(pas$)
pass$ = pas$
For i = 1 To x
a1$ = Mid$(a$, Len(pass$) * (i - 1) + 1, Len(pass$))
For j = 1 To Len(pass$)
rah$ = Chr$(Asc(Mid$(a1$, j, 1)))
rah$ = Chr$(Asc(rah$) Xor Asc(Mid$(pass$, j, 1)))
rah$ = Chr$(Asc(rah$) + 32)
a2$ = a2$ + rah$
Next
p1 = (Asc(Mid$(pass$, 2, 1)) + Asc(Right$(pass$, 1))) Xor 255
p1 = p1 Mod 159
pas1$ = Chr$(p1)
For j = 2 To Len(pass$) - 1
p1 = (Asc(Mid$(pass$, j - 1, 1)) + Asc(Mid$(pass$, j + 1, 1))) Xor 255
p1 = p1 Mod 159
pas1$ = pas1$ + Chr$(p1)
Next
p1 = (Asc(Mid$(pass$, 1, 1)) + Asc(Mid$(pass$, Len(pass$) - 1, 1))) Xor 255
p1 = p1 Mod 159
pas1$ = pas1$ + Chr$(p1)
pass$ = pas1$
Next
a1$ = Mid$(a$, Len(pas$) * x + 1, Len(pas$))
For j = 1 To y
a2$ = a2$ + Chr$(Asc(Mid$(a1$, j, 1)) Xor Asc(Mid$(pass$, j, 1)))
Next
'a3$ = ""
a$ = a2$
Dcode = a$
End Function
