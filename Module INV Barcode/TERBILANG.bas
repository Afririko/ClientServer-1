Attribute VB_Name = "TERBILANG"
Function TeksKeAngka(ByVal n As Double) As String
Dim sSatuan()
Dim s As String
sSatuan() = Array("Nol", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas", "Dua Belas", "Tiga Belas", "Empat Belas", "Lima Belas", "Enam Belas", "Tujuh Belas", "Delapan Belas", "Sembilan Belas", "Dua Puluh")
s = ""
If n > 999 And n < 200 Then
s = "Seribu"
n = n Mod 1000
If n = 0 Then
TeksKeAngka = s
Exit Function
End If
s = s & " "
End If
If n > 199 Then
s = s & sSatuan(Fix(n / 100)) & " Ratus"
n = n Mod 100
If n = 0 Then
TeksKeAngka = s
Exit Function
End If
s = s & " "
End If
If n > 99 And n < 200 Then
s = "Seratus"
n = n Mod 100
If n = 0 Then
TeksKeAngka = s
Exit Function
End If
s = s & " "
End If
If n > 20 And n < 100 Then
s = s & sSatuan(Fix(n / 10)) & " Puluh"
n = n Mod 10
If n = 0 Then
TeksKeAngka = s
Exit Function
End If
s = s & " "
End If
TeksKeAngka = s & sSatuan(n)
End Function

Function terbilang(ByVal n) As String
Dim sBil()
sBil = Array("", "Ribu", "Juta", "Milyar", "Triliun", "Quadriliun")
Dim i As Integer
Dim iInt As Integer
Dim s As String
Dim dInt As Double
Dim t As String
t = " Rupiah"
dInt = Fix(n)
If (dInt < 2000) Then
terbilang = TeksKeAngka(CInt(dInt))
Exit Function
End If
i = 0
s = ""
Do While dInt > 0
iInt = CInt(dInt - Fix(dInt / 1000) * 1000)
If iInt <> 0 Then
If Len(s) > 0 Then s = " " & s
s = TeksKeAngka(iInt) & " " & sBil(i) & s
End If
i = i + 1
dInt = Fix(dInt / 1000)
Loop
terbilang = s & t
End Function

