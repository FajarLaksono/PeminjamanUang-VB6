Attribute VB_Name = "mdlSystem"
Option Explicit
Public getDaftarPaket(14) As String
Public besarBunga As Integer

Sub setDaftarPaket()
    getDaftarPaket(0) = "Pelunasan 1 Bulan"
    getDaftarPaket(1) = "Pelunasan 2 Bulan"
    getDaftarPaket(2) = "Pelunasan 3 Bulan"
    getDaftarPaket(3) = "Pelunasan 4 Bulan"
    getDaftarPaket(4) = "Pelunasan 5 Bulan"
    getDaftarPaket(5) = "Pelunasan 6 Bulan"
    getDaftarPaket(6) = "Pelunasan 7 Bulan"
    getDaftarPaket(7) = "Pelunasan 8 Bulan"
    getDaftarPaket(8) = "Pelunasan 9 Bulan"
    getDaftarPaket(9) = "Pelunasan 10 Bulan"
    getDaftarPaket(10) = "Pelunasan 11 Bulan"
    getDaftarPaket(11) = "Pelunasan 12 Bulan"
    getDaftarPaket(12) = "Pelunasan 24 Bulan"
    getDaftarPaket(13) = "Pelunasan 36 Bulan"
End Sub

Function algo(var As Double, bulan As Double, idPeminjaman As Integer) 'algo untuk mencari jumlah uang + bunga
On Error GoTo errHandler 'kalo error ke errhandler di bawah
    Dim persen As Double
    Dim result As Double
    If idPeminjaman = 0 Or idPeminjaman = Null Then
        persen = CDbl(5) 'defaultnya 5 persen
    Else
        freelance.Open "SELECT besar_bunga FROM tblpeminjaman WHERE id_peminjaman = '" & idPeminjaman & "' ", dbConn 'atau cek di database kalo parameter id peminjaman terisi
        persen = CDbl(freelance.Fields("besar_bunga"))
        freelance.Close
    End If
    'rumus
    result = CDbl(var) \ CDbl(100)
    result = CDbl(result) * CDbl(persen)
    result = CDbl(var) + CDbl(result)
    result = CDbl(result) \ CDbl(bulan)
    algo = CDbl(result)
    Exit Function
errHandler: 'sering error jika melebihi batas tipe data
    MsgBox FormatCurrency(var, 2, True, True, True) + " melebihi maxsimal peminjaman, ketentuan dari perusahaan kami !", vbOKOnly + vbInformation, "Maxsimal Peminjaman"
End Function
