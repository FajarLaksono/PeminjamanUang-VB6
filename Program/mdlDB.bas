Attribute VB_Name = "mdlDB"
Option Explicit 'pernyataan ini akan mengharuskan kita untuk mendirikan variabel terlebih dahulu untuk penyimpanan

'mendirikan variabel yang berfungsi untuk melakukan koneksi kepada database
Public dbConn As ADODB.Connection
Public tblJaminan As ADODB.Recordset
Public tblAnggota As ADODB.Recordset
Public tblPeminjaman As ADODB.Recordset
Public tblPengembalian As ADODB.Recordset
Public tblPetugas As ADODB.Recordset
Public freelance As ADODB.Recordset

Sub setConn()
'inisialiasi semua variabel di atas dan buka koneksi
    Set dbConn = New ADODB.Connection
    Set tblJaminan = New ADODB.Recordset
    Set tblAnggota = New ADODB.Recordset
    Set tblPeminjaman = New ADODB.Recordset
    Set tblPengembalian = New ADODB.Recordset
    Set tblPetugas = New ADODB.Recordset
    Set freelance = New ADODB.Recordset
    dbConn.ConnectionString = "driver=MySQL ODBC 3.51 Driver;server=localhost;uid=root;db=peminjaman_uang;"
    dbConn.Open
End Sub

Sub unsetConn()
    'tutup Semua koneksi
    tblJaminan.Close
    tblAnggota.Close
    tblPeminjaman.Close
    tblPengembalian.Close
    tblPetugas.Close
    freelance.Close
    dbConn.Close
End Sub
