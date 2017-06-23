VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTablePeminjaman 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table anggota meminjam uang"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12405
   Icon            =   "frmTablePeminjaman.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLihat 
      Caption         =   "Lihat"
      Height          =   375
      Left            =   9480
      TabIndex        =   5
      ToolTipText     =   "Lihat lebih detail"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbLunas 
      Height          =   315
      Left            =   8160
      TabIndex        =   4
      Text            =   "Lunas"
      ToolTipText     =   "Lihat daftar anggota yang sudah lunas atau hutang"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "KTP"
      ToolTipText     =   "Pilih Jenis Pencarian"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtSearch 
      Height          =   350
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Tuliskan Kata Kunci"
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdCari 
      Caption         =   "Cari"
      Height          =   350
      Left            =   4440
      TabIndex        =   2
      ToolTipText     =   "Cari berdasarkan kata kunci"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "X"
      Height          =   350
      Left            =   5280
      TabIndex        =   3
      ToolTipText     =   "Keluar dari pencarian"
      Top             =   120
      Width           =   350
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   10800
      TabIndex        =   6
      ToolTipText     =   "Keluar"
      Top             =   120
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   4695
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   4
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
End
Attribute VB_Name = "frmTablePeminjaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lunas As Integer '0=hutang 1=lunas

Private Sub cmbLunas_click()
    Select Case cmbLunas.Text
    Case "Lunas"
        lunas = 1
    Case "Hutang"
        lunas = 0
    Case Else
        cmbLunas.Text = "Hutang"
        lunas = 0
    End Select
    
    Call myLoadData
End Sub

'Private Sub cmdHapus_Click()
'    opt = MsgBox("Anda Yakin untuk menghapus data peminjaman dengan ID Peminjaman = " & CDbl(grid.TextMatrix(grid.Row, 1)) & " ?", vbYesNo + vbInformation, "Konfirmasi")
'    If opt = vbYes Then
'        dbConn.Execute "DELETE FROM tblpeminjamans WHERE id_peminjaman` = '" & (CDbl(grid.TextMatrix(grid.Row, 1))) & "'"
'        MsgBox "Deleted", vbOKOnly + vbInformation, "Konfirmasi"
'    Else
'        Exit Sub
'    End If
'End Sub

Private Sub cmdLihat_Click()
    frmReportHutang.showMe (CDbl(grid.TextMatrix(grid.Row, 1)))
End Sub

Private Sub Form_Load()
    Call setConn
    Call inisialisasi
    Call myLoadData
End Sub

Function myLoadData()
    tblPeminjaman.Open "SELECT * FROM tblpeminjaman WHERE lunas = '" & lunas & "' ORDER BY tanggal_meminjam", dbConn
        tblPeminjaman.MoveFirst
        Call initStyleGrid
        grid.Rows = 1
        Dim i As Integer
        i = 1
        Do While Not tblPeminjaman.EOF
            grid.Rows = grid.Rows + 1
            grid.TextMatrix(Val(i), 1) = tblPeminjaman.Fields("id_peminjaman")
            grid.TextMatrix(Val(i), 2) = tblPeminjaman.Fields("no_ktp")
            grid.TextMatrix(Val(i), 3) = tblPeminjaman.Fields("paket")
            grid.TextMatrix(Val(i), 4) = tblPeminjaman.Fields("tanggal_meminjam")
            grid.TextMatrix(Val(i), 5) = getTanggalKembali(tblPeminjaman.Fields("tanggal_meminjam"), tblPeminjaman.Fields("paket"))
            grid.TextMatrix(Val(i), 6) = FormatCurrency(tblPeminjaman.Fields("hutang"), 2, True, True, True)
            grid.TextMatrix(Val(i), 7) = FormatCurrency(getCicilan(tblPeminjaman.Fields("hutang"), tblPeminjaman.Fields("paket"), tblPeminjaman.Fields("id_peminjaman")), 2, True, True, True)
            grid.TextMatrix(Val(i), 8) = tblPeminjaman.Fields("id_jaminan")
            If tblPeminjaman.Fields("lunas") = 0 Then
                grid.TextMatrix(Val(i), 9) = "Hutang"
            Else
                grid.TextMatrix(Val(i), 9) = "Lunas"
            End If
            grid.TextMatrix(Val(i), 10) = tblPeminjaman.Fields("nik")
            tblPeminjaman.MoveNext
            i = i + 1
        Loop
        Call initStyleGrid
    tblPeminjaman.Close
End Function

Function getCicilan(besarMeminjam As Double, paket As String, idPeminjaman As Integer)
    Select Case paket
    Case "Pelunasan 1 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 1, idPeminjaman))
    Case "Pelunasan 2 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 2, idPeminjaman))
    Case "Pelunasan 3 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 3, idPeminjaman))
    Case "Pelunasan 4 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 4, idPeminjaman))
    Case "Pelunasan 5 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 5, idPeminjaman))
    Case "Pelunasan 6 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 6, idPeminjaman))
    Case "Pelunasan 7 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 7, idPeminjaman))
    Case "Pelunasan 8 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 8, idPeminjaman))
    Case "Pelunasan 9 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 9, idPeminjaman))
    Case "Pelunasan 10 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 10, idPeminjaman))
    Case "Pelunasan 11 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 11, idPeminjaman))
    Case "Pelunasan 12 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 12, idPeminjaman))
    Case "Pelunasan 24 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 24, idPeminjaman))
    Case "Pelunasan 36 Bulan"
        getCicilan = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 36, idPeminjaman))
    End Select
End Function

Function getTanggalKembali(tanggalMulai As String, paket As String)
    Select Case paket
    Case "Pelunasan 1 Bulan"
        getTanggalKembali = DateAdd("m", 1, tanggalMulai)
    Case "Pelunasan 2 Bulan"
        getTanggalKembali = DateAdd("m", 2, tanggalMulai)
    Case "Pelunasan 3 Bulan"
        getTanggalKembali = DateAdd("m", 3, tanggalMulai)
    Case "Pelunasan 4 Bulan"
        getTanggalKembali = DateAdd("m", 4, tanggalMulai)
    Case "Pelunasan 5 Bulan"
        getTanggalKembali = DateAdd("m", 5, tanggalMulai)
    Case "Pelunasan 6 Bulan"
        getTanggalKembali = DateAdd("m", 6, tanggalMulai)
    Case "Pelunasan 7 Bulan"
        getTanggalKembali = DateAdd("m", 7, tanggalMulai)
    Case "Pelunasan 8 Bulan"
        getTanggalKembali = DateAdd("m", 8, tanggalMulai)
    Case "Pelunasan 9 Bulan"
        getTanggalKembali = DateAdd("m", 9, tanggalMulai)
    Case "Pelunasan 10 Bulan"
        getTanggalKembali = DateAdd("m", 10, tanggalMulai)
    Case "Pelunasan 11 Bulan"
        getTanggalKembali = DateAdd("m", 11, tanggalMulai)
    Case "Pelunasan 12 Bulan"
        getTanggalKembali = DateAdd("m", 12, tanggalMulai)
    Case "Pelunasan 24 Bulan"
        getTanggalKembali = DateAdd("m", 24, tanggalMulai)
    Case "Pelunasan 36 Bulan"
        getTanggalKembali = DateAdd("m", 36, tanggalMulai)
    End Select
End Function

Sub inisialisasi()
    Me.Top = 800
    Me.Left = 4000

    cmbSearch.Clear
    cmbSearch.Text = "Semua"
    cmbSearch.AddItem "Semua"
    cmbSearch.AddItem "ID Pinjam"
    cmbSearch.AddItem "No KTP"
    cmbSearch.AddItem "Tanggal Pinjam"
    cmbSearch.AddItem "Paket"
    cmbSearch.AddItem "Petugas"
    
    cmbLunas.Clear
    cmbLunas.AddItem "Hutang"
    cmbLunas.AddItem "Lunas"
    cmbLunas.Text = "Hutang"
    lunas = 0
End Sub

Sub initStyleGrid()
    s = " | ID peminjaman | No KTP | Paket | Tanggal Meminjam | Tanggal Kembali | Hutang | Besar Cicilan | ID jaminan | Lunas | Petugas"
    grid.FormatString = s
    grid.ColWidth(0) = 0
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1500
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 2000
    grid.ColWidth(5) = 2000
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 1500
    grid.ColWidth(8) = 1500
    grid.ColWidth(9) = 1500
    grid.ColWidth(10) = 1500
    grid.FixedRows = 1
End Sub

Private Sub cmdCari_Click()
    Dim searchSql As String

    Select Case cmbSearch.Text
    Case "Semua"
        searchSql = "SELECT * FROM tblpeminjaman WHERE id_peminjaman LIKE '%" & txtSearch.Text & "%' OR no_ktp LIKE '%" & txtSearch.Text & "%' OR tanggal_meminjam LIKE '%" & txtSearch.Text & "%' OR paket LIKE '%" & txtSearch.Text & "%' OR nik LIKE '%" & txtSearch.Text & "%' AND lunas = '" & lunas & "'"
    Case "ID Pinjam"
        searchSql = "SELECT * FROM tblpeminjaman WHERE id_peminjaman LIKE '%" & txtSearch.Text & "%' AND lunas = '" & lunas & "'"
    Case "No KTP"
        searchSql = "SELECT * FROM tblpeminjaman WHERE no_ktp LIKE '%" & txtSearch.Text & "%' AND lunas = '" & lunas & "'"
    Case "Tanggal Pinjam"
        searchSql = "SELECT * FROM tblpeminjaman WHERE tanggal_meminjam LIKE '%" & txtSearch.Text & "%' AND lunas = '" & lunas & "'"
    Case "Paket"
        searchSql = "SELECT * FROM tblpeminjaman WHERE paket LIKE '%" & txtSearch.Text & "%' AND lunas = '" & lunas & "'"
    Case "Petugas"
        searchSql = "SELECT * FROM tblpeminjaman WHERE nik LIKE '%" & txtSearch.Text & "%' AND lunas = '" & lunas & "'"
    Case Else
        MsgBox "Harap pilih jenis pencarian", vbOKOnly + vbInformation, "Peringatan"
        Exit Sub
    End Select
    
    grid.Clear
    tblPeminjaman.Open searchSql, dbConn
        If Not tblPeminjaman.EOF Then
            tblPeminjaman.MoveFirst
            Call initStyleGrid
            grid.Rows = 1
            Dim i As Integer
            i = 1
            Do While Not tblPeminjaman.EOF
                grid.Rows = grid.Rows + 1
                grid.TextMatrix(Val(i), 1) = tblPeminjaman.Fields("id_peminjaman")
                grid.TextMatrix(Val(i), 2) = tblPeminjaman.Fields("no_ktp")
                grid.TextMatrix(Val(i), 3) = tblPeminjaman.Fields("paket")
                grid.TextMatrix(Val(i), 4) = tblPeminjaman.Fields("tanggal_meminjam")
                grid.TextMatrix(Val(i), 5) = getTanggalKembali(tblPeminjaman.Fields("tanggal_meminjam"), tblPeminjaman.Fields("paket"))
                grid.TextMatrix(Val(i), 6) = FormatCurrency(tblPeminjaman.Fields("hutang"), 2, True, True, True)
                grid.TextMatrix(Val(i), 7) = FormatCurrency(getCicilan(tblPeminjaman.Fields("hutang"), tblPeminjaman.Fields("paket"), tblPeminjaman.Fields("id_peminjaman")), 2, True, True, True)
                grid.TextMatrix(Val(i), 8) = tblPeminjaman.Fields("id_jaminan")
                If tblPeminjaman.Fields("lunas") = 0 Then
                    grid.TextMatrix(Val(i), 9) = "Hutang"
                Else
                    grid.TextMatrix(Val(i), 9) = "Lunas"
                End If
                grid.TextMatrix(Val(i), 10) = tblPeminjaman.Fields("nik")
                tblPeminjaman.MoveNext
                i = i + 1
            Loop
        End If
        Call initStyleGrid
    tblPeminjaman.Close
    initStyleGrid
End Sub

Private Sub cmdClear_Click()
    grid.Clear
    Call myLoadData
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub
