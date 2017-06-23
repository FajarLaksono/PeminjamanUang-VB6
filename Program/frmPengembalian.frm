VERSION 5.00
Begin VB.Form frmPengembalian 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pengembalian"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8910
   Icon            =   "frmPengembalian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbJenisPelunasan 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Text            =   "- Jenis Pelunasan -"
      ToolTipText     =   "Pilih jenis pelunasan"
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox txtNama 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   350
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Nama Anggota"
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   350
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Total harus bayar"
      Top             =   2520
      Width           =   4215
   End
   Begin VB.ComboBox cmbNoKTP 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Text            =   "KTP"
      ToolTipText     =   "Nomer KTP Peminjam"
      Top             =   200
      Width           =   4215
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   6075
      TabIndex        =   8
      ToolTipText     =   "Simpan data"
      Top             =   3000
      Width           =   2565
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Close"
      Height          =   375
      Left            =   6075
      TabIndex        =   9
      ToolTipText     =   "Keluar"
      Top             =   3480
      Width           =   2565
   End
   Begin VB.TextBox txtUangKembali 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   350
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Uang kembali"
      Top             =   3480
      Width           =   4215
   End
   Begin VB.TextBox txtBayar 
      Appearance      =   0  'Flat
      Height          =   350
      Left            =   1680
      TabIndex        =   6
      ToolTipText     =   "Jumlah bayar"
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox txtSisaHutang 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   350
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Sisa Hutang"
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox txtPaket 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   350
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Paket"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label lblNama 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nama"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2550
      Width           =   1335
   End
   Begin VB.Image imgAnggota 
      Height          =   2565
      Left            =   6120
      Picture         =   "frmPengembalian.frx":048A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2565
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Uang Kembali"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3510
      Width           =   1335
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sisa Hutang"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bayar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblJenisPelunasan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Jenis Pelunasan"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2055
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Paket"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1125
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "No KTP"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   200
      Width           =   1335
   End
End
Attribute VB_Name = "frmPengembalian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ini paling ribet, yang buat juga sampe pusing nyelesain ini

'di cmbNoKTP hanya akan ada no ktp yang sedang meminjam
'untuk penghitungan, di form ini akan ditampilkan hutang + bunga dan jumlah bayar + bunga
'tapi saat dimasukan ke database dan data yang ada di database hanya akan hutang dan jumlah bayar

Dim bayarTotalDenganPersen As Double
Dim currDate As String 'hari ini

Private Sub cmbJenisPelunasan_Click()
    'atur perubahan setelah memilih pilihan dari jenis pelunasan
    'kemungkinan sub ini akan dipanggil oleh sub lain ketika mereka juga mengalami perubahan
    tblPeminjaman.Open "SELECT hutang, paket, id_peminjaman FROM tblpeminjaman WHERE no_ktp ='" & cmbNoKTP.Text & "' AND lunas= '0'", dbConn
    Select Case Me.cmbJenisPelunasan.Text
        Case "Bayar Cicilan"
            txtTotal.Text = FormatCurrency(getCicilan(Val(tblPeminjaman.Fields("hutang")), tblPeminjaman.Fields("paket"), tblPeminjaman.Fields("id_peminjaman")), 2, True, True, True)
        Case "Bayar Penuh"
            txtTotal.Text = txtSisaHutang.Text
        Case "- Pilih Jenis Pelunasan -"
            cmbJenisPelunasan.Text = "- Pilih Jenis Pelunasan -"
    End Select
    tblPeminjaman.Close
End Sub

'untuk penghitunga jumlah bayar(penuh atau cicilan) memakai rumus
'bunga = hutang / 100 %
'jumlahbayar = hutang + bunga
'kalo penuh kita cuma kalikan jumlah bayar dengan bulan

Function getBayarPenuh(besarMeminjam As Double, paket As String, idPeminjaman As Integer)
'dibawah merupakan pengantar ke algirutma pencarian bunnga yang ada di mdlSystem
    Select Case paket
    Case "Pelunasan 1 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 1, idPeminjaman)) * 1 'hasil dikali dengan berapa bulan
    Case "Pelunasan 2 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 2, idPeminjaman)) * 2
    Case "Pelunasan 3 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 3, idPeminjaman)) * 3
    Case "Pelunasan 4 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 4, idPeminjaman)) * 4
    Case "Pelunasan 5 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 5, idPeminjaman)) * 5
    Case "Pelunasan 6 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 6, idPeminjaman)) * 6
    Case "Pelunasan 7 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 7, idPeminjaman)) * 7
    Case "Pelunasan 8 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 8, idPeminjaman)) * 8
    Case "Pelunasan 9 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 9, idPeminjaman)) * 9
    Case "Pelunasan 10 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 10, idPeminjaman)) * 10
    Case "Pelunasan 11 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 11, idPeminjaman)) * 11
    Case "Pelunasan 12 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 12, idPeminjaman)) * 12
    Case "Pelunasan 24 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 24, idPeminjaman)) * 24
    Case "Pelunasan 36 Bulan"
        getBayarPenuh = CDbl(mdlSystem.algo(CDbl(besarMeminjam), 36, idPeminjaman)) * 36
    End Select
End Function


Function getCicilan(besarMeminjam As Double, paket As String, idPeminjaman As Integer)
'salah satu pengantar ke algo di sistem
'karena menyicil maka hasil tidak dikali dengan jumlah bulan
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


Function getBulan(paket As String)
'cuma ambil angka dari pilihan paket
    Select Case paket
    Case "Pelunasan 1 Bulan"
        getBulan = 1
    Case "Pelunasan 2 Bulan"
        getBulan = 2
    Case "Pelunasan 3 Bulan"
        getBulan = 3
    Case "Pelunasan 4 Bulan"
        getBulan = 4
    Case "Pelunasan 5 Bulan"
        getBulan = 5
    Case "Pelunasan 6 Bulan"
        getBulan = 6
    Case "Pelunasan 7 Bulan"
        getBulan = 7
    Case "Pelunasan 8 Bulan"
        getBulan = 8
    Case "Pelunasan 9 Bulan"
        getBulan = 9
    Case "Pelunasan 10 Bulan"
        getBulan = 10
    Case "Pelunasan 11 Bulan"
        getBulan = 11
    Case "Pelunasan 12 Bulan"
        getBulan = 12
    Case "Pelunasan 24 Bulan"
        getBulan = 24
    Case "Pelunasan 36 Bulan"
        getBulan = 36
    End Select
End Function

Private Sub cmbNoKTP_Click()
    tblPeminjaman.Open "SELECT * FROM tblpeminjaman WHERE no_ktp ='" & cmbNoKTP.Text & "' AND lunas= '0'", dbConn 'mencari hutang anggota di tabel peminjaman
        tblAnggota.Open "SELECT nama, foto FROM tblanggota WHERE no_ktp ='" & cmbNoKTP.Text & "'", dbConn 'ambil nama dan foto untuk ditampilkan
            'set control nama, paket dan foto
            txtNama.Text = tblAnggota.Fields("nama")
            txtPaket.Text = tblPeminjaman.Fields("paket")
            
            dirFotoAnggota = "/images/anggota/"
            If Not tblAnggota.Fields("foto") = "" Then
                'periksa jika foto ada pada direktori, ubah ke default jika foto tidak ditemukan pada direktori
                If Dir(App.Path + dirFotoAnggota + tblAnggota.Fields("foto")) <> "" Then
                    'foto ditemukan
                    imgAnggota.Picture = LoadPicture(App.Path & dirFotoAnggota & tblAnggota.Fields("foto"))
                Else
                    'foto tidak ditemukan
                    MsgBox "Terjadi kesalahan dalam pencarian file gambar anggota !", vbCritical + vbOKOnly, "Peminjaman Uang"
                    imgAnggota.Picture = LoadPicture(App.Path & dirFotoAnggota & "default.JPG")
                End If
            Else
                'rekaman foto kosong
                MsgBox "Terjadi kesalahan dalam pencarian file gambar anggota !", vbCritical + vbOKOnly, "Peminjaman Uang"
                imgAnggota.Picture = LoadPicture(App.Path & dirFotoAnggota & "default.JPG")
            End If

            'Mencari Sisa Hutang
            tblPengembalian.Open "SELECT * FROM tblpengembalian WHERE id_peminjaman = '" & tblPeminjaman.Fields("id_peminjaman") & "'", dbConn 'mencari data pengembalian yang memiliki id_peminjaman tersebut
                'rumus mencari sisa hutang di mulai
                'caranya hanya masuk ke database dan mencari pengembalian dengan id_peminjaman tersebut.
                'ambil field jumlah bayarya dan jumlahkan mereka semua
                Dim totalSudahBayar As Double
                totalSudahBayar = 0
                If Not tblPengembalian.EOF Then 'cek keberadaan record pengembalian
                tblPengembalian.MoveFirst
                    Do While Not tblPengembalian.EOF
                        totalSudahBayar = totalSudahBayar + CDbl(tblPengembalian.Fields("uang_bayar")) 'mencari total sudah bayar tanpa persen
                        tblPengembalian.MoveNext
                    Loop
                End If
                'lalu di persen
                Dim persenSudahBayar As Double
                persenSudahBayar = (totalSudahBayar / 100) * tblPeminjaman.Fields("besar_bunga") 'mencari bunga dari sisa hutang
                totalSudahBayar = totalSudahBayar + persenSudahBayar 'menambahkan kedua variabel tersebut
                txtSisaHutang.Text = FormatCurrency(CDbl(getBayarPenuh(tblPeminjaman.Fields("hutang"), tblPeminjaman.Fields("paket"), tblPeminjaman.Fields("id_peminjaman"))) - CDbl(totalSudahBayar), 2, True, True, True) 'pemanggilan function getbayarpenuh dan dibungkus dengan formatcurrencty (untuk mengubah angka menjadi format mata uang)
                txtSisaHutang.ToolTipText = "Sudah Bayar : " & FormatCurrency(CDbl(totalSudahBayar), 2, True, True, True)
            tblPengembalian.Close
        tblAnggota.Close
    tblPeminjaman.Close
    cmbJenisPelunasan_Click
End Sub

Private Sub cmdKeluar_Click()
    Unload Me 'keluar
End Sub

Private Sub cmdSimpan_Click()
'cek jika semua sudah terisi dengan benar
    If cmbNoKTP.Text = "" Or cmbNoKTP.Text = Null Or cmbNoKTP.Text = "- Pilih KTP -" Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        cmbNoKTP.SetFocus
        Exit Sub
    End If
    
    If txtNama.Text = "" Or txtNama.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtNama.SetFocus
        Exit Sub
    End If
    
    If txtPaket.Text = "" Or txtPaket.Text = Null Or txtSisaHutang.Text = "0" Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtPaket.SetFocus
        Exit Sub
    End If
    
    If txtSisaHutang.Text = "" Or txtSisaHutang.Text = Null Or txtSisaHutang.Text = "0" Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtSisaHutang.SetFocus
        Exit Sub
    End If

    If cmbJenisPelunasan.Text = "" Or cmbJenisPelunasan.Text = Null Or cmbJenisPelunasan.Text = "- Jenis Pelunasan -" Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        cmbJenisPelunasan.SetFocus
        Exit Sub
    End If
    
    If txtTotal.Text = "" Or txtTotal.Text = Null Or CDbl(txtTotal.Text) = 0 Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        cmbJenisPelunasan.SetFocus
        Exit Sub
    End If
    
    If txtBayar.Text = "" Or txtBayar.Text = Null Or CDbl(txtBayar.Text) = 0 Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtBayar.SetFocus
        Exit Sub
    End If
    
    If CDbl(txtBayar.Text) < CDbl(txtTotal.Text) Then
        MsgBox "Jumlah bayar tidak senilai", vbOKOnly + vbInformation, "Data belum diisi"
        txtBayar.SetFocus
        Exit Sub
    End If
    
    Call inputToDatabase
    Call checkLunas
    MsgBox "Tersimpan", vbInformation + vbOKOnly, "Peminjaman Uang"
    Unload Me
End Sub

Private Sub Form_Load()
    Call setConn
    Me.Top = 800
    Me.Left = 6000
    
    bayarTotalDenganPersen = 0
    Call initAll
End Sub

Function initAll()
    cmbNoKTP.Clear
    cmbNoKTP.Text = "- Pilih KTP -"
    tblPeminjaman.Open "SELECT no_ktp FROM tblpeminjaman WHERE lunas = '0'", dbConn
    If Not tblPeminjaman.EOF Then
        tblPeminjaman.MoveFirst
        Do While Not tblPeminjaman.EOF
            cmbNoKTP.AddItem tblPeminjaman.Fields("no_ktp")
            tblPeminjaman.MoveNext
        Loop
    End If
    tblPeminjaman.Close
    
    cmbJenisPelunasan.Clear
    cmbJenisPelunasan.Text = "- Pilih Jenis Pelunasan -"
    cmbJenisPelunasan.AddItem "Bayar Cicilan"
    cmbJenisPelunasan.AddItem "Bayar Penuh"
    
    txtPaket.Text = "0"
    txtSisaHutang.Text = "0"
    txtTotal.Text = "0"
    txtBayar.Text = "0"
    txtUangKembali.Text = "0"
End Function

Private Sub txtBayar_KeyPress(KeyAscii As Integer)
'membatasi tombol dan hanya 1-9 dan backspace
    Dim strValid As String
    strValid = "0123456789"
    strValid = strValid & Chr(8)
    If KeyAscii = vbKeyReturn Then
        txtBayar_lostfocus
    Else
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtBayar_lostfocus()
    txtBayar.Text = FormatCurrency(CDbl(txtBayar.Text), 2, True, True, True)
    If CDbl(txtBayar.Text) <= CDbl(txtTotal.Text) Then
        txtUangKembali.Text = 0
    Else
        txtUangKembali.Text = FormatCurrency(CDbl(txtBayar.Text) - CDbl(txtTotal.Text), 2, tue, True, True)
    End If
End Sub

Function inputToDatabase()
    'get variables
    Dim bayarTanpaPersen As Double
    
    bayarTanpaPersen = 0
    currDate = Format(Now, "YYYY-MM-dd")
    
    tblPeminjaman.Open "SELECT * FROM tblpeminjaman WHERE no_ktp ='" & cmbNoKTP.Text & "' AND lunas= '0'", dbConn
        Select Case Me.cmbJenisPelunasan.Text
        'yang ditampilkan mungkin menggunakan persen/bunga
        'tapi di sini, di database pembayaran tidak termasuk bunga.
             Case "Bayar Cicilan"
                 bayarTanpaPersen = CDbl(tblPeminjaman.Fields("hutang")) / CDbl(getBulan(tblPeminjaman.Fields("paket")))
             Case "Bayar Penuh"
                 tblPengembalian.Open "SELECT * FROM tblpengembalian WHERE id_peminjaman = '" & tblPeminjaman.Fields("id_peminjaman") & "'", dbConn
                 If tblPengembalian.EOF Then 'jika blm pernah melakukan pengembalian maka
                     bayarTanpaPersen = tblPeminjaman.Fields("hutang")
                 Else 'jika sudah maka
                    'mencari sisa hutang dengan tambahkan semua uang yang pernah di bayar untuk mengurangi total hutang
                     Dim hutangTerbayar As Double
                     hutangTerbayar = 0
                     Do While Not tblPengembalian.EOF
                         hutangTerbayar = hutangTerbayar + tblPengembalian.Fields("uang_bayar")
                     Loop
                     bayarTanpaPersen = CDbl(tblPeminjaman.Fields("hutang")) - CDbl(hutangTerbayar)
                 End If
                 tblPengembalian.Close
         End Select
         dbConn.Execute "INSERT INTO tblpengembalian VALUES('', '" & tblPeminjaman.Fields("id_peminjaman") & "', '" & cmbNoKTP.Text & "', '" & bayarTanpaPersen & "', '" & currDate & "','" & frmMenuUtama.strBar.Panels(3).Text & "')"
    tblPeminjaman.Close
End Function

Function checkLunas()
    'cek lunas jika memang terbukti membayar sesuai hutang maka akan terbilang lunas
    tblPeminjaman.Open "SELECT id_peminjaman, hutang FROM tblpeminjaman WHERE no_ktp ='" & cmbNoKTP.Text & "' AND lunas= '0'", dbConn
        tblPengembalian.Open "SELECT id_pengembalian, uang_bayar, tanggal_bayar FROM tblpengembalian WHERE id_peminjaman = '" & tblPeminjaman.Fields("id_peminjaman") & "'", dbConn
        If tblPengembalian.EOF Then
            Exit Function
        Else
            'mulai perhitungan sudah berapa banyak dia membayar
            Dim hutangTerbayar As Double
            hutangTerbayar = 0
            tblPengembalian.MoveFirst
            Do While Not tblPengembalian.EOF
                hutangTerbayar = hutangTerbayar + tblPengembalian.Fields("uang_bayar")
                tblPengembalian.MoveNext
            Loop
            If CDbl(tblPeminjaman.Fields("hutang")) <= CDbl(hutangTerbayar) Then 'jika yang sudah di bayar sesuai dengan hutangnya maka
                'update lunas menjadi true pada tbl peminjaman
                dbConn.Execute "UPDATE tblpeminjaman SET lunas = '1' WHERE id_peminjaman = '" & tblPeminjaman.Fields("id_peminjaman") & "'"
                freelance.CursorType = adOpenDynamic 'set pembacaan record dinamis dapat dibaca mulai dari mana saja
                freelance.Open "SELECT * FROM tblpengembalian WHERE id_peminjaman = '" & tblPeminjaman.Fields("id_peminjaman") & "' ORDER BY id_pengembalian", dbConn
                    freelance.MoveLast
                    Dim idTerakhir As Integer
                    idTerakhir = freelance.Fields("id_peminjaman")
                    Dim myTgl As String
                    myTgl = Format(freelance.Fields("tanggal_bayar"), "yyyy-mm-dd")
                'update tanggal_lunas pada tblpeminjaman dengan nilai tanggal hari ini
                dbConn.Execute "UPDATE tblpeminjaman SET tanggal_lunas = '" & myTgl & "' WHERE id_peminjaman = '" & idTerakhir & "'"
            Else
                Exit Function
            End If
        End If
        tblPengembalian.Close
    tblPeminjaman.Close
End Function
