VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPeminjaman 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Peminjaman Uang"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610
   Icon            =   "frmPeminjaman.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog commonDialog 
      Left            =   7320
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtNama 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   350
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Nama calon peminjam"
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox txtJaminan 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "Jenis jaminan calon peminjam"
      Top             =   1800
      Width           =   5175
   End
   Begin VB.TextBox txtTglAkhir 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   350
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Tanggal Akhir"
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txtTglMulai 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   350
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Tanggal mulai"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Close"
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      ToolTipText     =   "Close"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      ToolTipText     =   "Simpan"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtBesarCicilan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   350
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Besar Cicilan"
      Top             =   3720
      Width           =   5175
   End
   Begin VB.ComboBox cmbPaket 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Text            =   "- Pilih --"
      ToolTipText     =   "Paket Meminjam"
      Top             =   2760
      Width           =   5175
   End
   Begin VB.TextBox txtBesarMeminjam 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """Rp""#.##0;(""Rp""#.##0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   0
      EndProperty
      Height          =   350
      Left            =   1440
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Besar peminjaman"
      Top             =   2280
      Width           =   5175
   End
   Begin VB.CommandButton cmdGambarJaminan 
      Caption         =   "Pilih Foto Jaminan"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      ToolTipText     =   "Pilih Bukti Foto jaminan"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ComboBox cmbNoKTP 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Text            =   "KTP"
      ToolTipText     =   "Pilih KTP calon peminjam"
      Top             =   480
      Width           =   5175
   End
   Begin VB.Image imgAnggota 
      Height          =   1440
      Left            =   6720
      Picture         =   "frmPeminjaman.frx":000C
      Stretch         =   -1  'True
      ToolTipText     =   "Foto Anggota"
      Top             =   240
      Width           =   1680
   End
   Begin VB.Label lblJangka 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sampai"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblMaxTanggal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Jangka Hutang"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblBesarCicilan 
      BackStyle       =   0  'Transparent
      Caption         =   "Besar Cicilan"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblPaket 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Paket"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Besar Meminjam"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Image imgJaminan 
      Height          =   1305
      Left            =   6720
      Picture         =   "frmPeminjaman.frx":0F5B
      Stretch         =   -1  'True
      ToolTipText     =   "Foto Bukti Jaminan"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Jaminan"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1005
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No KTP"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   540
      Width           =   1095
   End
End
Attribute VB_Name = "frmPeminjaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim alamatFoto As String '3 var ini digunakan untuk menyimpan alamat foto dan ditampilkan pada form
Dim defaultFotoAnggota As String 'untuk menyimpan gambar default
Dim defaultFotoJaminan As String
Dim dirFotoJaminan As String 'alamat penyimpanan foto jaminan
Dim getPicName As String 'var ini berfungsi menyimpan nama file foto dan untuk berhubungan dan menyimpan ke database

Function initAll()
    'inisialisasi untuk cmbKTP
    Me.cmbNoKTP.Clear
    tblAnggota.Open "SELECT no_ktp FROM tblanggota", dbConn
    If Not tblAnggota.EOF Then 'jika tblanggota bukan End of File (bukan kosong) maka
        tblAnggota.MoveFirst 'pindah posisi pembacaan ke awal rekaman
        Do While Not tblAnggota.EOF 'jika tblanggota bukan End of File (bukan kosong) maka melakukan pengulangan
            Me.cmbNoKTP.AddItem tblAnggota.Fields("no_ktp") 'tambahkan no_ktp pada combo box ini
            tblAnggota.MoveNext 'pindah ke rekaman selanjutnya
        Loop
    End If
    tblAnggota.Close 'tutup jika sudah selesai dengan tablenya
    
    'inisialisasi ganti warna dan kunci
    'dibawah merupakan control yang memang tidak boleh dirubah jadi kita kunci dan bedakan dengan warna
    txtNama.Locked = True
    txtNama.BackColor = &H80000004
    
    txtTglMulai.Locked = True
    txtTglMulai.BackColor = &H80000004
    
    txtTglAkhir.Locked = True
    txtTglAkhir.BackColor = &H80000004
    
    txtBesarCicilan.Locked = True
    txtBesarCicilan.BackColor = &H80000004
    
    'inisialiasasi gambar
    defaultFotoAnggota = App.Path + "\images\Anggota\default.jpg"
    defaultFotoJaminan = App.Path + "\images\jaminan\default.jpg"
    dirFotoJaminan = "\images\jaminan\"
    getPicName = ""
    
    imgAnggota.Picture = LoadPicture(defaultFotoAnggota)
    imgJaminan.Picture = LoadPicture(defaultFotoJaminan)
    
    'set daftar paket
    Call mdlSystem.setDaftarPaket
    For i = 0 To Val(UBound(getDaftarPaket)) - 1 'ubound = hitung jumlah array
        Me.cmbPaket.AddItem mdlSystem.getDaftarPaket(i)
    Next i
    
    'inisialiasi
    Me.cmbPaket.Text = "Pelunasan 1 Bulan"
    Me.txtTglMulai.Text = Format(Date, "DD/MM/YYYY") 'hari ini
    
    txtBesarCicilan.Text = "0"
End Function

Function Hitung() 'ini algoritma penghitungan cicilan + bunga
'ada function khusus untuk penghitungan yaitu algo pada mdlsystem
'dan function ini adalah pengantarnya, argumen akan berbeda2 sesuai dengan kondisi(cmbpaket.text) di bawah
    Select Case Me.cmbPaket.Text
    Case "Pelunasan 1 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 1, Me.txtTglMulai.Text) 'menambahkan jumlah bulan pada tampilan tanggal
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 1, 0)) '0 = bunga umum yaitu 5 persen, bisa di atur di mdlSystem
    Case "Pelunasan 2 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 2, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 2, 0))
    Case "Pelunasan 3 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 3, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 3, 0))
    Case "Pelunasan 4 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 4, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 4, 0))
    Case "Pelunasan 5 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 5, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 5, 0))
    Case "Pelunasan 6 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 6, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 6, 0))
    Case "Pelunasan 7 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 7, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 7, 0))
    Case "Pelunasan 8 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 8, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 8, 0))
    Case "Pelunasan 9 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 9, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 9, 0))
    Case "Pelunasan 10 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 10, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 10, 0))
    Case "Pelunasan 11 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 11, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 11, 0))
    Case "Pelunasan 12 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 12, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 12, 0))
    Case "Pelunasan 24 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 24, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 24, 0))
    Case "Pelunasan 36 Bulan"
        Me.txtTglAkhir.Text = DateAdd("m", 36, Me.txtTglMulai.Text)
        txtBesarCicilan.Text = CDbl(mdlSystem.algo(CDbl(txtBesarMeminjam.Text), 36, 0))
    End Select
End Function

Function setPicToDir()
    'function ini digunakan untuk menyalin gambar di suatu alamat yang dipilih oleh user ke alamat yang sudah diatur
    'oleh kita agar mudah ditemukan oleh program, tidak hanya menyalin kita juga merubah namanya agar tidak terjadi dulikat nama jika user memilih foto
    'yang mempunyai nama sama pada foto sebelumnya. namanya di setel menggunakan jaminan+tanggal+waktu dengan itu gambar yang diinputkan namanya tidak akan sama
    'setelah disalin ke alamat yang sudah ditentuka kita ambil hasil nama foto tersebut ke getpicname, dan getpicname akan digunakan untuk menyimpan sementara nama foto yang selanjutnya akan diserahkan ke database
    Dim tanggal As String
    Dim waktu As String
    Dim namaFile As String
    'ubah nama file dengan id khusus yaitu jaminan[tanggal]-[jam] digunakan untuk mengantisipasi duplikasi nama pada direktori
    'dibawah menyetel namanya terlebih dahulu
    tanggal = Format(Date, "d-mmmm-yyyy")
    waktu = Format(Time, "h-m-s")
    namaFile = "jaminan" + tanggal + "-" + waktu
    'copy dari alamat asal ke alamat yang mudah ditemukan program yaitu app.path + /images/jaminan/
    FileCopy commonDialog.FileName, App.Path + dirFotoJaminan + namaFile + commonDialog.FileTitle 'copy gambar yang dipilih ke dalam forder jaminan dan mengganti namanya
    getPicName = namaFile + commonDialog.FileTitle 'dipake untuk penyimpanan sementara
End Function

Function inputToDatabase()
    'function ini berfungsi untuk berhubungan langsung ke server > database
    'dibawah untuk pemyinpanan pada table jaminan dengan beberapa argument yang sudah di inputkan
    dbConn.Execute "INSERT INTO tbljaminan VALUES ('', '" & Me.txtJaminan.Text & "', '" & getPicName & "')" 'input ke tbljaminan
    mdlDB.tblJaminan.Open "SELECT id_jaminan FROM tbljaminan WHERE foto = '" & getPicName & "'", dbConn 'ambil idnya
        dbConn.Execute "INSERT INTO tblpeminjaman VALUES ('', '" & cmbNoKTP.Text & "', '" & Format(Date, "YYYY-MM-DD") & "', '" & cmbPaket.Text & "', '" & CDbl(txtBesarMeminjam.Text) & "', '" & tblJaminan.Fields("id_jaminan") & "','0', '', '" & frmMenuUtama.strBar.Panels(3).Text & "', '5')" 'input ke tblpeminjaman
    mdlDB.tblJaminan.Close ' tutup kalo sudah selesai
    MsgBox "Data Telah Dimasukan.", vbOKOnly + vbInformation, "Konfirmasi"
End Function

Private Sub cmdGambarJaminan_Click()
'ini berfungsi untuk membuka commonDialog yang berfungsi untuk memilih file, disini kita setel defaultnya untuk .jpg
    commonDialog.FileName = "" 'setel awal nama file
    commonDialog.Filter = "JPEG Files|*.jpg|All Files|*.*" 'set untuk menyaring format pengambilan file yaitu hanya semua JPG
    commonDialog.ShowOpen 'buka dialog pemilihan file dari windows
    alamatFoto = commonDialog.FileName 'ambil alamat gambar ke variabel alamatFoto
    
    'dibawah adalah manipulasi string, check apakah file yang diambil adalah "" / 0
    If Len(Trim(alamatFoto)) < 1 Then
        Exit Sub 'jika kosong maka tidak perlu melakukan pernyataan di bawah
    End If
    'set imgJaminan sesuai gambar yang dipilih
    Me.imgJaminan.Picture = LoadPicture(alamatFoto)
End Sub

Private Sub cmdKeluar_Click()
    Unload Me 'tidak perlu penjelasan kau tau ini
End Sub

Private Sub cmdSimpan_Click()
'fungsi ini dugunakan untuk mengecek semua data dan menyimpanya
'pertama kita check semua data apakah user sudah menginputkan semua data yang diinputkan jika ada yang masih kosong maka akan dikeluarkan dari fungsi ini
'jika sudah terisi maka akan melakukan penyimpanan dengan hanya memanggil function2 dengan fungsi menyimpan yang sudah kita buat di atas
     
     If cmbNoKTP.Text = "" Or cmbNoKTP.Text = Null Then 'cek jika data tersebut kosong maka
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi" 'tampilkan peringatan
        cmbNoKTP.SetFocus 'set fokus ke kontrol yang masih ksong
        Exit Sub 'keluarkan jalur eksekusi
    End If
    
    If Me.txtJaminan.Text = "" Or txtJaminan.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtJaminan.SetFocus
        Exit Sub
    End If
    
    If alamatFoto = defaultFotoJaminan Or commonDialog.FileName = "" Or commonDialog.FileTitle = "" Then 'jika foto belum dipilih atau masih sama seperti default foto maka
        MsgBox "Foto Bukti Jaminan Anggota Belum Dimasukan !", vbInformasi, "Pemberitahuan"
        Me.cmdGambarJaminan.SetFocus
        Exit Sub
    End If
    
    If txtBesarMeminjam.Text = "" Or txtBesarMeminjam.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtBesarMeminjam.SetFocus
        Exit Sub
    End If
    
    If cmbPaket.Text = "" Or cmbPaket.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        cmbPaket.SetFocus
        Exit Sub
    End If
    
    If CDbl(txtBesarCicilan.Text) = 0 Or txtBesarCicilan.Text = "" Or txtBesarCicilan.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtBesarMeminjam.SetFocus
        Exit Sub
    End If
    
    'dibawah untuk memeriksa keberadaan anggota dan apakah anggota sedang berhutang
    tblAnggota.Open "SELECT no_ktp FROM tblanggota WHERE no_ktp='" & cmbNoKTP.Text & "'", dbConn
    If tblAnggota.EOF Then 'jika tidak ada anggota yang dimaksud maka
        MsgBox "Anggota berdasarkan Nomer KTP yang anda masukan tidak ditemukan! silahkan periksa kembali.", vbOKOnly + vbInformation, "Pemberitahuan"
        tblAnggota.Close
    Else 'jika anggota ada maka
        'cek jika anggota sedang berhutang
        tblPeminjaman.Open "SELECT hutang FROM tblpeminjaman WHERE no_ktp = '" & cmbNoKTP.Text & "' AND lunas = '0'", dbConn
            If Not tblPeminjaman.EOF Then 'jika sedang berhutang maka
                MsgBox "Anggota tidak bisa meminjam uang karena anggota sedang meminjam uang sebesar " & FormatCurrency(tblPeminjaman.Fields("hutang"), 2, True, True, True), vbInformation + vbOKOnly, "Anggota tidak bisa meminjam uang"
                tblAnggota.Close
                tblPeminjaman.Close
            Else 'jika tidak sedang berhutang maka
                Call setPicToDir
                Call inputToDatabase
                tblAnggota.Close
                tblPeminjaman.Close
                Unload Me
            End If
    End If
End Sub

Private Sub cmbNoKTP_Click()
'digunakan untuk menampilkan data anggota (nama dan foto anggota) setelah memilih no ktp
    tblAnggota.Open "SELECT no_ktp, nama, foto FROM tblanggota WHERE no_ktp='" & cmbNoKTP.Text & "'", dbConn
    If Not tblAnggota.EOF Then
        txtNama.Text = tblAnggota.Fields("nama")
        imgAnggota.Picture = LoadPicture(App.Path + "\images\anggota\" + tblAnggota.Fields("foto"))
    End If
    tblAnggota.Close
End Sub

Private Sub cmbPaket_Click()
    'panggil function tersebut setelah memilih paket
    Call Hitung
End Sub

Private Sub txtBesarCicilan_Change()
'melakukan pernyataan dibawah setiap perubahan pada txtBesarCicilan
    Me.txtBesarCicilan.Text = FormatCurrency(Me.txtBesarCicilan.Text, 2, True, True, True) 'mengubah semua perubahan menjadi format uang
End Sub

Private Sub txtBesarMeminjam_Change()
    'jika user mengisi "" atau kosong maka akan diisi dengan angka 0
    If Me.txtBesarMeminjam.Text = Null Or Me.txtBesarMeminjam.Text = "" Then
        Me.txtBesarMeminjam.Text = 0
        Exit Sub 'keluarkan dari sub tapi jika tidak kosong maka jalankan
    End If
    Call Hitung 'pernyataan ini tanpa menjalankan pernyataan di dalam penyeleksian di atas
End Sub

Private Sub txtBesarMeminjam_KeyPress(KeyAscii As Integer)
    'membatasi tombol pada input besar meminjam
    Dim strValid As String
    strValid = "0123456789" 'hanya bisa tombol 0 sampai 9
    strValid = strValid & Chr(8) 'ditambah dengan tombol del
    If KeyAscii = vbKeyReturn Then 'jika tombol enter ditekan maka
        Call Hitung
        Me.cmbPaket.SetFocus
    Else 'jika tidak
        If InStr(strValid, Chr(KeyAscii)) = 0 Then 'jika tombol yang ditekan bukan termasuk strvalid maka
            KeyAscii = 0 'kosongkan kuncinya agar tidak ada yang terjadi
        End If
    End If
End Sub

Private Sub txtBesarMeminjam_LostFocus()
'ubah besarmeminjam menjadi 0 jika besar cicilan juga 0
    If txtBesarCicilan.Text <= 0 Or txtBesarCicilan.Text = "" Then
        txtBesarMeminjam.Text = 0
    Else
        txtBesarMeminjam.Text = FormatCurrency(txtBesarMeminjam.Text, 2, True, True, True)
    End If
End Sub

Private Sub Form_Load()
    Call setConn
    
    'Set Lokasi Muncul jendela
    Me.Top = 800
    Me.Left = 6000
    
    Call initAll
End Sub
