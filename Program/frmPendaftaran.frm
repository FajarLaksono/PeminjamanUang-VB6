VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPendaftaran 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendaftaran Anggota"
   ClientHeight    =   6690
   ClientLeft      =   5730
   ClientTop       =   375
   ClientWidth     =   8685
   Icon            =   "frmPendaftaran.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8685
   Begin VB.Frame frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H00C0C0C0&
      Height          =   6615
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   8655
      Begin VB.OptionButton optKelamin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Perempuan"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtLahir 
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         ToolTipText     =   "Tanggal lahir calon anggota"
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   188022785
         CurrentDate     =   -328716
         MinDate         =   -328716
      End
      Begin VB.CommandButton cmdKeluar 
         Appearance      =   0  'Flat
         Caption         =   "Keluar"
         Height          =   375
         Left            =   7200
         TabIndex        =   16
         ToolTipText     =   "Keluar"
         Top             =   6120
         Width           =   975
      End
      Begin MSComDlg.CommonDialog commonDialog 
         Left            =   7080
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdPilihFoto 
         Caption         =   "Pilih Foto"
         Height          =   435
         Left            =   6360
         TabIndex        =   7
         ToolTipText     =   "Pilih foto calon anggota"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtNomerKTP 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   0
         ToolTipText     =   "Nomer KTP calon anggota"
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   1
         ToolTipText     =   "Nama calon anggota"
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtTempatLahir 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   2
         ToolTipText     =   "Tempat lahir calon anggota"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.OptionButton optKelamin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Laki - laki"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   4
         Top             =   1680
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtNoTelepon 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   6
         ToolTipText     =   "No. Telepon calon anggota"
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtAlamat 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   8
         ToolTipText     =   "Alamat calon anggota"
         Top             =   2625
         Width           =   5685
      End
      Begin VB.TextBox txtRtRw 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   9
         ToolTipText     =   "RT dan RW calon anggota"
         Top             =   3120
         Width           =   5685
      End
      Begin VB.TextBox txtKelDesa 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   10
         ToolTipText     =   "Kelurahan atau desa calon anggota"
         Top             =   3600
         Width           =   5685
      End
      Begin VB.TextBox txtKecamatan 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   11
         ToolTipText     =   "Kecamatan calon anggota"
         Top             =   4080
         Width           =   5685
      End
      Begin VB.TextBox txtKabupaten 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   12
         ToolTipText     =   "Kabupaten calon anggota"
         Top             =   4560
         Width           =   5685
      End
      Begin VB.TextBox txtKodePos 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   13
         ToolTipText     =   "Kode Pos calon anggota"
         Top             =   5040
         Width           =   5685
      End
      Begin VB.TextBox txtPekerjaan 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   14
         ToolTipText     =   "Pekerjaan calon anggota"
         Top             =   5520
         Width           =   5685
      End
      Begin VB.CommandButton cmdSimpan 
         Appearance      =   0  'Flat
         Caption         =   "Simpan"
         Height          =   375
         Left            =   6000
         TabIndex        =   15
         ToolTipText     =   "Simpan Data"
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Image imgAnggota 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   6360
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label labelNomerKTP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nomor KTP "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   285
         Width           =   2175
      End
      Begin VB.Label labelNama 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nama"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   765
         Width           =   2175
      End
      Begin VB.Label labelTempatTanggalLahir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tempat / Tanggal Lahir"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label labelJenisKelamin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jenis Kelamin"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label labelAlamat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alamat"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label labekNoTelepon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No Telepon"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label labelPekerjaan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pekerjaan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   5550
         Width           =   2175
      End
      Begin VB.Label labelRtRw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "RT / RW"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label labelKelDesa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kel / Desa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label labelKecamatan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kecamatan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label labelKabupaten 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kabupaten"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label labelKodePos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kode Pos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   5055
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmPendaftaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim today As Date 'digunakan untuk penyimpanan sementara tanggal hari ini
Dim defaultFoto As String 'menyimpan alamat foto default
Dim alamatFotoAnggota As String 'untuk menyimpan alamat Foto anggota yang dipilih
Dim dirFotoAnggota As String 'untuk menyimpan alamat forder foto anggota yaitu app.path images/anggota
Dim getPicName As String 'variabel untuk penyimpanan nama foto sementara yang akan berhubungan langsung dengan database

Function saveData()
'function ini berfngsi untuk memeriksa sumua data, cek apakah data masih kosong, jika
'kosong makan akan mengeluarkan txtbox dan jalur eksekusi akan dikeluarkan dari function ini
    If txtNomerKTP.Text = "" Or txtNomerKTP.Text = Null Then 'jika control tersebut masih kosong maka
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtNomerKTP.SetFocus 'fokus ke control yang masih kosong
        Exit Function 'keluar
    End If

    If txtNama.Text = "" Or txtNama.Text = Null Then '
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtNama.SetFocus
        Exit Function
    End If

    If txtTempatLahir.Text = "" Or txtTempatLahir.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtTempatlLahir.SetFocus
        Exit Function
    End If

    If dtLahir.Value = today Then 'jika tanggal lahir yang dipilih adalah tanggal sekarang maka
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        dtLahir.SetFocus
        Exit Function
    End If
    
    If txtNoTelepon.Text = "" Or txtNoTelepon.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtNoTelepon.SetFocus
        Exit Function
    End If
    
    If txtAlamat.Text = "" Or txtAlamat.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtAlamat.SetFocus
        Exit Function
    End If
    
    If txtRtRw.Text = "" Or txtRtRw.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtRtRw.SetFocus
        Exit Function
    End If
    
    If txtKelDesa.Text = "" Or txtKelDesa.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtKelDesa.SetFocus
        Exit Function
    End If
    
    If txtKecamatan.Text = "" Or txtKecamatan.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtKecamatan.SetFocus
        Exit Function
    End If
    
    If txtKabupaten.Text = "" Or txtKabupaten.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtKabupaten.SetFocus
        Exit Function
    End If
    
    If txtKodePos.Text = "" Or txtKodePos.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtKodePos.SetFocus
        Exit Function
    End If
    
    If txtPekerjaan.Text = "" Or txtPekerjaan.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtPekerjaan.SetFocus
        Exit Function
    End If
    
    If alamatFotoAnggota = defaultFotoAnggota Or commonDialog.FileName = "" Or commonDialog.FileTitle = "" Then 'periksa jika alamatfoto bukan foto default
        MsgBox "Foto Calon Anggota Belum Dimasukan !", vbInformasi, "Pemberitahuan"
        Exit Function
    End If
    
    'periksa apakah anggota sudah pernah mendaftar
    tblAnggota.Open "SELECT no_ktp FROM tblanggota WHERE no_ktp='" & txtNomerKTP.Text & "'", dbConn
    If Not tblAnggota.EOF Then
        tblAnggota.Close
        MsgBox "Nomer KTP yang dimasukan sudah ada dalam database, kemungkinan calon anggota sudah pernah mendaftar", vbOKOnly + vbInformation, "Pemberitahuan"
    Else
        'jika belum maka lanjut ke penyimpanan
        tblAnggota.Close
        setPicToDir
        inputToDatabase
    End If
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
    namaFile = "FotoAnggota" + tanggal + "-" + waktu
    'copy dari alamat asal ke alamat yang mudah ditemukan program yaitu app.path + /images/anggota/
    FileCopy commonDialog.FileName, App.Path + dirFotoAnggota + namaFile + commonDialog.FileTitle
    getPicName = namaFile + commonDialog.FileTitle 'dipake untuk penyimpanan sementara
End Function

Function inputToDatabase()
    'cari tau bahwa user memilih laki-laki atau perempuan, hasilnya akan disimpan sebagai string di getjeniskelamin
    Dim getJenisKelamin As String
    If optKelamin(0).Value = True Then
        getJenisKelamin = "Laki-Laki"
    Else
        getJenisKelamin = "Perempuan"
    End If
    'menyimpan data ke database
    dbConn.Execute "INSERT INTO tblanggota VALUES ('" & txtNomerKTP.Text & "','" & txtNama.Text & "','" & txtTempatLahir.Text & "','" & dtLahir.Year & "-" & dtLahir.Month & "-" & dtLahir.Day & "','" & getJenisKelamin & "','" & txtPekerjaan.Text & "','" & txtNoTelepon.Text & "','" & txtAlamat.Text & " " & txtRtRw.Text & " " & txtKelDesa.Text & " " & txtKecamatan.Text & " " & txtKabupaten.Text & "','" & txtKodePos.Text & "','" & getPicName & "')"
    MsgBox "Data Telah Dimasukan.", vbOKOnly + vbInformation, "Konfirmasi"
    Unload Me
End Function

Private Sub Form_Load()
'inisialisasi
    Call setConn
    Me.Top = 800
    Me.Left = 6000
    
    today = Format(Date, "dd/mm/yyyy") 'inisialisasi dengan hari ini
    dtLahir.Value = today 'atur dengan nilai hari ini
    optKelamin(0).Value = True 'sebagai default
    
    'inisialisasi penangkal file gambar
    defaultFoto = App.Path & "\images\anggota\default.jpg"
    imgAnggota.Picture = LoadPicture(defaultFoto)
    dirFotoAnggota = "\images\anggota\"
End Sub

Private Sub cmdPilihFoto_Click()
'ini berfungsi untuk membuka commonDialog yang berfungsi untuk memilih file, disini kita setel defaultnya untuk .jpg
    commonDialog.FileName = "" 'setel awal nama file
    commonDialog.Filter = "JPEG Files|*.jpg|All Files|*.*" 'set untuk menyaring format pengambilan file yaitu hanya semua JPG
    commonDialog.ShowOpen 'buka dialog pemilihan file dari windows
    alamatFotoAnggota = commonDialog.FileName 'ambil alamat gambar ke variabel alamatFoto
    
    'dibawah adalah manipulasi string, check apakah file yang diambil adalah "" / 0
    If Len(Trim(alamatFotoAnggota)) < 1 Then
        Exit Sub 'jika kosong maka tidak perlu melakukan pernyataan di bawah
    End If
    imgAnggota.Picture = LoadPicture(alamatFotoAnggota) 'set imageFotoAnggota sesuai gambar yang dipilih
End Sub

Private Sub cmdSimpan_Click()
    Call saveData 'panggil function yang bertugas untuk menyimpan data
End Sub

Private Sub cmdKeluar_Click()
    Unload Me 'keluar dari form
End Sub
