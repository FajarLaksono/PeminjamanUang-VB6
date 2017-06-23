VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReportAnggota 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Anggota"
   ClientHeight    =   5460
   ClientLeft      =   5385
   ClientTop       =   2745
   ClientWidth     =   8640
   Icon            =   "frmReportAnggota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8640
   Begin VB.Frame frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H00C0C0C0&
      Height          =   5775
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8655
      Begin Crystal.CrystalReport crystalReport 
         Left            =   120
         Top             =   4920
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2640
         TabIndex        =   11
         ToolTipText     =   "Status aktivitas anggota"
         Top             =   4560
         Width           =   5655
      End
      Begin VB.CommandButton cmdCetak 
         Appearance      =   0  'Flat
         Caption         =   "Cetak"
         Height          =   375
         Left            =   4800
         TabIndex        =   12
         ToolTipText     =   "Cetak Infotmasi anggota"
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txtPekerjaan 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   10
         ToolTipText     =   "Pekerjaan anggota"
         Top             =   4095
         Width           =   5685
      End
      Begin VB.TextBox txtKodePos 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   9
         ToolTipText     =   "Kode Pos anggota"
         Top             =   3615
         Width           =   5685
      End
      Begin VB.TextBox txtAlamat 
         Appearance      =   0  'Flat
         Height          =   825
         Left            =   2640
         TabIndex        =   8
         ToolTipText     =   "Alamat anggota"
         Top             =   2640
         Width           =   5685
      End
      Begin VB.CommandButton cmdEditSimpan 
         Appearance      =   0  'Flat
         Caption         =   "Simpan"
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         ToolTipText     =   "Edit / Simpan Data"
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txtNoTelepon 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   6
         ToolTipText     =   "Nomer telepon anggota"
         Top             =   2160
         Width           =   3495
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
         ToolTipText     =   "Laki-Laki"
         Top             =   1680
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtTempatLahir 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   2
         ToolTipText     =   "Tempat Lahir Anggota"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   1
         ToolTipText     =   "Nama Anggota"
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtNomerKTP 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   0
         ToolTipText     =   "Nomer KTP anggota"
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton cmdPilihFoto 
         Caption         =   "Pilih Foto"
         Height          =   435
         Left            =   6360
         TabIndex        =   7
         ToolTipText     =   "Pilih foto anggota"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton cmdKeluar 
         Appearance      =   0  'Flat
         Caption         =   "Keluar"
         Height          =   375
         Left            =   7200
         TabIndex        =   14
         ToolTipText     =   "Keluar"
         Top             =   4920
         Width           =   1095
      End
      Begin VB.OptionButton optKelamin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Perempuan"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   5
         ToolTipText     =   "Perempuan"
         Top             =   1680
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtLahir 
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         ToolTipText     =   "Tanggal Lahir Anggota"
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   113115137
         CurrentDate     =   -328716
         MinDate         =   -328716
      End
      Begin MSComDlg.CommonDialog commonDialog 
         Left            =   7080
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label labelKodePos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kode Pos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3630
         Width           =   2175
      End
      Begin VB.Label labelPekerjaan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pekerjaan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   4125
         Width           =   2175
      End
      Begin VB.Label labelAlamat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alamat"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2655
         Width           =   2175
      End
      Begin VB.Label labekNoTelepon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No Telepon"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label labelJenisKelamin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jenis Kelamin"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label labelTempatTanggalLahir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tempat / Tanggal Lahir"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label labelNama 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nama"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   765
         Width           =   2175
      End
      Begin VB.Label labelNomerKTP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nomor KTP "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   285
         Width           =   2175
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
   End
End
Attribute VB_Name = "frmReportAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frmToggle As Boolean '0 = view mode / 1 = edit mode
Dim currID As Double
Dim alamatFotoAnggota As String
Dim currFotoAnggota As String
Dim hutang As Double
Dim dirFotoAnggota As String

Function getArg()
    frmToggle = frmTableAnggota.myFrmToggle
    currID = CDbl(frmTableAnggota.myCurrID)
End Function

Private Sub cmdCetak_Click()
    crystalReport.SelectionFormula = "{tblanggota.no_ktp}='" & currID & "' "
    crystalReport.ReportFileName = App.Path & "/reports/rptAnggota.rpt"
    crystalReport.WindowState = crptMaximized
    crystalReport.RetrieveDataFiles
    crystalReport.Action = 1
End Sub

Private Sub cmdEditSimpan_Click()
    If cmdEditSimpan.Caption = "Edit" Then
        Call initAsEdit
    Else
        Call simpan
    End If
End Sub

Function savePic()
    'Ganti nama file
    Dim tanggal As String
    Dim waktu As String
    Dim namaFile As String
    
    If Not commonDialog.FileName = "" And Not commonDialog.FileTitle = "" And Not alamatFotoAnggota = currFotoAnggota Then
        tanggal = Format(Date, "d-mmmm-yyyy")
        waktu = Format(Time, "h-m-s")
        namaFile = "FotoAnggota" + tanggal + "-" + waktu
        FileCopy commonDialog.FileName, App.Path + dirFotoAnggota + namaFile + commonDialog.FileTitle
        getPicName = namaFile + commonDialog.FileTitle
        dbConn.Execute "UPDATE tblanggota SET foto = '" & getPicName & "' WHERE no_ktp = '" & txtNomerKTP.Text & "';"
    End If
End Function

Function simpan()
    If txtNomerKTP.Text = "" Or txtNomerKTP = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtNomerKTP.SetFocus
        Exit Function
    End If
    
    If txtNama.Text = "" Or txtNama.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtNama.SetFocus
        Exit Function
    End If
    
    If txtTempatLahir.Text = "" Or txtTempatLahir.Text = Null Then
        MsgBox "Harap isi data dengan benar.", vbOKOnly + vbInformation, "Data belum diisi"
        txtTempatLahir.SetFocus
        Exit Function
    End If
    
    Dim currtanggal As Date
    currtanggal = Format(Date, "dd/mm/yy")
    
    If dtLahir >= currtanggal Then
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
    
    Call savePic
    Call updateData
    MsgBox "Tersimpan", vbInformation + vbOKOnly, "Konfirmasi"
    frmTableAnggota.loadData
    Unload Me
End Function

Function updateData()
    Dim jenisKelamin As String
    
    If optKelamin(0) = True Then
        jenisKelamin = "Laki-Laki"
    Else
        jenisKelamin = "Perempuan"
    End If
    dbConn.Execute "UPDATE tblanggota SET nama = '" & txtNama.Text & "', tanggal_lahir = '" & dtLahir.Year & "-" & dtLahir.Month & "-" & dtLahir.Day & "', tempat_lahir = '" & txtTempatLahir.Text & "', jenis_kelamin = '" & jenisKelamin & "', pekerjaan = '" & txtPekerjaan.Text & "', telepon = '" & txtNoTelepon.Text & "', alamat = '" & txtAlamat.Text & "', kode_pos = '" & txtKodePos.Text & "' WHERE no_ktp = '" & txtNomerKTP.Text & "'"
End Function

Private Sub cmdKeluar_Click()
    If cmdKeluar.Caption = "Keluar" Then
        Unload Me
    Else
        Call initAsReport
    End If
End Sub

Private Sub cmdPilihFoto_Click()
    'dibawah adalah cara untuk mendapatkan file, kami buat hanya bisa memilih file JPEG untuk gambar
    commonDialog.FileName = "" 'inisialisasi
    commonDialog.Filter = "JPEG Files|*.jpg|All Files|*.*" 'set format pengambilan file
    commonDialog.ShowOpen 'buka dialog pemilihan dari windows
    alamatFotoAnggota = commonDialog.FileName 'ambil alamat gambar ke variabel alamatFotoAnggota
    
    'dibawah adalah manipulasi string, check apakah file yang diambil adalah "" / 0
    If Len(Trim(alamatFotoAnggota)) < 1 Then
        Exit Sub
    End If
    'set imageFotoAnggota sesuai gambar
    imgAnggota.Picture = LoadPicture(alamatFotoAnggota)
End Sub

Private Sub Form_Load()
    Call mdlDB.setConn
    Call getArg
    
    Me.Top = 2000
    Me.Left = 5500
    
    If frmToggle = True Then
        Call initAsEdit
    Else
        Call initAsReport
    End If
    
    Call loadData
End Sub

Function initAsEdit()
    txtNomerKTP.Enabled = False
    txtNama.Enabled = True
    txtTempatLahir.Enabled = True
    dtLahir.Enabled = True
    optKelamin(0).Enabled = True
    optKelamin(1).Enabled = True
    txtNoTelepon.Enabled = True
    txtAlamat.Enabled = True
    txtPekerjaan.Enabled = True
    txtKodePos.Enabled = True
    cmdPilihFoto.Visible = True
    cmdCetak.Visible = False
    cmdEditSimpan.Caption = "Simpan"
    cmdKeluar.Caption = "Cancle"
    txtStatus.Visible = False
    txtStatus.Enabled = False
    
    txtNomerKTP.BackColor = &H80000005
    txtNama.BackColor = &H80000005
    txtTempatLahir.BackColor = &H80000005
    txtNoTelepon.BackColor = &H80000005
    txtAlamat.BackColor = &H80000005
    txtPekerjaan.BackColor = &H80000005
    txtKodePos.BackColor = &H80000005
    
End Function

Function initAsReport()
    txtNomerKTP.Enabled = False
    txtNama.Enabled = False
    txtTempatLahir.Enabled = False
    dtLahir.Enabled = False
    optKelamin(0).Enabled = False
    optKelamin(1).Enabled = False
    txtNoTelepon.Enabled = False
    txtAlamat.Enabled = False
    txtPekerjaan.Enabled = False
    txtKodePos.Enabled = False
    cmdPilihFoto.Visible = False
    cmdCetak.Visible = True
    cmdEditSimpan.Caption = "Edit"
    cmdKeluar.Caption = "Keluar"
    txtStatus.Visible = True
    txtStatus.Enabled = False
    
    txtNomerKTP.BackColor = &H80000004
    txtNama.BackColor = &H80000004
    txtTempatLahir.BackColor = &H80000004
    txtNoTelepon.BackColor = &H80000004
    txtAlamat.BackColor = &H80000004
    txtPekerjaan.BackColor = &H80000004
    txtKodePos.BackColor = &H80000004
    txtStatus.BackColor = &H80000004
End Function

Function loadData()
    tblAnggota.Open "SELECT * FROM tblanggota WHERE no_ktp='" & currID & "'", dbConn
        txtNomerKTP.Text = tblAnggota.Fields("no_ktp")
        txtNama.Text = tblAnggota.Fields("nama")
        txtTempatLahir.Text = tblAnggota.Fields("tempat_lahir")
        dtLahir = tblAnggota.Fields("tanggal_lahir")
        If tblAnggota.Fields("jenis_kelamin") = "Laki-Laki" Then
            optKelamin(0) = True
        Else
            optKelamin(1) = True
        End If
        txtNoTelepon.Text = tblAnggota.Fields("telepon")
        txtAlamat.Text = tblAnggota.Fields("alamat")
        txtKodePos.Text = tblAnggota.Fields("kode_pos")
        txtPekerjaan.Text = tblAnggota.Fields("pekerjaan")
        
        'Check Foto keberadaan foto
        dirFotoAnggota = "/images/anggota/"
        If Not tblAnggota.Fields("foto") = "" Then
            'periksa jika foto ada pada direktori, ubah ke default jika foto tidak ditemukan pada direktori
            If Dir(App.Path + dirFotoAnggota + tblAnggota.Fields("foto")) <> "" Then
                'foto ditemukan
                imgAnggota.Picture = LoadPicture(App.Path & dirFotoAnggota & tblAnggota.Fields("foto"))
                alamatFotoAnggota = App.Path & dirFotoAnggota & tblAnggota.Fields("foto")
                currFotoAnggota = App.Path & dirFotoAnggota & tblAnggota.Fields("foto")
            Else
                'foto tidak ditemukan
                MsgBox "Terjadi kesalahan dalam pencarian file gambar anggota !", vbCritical + vbOKOnly, "Peminjaman Uang"
                imgAnggota.Picture = LoadPicture(App.Path & dirFotoAnggota & "default.JPG")
                alamatFotoAnggota = App.Path & dirFotoAnggota & "default.JPG"
                currFotoAnggota = App.Path & dirFotoAnggota & "default.JPG"
            End If
        Else
            'rekaman foto kosong
            MsgBox "Terjadi kesalahan dalam pencarian file gambar anggota !", vbCritical + vbOKOnly, "Peminjaman Uang"
            imgAnggota.Picture = LoadPicture(App.Path & dirFotoAnggota & "default.JPG")
            alamatFotoAnggota = App.Path & dirFotoAnggota & "default.JPG"
            currFotoAnggota = App.Path & dirFotoAnggota & "default.JPG"
        End If

    tblAnggota.Close
    
    tblPeminjaman.Open "SELECT hutang FROM tblpeminjaman WHERE no_ktp='" & currID & "' and lunas='0'", dbConn
        If Not tblPeminjaman.EOF Then
            txtStatus.Text = "Akun ini sedang meminjam sebesar " & FormatCurrency(tblPeminjaman.Fields("hutang"), 2, True, True, True)
        Else
            txtStatus.Text = "Akun ini sedang tidak meminjam"
        End If
    tblPeminjaman.Close
End Function
