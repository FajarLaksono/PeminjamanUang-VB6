VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMenuUtama 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   Caption         =   "Peminjaman Uang"
   ClientHeight    =   3120
   ClientLeft      =   8265
   ClientTop       =   4260
   ClientWidth     =   9255
   Icon            =   "frmMenuUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar strBar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2625
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "6/23/2017"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "4:05 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IconList 
      Left            =   1920
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuUtama.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuUtama.frx":045E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuUtama.frx":08B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuUtama.frx":0D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuUtama.frx":1154
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1482
      ButtonWidth     =   3175
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "IconList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pendaftaran Anggota"
            Key             =   "Pendaftaran"
            Object.ToolTipText     =   "Pendaftaran anggota"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Peminjaman"
            Key             =   "Peminjaman"
            Object.ToolTipText     =   "Peminjaman uang"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pengembalian"
            Key             =   "Pengembalian"
            Object.ToolTipText     =   "Pengembalian uang"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Daftar Peminjam"
            Key             =   "Daftar"
            Object.ToolTipText     =   "Daftar Anggota yang meminjam"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Keluar"
            Key             =   "Logout"
            Object.ToolTipText     =   "Keluar"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPendaftaran 
      Caption         =   "Pendaftaran"
      Begin VB.Menu mnuPendafaranAnggota 
         Caption         =   "Pendaftaran Anggota"
      End
      Begin VB.Menu mnuAnggotaTerdaftar 
         Caption         =   "Anggota Terdaftar"
      End
   End
   Begin VB.Menu mnuPeminjaman 
      Caption         =   "Peminjaman"
      Begin VB.Menu mnuPeminjamanUang 
         Caption         =   "Peminjaman Uang"
      End
      Begin VB.Menu mnuPengembalianUang 
         Caption         =   "Pengembalian Uang"
      End
      Begin VB.Menu mnuDaftarAnggotaMeminjam 
         Caption         =   "Daftar Anggota Meminjam"
      End
   End
   Begin VB.Menu mnuAkun 
      Caption         =   "Akun"
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu mnuTentangProgram 
         Caption         =   "Tentang Program"
      End
   End
End
Attribute VB_Name = "frmMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'app.path berfungsi untuk mendeteksi alamat dimana program itu di simpan untuk sekarang ini. ini dapat kita gunakan
'untuk membantu kita untuk membuat program yang dinamis, dapat di pindah2 tempat penyimpananya tanpa kawatir program akan error karena kehilangan alamat file seperti gambar

Private Sub MDIForm_Load()
    'inisialiasi listimage
    'listimages adalah control yang berfungsi untuk menyimpan daftar gambar
    Me.IconList.ListImages.Clear 'bersihkan daftar gambar dan load satu per satu seperti dibawah ini
    Me.IconList.ListImages.Add , "Pendaftaran", LoadPicture(App.Path + "\images\icons\MAIL07.ico")
    Me.IconList.ListImages.Add , "Peminjaman", LoadPicture(App.Path + "\images\icons\CLIP06.ico")
    Me.IconList.ListImages.Add , "Pengembalian", LoadPicture(App.Path + "\images\icons\FOLDER05.ico")
    Me.IconList.ListImages.Add , "Daftar", LoadPicture(App.Path + "\images\icons\CRDFLE03.ico")
    Me.IconList.ListImages.Add , "Logout", LoadPicture(App.Path + "\images\icons\W95MBX01.ico")
    'di sini listimages digunakan untuk menyimpan gambar icon untuk toolbar
    
    'loading pertama
    frmlLoadingBarMainMenu.Show
    frmlLoadingBarMainMenu.loading
End Sub

'dibawah fungsinya sama semua, membuka form kalo tombol di klik dan juga jalankan animasi loading sebelum form terbuka
Private Sub mnuAnggotaTerdaftar_Click()
    frmlLoadingBarMainMenu.loading
    frmTableAnggota.Show
End Sub

Private Sub mnuDaftarAnggotaMeminjam_Click()
    frmlLoadingBarMainMenu.loading
    frmTablePeminjaman.Show
End Sub

Private Sub mnuLogout_Click()
    frmlLoadingBarMainMenu.loading
    frmLogin.Show
    Unload Me
End Sub

Private Sub mnuPeminjamanUang_Click()
    frmlLoadingBarMainMenu.loading
    frmPeminjaman.Show
End Sub

Private Sub mnuPendafaranAnggota_Click()
    frmlLoadingBarMainMenu.loading
    frmPendaftaran.Show
End Sub

Private Sub mnuPengembalianUang_Click()
    frmlLoadingBarMainMenu.loading
    frmPengembalian.Show
End Sub

Private Sub mnuTentangProgram_Click()
    frmlLoadingBarMainMenu.loading
    frmTentang.Show
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'penyeleksian check apa yang diklik, bisa kita bedakan berdasarkan .key yang telah kita setel tadi
    Select Case Button.Key
        Case "Pendaftaran"
            Call mnuPendafaranAnggota_Click
        Case "Peminjaman"
            Call mnuPeminjamanUang_Click
        Case "Pengembalian"
            Call mnuPengembalianUang_Click
        Case "Daftar"
            Call mnuDaftarAnggotaMeminjam_Click
        Case "Logout"
            Call mnuLogout_Click
    End Select
End Sub
