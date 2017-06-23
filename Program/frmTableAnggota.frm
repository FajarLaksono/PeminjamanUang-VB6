VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTableAnggota 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table Anggota Terdaftar"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12435
   Icon            =   "frmTableAnggota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   12435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   10800
      TabIndex        =   7
      ToolTipText     =   "Keluar"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdUbah 
      Caption         =   "Ubah"
      Height          =   375
      Left            =   9360
      TabIndex        =   6
      ToolTipText     =   "Ubah informasi rekaman"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdLihat 
      Caption         =   "Lihat"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      ToolTipText     =   "Lihat rekaman lebih rinci"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "X"
      Height          =   350
      Left            =   5160
      TabIndex        =   4
      ToolTipText     =   "Keluar dari pencarian"
      Top             =   120
      Width           =   350
   End
   Begin VB.CommandButton cmdCari 
      Caption         =   "Cari"
      Height          =   350
      Left            =   4320
      TabIndex        =   3
      ToolTipText     =   "Cari berdasarkan kata kunci dan jenis pencarian"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtSearch 
      Height          =   350
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Tulis kata kunci"
      Top             =   120
      Width           =   2415
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "KTP"
      ToolTipText     =   "Pilih jenis pencarian"
      Top             =   120
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8281
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmTableAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myCurrID As Double
Public myFrmToggle As Boolean

'Private Sub cmdHapus_Click()
'    opt = MsgBox("Anda yakin ingin menghapus anggota dengan no ktp = " & CDbl(grid.TextMatrix(grid.Row, 1)) & " ? ", vbYesNo + vbInformation, "Konfirmasi")
'    If opt = vbYes Then
'        dbConn.Execute "DELETE FROM tblanggota WHERE no_ktp` = '" & (CDbl(grid.TextMatrix(grid.Row, 1))) & "'"
'        MsgBox "Deleted", vbOKOnly + vbInformation, "Konfirmasi"
'    Else
'        Exit Sub
'    End If
'End Sub

Private Sub cmdLihat_Click()
    myCurrID = CDbl(grid.TextMatrix(grid.Row, 1))
    myFrmToggle = False
    frmReportAnggota.Show
End Sub

Private Sub cmdUbah_Click()
    myCurrID = CDbl(grid.TextMatrix(grid.Row, 1))
    myFrmToggle = True
    frmReportAnggota.Show
End Sub

Private Sub Form_Load()
    Call setConn
    Call inisialisasi
    Call loadData
End Sub

Sub inisialisasi()
    Me.Top = 800
    Me.Left = 4000

    cmbSearch.Clear
    cmbSearch.Text = "Semua"
    cmbSearch.AddItem "Semua"
    cmbSearch.AddItem "KTP"
    cmbSearch.AddItem "Nama"
    cmbSearch.AddItem "Tempat Lahir"
    cmbSearch.AddItem "Tanggal Lahir"
    cmbSearch.AddItem "Pekerjaan"
    cmbSearch.AddItem "Telepon"
    cmbSearch.AddItem "Alamat"
    cmbSearch.AddItem "Kode Pos"
    
    myCurrID = 0
    myFrmToggle = True
End Sub

Sub initStyleGrid()
    s = " | KTP | Nama | Tempat Lahir | Tanggal_Lahir | Jenis Kelamin | Pekerjaan | Telepon | Alamat | Kode Pos"
    grid.FormatString = s
    grid.ColWidth(0) = 0
    grid.ColWidth(1) = 3000
    grid.ColWidth(2) = 3000
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 1500
    grid.ColWidth(5) = 1500
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 1500
    grid.ColWidth(8) = 1500
    grid.ColWidth(9) = 1500
End Sub

Sub loadData()
    tblAnggota.Open "SELECT no_ktp, nama, tempat_lahir, tanggal_lahir, jenis_kelamin, pekerjaan, telepon, alamat, kode_pos  FROM tblanggota order by no_ktp", dbConn
    Set grid.DataSource = tblAnggota
    Call initStyleGrid
    tblAnggota.Close
End Sub

Private Sub cmdCari_Click()
    Dim searchSql As String

    Select Case cmbSearch.Text
    Case "Semua"
        searchSql = "SELECT no_ktp, nama, tempat_lahir, tanggal_lahir, jenis_kelamin, pekerjaan, telepon, alamat, kode_pos FROM tblAnggota WHERE no_ktp LIKE '%" & txtSearch.Text & "%' OR nama LIKE '%" & txtSearch.Text & "%' OR tempat_lahir LIKE '%" & txtSearch.Text & "%' OR tanggal_lahir LIKE '%" & txtSearch.Text & "%' OR pekerjaan LIKE '%" & txtSearch.Text & "%' OR telepon LIKE '%" & txtSearch.Text & "%' OR alamat LIKE '%" & txtSearch.Text & "%' OR kode_pos LIKE '%" & txtSearch.Text & "%' "
    Case "KTP"
        searchSql = "SELECT no_ktp, nama, tempat_lahir, tanggal_lahir, jenis_kelamin, pekerjaan, telepon, alamat, kode_pos FROM tblAnggota WHERE no_ktp LIKE '%" & txtSearch.Text & "%'"
    Case "Nama"
        searchSql = "SELECT no_ktp, nama, tempat_lahir, tanggal_lahir, jenis_kelamin, pekerjaan, telepon, alamat, kode_pos FROM tblAnggota WHERE nama LIKE '%" & txtSearch.Text & "%'"
    Case "Tempat Lahir"
        searchSql = "SELECT no_ktp, nama, tempat_lahir, tanggal_lahir, jenis_kelamin, pekerjaan, telepon, alamat, kode_pos FROM tblAnggota WHERE tempat_lahir LIKE '%" & txtSearch.Text & "%'"
    Case "Tanggal Lahir"
        searchSql = "SELECT no_ktp, nama, tempat_lahir, tanggal_lahir, jenis_kelamin, pekerjaan, telepon, alamat, kode_pos FROM tblAnggota WHERE tanggal_lahir LIKE '%" & txtSearch.Text & "%'"
    Case "Pekerjaan"
        searchSql = "SELECT no_ktp, nama, tempat_lahir, tanggal_lahir, jenis_kelamin, pekerjaan, telepon, alamat, kode_pos FROM tblAnggota WHERE pekerjaan LIKE '%" & txtSearch.Text & "%'"
    Case "Telepon"
        searchSql = "SELECT no_ktp, nama, tempat_lahir, tanggal_lahir, jenis_kelamin, pekerjaan, telepon, alamat, kode_pos FROM tblAnggota WHERE telepon LIKE '%" & txtSearch.Text & "%'"
    Case "Alamat"
        searchSql = "SELECT no_ktp, nama, tempat_lahir, tanggal_lahir, jenis_kelamin, pekerjaan, telepon, alamat, kode_pos FROM tblAnggota WHERE alamat LIKE '%" & txtSearch.Text & "%'"
    Case "Kode Pos"
        searchSql = "SELECT no_ktp, nama, tempat_lahir, tanggal_lahir, jenis_kelamin, pekerjaan, telepon, alamat, kode_pos FROM tblAnggota WHERE kode_pos LIKE '%" & txtSearch.Text & "%'"
    Case Else
        MsgBox "Harap pilih jenis pencarian", vbOKOnly + vbInformation, "Peringatan"
        Exit Sub
    End Select
    
    tblAnggota.Open searchSql, dbConn
    Set grid.DataSource = tblAnggota
    grid.Refresh
    initStyleGrid
    tblAnggota.Close
End Sub

Private Sub cmdClear_Click()
    Call loadData
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub
