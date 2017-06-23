VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReportHutang 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rincian Hutang"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9195
   Icon            =   "frmReportHutang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9195
   Begin VB.TextBox txtLunas 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "status"
      ToolTipText     =   "Status peminjaman"
      Top             =   7680
      Width           =   4455
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      ToolTipText     =   "Cetak informasi"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txtPaket 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   350
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Paket yang dipilih"
      Top             =   2760
      Width           =   5175
   End
   Begin VB.TextBox txtNoKTP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   350
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Nomer KTP Anggota"
      Top             =   480
      Width           =   5175
   End
   Begin VB.Frame frm 
      Caption         =   "Foto Bukti Jaminan"
      Height          =   2175
      Left            =   6960
      TabIndex        =   22
      Top             =   2400
      Width           =   2055
      Begin VB.Image imgJaminan 
         Height          =   1800
         Left            =   120
         Picture         =   "frmReportHutang.frx":048A
         Stretch         =   -1  'True
         ToolTipText     =   "Foto bukti jaminan"
         Top             =   240
         Width           =   1800
      End
   End
   Begin VB.Frame frmPengembalian 
      Caption         =   "Pelunasan"
      Height          =   2895
      Left            =   240
      TabIndex        =   21
      Top             =   4680
      Width           =   8775
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mygrid 
         Height          =   2535
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Daftar riwayat pelunasan"
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   4
         BackColorBkg    =   -2147483634
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
   End
   Begin VB.TextBox txtBesarBunga 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   350
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Besar bunga"
      Top             =   3720
      Width           =   5175
   End
   Begin VB.TextBox txtBesarMeminjam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
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
      Left            =   1560
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Besar Meminjam"
      Top             =   2280
      Width           =   5175
   End
   Begin VB.TextBox txtBesarCicilan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   350
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Besar cicilan"
      Top             =   4200
      Width           =   5175
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Close"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      ToolTipText     =   "Keluar"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txtTglMulai 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   350
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Hari dimulai hutang"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtTglAkhir 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   350
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Hari akhir jangka hutang"
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txtJaminan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Jaminan Meminjam"
      Top             =   1800
      Width           =   5175
   End
   Begin VB.TextBox txtNama 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   350
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Nama Anggota"
      Top             =   960
      Width           =   5175
   End
   Begin Crystal.CrystalReport crystalReport 
      Left            =   5760
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblBesarBunga 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Besar Bunga"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3750
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No KTP"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1005
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Jaminan"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   1850
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Besar Meminjam"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   2300
      Width           =   1335
   End
   Begin VB.Label lblPaket 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Paket"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2800
      Width           =   1335
   End
   Begin VB.Label lblBesarCicilan 
      BackStyle       =   0  'Transparent
      Caption         =   "Besar Cicilan"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4220
      Width           =   1335
   End
   Begin VB.Label lblMaxTanggal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Jangka Hutang"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Label lblJangka 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sampai"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   3240
      Width           =   855
   End
   Begin VB.Image imgAnggota 
      Height          =   2000
      Left            =   6960
      Picture         =   "frmReportHutang.frx":1FDD
      Stretch         =   -1  'True
      ToolTipText     =   "Foto Anggota"
      Top             =   240
      Width           =   2000
   End
End
Attribute VB_Name = "frmReportHutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currID As Double

Private Sub cmdCetak_Click()
    tblPengembalian.Open "SELECT * FROM tblpengembalian WHERE id_peminjaman = '" & currID & "'"
        crystalReport.SelectionFormula = "tonumber({tblpeminjaman.id_peminjaman})=" & currID
        If Not tblPengembalian.EOF Then
            crystalReport.ReportFileName = App.Path & "/reports/rptTransaksi.rpt"
        Else
            crystalReport.ReportFileName = App.Path & "/reports/rptTransaksi2.rpt"
        End If
        crystalReport.WindowState = crptMaximized
        crystalReport.RetrieveDataFiles
        crystalReport.Action = 1
    tblPengembalian.Close
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Function showMe(myCurrID As Double)
    currID = CDbl(myCurrID)
    Me.Show
End Function

Function loadData()
    tblPeminjaman.Open "SELECT * FROM tblpeminjaman WHERE id_peminjaman ='" & currID & "'", dbConn
        tblAnggota.Open "SELECT no_ktp, nama, foto FROM tblanggota WHERE no_ktp ='" & tblPeminjaman.Fields("no_ktp") & "'", dbConn
            txtNoKTP.Text = tblAnggota.Fields("no_ktp")
            txtNama.Text = tblAnggota.Fields("nama")
            
            'Check Foto keberadaan foto
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
                
        tblAnggota.Close
        tblJaminan.Open "SELECT * FROM tbljaminan WHERE id_jaminan='" & tblPeminjaman.Fields("id_jaminan") & "'", dbConn
            txtJaminan.Text = tblJaminan.Fields("jenis")
            
            'Check Foto keberadaan foto
                dirFotoJaminan = "/images/jaminan/"
                If Not tblJaminan.Fields("foto") = "" Then
                    'periksa jika foto ada pada direktori, ubah ke default jika foto tidak ditemukan pada direktori
                    If Dir(App.Path + dirFotoJaminan + tblJaminan.Fields("foto")) <> "" Then
                        'foto ditemukan
                        imgJaminan.Picture = LoadPicture(App.Path & dirFotoJaminan & tblJaminan.Fields("foto"))
                    Else
                        'foto tidak ditemukan
                        MsgBox "Terjadi kesalahan dalam pencarian file gambar jaminan !", vbCritical + vbOKOnly, "Peminjaman Uang"
                        imgJaminan.Picture = LoadPicture(App.Path & dirFotoJaminan & "default.JPG")
                    End If
                Else
                    'rekaman foto kosong
                    MsgBox "Terjadi kesalahan dalam pencarian file gambar jaminan !", vbCritical + vbOKOnly, "Peminjaman Uang"
                    imgAnggota.Picture = LoadPicture(App.Path & dirFotoJaminan & "default.JPG")
                End If
                
        tblJaminan.Close
        txtBesarMeminjam.Text = tblPeminjaman.Fields("hutang")
        txtPaket.Text = tblPeminjaman.Fields("paket")
        txtTglMulai = tblPeminjaman.Fields("tanggal_meminjam")
        txtTglAkhir = getTglAkhir(tblPeminjaman.Fields("tanggal_meminjam"), tblPeminjaman.Fields("paket"))
        txtBesarBunga.Text = tblPeminjaman.Fields("besar_bunga") & " %"
        txtBesarCicilan.Text = algo(tblPeminjaman.Fields("hutang"), getBulan(tblPeminjaman.Fields("paket")), tblPeminjaman.Fields("id_peminjaman"))
        If tblPeminjaman.Fields("lunas") = 0 Then
            txtLunas = "Hutang"
        Else
            txtLunas = "Lunas pada tanggal " & tblPeminjaman.Fields("tanggal_lunas")
        End If
    tblPeminjaman.Close
End Function

Function getTglAkhir(awal As Date, paket As String)
    Select Case paket
    Case "Pelunasan 1 Bulan"
        getTglAkhir = DateAdd("m", 1, awal)
    Case "Pelunasan 2 Bulan"
        getTglAkhir = DateAdd("m", 2, awal)
    Case "Pelunasan 3 Bulan"
        getTglAkhir = DateAdd("m", 3, awal)
    Case "Pelunasan 4 Bulan"
        getTglAkhir = DateAdd("m", 4, awal)
    Case "Pelunasan 5 Bulan"
        getTglAkhir = DateAdd("m", 5, awal)
    Case "Pelunasan 6 Bulan"
        getTglAkhir = DateAdd("m", 6, awal)
    Case "Pelunasan 7 Bulan"
        getTglAkhir = DateAdd("m", 7, awal)
    Case "Pelunasan 8 Bulan"
        getTglAkhir = DateAdd("m", 8, awal)
    Case "Pelunasan 9 Bulan"
        getTglAkhir = DateAdd("m", 9, awal)
    Case "Pelunasan 10 Bulan"
        getTglAkhir = DateAdd("m", 10, awal)
    Case "Pelunasan 11 Bulan"
        getTglAkhir = DateAdd("m", 11, awal)
    Case "Pelunasan 12 Bulan"
        getTglAkhir = DateAdd("m", 12, awal)
    Case "Pelunasan 24 Bulan"
        getTglAkhir = DateAdd("m", 24, awal)
    Case "Pelunasan 36 Bulan"
        getTglAkhir = DateAdd("m", 36, awal)
    End Select
End Function

Function getBulan(paket As String)
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

Private Sub Form_Load()
    Call mdlDB.setConn
    
    Me.Top = 100
    Me.Left = 5500
    
    Call loadData
    Call loadDataGrid
    Call initStyleGrid
End Sub

Function loadDataGrid()
    tblPengembalian.Open "SELECT * FROM tblpengembalian WHERE id_peminjaman = '" & currID & "' ORDER BY id_peminjaman", dbConn
        If Not tblPengembalian.EOF Then
            tblPengembalian.MoveFirst
            mygrid.Rows = 2
            Call initStyleGrid
            Dim i As Integer
            i = 1
            Do While Not tblPengembalian.EOF
                mygrid.Rows = mygrid.Rows + 1
                mygrid.TextMatrix(Val(i), 1) = tblPengembalian.Fields("nik")
                mygrid.TextMatrix(Val(i), 2) = FormatCurrency(tblPengembalian.Fields("uang_bayar"), 2, True, True, True)
                mygrid.TextMatrix(Val(i), 3) = tblPengembalian.Fields("tanggal_bayar")
                tblPengembalian.MoveNext
                i = i + 1
            Loop
            Call initStyleGrid
        End If
    tblPengembalian.Close
End Function

Function initStyleGrid()
    Dim header As String
    header = " | Petugas | Uang bayar | Tanggal Bayar "
    mygrid.FormatString = header
    mygrid.ColWidth(0) = 0
    mygrid.ColWidth(1) = 2800
    mygrid.ColWidth(2) = 2900
    mygrid.ColWidth(3) = 2800
    mygrid.FixedRows = 1
End Function

Private Sub grid_Click()
    MsgBox grid.Row
End Sub
