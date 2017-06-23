VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peminjaman Uang : Login Petugas"
   ClientHeight    =   4860
   ClientLeft      =   5235
   ClientTop       =   2625
   ClientWidth     =   9165
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":10CA
   ScaleHeight     =   4860
   ScaleWidth      =   9165
   Begin VB.TextBox txtNIK 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000006&
      Height          =   300
      Left            =   5400
      TabIndex        =   0
      Text            =   "NIK"
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000006&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5400
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Password"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Shape loadingBar 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.Label cmdLogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      Left            =   7260
      TabIndex        =   2
      Top             =   2760
      Width           =   600
   End
   Begin VB.Label errMsg 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "ErrConditionLogin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   4200
      Width           =   8295
   End
   Begin VB.Shape shapeLogin 
      BorderColor     =   &H80000004&
      Height          =   1695
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Shape shapeButton 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   2715
      Width           =   1200
   End
   Begin VB.Shape shapeMsg 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   8775
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim un_touched As Boolean 'variabel boolean untuk memeriksa kekosongan textbox
Dim un_touched_pass As Boolean 'fungsi sama seperti di atas tapi untuk txtpassword

'Function Tambahan Start
Function clear_unInputForm() 'ditujukan untuk tampilan form
    'fungsi ini untuk memudahkan kita memeriksa kekosongan pada textbox
    'agar jika kosong bisa kita isikan NIK dan PASS
    'dan ketika disentuh maka tulisan tersebut hilang
    'jika textbox kosong maka akan di isi dengan NIK atau * tapi jika isi maka akan tetap dengan isi tersebut
    If txtNIK.Text = "" Then
        txtNIK.Text = "NIK"
        un_touched = True
    End If
    
    If txtPass.Text = "" Then
        txtPass.Text = "*"
        un_touched_pass = True
    End If
End Function

Function enterHited(KeyAscii As Integer) 'kubuat function biar g nulis ulang 3 baris berkali2
    If KeyAscii = 13 Then 'jika tombol enter ditekan maka
        Call cmdLogin_Click
    End If
End Function

Function loading() 'ditujukan untuk tampilan form
'ini tidak ada hubunganya sama frmLoadingBarMainMenu, frmLogin mempunyai loading bar sendiri
    loadingBarInisialisasi 'menyetel loading pada titik awal
    loadingBarTime_Timer 'jalankan animasi
    loadingBarInisialisasi 'setel ulang
End Function

Private Sub loadingBarTime_Timer() 'ditujukan untuk tampilan form
    'animasi loading bar untuk form login
    Dim i As Integer
    For i = 1 To 10000
        loadingBar.Width = i
    Next i
End Sub

Function loadingBarInisialisasi() 'ditujukan untuk tampilan form
    loadingBar.Width = 1 'taruh loading bar pada garis awal mulai
End Function

'SUB catagory Start
Private Sub txtNIK_GotFocus() 'ditujukan untuk tampilan form
    'akan segera mengkosongkan jika user fokus pada textbox dan untouched bernilai false
    If un_touched = True Then
        txtNIK.Text = ""
        un_touched = False
    End If
End Sub

Private Sub txtPass_GotFocus() 'ditujukan untuk tampilan form
    'akan segera mengkosongkan jika user fokus pada textbox dan untouched bernilai false
    If un_touched_pass = True Then
        txtPass.Text = ""
        un_touched_pass = False
    End If
End Sub

Private Sub txtNIK_LostFocus() 'ditujukan untuk tampilan form
    clear_unInputForm 'jika textbox kosong maka akan di isi dengan NIK atau * tapi jika isi maka akan tetap dengan isi tersebut
End Sub

Private Sub txtPass_LostFocus() 'ditujukan untuk tampilan form
    clear_unInputForm 'jika textbox kosong maka akan di isi dengan NIK atau * tapi jika isi maka akan tetap dengan isi tersebut
End Sub

Private Sub txtNIK_KeyPress(KeyAscii As Integer)
    enterHited (KeyAscii) 'tinggal panggil, parameter keyascii dilempar kembali menjadi argumen di enterHited
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    enterHited (KeyAscii) 'idem
End Sub

Private Sub cmdLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'ditujukan untuk tampilan form
    'berfungsi untuk mengubah warna tombol saat mouse berada tepat di atas tombol
    shapeButton.BackStyle = 1
    shapeButton.BackColor = &H80000005
    cmdLogin.ForeColor = &H80000006
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'berfungsi untuk mengubah warna tobol saat mouse keluar dari atas tombol (di atas form)
    shapeButton.BackStyle = 0
    shapeButton.BackColor = &H80000006
    cmdLogin.ForeColor = &H80000005
End Sub

Private Sub cmdLogin_Click()
retry: 'label retry
On Error GoTo errHandler 'jika terjadi error maka loncat dan menuju label errhandler
    errMsg.Visible = False 'inisialisasi errmsg dan shapemsg
    shapeMsg.Visible = False
    Call loading 'animasi loading di jalankan di frmlogin
    tblPetugas.Open "SELECT * FROM tblpetugas WHERE nik='" & txtNIK.Text & "'", dbConn 'cek keberadaan NIK
        If Not tblPetugas.EOF Then 'jika ada
            If tblPetugas!Password = txtPass.Text Then 'cek kebenaran password, jika benar maka
                'menyetel nilai2 yang ada di statusbar frmMenuUtama
                frmMenuUtama.strBar.Panels(3).Text = tblPetugas!nik
                frmMenuUtama.strBar.Panels(4).Text = tblPetugas!nama
                tblPetugas.Close 'tutup kalo udah selesai biar nanti bisa dipake lg
                frmMenuUtama.Show 'buka frmMenuUtama
                Unload frmLogin 'Tutup form ini
                Exit Sub
            Else
                errMsg.Visible = True 'fungsi hampir sama kayak msgbox, tapi biar keren kubuat dengan caraku sendiri, ini untuk menampilkan pesan error di bawah frmlogin
                shapeMsg.Visible = True 'aktifkan juga box yang dibelakangnya
                errMsg.Caption = "Anda salah memasukan Password !" 'beri nilai ke msgbox sebagai pesan error untuk user
                txtPass.SetFocus 'fokus ke kesalahan user yaitu di txtpass
            End If
        Else
            errMsg.Visible = True 'idem bro
            shapeMsg.Visible = True
            errMsg.Caption = "Akun Tidak ditemukan"
            txtNIK.SetFocus 'fokus ke NIK
        End If
    tblPetugas.Close 'tutup kalo udah selesai biar nanti bisa dipake lg
    Exit Sub 'exitsub saya gunakan untuk mengeluarkan jalur eksekusi jika eksekusi berjalan lancar, saya keluarkan karena
    'dihindarkan agar jalur eksekusi tidak eksekusi beberapa pernyataan di bawah (di bawah label errHandler)
    'karena beberapa pertanyaan tersebut dikhususkan jika eksekusi tidak berjalan lancar
    
errHandler: '
    s = MsgBox("[Error] kesalahan saat mencoba untuk menghubungkan ke server, silahkan coba lagi nanti", vbRetryCancel + vbInformation, "Kesalahan")
    If s = vbRetry Then GoTo retry ' biasanya error berada pada koneksi program ke server, jika user klik retry maka akan kembali ke label retry di atas
End Sub

Private Sub Form_Load()
    'inisialisasi
    Me.Top = 2200
    Me.Left = 5200
     
    Call setConn ' ciptakan variabel dan inisialiasi variabel yang ada pada module mdDB
    shapeMsg.Visible = False 'inisialiasi
    errMsg.Visible = False
    errMsg.Caption = ""
    
    'inisialisasi variabel un_touched
    un_touched = True
    un_touched_pass = True
    
    'Mengatur ToolTip
    txtNIK.ToolTipText = "Masukan NIK"
    txtPass.ToolTipText = "Masukan Password"
    cmdLogin.ToolTipText = "Masuk"
    
    Call loading
End Sub
