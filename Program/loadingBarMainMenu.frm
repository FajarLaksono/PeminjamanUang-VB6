VERSION 5.00
Begin VB.Form frmlLoadingBarMainMenu 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "loading"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   -1125
   ClientWidth     =   24000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   90
   ScaleWidth      =   24000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer LoadingBarMainMenuTime 
      Left            =   3120
      Top             =   0
   End
   Begin VB.Shape loadingBarMainMenu 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "frmlLoadingBarMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'form ini digunakan hanya untuk penampilan form frmMenuUtama
'Untuk menambahkan animasi loading bar di bawah toolbar

Function loading() 'jika ingin menjalankan animasi cukup panggil function ini
    Call Me.LoadingBarMainMenuTime_Timer 'memanggil function tersebut untuk menjalankan animasi
    Call Me.loadingBarInisialisasi 'Set ulang nilai loading
End Function

Function loadingBarInisialisasi()
    loadingBarMainMenu.Width = 1 'taruh loading bar ke posisi awal
End Function

Public Sub LoadingBarMainMenuTime_Timer()
'digunakan untuk menjalankan animasi dengan menggunakan pengulangan for, nilai akan selalu ditambah pada pengulangan
    Dim i As Integer 'indek pengulangan
    For i = 0 To 25000
        i = i + 25 'biar loadingnya gk kelamaan jadi ku taruh 25 langsung
        loadingBarMainMenu.Width = i
    Next i
End Sub

Private Sub Form_Load()
    Me.Top = 0 'set loakasi
    Me.Left = 0 'set lokasi
    LoadingBarMainMenuTime.Interval = 0
    'Pojok Kiri Atas
    
    Call loadingBarInisialisasi 'inisialisasi loading bar saat form terbuka
End Sub
