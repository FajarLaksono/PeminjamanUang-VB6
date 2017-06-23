VERSION 5.00
Begin VB.Form frmTentang 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tentang Program"
   ClientHeight    =   4320
   ClientLeft      =   7200
   ClientTop       =   3885
   ClientWidth     =   5880
   Icon            =   "frmTentang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   1
      Top             =   3705
      Width           =   1260
   End
   Begin VB.ListBox listBoxCreditsList 
      Height          =   1815
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.5"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   3885
   End
   Begin VB.Image picIcon 
      Height          =   855
      Left            =   240
      Picture         =   "frmTentang.frx":048A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Peminjaman Uang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1290
      TabIndex        =   3
      Top             =   360
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Creator :"
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   5684
      Y1              =   3600
      Y2              =   3600
   End
End
Attribute VB_Name = "frmTentang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me 'keluar form
End Sub

Private Sub Form_Load()
'set posisi jendela form
    Me.Top = 2000
    Me.Left = 7000
    
'isi listbox
    listBoxCreditsList.Clear
    listBoxCreditsList.AddItem "Fajar Aziz Laksono"
	listBoxCreditsList.AddItem "email : Fajarazizlaksono@gmail.com"
End Sub

