VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "BERANDA"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14385
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   14385
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "KELUAR"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   6360
      TabIndex        =   0
      Top             =   3000
      Width           =   6255
      Begin VB.CommandButton Command2 
         Caption         =   "STOK BARANG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         TabIndex        =   2
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "KASIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   1
         Top             =   1440
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Hide
    Form3.Show
End Sub

Private Sub Command2_Click()
    Form2.Hide
    Form4.Show
End Sub

Private Sub Command3_Click()
    Dim pesan As String
    pesan = MsgBox("Ingin Menutup Aplikasi ?", vbOKCancel, "Tutup Aplikasi")
    If pesan = vbOK Then End
End Sub

