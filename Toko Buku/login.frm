VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "LOGIN"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   14400
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "KELUAR"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "LOGIN"
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
      Left            =   5880
      TabIndex        =   0
      Top             =   2760
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   "Masuk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   5
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         TabIndex        =   3
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim user, password As String
    user = Text1.Text
    pass = Text2.Text

    If user = "admin" And pass = "admin" Then
        pesan = MsgBox("Login Berhasil", vbInformation, "Sukses")
        Form1.Hide
        Form2.Show
    Else
        pesan = MsgBox("Login Gagal", vbCritical, "Error")
        Text1.SetFocus
    End If

    'Dim rs As New ADODB.Recordset
    'Set rs = panggil("SELECT * FROM tb_admin WHERE username = '" & Text1.Text & "' And password = '" & Text2.Text & "'")
    
    'If rs.RecordCount = 0 Then
        'MsgBox "Login Gagal!", vbCritical + vbOKOnly, "Error"
        'Text1.SetFocus
        'Exit Sub
    'Else
        'MsgBox "Login Berhasil", vbInformation, "Success"
        'Form2.Show
        'Form1.Hide
    'End If


    'Dim rs As New ADODB.Recordset
    'Set rs = JalankanSQL("SELECT * FROM tb_admin WHERE username = '" & Text1.Text & "' And password = '" & Text2.Text & "'")
    
    'If rs.RecordCount = 0 Then
        'MsgBox "Login Gagal!", vbCritical + vbOKOnly, "Error"
        'Text1.SetFocus
        'Exit Sub
    'Else
        'MsgBox "Login Berhasil", vbInformation, "Success"
        'Form2.Show
        'Form1.Hide
    'End If
End Sub

Private Sub Command2_Click()
    Dim pesan As String
    pesan = MsgBox("Ingin Menutup Aplikasi ?", vbOKCancel, "Tutup Aplikasi")
    If pesan = vbOK Then End
End Sub

Private Sub Image1_Click()

End Sub
