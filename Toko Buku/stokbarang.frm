VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form4 
   Caption         =   "STOK BARANG"
   ClientHeight    =   9915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   LinkTopic       =   "Form4"
   ScaleHeight     =   9915
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   9840
      Top             =   9360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PRINT"
      Height          =   435
      Left            =   11880
      TabIndex        =   23
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIEW"
      Height          =   435
      Left            =   10440
      TabIndex        =   22
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   9240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   480
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   16
      Top             =   5160
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7695
      Left            =   10440
      TabIndex        =   15
      Top             =   1440
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   13573
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   14
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   5160
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "stokbarang.frx":0000
      Left            =   2640
      List            =   "stokbarang.frx":0019
      TabIndex        =   11
      Text            =   "Pilih"
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Text3 
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
      Left            =   2640
      TabIndex        =   10
      Top             =   4200
      Width           =   3495
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
      Left            =   2640
      TabIndex        =   9
      Top             =   3600
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
      Left            =   2640
      TabIndex        =   8
      Top             =   3000
      Width           =   3495
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Left            =   2640
      TabIndex        =   7
      Top             =   2040
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
      _ExtentY        =   556
      Calendar        =   "stokbarang.frx":0053
      Caption         =   "stokbarang.frx":016B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "stokbarang.frx":01D7
      Keys            =   "stokbarang.frx":01F5
      Spin            =   "stokbarang.frx":0253
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   1835335681
      Value           =   44371
      CenturyMode     =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BERANDA"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "TOKO BUKU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1920
      TabIndex        =   21
      Top             =   1080
      Width           =   2280
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "No. Buku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   19
      Top             =   9360
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label ljam 
      AutoSize        =   -1  'True
      Caption         =   "jam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9000
      TabIndex        =   18
      Top             =   0
      Width           =   510
   End
   Begin VB.Label ltanggal 
      AutoSize        =   -1  'True
      Caption         =   "tanggal"
      Height          =   195
      Left            =   8760
      TabIndex        =   17
      Top             =   480
      Width           =   525
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Buku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   4320
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Harga Buku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nama Buku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Jenis Buku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Stok Barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form4.Hide
    Form2.Show
End Sub

Sub bersih()
    TDBDate1.Text = ""
    Text4.Text = ""
    Combo1.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub

Private Sub Command2_Click()
    If Combo1.Text = "" Then
        MsgBox "Data Yang dimasukkan Belum Lengkap", vbInformation, "Konfirmasi"
        Combo1.SetFocus
        Exit Sub
    ElseIf Text1.Text = "" Then
        MsgBox "Data Yang dimasukkan Belum Lengkap", vbInformation, "Konfirmasi"
        Text1.SetFocus
        Exit Sub
    ElseIf Text2.Text = "" Then
        MsgBox "Data Yang dimasukkan Belum Lengkap", vbInformation, "Konfirmasi"
        Text2.SetFocus
        Exit Sub
    ElseIf Text3.Text = "" Then
        MsgBox "Data Yang dimasukkan Belum Lengkap", vbInformation, "Konfirmasi"
        Text3.SetFocus
        Exit Sub
    End If

    sql = "INSERT INTO tb_stokbarang (tgl,jenisBuku,namaBuku,hargaBuku,jumlahBuku) " _
    & "VALUES ('" & TDBDate1.Text _
    & "','" & Combo1.Text _
    & "','" & Text1.Text _
    & "','" & Text2.Text _
    & "','" & Text3.Text & "')"
    
    con.Execute sql
    MsgBox "Simpan Data", vbInformation, "Konfirmasi"
    tampil
    bersih
End Sub

Private Sub Command3_Click()
    sql = "UPDATE tb_stokbarang SET tgl = '" & TDBDate1.Text _
    & "', namaBuku = '" & Text1.Text _
    & "', jenisBuku = '" & Combo1.Text _
    & "', hargaBuku = '" & Text2.Text _
    & "', jumlahBuku = '" & Text3.Text _
    & "' WHERE no = '" & Text4.Text & "'"
    
    con.Execute sql
    MsgBox "Update Data", vbInformation, "Konfirmasi"
    tampil
    bersih
End Sub

Private Sub Command4_Click()
    sql = "DELETE FROM tb_stokbarang WHERE namaBuku = '" & Text1.Text & "'"
    
    con.Execute sql
    MsgBox "Hapus Data", vbInformation, "Konfirmasi"
    tampil
    bersih
End Sub

Private Sub Command5_Click()
    bersih
End Sub

Private Sub Command6_Click()
    'CR.ReportFileName = "D:\Kuliah-Wahyu\SEMESTER 4\Praktek Pemrograman Visual 1\Tugas Akhir\Toko Buku\laporanStokBarang.rpt"
    CR.ReportFileName = direktoriExe & "laporanStokBarang.rpt"
    CR.WindowState = crptMaximized
    CR.Destination = crptToWindow
    CR.DiscardSavedData = True
    CR.WindowShowCloseBtn = True
    CR.Action = 1
End Sub

Private Sub Command7_Click()
    'CR.ReportFileName = "D:\Kuliah-Wahyu\SEMESTER 4\Praktek Pemrograman Visual 1\Tugas Akhir\Toko Buku\laporanStokBarang.rpt"
    CR.ReportFileName = direktoriExe & "laporanStokBarang.rpt"
    CR.WindowState = crptMaximized
    CR.Destination = crptToPrinter
    CR.DiscardSavedData = True
    CR.WindowShowCloseBtn = True
    CR.Action = 1
End Sub

Private Sub DataGrid1_Click()
    Text4.Text = DataGrid1.Columns(0)
    TDBDate1.Text = DataGrid1.Columns(1)
    Combo1.Text = DataGrid1.Columns(2)
    Text1.Text = DataGrid1.Columns(3)
    Text2.Text = DataGrid1.Columns(4)
    Text3.Text = DataGrid1.Columns(5)
End Sub

Private Sub Form_Load()
    bukakoneksi
    tampil
End Sub

Sub tampil()
    con.CursorLocation = adUseClient
    sql = "SELECT * FROM tb_stokbarang"
    Set tabel = con.Execute(sql)
    Set DataGrid1.DataSource = tabel
End Sub

Private Sub Timer1_Timer()
    ljam.Caption = Time
    ltanggal.Caption = Format(Date, "dddd, d mmmm, yyyy")
End Sub
