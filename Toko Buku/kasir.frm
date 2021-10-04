VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form3 
   Caption         =   "KASIR (PEMBAYARAN)"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18915
   LinkTopic       =   "Form3"
   ScaleHeight     =   10200
   ScaleWidth      =   18915
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   7680
      Top             =   9360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PRINT"
      Height          =   435
      Left            =   9720
      TabIndex        =   27
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PREVIEW"
      Height          =   435
      Left            =   8280
      TabIndex        =   26
      Top             =   9360
      Width           =   1215
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Left            =   2880
      TabIndex        =   25
      Top             =   2040
      Width           =   3615
      _Version        =   65536
      _ExtentX        =   6376
      _ExtentY        =   556
      Calendar        =   "kasir.frx":0000
      Caption         =   "kasir.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "kasir.frx":0184
      Keys            =   "kasir.frx":01A2
      Spin            =   "kasir.frx":0200
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
      ValueVT         =   1866661889
      Value           =   44371
      CenturyMode     =   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "kasir.frx":0228
      Left            =   2880
      List            =   "kasir.frx":0241
      TabIndex        =   23
      Text            =   "Pilih"
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   480
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "kasir.frx":027B
      Left            =   2880
      List            =   "kasir.frx":0294
      TabIndex        =   20
      Text            =   "Pilih"
      Top             =   4800
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BERANDA"
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7575
      Left            =   8280
      TabIndex        =   17
      Top             =   1440
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   13361
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
   Begin VB.TextBox Text8 
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   7680
      Width           =   3615
   End
   Begin VB.TextBox Text7 
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   7080
      Width           =   3615
   End
   Begin VB.TextBox Text6 
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
      Left            =   2880
      TabIndex        =   14
      Top             =   5280
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "INPUT"
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
      TabIndex        =   9
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
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
      Left            =   2040
      TabIndex        =   8
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox Text4 
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
      Left            =   2880
      TabIndex        =   7
      Top             =   4200
      Width           =   3615
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
      Left            =   2880
      TabIndex        =   6
      Top             =   3600
      Width           =   3615
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
      Left            =   2880
      TabIndex        =   5
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Pembelian"
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
      Left            =   240
      TabIndex        =   24
      Top             =   2040
      Width           =   1785
   End
   Begin VB.Label ltanggal 
      AutoSize        =   -1  'True
      Caption         =   "tanggal"
      Height          =   195
      Left            =   8760
      TabIndex        =   22
      Top             =   480
      Width           =   525
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
      TabIndex        =   21
      Top             =   0
      Width           =   510
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Data Penjualan"
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
      Left            =   8280
      TabIndex        =   18
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Kembalian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   13
      Top             =   7800
      Width           =   1110
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Diskon %"
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
      Left            =   240
      TabIndex        =   12
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   11
      Top             =   7200
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah dibayar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   5400
      Width           =   1605
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Harga Per/Buku"
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
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   1440
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   1050
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   2280
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
    Combo1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Combo2.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
End Sub

Private Sub Command1_Click()
    bersih
End Sub

Private Sub Command2_Click()
    Dim jenisBuku, namaBuku, jenisDiskon As String
    Dim hargaBuku, jumlahBuku, diskon, jumlahBayar As Double
    
    jenisBuku = Combo1.Text
    namaBuku = Text2.Text
    hargaBuku = Text3.Text
    jumlahBuku = Text4.Text
    jenisDiskon = Combo2.Text
    jumlahBayar = Text6.Text
    
    If Combo1.Text = "" Then
        MsgBox "Data Yang dimasukkan Belum Lengkap", vbInformation, "Konfirmasi"
        Combo1.SetFocus
        Exit Sub
    ElseIf Text2.Text = "" Then
        MsgBox "Data Yang dimasukkan Belum Lengkap", vbInformation, "Konfirmasi"
        Text2.SetFocus
        Exit Sub
    ElseIf Text3.Text = "" Then
        MsgBox "Data Yang dimasukkan Belum Lengkap", vbInformation, "Konfirmasi"
        Text3.SetFocus
        Exit Sub
    ElseIf Text4.Text = "" Then
        MsgBox "Data Yang dimasukkan Belum Lengkap", vbInformation, "Konfirmasi"
        Text4.SetFocus
        Exit Sub
    ElseIf Combo2.Text = "" Then
        MsgBox "Data Yang dimasukkan Belum Lengkap", vbInformation, "Konfirmasi"
        Combo2.SetFocus
        Exit Sub
    ElseIf Text6.Text = "" Then
        MsgBox "Data Yang dimasukkan Belum Lengkap", vbInformation, "Konfirmasi"
        Text6.SetFocus
        Exit Sub
    End If
    
    If jenisDiskon = "5%" Then
        diskon = 0.05 * hargaBuku
    ElseIf jenisDiskon = "10%" Then
        diskon = 0.1 * hargaBuku
    ElseIf jenisDiskon = "15%" Then
        diskon = 0.15 * hargaBuku
    ElseIf jenisDiskon = "20%" Then
        diskon = 0.2 * hargaBuku
    ElseIf jenisDiskon = "25%" Then
        diskon = 0.25 * hargaBuku
    ElseIf jenisDiskon = "30%" Then
        diskon = 0.3 * hargaBuku
    Else: jenisDiskon = 0
    End If
    
    Text7.Text = (hargaBuku - diskon) * jumlahBuku 'Total Harga'
    Text8.Text = jumlahBayar - Text7 'Kembalian'
    
    sql = "INSERT INTO tb_kasir (tglPembelian,jenisBuku,namaBuku,hargaBuku,jumlahBuku,jenisDiskon,totalHarga) " _
    & "VALUES ('" & TDBDate1.Text _
    & "','" & Combo1.Text _
    & "','" & Text2.Text _
    & "','" & Text3.Text _
    & "','" & Text4.Text _
    & "','" & Combo2.Text _
    & "','" & Text7.Text & "')"
    
    con.Execute sql
    MsgBox "Simpan Data", vbInformation, "Konfirmasi"
    tampil
    bersih
End Sub

Private Sub Command3_Click()
    Form3.Hide
    Form2.Show
End Sub

Private Sub Command4_Click()
    'CR.ReportFileName = "D:\Kuliah-Wahyu\SEMESTER 4\Praktek Pemrograman Visual 1\Tugas Akhir\Toko Buku\laporanPenjualan.rpt"
    CR.ReportFileName = direktoriExe & "laporanPenjualan.rpt"
    CR.WindowState = crptMaximized
    CR.Destination = crptToWindow
    CR.DiscardSavedData = True
    CR.WindowShowCloseBtn = True
    CR.Action = 1
End Sub

Private Sub Command5_Click() 'printer'
    'CR.ReportFileName = "D:\Kuliah-Wahyu\SEMESTER 4\Praktek Pemrograman Visual 1\Tugas Akhir\Toko Buku\laporanPenjualan.rpt"
    CR.ReportFileName = direktoriExe & "laporanPenjualan.rpt"
    CR.WindowState = crptMaximized
    CR.Destination = crptToPrinter
    CR.DiscardSavedData = True
    CR.WindowShowCloseBtn = True
    CR.Action = 1
End Sub

Private Sub Form_Load()
    bukakoneksi
    tampil
End Sub

Sub tampil()
    con.CursorLocation = adUseClient
    sql = "SELECT * FROM tb_kasir"
    Set tabel = con.Execute(sql)
    Set DataGrid1.DataSource = tabel
End Sub

Private Sub Timer1_Timer()
    ljam.Caption = Time
    ltanggal.Caption = Format(Date, "dddd, d mmmm, yyyy")
End Sub
