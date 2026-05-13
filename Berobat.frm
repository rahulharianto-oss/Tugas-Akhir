VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF80FF&
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Berobat.frx":0000
         Height          =   3615
         Left            =   240
         TabIndex        =   30
         Top             =   3600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6376
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "noantrian2320006"
            Caption         =   "noantrian2320006"
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
            DataField       =   "tanggal2320006"
            Caption         =   "tanggal2320006"
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
         BeginProperty Column02 
            DataField       =   "namapasien2320006"
            Caption         =   "namapasien2320006"
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
         BeginProperty Column03 
            DataField       =   "umur2320006"
            Caption         =   "umur2320006"
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
         BeginProperty Column04 
            DataField       =   "namadokter2320006"
            Caption         =   "namadokter2320006"
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
         BeginProperty Column05 
            DataField       =   "biayadokter2320006"
            Caption         =   "biayadokter2320006"
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
         BeginProperty Column06 
            DataField       =   "biayaobat2320006"
            Caption         =   "biayaobat2320006"
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
         BeginProperty Column07 
            DataField       =   "totalbiaya2320006"
            Caption         =   "totalbiaya2320006"
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
         BeginProperty Column08 
            DataField       =   "uangbayar2320006"
            Caption         =   "uangbayar2320006"
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
         BeginProperty Column09 
            DataField       =   "kembalian2320006"
            Caption         =   "kembalian2320006"
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
               ColumnWidth     =   1305,071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   7320
         Top             =   7560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\DELL\Documents\PTI2026\Berobat2320006.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\DELL\Documents\PTI2026\Berobat2320006.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "berobat2320006"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   6240
         TabIndex        =   29
         Top             =   1440
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         Format          =   129040385
         CurrentDate     =   46028
      End
      Begin VB.CommandButton btnKeluar 
         Caption         =   "KELUAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   28
         Top             =   7080
         Width           =   1335
      End
      Begin VB.TextBox txtKembali 
         Height          =   495
         Left            =   11640
         TabIndex        =   27
         Top             =   6240
         Width           =   1935
      End
      Begin VB.TextBox txtUang 
         Height          =   495
         Left            =   11640
         TabIndex        =   25
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton btnBatal 
         Caption         =   "BATAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9120
         TabIndex        =   23
         Top             =   6840
         Width           =   1575
      End
      Begin VB.CommandButton btnHapus 
         Caption         =   "HAPUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9120
         TabIndex        =   22
         Top             =   6120
         Width           =   1575
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "EDIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9120
         TabIndex        =   21
         Top             =   5400
         Width           =   1575
      End
      Begin VB.CommandButton btnSimpan 
         Caption         =   "SIMPAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9120
         TabIndex        =   20
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton btnTambah 
         Caption         =   "TAMBAH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9120
         TabIndex        =   19
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox txtBayar 
         Height          =   405
         Left            =   11880
         TabIndex        =   18
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtObat 
         Height          =   375
         Left            =   9360
         TabIndex        =   16
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txtbiayadokter 
         Height          =   375
         Left            =   6960
         TabIndex        =   14
         Top             =   3120
         Width           =   1815
      End
      Begin VB.ComboBox cmbnamadokter 
         Height          =   315
         Left            =   4200
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox txtUmur 
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtPasien 
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox txtCari 
         Height          =   375
         Left            =   10800
         TabIndex        =   6
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtAntrian 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label12 
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
         Height          =   375
         Left            =   11640
         TabIndex        =   26
         Top             =   5760
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Uang Bayar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         TabIndex        =   24
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Total Bayar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         TabIndex        =   17
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Biaya Obat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   15
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Biaya Dokter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   13
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Nama Dokter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Umur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Nama Pasien"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Cari "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Tanggal :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "No Antrian :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Caption         =   "APLIKASI BEROBAT BERSAMA RAHUL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   480
         Width           =   6975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kon As ADODB.Connection
Dim rs As ADODB.Recordset
Dim SQL As String

'================ DATABASE =================
Sub bukadb()
    If kon Is Nothing Then
        Set kon = New ADODB.Connection
        kon.Provider = "Microsoft.Jet.OLEDB.4.0"
        kon.CursorLocation = adUseClient
        kon.Open App.Path & "\berobat2320006.mdb"
    End If

    If rs Is Nothing Then
        Set rs = New ADODB.Recordset
    End If
End Sub

'================ TAMPIL DATA =================
Sub tampil()
    bukadb
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM berobat2320006", kon, adOpenStatic, adLockReadOnly
    Set DataGrid1.DataSource = rs
End Sub

'================ FORM CONTROL =================
Sub bersih()
    txtAntrian.Text = ""
    txtPasien.Text = ""
    txtUmur.Text = ""
    cmbnamadokter.Text = "==Pilih=="
    txtObat.Text = ""
    txtBayar.Text = ""
    txtUang.Text = ""
    txtKembali.Text = ""
End Sub

Sub aktif()
    txtAntrian.Enabled = True
    txtPasien.Enabled = True
    txtUmur.Enabled = True
    cmbnamadokter.Enabled = True
    txtbiayadokter.Enabled = True

    txtObat.Enabled = True
    txtBayar.Locked = True
    txtKembali.Enabled = True
    txtKembali.Locked = True
End Sub

Sub tidakaktif()
    txtAntrian.Enabled = False
    txtPasien.Enabled = False
    txtUmur.Enabled = False
    cmbnamadokter.Enabled = False
    txtbiayadokter.Enabled = False
    txtObat.Enabled = False
    txtBayar.Enabled = False
    txtKembali.Enabled = False
End Sub

'================ KODE OTOMATIS =================
Sub buatkode()
    bukadb
    If rs.State = adStateOpen Then rs.Close

    rs.Open "SELECT * FROM berobat2320006 ORDER BY noantrian2320006", kon

    If rs.EOF Then
        txtAntrian.Text = "T-001"
    Else
        rs.MoveLast
        txtAntrian.Text = "T-" & Format(Val(Right(rs!noantrian2320006, 3)) + 1, "000")
    End If
End Sub

'================ HITUNG =================
Sub hitungtotal()
    txtBayar.Text = Val(txtbiayadokter.Text) * Val(txtObat.Text)
End Sub

Sub uangkembali()
    txtKembali.Text = Val(txtUang.Text) - Val(txtBayar.Text)
End Sub

Private Sub btnKeluar_Click()
Unload
End Sub

'================ BUTTON =================
Private Sub btnTambah_Click()
    aktif
    bersih
    buatkode
    txtPasien.SetFocus
    btnTambah.Enabled = False
    btnSimpan.Enabled = True
End Sub
Private Sub btnBatal_Click()
 bersih
 tidakaktif
 btnBatal.Enabled = False
 btnTambah.Enabled = True
 btnSimpan.Enabled = False
 btnHapus.Enabled = False
 btnEdit.Enabled = False
End Sub

Private Sub btnEdit_Click()
bukadb
 SQL = "UPDATE berobat2320006 SET " & _
 "tanggal2320006 = #" & Format(DTPicker1.Value, "yyyy-MM-dd") & "#, " & _
 "namapasien2320006 = '" & txtPasien.Text & "', " & _
 "umur2320006 = '" & txtUmur.Text & "', " & _
 "namadokter2320006 = '" & cmbnamadokter.Text & "', " & _
 "biayadokter2320006 = " & Val(txtbiayadokter.Text) & ", " & _
 "biayaobat2320006 = " & Val(txtObat.Text) & ", " & _
 "totalbiaya2320006 = " & Val(txtBayar.Text) & ", " & _
 "uangbayar2320006 = " & Val(txtUang.Text) & ", " & _
 "kembalian2320006 = " & Val(txtKembali.Text) & " " & _
 "WHERE noantrian2320006 = '" & txtAntrian.Text & "'"
 kon.Execute SQL
 tampil
 bersih
 btnTambah.Enabled = True
 btnSimpan.Enabled = False
 btnBatal.Enabled = False
 btnHapus.Enabled = False
 btnEdit.Enabled = False
 tidakaktif
End Sub

Private Sub btnSimpan_Click()
    bukadb

    SQL = "INSERT INTO berobat2320006 (" & _
          "noantrian2320006, [tanggal2320006], namapasien2320006, umur2320006, " & _
          "namadokter2320006, biayadokter2320006, biayaobat2320006, " & _
          "totalbiaya2320006, uangbayar2320006, kembalian2320006) VALUES (" & _
          "'" & txtAntrian.Text & "', " & _
          "#" & Format(DTPicker1.Value, "yyyy/mm/dd") & "#, " & _
          "'" & txtPasien.Text & "', " & _
          Val(txtUmur.Text) & ", " & _
          "'" & cmbnamadokter.Text & "', " & _
          Val(txtbiayadokter.Text) & ", " & _
          Val(txtObat.Text) & ", " & _
          Val(txtBayar.Text) & ", " & _
          Val(txtUang.Text) & ", " & _
          Val(txtKembali.Text) & ")"

    kon.Execute SQL, , adExecuteNoRecords

    MsgBox "Data berhasil disimpan", vbInformation

    tampil
    bersih
    tidakaktif

    btnTambah.Enabled = True
    btnSimpan.Enabled = False
End Sub

Private Sub btnHapus_Click()
    If MsgBox("Yakin data akan dihapus?", vbYesNo + vbQuestion) = vbYes Then
        bukadb
        kon.Execute "DELETE FROM berobat2320006 WHERE noantrian2320006='" & txtAntrian.Text & "'"
        tampil
        bersih
        tidakaktif
    End If
End Sub

'================ EVENT =================
Private Sub txtObat_Change()
    hitungtotal
End Sub

Private Sub txtCari_Change()
    bukadb
    If rs.State = adStateOpen Then rs.Close

    rs.Open "SELECT * FROM berobat2320006 WHERE " & _
            "noantrian2320006 LIKE '%" & txtCari.Text & "%' OR " & _
            "namapasien2320006 LIKE '%" & txtCari.Text & "%'", kon

    Set DataGrid1.DataSource = rs
End Sub

Private Sub DataGrid1_Click()
    txtAntrian.Text = DataGrid1.Columns(0).Text
    DTPicker1.Value = DataGrid1.Columns(1).Text
    txtPasien.Text = DataGrid1.Columns(2).Text
    txtUmur.Text = DataGrid1.Columns(3).Text
    cmbnamadokter.Text = DataGrid1.Columns(4).Text
    txtbiayadokter.Text = DataGrid1.Columns(5).Text
    txtObat.Text = DataGrid1.Columns(6).Text
    txtBayar.Text = DataGrid1.Columns(7).Text
    txtUang.Text = DataGrid1.Columns(8).Text
    txtKembali.Text = DataGrid1.Columns(9).Text

    aktif
    btnHapus.Enabled = True
    btnSimpan.Enabled = False
End Sub

'================ KETENTUAN JENIS SERVIS =================
Private Sub cmbnamadokter_Click()
    Select Case cmbnamadokter.Text
        Case "Mulyono"
            txtbiayadokter.Text = 75000
        Case "Zainudin"
            txtbiayadokter.Text = 150000
        Case "Megachan"
            txtbiayadokter.Text = 175000
        Case "Wolveriness"
            txtbiayadokter.Text = 200000
    End Select

    hitungtotal
End Sub

'================ FORM =================
Private Sub Form_Load()
    With cmbnamadokter
        .Clear
        .AddItem "Mulyono"
        .AddItem "Zainudin"
        .AddItem "Megachan"
        .AddItem "Wolveriness"
        .Text = "==Pilih=="
    End With
End Sub

Private Sub Form_Activate()
    tampil
    bersih
    tidakaktif
    btnTambah.Enabled = True
    btnSimpan.Enabled = False
    btnHapus.Enabled = False
End Sub




Private Sub txtUang_Change()
uangkembali
End Sub
