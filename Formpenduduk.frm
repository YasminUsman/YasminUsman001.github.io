VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Formpenduduk 
   Caption         =   "DATA PENDUDUK"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17955
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   17955
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   7920
      TabIndex        =   22
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5741
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000A&
      Height          =   1335
      Left            =   0
      TabIndex        =   17
      Top             =   5880
      Width           =   7815
      Begin VB.CommandButton Command1 
         Caption         =   "NEXT"
         Height          =   495
         Left            =   3120
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton commhapus 
         Caption         =   "HAPUS"
         Height          =   375
         Left            =   6240
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton commbatal 
         Caption         =   "BATAL"
         Height          =   375
         Left            =   4320
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton commubah 
         Caption         =   "UBAH"
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton commsimpan 
         Caption         =   "SIMPAN"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   7815
      Begin VB.ComboBox comjk 
         Height          =   315
         Left            =   2520
         TabIndex        =   16
         Text            =   "Jenis Kelamin"
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox txtrw 
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox txtrt 
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtalamat 
         Height          =   975
         Left            =   2520
         TabIndex        =   13
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox txtnamakk 
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox txtnikk 
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtkdalternatif 
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "RW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "RT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "JENIS KELAMIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "ALAMAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "NAMA KEPALA KELUARGA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "NIKK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "KODE ALTERNATIF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "DATA PENDUDUK"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   120
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Formpenduduk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim penduduk As New ADODB.Recordset

Private Sub Command1_Click()
Formkerja.Show
End Sub

Private Sub commbatal_Click()
Call kosong
End Sub

Private Sub Commhapus_Click()
koneksidb.Execute "delete from tb_penduduk where kd_alternatif='" & txtkdalternatif & "'"
Call refreshh
Call kosong
txtkdalternatif.SetFocus
End Sub

Private Sub Commsimpan_Click()
If txtkdalternatif = "" Then
MsgBox "Kode Kosong", vbExclamation, "pesan"
txtkdalternatif.SetFocus
Exit Sub
End If
    If txtnikk = "" Then
    MsgBox "NIKK Kosong", vbExclamation, "pesan"
    txtnikk.SetFocus
    Exit Sub
    End If
If txtnamakk = "" Then
MsgBox "Nama Kosong", vbExclamation, "pesan"
txtnamakk.SetFocus
Exit Sub
End If
    If txtalamat = "" Then
    MsgBox "Alamat Kosong", vbExclamation, "pesan"
    txtalamat.SetFocus
    Exit Sub
    End If
If comjk = "" Then
MsgBox "Jenis Kelamin Kosong", vbExclamation, "pesan"
comjk.SetFocus
Exit Sub
End If
    If txtrt = "" Then
    MsgBox "RT Kosong", vbExclamation, "pesan"
    txtrt.SetFocus
    Exit Sub
    End If
If txtrw = "" Then
MsgBox "RW Kosong", vbExclamation, "pesan"
txtrw.SetFocus
Exit Sub
End If
Set penduduk = New ADODB.Recordset
penduduk.Open "select*from tb_penduduk where kd_alternatif='" & txtkdalternatif & "'", koneksidb
If Not penduduk.EOF Then
MsgBox "Kode Sudah Ada", vbCritical, "pesan"
txtkdalternatif = ""
txtkdalternatif.SetFocus
Exit Sub
Else
koneksidb.Execute "insert into tb_penduduk(kd_alternatif,nikk,namakk,alamat,jk,rt,rw) value ('" & txtkdalternatif & "','" & txtnikk & "','" & txtnamakk & "','" & txtalamat & "','" & comjk & "','" & txtrt & "','" & txtrw & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = penduduk
With DataGrid1
End With
Call edit_grid
End If
End Sub

Private Sub commubah_Click()
koneksidb.Execute "update tb_penduduk set nikk='" & txtnikk & "',namakk='" & txtnamakk & "',alamat='" & txtalamat & "',jk='" & comjk & "',rt='" & txtrt & "',rw='" & txtrw & "' where kd_alternatif='" & txtkdalternatif & "'"
Call update
Call edit_grid
Call kosong
End Sub

Private Sub DataGrid1_Click()
txtkdalternatif.Text = penduduk!kd_alternatif
txtnikk.Text = penduduk!nikk
txtnamakk.Text = penduduk!namakk
txtalamat.Text = penduduk!alamat
comjk.Text = penduduk!jk
txtrt.Text = penduduk!rt
txtrw.Text = penduduk!rw
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = penduduk
With penduduk
End With
Call edit_grid
With comjk
    .AddItem "laki-laki"
    .AddItem "perempuan"
End With
End Sub
Sub tampil_data()
Set penduduk = New ADODB.Recordset
penduduk.ActiveConnection = koneksidb
penduduk.CursorLocation = adUseClient
penduduk.LockType = adLockOptimistic
penduduk.Source = "select * from tb_penduduk"
penduduk.Open
End Sub
Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = penduduk
With DataGrid1
End With
Call edit_grid
End Sub
Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = penduduk
With DataGrid1
End With
End Sub
Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "KODE"
    .Columns(1).Caption = "NIKK"
    .Columns(2).Caption = "NAMA KEPALA KELUARGA"
    .Columns(3).Caption = "ALAMAT"
    .Columns(4).Caption = "JENIS KELAMIN"
    .Columns(5).Caption = "RT"
    .Columns(6).Caption = "RW"
    .Columns(0).Width = 1200
    .Columns(1).Width = 2000
    .Columns(2).Width = 2000
    .Columns(3).Width = 2000
    .Columns(4).Width = 1200
    .Columns(5).Width = 1200
    .Columns(6).Width = 1200
End With
End Sub
Sub kosong()
txtkdalternatif.Text = ""
txtnikk.Text = ""
txtnamakk.Text = ""
txtalamat.Text = ""
comjk.Text = "Jenis Kelamin"
txtrt.Text = ""
txtrw.Text = ""
End Sub

