VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Formkerja 
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   1935
   ClientWidth     =   19680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   19680
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   7560
      TabIndex        =   24
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   3625
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
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   19
      Top             =   5880
      Width           =   7455
      Begin VB.CommandButton Commbatal 
         Caption         =   "BATAL"
         Height          =   495
         Left            =   5760
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Commhapus 
         Caption         =   "HAPUS"
         Height          =   495
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Commubah 
         Caption         =   "UBAH"
         Height          =   495
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Commsimpan 
         BackColor       =   &H80000012&
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   7455
      Begin VB.ComboBox Combdinding 
         Height          =   315
         Left            =   2640
         TabIndex        =   18
         Top             =   4560
         Width           =   2655
      End
      Begin VB.ComboBox Comblantai 
         Height          =   315
         Left            =   2640
         TabIndex        =   17
         Top             =   3840
         Width           =   2655
      End
      Begin VB.ComboBox Combmilik 
         Height          =   315
         Left            =   2640
         TabIndex        =   16
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox txtjumlah 
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox Combpenghasilan 
         Height          =   315
         Left            =   2640
         TabIndex        =   14
         Top             =   2040
         Width           =   2655
      End
      Begin VB.ComboBox Combkerja 
         Height          =   315
         Left            =   2640
         TabIndex        =   13
         Top             =   1440
         Width           =   2655
      End
      Begin VB.ComboBox Combkd 
         Height          =   315
         Left            =   2640
         TabIndex        =   12
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtidkerja 
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "DIDING"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "LANTAI"
         BeginProperty Font 
            Name            =   "MS Serif"
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
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "KEPEMILIKAN"
         BeginProperty Font 
            Name            =   "MS Serif"
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
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "JUMLAH TANGGUNGAN"
         BeginProperty Font 
            Name            =   "MS Serif"
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
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "PENGHASILAN"
         BeginProperty Font 
            Name            =   "MS Serif"
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
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "PEKERJAAN"
         BeginProperty Font 
            Name            =   "MS Serif"
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
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "KODE_PENDUDUK"
         BeginProperty Font 
            Name            =   "MS Serif"
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
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "ID PEKERJAAN"
         BeginProperty Font 
            Name            =   "MS Serif"
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
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "INPUT DATA PENDUDUK 2"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         Width           =   5055
      End
   End
End
Attribute VB_Name = "Formkerja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kerja As New ADODB.Recordset
Dim penduduk As New ADODB.Recordset
Sub tampil_data()
Set kerja = New ADODB.Recordset
kerja.ActiveConnection = koneksidb
kerja.CursorLocation = adUseClient
kerja.LockType = adLockOptimistic
kerja.Source = "select * from tb_pekerjaan"
kerja.Open
End Sub
Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = kerja
With DataGrid1
End With
Call edit_grid
End Sub
Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = kerja
With DataGrid1
End With
End Sub
Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "KODE KERJA"
    .Columns(1).Caption = "KODE PENDUDUK"
    .Columns(2).Caption = "C1"
    .Columns(3).Caption = "C2"
    .Columns(4).Caption = "C3"
    .Columns(5).Caption = "C4"
    .Columns(6).Caption = "C5"
    .Columns(7).Caption = "C6"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 1200
    .Columns(3).Width = 1200
    .Columns(4).Width = 1200
    .Columns(5).Width = 1200
    .Columns(6).Width = 1200
    .Columns(7).Width = 1200
End With
End Sub
Sub kosong()
txtidkerja = ""
Combkd = ""
Combkerja = ""
Combpenghasilan = ""
txtjumlah = ""
Combmilik = ""
Comblantai = ""
Combdinding = ""
txtidkerja.SetFocus
End Sub

Private Sub Commsimpan_Click()
koneksidb.Execute "insert into tb_pekerjaan(id_pekerjaan,kd_alternatif,pekerjaan,penghasilan,jm_tanggungan,kepemilikan,lantai,dinding) value ('" & txtidkerja & "','" & Combkd & "','" & Combkerja & "','" & Combpenghasilan & "','" & txtjumlah & "','" & Combmilik & "','" & Comblantai & "','" & Combdinding & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = kerja
With DataGrid1
End With
Call edit_grid
End Sub

Private Sub DataGrid1_Click()
txtidkerja.Text = kerja!id_pekerjaan
Combkd.Text = kerja!kd_alternatif
Combkerja.Text = kerja!pekerjaan
Combpenghasilan.Text = kerja!penghasilan
txtjumlah.Text = kerja!jm_tanggungan
Combmilik.Text = kerja!kepemilikan
Comblantai.Text = kerja!lantai
Combdinding.Text = kerja!dinding
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = kerja
With kerja
End With
With DataGrid1
End With
Call edit_grid
If penduduk.State = adStateOpen Then penduduk.Close
penduduk.Open "select kd_alternatif from tb_penduduk", koneksidb
Do While Not penduduk.EOF
Combkd.AddItem penduduk!kd_alternatif
penduduk.MoveNext
Loop
With Combkerja
    .AddItem "PNS"
    .AddItem "Wiraswasta"
    .AddItem "Petani"
    .AddItem "Buruh"
    .AddItem "Petani"
    .AddItem "Penganguran"
End With
With Combpenghasilan
    .AddItem "< 500.000"
    .AddItem "500.000-1.000.000"
    .AddItem "1.000.000-2.500.000"
    .AddItem "2.500.000-5.000.000"
    .AddItem "> 5.000.000"
End With
With Combmilik
    .AddItem "Milik Sendiri"
    .AddItem "Sewa"
    .AddItem "Menumpang"
    .AddItem "Tidak Punya Rumah"
End With
With Comblantai
    .AddItem "Keramik"
    .AddItem "Semen"
    .AddItem "Kayu"
    .AddItem "Tanah"
End With
With Combdinding
    .AddItem "Tembok"
    .AddItem "Kayu"
    .AddItem "Bambu"
End With
End Sub
