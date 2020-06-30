VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Formkriteria 
   Caption         =   "KRITERIA"
   ClientHeight    =   6735
   ClientLeft      =   5520
   ClientTop       =   1935
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   8895
   Begin VB.Frame Frame2 
      BackColor       =   &H80000011&
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8775
      Begin VB.CommandButton Commhapus 
         Caption         =   "HAPUS"
         Height          =   495
         Left            =   3600
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2415
         Left            =   360
         TabIndex        =   14
         Top             =   2880
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4260
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
      Begin VB.CommandButton commbatal 
         Caption         =   "BATAL"
         Height          =   495
         Left            =   6360
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton commubah 
         Caption         =   "UBAH"
         Height          =   495
         Left            =   3720
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton commsimpan 
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   1080
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtbobot 
         Height          =   375
         Left            =   6960
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtjenis 
         Height          =   375
         Left            =   5040
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtnmkriteria 
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtkdkriteria 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         Caption         =   "BOBOT"
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
         Left            =   6960
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         Caption         =   "JENIS"
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
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         Caption         =   "NAMA KRITERIA"
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
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         Caption         =   "KODE"
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
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000011&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         Caption         =   "FORM KRITERIA"
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
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "Formkriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KRITERIA As New ADODB.Recordset
Sub tampil_data()
Set KRITERIA = New ADODB.Recordset
KRITERIA.ActiveConnection = koneksidb
KRITERIA.CursorLocation = adUseClient
KRITERIA.LockType = adLockOptimistic
KRITERIA.Source = "select * from tb_kriteria"
KRITERIA.Open
End Sub
Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = KRITERIA
With DataGrid1
End With
Call edit_grid
End Sub
Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = KRITERIA
With DataGrid1
End With
End Sub
Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "KODE KRITERIA"
    .Columns(1).Caption = "NAMA KRITERIA"
    .Columns(2).Caption = "JENIS"
    .Columns(3).Caption = "BOBOT"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 1200
    .Columns(3).Width = 1200
End With
End Sub
Sub kosong()
txtkdkriteria = ""
txtnmkriteria = ""
txtjenis = ""
txtbobot = ""
txtkdkriteria.SetFocus
End Sub

Private Sub commbatal_Click()
Call kosong
End Sub

Private Sub Commhapus_Click()
koneksidb.Execute "delete from tb_kriteria where kd_kriteria='" & txtkdkriteria & "'"
Call refreshh
Call kosong
txtkdkriteria.SetFocus
End Sub

Private Sub Commsimpan_Click()
If txtkdkriteria = "" Then
MsgBox "Kode Kosong", vbExclamation, "pesan"
txtkdkriteria.SetFocus
Exit Sub
End If
    If txtnmkriteria = "" Then
    MsgBox "Nama Kriteria Kosong", vbExclamation, "pesan"
    txtnmkriteria.SetFocus
    Exit Sub
    End If
If txtjenis = "" Then
MsgBox "Jenis Kriteria Kosong", vbExclamation, "pesan"
txtjenis.SetFocus
Exit Sub
End If
    If txtbobot = "" Then
    MsgBox "Bobot Kosong", vbExclamation, "pesan"
    txtbobot.SetFocus
    Exit Sub
    End If
Set KRITERIA = New ADODB.Recordset
KRITERIA.Open "select*from tb_kriteria where kd_kriteria='" & txtkdkriteria & "'", koneksidb
If Not KRITERIA.EOF Then
MsgBox "Kode Sudah Ada", vbCritical, "pesan"
txtkdkriteria = ""
txtkdkriteria.SetFocus
Exit Sub
Else
koneksidb.Execute "insert into tb_kriteria(kd_kriteria,nm_kriteria,jenis_kriteria,bobot) value ('" & txtkdkriteria & "','" & txtnmkriteria & "','" & txtjenis & "','" & txtbobot & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = KRITERIA
With DataGrid1
End With
Call edit_grid
End If
End Sub

Private Sub commubah_Click()
koneksidb.Execute "update tb_kriteria set nm_kriteria='" & txtnmkriteria & "',jenis_kriteria='" & txtjenis & "',bobot='" & txtbobot & "' where kd_kriteria='" & txtkdkriteria & "'"
Call update
Call edit_grid
Call kosong
End Sub

Private Sub DataGrid1_Click()
txtkdkriteria.Text = KRITERIA!kd_kriteria
txtnmkriteria.Text = KRITERIA!nm_kriteria
txtjenis.Text = KRITERIA!jenis_kriteria
txtbobot.Text = KRITERIA!bobot
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = KRITERIA
With KRITERIA
End With
Call edit_grid
End Sub
