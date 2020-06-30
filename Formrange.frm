VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Formrange 
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000A&
      Height          =   4335
      Left            =   5040
      TabIndex        =   15
      Top             =   960
      Width           =   8895
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3375
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5953
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Height          =   4335
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   4935
      Begin VB.CommandButton Command3 
         Caption         =   "BATAL"
         Height          =   495
         Left            =   3120
         TabIndex        =   14
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "UBAH"
         Height          =   495
         Left            =   1560
         TabIndex        =   13
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtnilai 
         Height          =   495
         Left            =   2160
         TabIndex        =   11
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtrange 
         Height          =   495
         Left            =   2160
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtnmkriteria 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtkdkriteria 
         Height          =   495
         Left            =   2160
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtkdrange 
         Height          =   495
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000A&
         Caption         =   "NILAI RANGE"
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
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000A&
         Caption         =   "RANGE KRITERIA"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Caption         =   "KODE KRITERIA"
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
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000A&
         Caption         =   "KODE RANGE"
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
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "FORM RANGE KRITERIA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   6135
      End
   End
End
Attribute VB_Name = "Formrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim range As New ADODB.Recordset

Private Sub Command1_Click()
koneksidb.Execute "insert into tb_rangekriteria(kd_range,kd_kriteria,pekerjaan,range_kriteria,nilairange) value ('" & txtkdrange & "','" & txtkdkriteria & "','" & txtnmkriteria & "','" & txtrange & "','" & txtnilai & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = range
With DataGrid1
End With
Call edit_grid
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = range
With range
End With
Call edit_grid
End Sub
Sub tampil_data()
Set range = New ADODB.Recordset
range.ActiveConnection = koneksidb
range.CursorLocation = adUseClient
range.LockType = adLockOptimistic
range.Source = "select * from tb_rangekriteria"
range.Open
End Sub
Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = range
With DataGrid1
End With
Call edit_grid
End Sub
Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = range
With DataGrid1
End With
End Sub
Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "KODE"
    .Columns(1).Caption = "KODE KRITERIA"
    .Columns(2).Caption = "NAMA KRITERIA"
    .Columns(3).Caption = "RANGE KRITERIA"
    .Columns(4).Caption = "NILAI RANGE"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 2000
    .Columns(3).Width = 2000
    .Columns(4).Width = 1200
End With
End Sub
Sub kosong()
txtkdrange = ""
txtkdkriteria = ""
txtnmkriteria = ""
txtrange = ""
txtnilai = ""
txtkdrange.SetFocus
End Sub
