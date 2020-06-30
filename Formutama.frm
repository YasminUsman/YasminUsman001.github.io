VERSION 5.00
Begin VB.Form Formutama 
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   1800
   ClientTop       =   1500
   ClientWidth     =   15900
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   15900
   Begin VB.Image Image1 
      Height          =   8355
      Left            =   0
      Picture         =   "Formutama.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   15930
   End
   Begin VB.Menu MENU 
      Caption         =   "MENU"
      Begin VB.Menu DATA_PENDUDUK 
         Caption         =   "DATA PENDUDUK"
      End
      Begin VB.Menu KRITERIA 
         Caption         =   "KRITERIA"
      End
      Begin VB.Menu RANGE_KRITERIA 
         Caption         =   "RANGE KRITERIA"
      End
   End
   Begin VB.Menu PERHITUNGAN 
      Caption         =   "PERHITUNGAN"
      Begin VB.Menu PERHITUNGAN_DATA 
         Caption         =   "PERHITUNGAN DATA"
      End
   End
   Begin VB.Menu HASIL 
      Caption         =   "HASIL"
   End
End
Attribute VB_Name = "Formutama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DATA_PENDUDUK_Click()
Formpenduduk.Show
End Sub

Private Sub KRITERIA_Click()
Formkriteria.Show
End Sub

Private Sub PERHITUNGAN_DATA_Click()
Formperhitungan.Show
End Sub

Private Sub RANGE_KRITERIA_Click()
Formrange.Show
End Sub
