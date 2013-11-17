VERSION 5.00
Begin VB.Form Yonetici 
   BackColor       =   &H00800000&
   Caption         =   "Yönetici Formu"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CýkýsButon 
      Caption         =   "Çýkýþ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5400
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton KargolanacakKitaplarButonu 
      Caption         =   "Kargolanacak Kitaplar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CommandButton KitabiGeriALButonu 
      Caption         =   "Kitabý Geri Al"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton KitaplarButon 
      Caption         =   "Kitaplar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton UyeEkleButon 
      Caption         =   "Üye Bilgileri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton YazarlarButon 
      Caption         =   "Yazarlar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label KargoBilgi 
      BackColor       =   &H00800000&
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   3000
   End
   Begin VB.Label KazanilanParaBilgi 
      BackColor       =   &H00800000&
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "Yonetici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Db As Database
Dim Rs As Recordset

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Adodc1.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
    
End Sub

Private Sub CýkýsButon_Click()
 End
End Sub

Private Sub Form_Load()
'Adodc1.CommandType = adCmdText
'Adodc1.RecordSource = "select * from Kullanicilar "
'Adodc1.Refresh
'DataGrid1.Refresh

Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Kiralananlar")
Fiyat = 0
Toplam = 0
Do Until Rs.EOF
Fiyat = Fiyat + Rs("Fiyat")
If Rs("Aktif") = 3 Then
Toplam = Toplam + 1
End If
Rs.MoveNext
Loop
KazanilanParaBilgi.Caption = " Þuana Kadar Toplam " & Fiyat & " TL Kazandýnýz"
KargoBilgi.Caption = " Kargolanacak Toplam " & Toplam & " kitabýnýz var"
Db.Close


End Sub


Private Sub KitapEkleButon_Click()
Yonetici.Hide
KitapEkle.Show
End Sub


Private Sub KargolanacakKitaplarButonu_Click()
Unload Me
Kargolanacak.Show

End Sub

Private Sub KitabiGeriALButonu_Click()
Unload Yonetici
GeriGelecekKitaplar.Show
End Sub

Private Sub KitaplarButon_Click()
Kitaplar.Show
Unload Yonetici
End Sub

Private Sub UyeEkleButon_Click()
Yonetici.Hide
UyeEkle.Show
End Sub



Private Sub YazarlarButon_Click()
Yazarlar.Show
Unload Yonetici
End Sub
