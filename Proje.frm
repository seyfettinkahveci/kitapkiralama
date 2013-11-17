VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "Sistem Giriþi"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CikisButon 
      Caption         =   "Çýkýþ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   9240
      TabIndex        =   0
      Top             =   1920
      Width           =   2000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Üye Giriþi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   3615
      Left            =   3600
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton ParolaButon 
         Caption         =   "Parolamý Unuttum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         TabIndex        =   7
         Top             =   2040
         Width           =   2000
      End
      Begin VB.CommandButton GirisButon 
         Caption         =   "Giriþ Yap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Picture         =   "Proje.frx":0000
         TabIndex        =   6
         Top             =   2040
         Width           =   2000
      End
      Begin VB.TextBox SifreText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox KullaniciAdiText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Þube Giriþi Yapmak Ýsterseniz Kullanýcý Adý Sube Sifre Sube giriniz."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   5055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Sistem Sahibi Olarak Giriþ Yapmak Ýçin Kullanýcý Adý Admin Þifre Admin giriniz."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Þifre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Kullanýcý Adý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Üye Kaydý"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   9240
      TabIndex        =   8
      Top             =   120
      Width           =   2000
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   120
      Picture         =   "Proje.frx":74BE8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Db As Database
Dim Rs As Recordset
Dim KitapAdi As String
Dim Fiyat As String
Public Kullanici As String
Public Adres As String

Private Sub CikisButon_Click()
    End
End Sub

Private Sub Command1_Click()
Form1.Hide
UyeKayit.Show

End Sub



Private Sub GirisButon_Click()
If KullaniciAdiText.Text = "" Then
dugme = MsgBox("Kullanýcý Adý Boþ Olamaz", 64, "Uyari")
ElseIf SifreText.Text = "" Then
dugme = MsgBox("Þifre Boþ Olamaz", 64, "Uyari")
Else
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Kullanicilar")
Do Until Rs.EOF
KullaniciAdi = Rs("KAdi")
Sifre = Rs("Sifre")
Yetki = Rs("Yetki")
If KullaniciAdi = KullaniciAdiText.Text And Sifre = SifreText.Text And Yetki = 2 Then
Yonetici.Show
Unload Form1
End If
If KullaniciAdi = KullaniciAdiText.Text And Sifre = SifreText.Text And Yetki = 1 Then
Kullanici = Rs("AdSoyad")
Adres = Rs("Adres")
Kirala.Show
Unload Form1
End If
Rs.MoveNext
Loop
Db.Close
End If
End Sub

Private Sub ParolaButon_Click()
Unload Me
ParolamiAnimsa.Show
End Sub

Private Sub Picture1_Click()

End Sub
