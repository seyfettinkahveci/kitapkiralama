VERSION 5.00
Begin VB.Form UyeKayit 
   BackColor       =   &H00800000&
   Caption         =   "Üye Kaydý"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form2"
   ScaleHeight     =   5160
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame UyeKaydiFrame 
      BackColor       =   &H00800000&
      Caption         =   "Üye Bilgileri"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7215
      Begin VB.TextBox SifreTekrarText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         IMEMode         =   3  'DISABLE
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2400
         Width           =   3000
      End
      Begin VB.TextBox AdresText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3000
         TabIndex        =   5
         Top             =   3000
         Width           =   3000
      End
      Begin VB.CommandButton UyaOlIptal 
         Caption         =   "Ýptal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3960
         TabIndex        =   0
         Top             =   4080
         Width           =   2000
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00800000&
         Caption         =   "Erkek"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   400
         Left            =   4320
         TabIndex        =   8
         Top             =   3600
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Kýz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   400
         Left            =   3000
         TabIndex        =   7
         Top             =   3600
         Width           =   1500
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
         Height          =   400
         IMEMode         =   3  'DISABLE
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1800
         Width           =   3000
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
         Height          =   400
         Left            =   3000
         TabIndex        =   2
         Top             =   1200
         Width           =   3000
      End
      Begin VB.TextBox AdSoyadText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3000
         TabIndex        =   1
         Top             =   600
         Width           =   3000
      End
      Begin VB.CommandButton UyeOlButon 
         Caption         =   "Üye Ol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1440
         TabIndex        =   9
         Top             =   4080
         Width           =   2000
      End
      Begin VB.Label Label6 
         BackColor       =   &H00800000&
         Caption         =   "Þifre Tekrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   405
         Left            =   1500
         TabIndex        =   15
         Top             =   2400
         Width           =   1500
      End
      Begin VB.Label label5 
         BackColor       =   &H00800000&
         Caption         =   "Adres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   405
         Left            =   1500
         TabIndex        =   14
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "Cinsiyet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   405
         Left            =   1440
         TabIndex        =   13
         Top             =   3600
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Þifre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   405
         Left            =   1560
         TabIndex        =   12
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Kullanýcý Adý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   405
         Left            =   1500
         TabIndex        =   11
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Ad Soyad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   400
         Left            =   1500
         TabIndex        =   10
         Top             =   600
         Width           =   1500
      End
   End
End
Attribute VB_Name = "UyeKayit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Db As Database
Dim Rs As Recordset

Private Sub UyaOlIptal_Click()
Form1.Show
UyeKayit.Hide

End Sub

Private Sub UyeOlButon_Click()
If AdSoyadText.Text = "" Then
dugme = MsgBox("Ad Soyad Boþ Olamaz", 64, "Uyari")
ElseIf KullaniciAdiText.Text = "" Then
dugme = MsgBox("Kullanýcý Adý Boþ Olamaz", 64, "Uyari")
ElseIf SifreText.Text <> SifreTekrarText.Text Then
dugme = MsgBox("Þifreler Eþleþmiyor", 64, "Uyari")
ElseIf SifreText.Text = "" Or SifreTekrarText.Text = "" Then
dugme = MsgBox("Þifreler Boþ Olamaz", 64, "Uyari")
ElseIf AdresText.Text = "" Then
dugme = MsgBox("Adres Boþ olamaz", 64, "Uyari")
ElseIf Option2.Value = False And Option1.Value = False Then
dugme = MsgBox("Cinsiyet Seçiniz", 64, "Uyari")
Else
Set Db = OpenDatabase("VeriTabani/VeriTabani.mdb")
Set Rs = Db.OpenRecordset("Kullanicilar")
Rs.AddNew
Rs!AdSoyad = AdSoyadText.Text
Rs!KAdi = KullaniciAdiText.Text
Rs!Sifre = SifreText.Text
If (Option2.Value = True) Then
Cinsiyet = "E"
End If
If (Option1.Value = True) Then
Cinsiyet = "K"
End If
Rs!Cinsiyet = Cinsiyet
Rs!Adres = AdresText.Text
Rs!Yetki = 1
Rs.Update
Db.Close
UyeKayit.Hide
Form1.Show
End If
End Sub
