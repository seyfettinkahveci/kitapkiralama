VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form UyeEkle 
   BackColor       =   &H00800000&
   Caption         =   "Üye Ekle"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   ScaleHeight     =   7500
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame UyeKaydiFrame 
      BackColor       =   &H00800000&
      Caption         =   "Üye Ekle"
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
      Height          =   5175
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   7455
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
         Left            =   3480
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2400
         Width           =   2500
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
         Left            =   3480
         TabIndex        =   6
         Top             =   3000
         Width           =   2500
      End
      Begin VB.ComboBox Yetki 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3500
         TabIndex        =   9
         Text            =   "Kullanýcýnýn Yetkisini Seçiniz"
         Top             =   3960
         Width           =   2500
      End
      Begin VB.CommandButton KaydetButon 
         Caption         =   "Kaydet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   10
         Top             =   4440
         Width           =   2000
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
         Left            =   3500
         TabIndex        =   2
         Top             =   600
         Width           =   2500
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
         Left            =   3500
         TabIndex        =   3
         Top             =   1200
         Width           =   2500
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
         Left            =   3480
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1800
         Width           =   2500
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
         ForeColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3600
         TabIndex        =   7
         Top             =   3480
         Width           =   1095
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
         ForeColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   4800
         TabIndex        =   8
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton KaydetIptal 
         Caption         =   "Anasayfaya Git"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   0
         Top             =   4440
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2000
         TabIndex        =   18
         Top             =   2400
         Width           =   1500
      End
      Begin VB.Label Adres 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2000
         TabIndex        =   17
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800000&
         Caption         =   "Yetki"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1995
         TabIndex        =   16
         Top             =   3960
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2000
         TabIndex        =   15
         Top             =   600
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
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2000
         TabIndex        =   14
         Top             =   1200
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2000
         TabIndex        =   13
         Top             =   1800
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2000
         TabIndex        =   12
         Top             =   3480
         Width           =   1500
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   7560
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   $"UyeEkle.frx":0000
      OLEDBString     =   $"UyeEkle.frx":00DC
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select AdSoyad,KAdi,Cinsiyet,Adres  from Kullanicilar"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "UyeEkle.frx":01B8
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   3625
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Kayýtlý Üyeler"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "AdSoyad"
         Caption         =   "AdSoyad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "KAdi"
         Caption         =   "KAdi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Cinsiyet"
         Caption         =   "Cinsiyet"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Adres"
         Caption         =   "Adres"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   555,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "UyeEkle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Adodc1.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End Sub

Private Sub Form_Load()
Yetki.AddItem "Kullanýcý"
Yetki.AddItem "Yönetici"
End Sub

Private Sub KaydetButon_Click()
If AdSoyadText.Text = "" Then
dugme = MsgBox("AdSoyad Boþ Olamaz", 64, "Uyari")
ElseIf KullaniciAdiText.Text = "" Then
dugme = MsgBox("Kullanýcý Adý Boþ Olamaz", 64, "Uyari")
ElseIf SifreTekrarText.Text = "" Or SifreText.Text = "" Then
dugme = MsgBox("Þifreler Boþ Olamaz", 64, "Uyari")
ElseIf SifreTekrarText.Text <> SifreText.Text Then
dugme = MsgBox("Þifreler Eþleþmiyor", 64, "Uyari")
ElseIf AdresText.Text = "" Then
dugme = MsgBox("Adres Boþ Olamaz", 64, "Uyari")
ElseIf Option2.Value = False And Option1.Value = False Then
dugme = MsgBox("Cinsiyet Seçiniz", 64, "Uyari")
ElseIf Yetki.Text = "" Or Yetki.Text = "Kullanýcýnýn Yetkisini Seçiniz" Then
dugme = MsgBox("Lütfen Üyenin Yetkisini Seçiniz ", 64, "Uyari")
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
Rs!Yetki = Yetki.ListIndex + 1
Rs.Update
Db.Close
Adodc1.Refresh
DataGrid1.Refresh
End If
End Sub



Private Sub KaydetIptal_Click()
Yonetici.Show
Unload UyeEkle
End Sub

