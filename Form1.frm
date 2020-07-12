VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menggabungkan Isi 2 Combobox yang Mirip"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Text            =   "Combo3"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   960
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strData As String  'Ini untuk menampung seluruh
                       'data
Private Sub Command1_Click()
  strData = "" 'Mula-mula masih kosong, selalu!
  'Ulangi sebanyak jumlah data di Combo1
  For i = 0 To Combo1.ListCount - 1
    'Tampung ke dalam variabel string, pisahkan dgn
    'koma
    strData = strData & Combo1.List(i) & ","
  Next i
  'Tampilkan data yang sudah digabung dalam satu string
  MsgBox strData, vbInformation, "Data di Combo1"
  'Berikut ini untuk memeriksa/membandingkan antara
  'data yang sudah ditampung di variabel string dengan
  'data yang ada di Combo2 (yang akan digabung)
  For i = 0 To Combo2.ListCount - 1
    'Jika data/item di Combo2 tidak terdapat di dalam
    'variabel string tadi, tambahkan di bagian akhir
    'dari variabel string (= join)
    If InStr(1, strData, Combo2.List(i)) < 1 Then
      'Tampilkan data yang tidak ada di variabel string
      MsgBox Combo2.List(i), vbInformation, _
             "Data di Combo2 yang tidak ada di Combo1"
      'Tambahkan di bagian akhir dari variabel string
      'dan dalam kasus ini, pisahkan dengan karakter
      'koma
      strData = strData & Combo2.List(i) & ","
    End If
  Next i
  'Berikut ini untuk mengambil data yang sudah digabung
  'seluruhnya (ingat, menggabungkan di sini artinya
  'sama dengan join; yaitu menambahkan data yang belum
  'ada, serta mengabaikan data yang sudah ada (sama))
  'dan membuang tanda koma di ujung paling kanan-->
  'untuk memudahkan dalam pemisahan data di Combo3)
  If Right(strData, 1) = "," Then
    strData = Left(strData, Len(strData) - 1)
  End If
  'Berikut ini untuk menampilkan data seluruhnya yang
  'sudah berhasil digabung ke dalam variabel string
  MsgBox strData, vbInformation, _
         "Data Hasil Gabung Combo1 dan Combo2"
End Sub

'Prosedur berikut untuk memisahkan data yang ada di 'dalam variabel string hasil penggabungan ke dalam 'Combo3. Agar hasilnya urut di Combo3, jangan lupa set 'property
'Sort milik Combo3 menjadi True saat "design-time"
'(Karena property Sort bersifat Read-Only, maka dia 'hanya dapat diset True saat "design-time". Jika Anda 'mengeset saat "run-time", maka akan terjadi error run-'time).
'(lihat pada Form_Load bagian bawah)
Private Sub Command2_Click()
  Dim i As Integer
  Dim arrData() As String
  arrData = Split(strData, ",")
  'Ulangi mulai batas bawah array sampai ke batas
  'atas array (untuk menampilkan data hasil
  'penggabungan).
  For i = LBound(arrData) To UBound(arrData)
     MsgBox arrData(i), vbInformation, _
          "Data Hasil Penggabungan di Combo3"
     Combo3.AddItem arrData(i)
  Next
MsgBox "Klik Combo3 u/ melihat hasil secara urut!", _
       vbInformation, "Hasil Gabung ada di Combo3"
End Sub

Private Sub Form_Load()
  'Berikut ini data yang ada di Combo1
  Combo1.Text = ""
  Combo1.AddItem "1"
  Combo1.AddItem "2"
  Combo1.AddItem "3"
  Combo1.AddItem "4"
  Combo1.AddItem "7"
  Combo1.AddItem "8"
  Combo1.Text = Combo1.List(0) 'Sorot data teratas
  
  'Berikut ini data yang ada di Combo2
  Combo2.Text = ""
  Combo2.AddItem "1"
  Combo2.AddItem "3"
  Combo2.AddItem "5"
  Combo2.AddItem "6"
  Combo2.AddItem "7"
  Combo2.AddItem "8"
  Combo2.Text = Combo2.List(0) 'Sorot data teratas
  
  'Sedangkan Combo3 mula-mula masih kosong,
  'dan akan dijadikan tempat untuk menggabung data.
  Combo3.Text = "" 'Tempat hasil penggabungan (Join)
  'Perintah di bawah akan menyebabkan error-run-time
  '(Can't assign to read-only property)
  'Combo3.Sorted = True '<-- ditutup, hanya bisa saat
                        '    design-time saja!
End Sub


