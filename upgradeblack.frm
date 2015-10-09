VERSION 5.00
Begin VB.Form upgradeblack 
   Caption         =   "Upgrade Black"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form2"
   ScaleHeight     =   3435
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   480
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   2040
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   3600
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      Height          =   975
      Left            =   5160
      ScaleHeight     =   915
      ScaleWidth      =   1260
      TabIndex        =   0
      Top             =   1440
      Width           =   1320
   End
End
Attribute VB_Name = "upgradeblack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim uindex As Integer
Function upblack(index As Integer)
Picture4.Picture = Form1.ImageList1.ListImages(8).Picture
Picture4.Tag = 220
Picture3.Picture = Form1.ImageList1.ListImages(9).Picture
Picture3.Tag = 230
Picture2.Picture = Form1.ImageList1.ListImages(10).Picture
Picture2.Tag = 240
Picture1.Picture = Form1.ImageList1.ListImages(11).Picture
Picture1.Tag = 250
uindex = index
End Function

Private Sub Picture1_Click()
Form1.Picture1(uindex).Tag = 250
Form1.Picture1(uindex).Picture = Picture1.Picture
upgradewhite.Hide
Form1.Show
Form1.Enabled = True
End Sub

Private Sub Picture2_Click()
Form1.Picture1(uindex).Tag = 240
Form1.Picture1(uindex).Picture = Picture2.Picture
upgradewhite.Hide
Form1.Show
Form1.Enabled = True
End Sub

Private Sub Picture3_Click()
Form1.Picture1(uindex).Tag = 230
Form1.Picture1(uindex).Picture = Picture3.Picture
upgradewhite.Hide
Form1.Show
Form1.Enabled = True
End Sub

Private Sub Picture4_Click()
Form1.Picture1(uindex).Tag = 220
Form1.Picture1(uindex).Picture = Picture4.Picture
upgradewhite.Hide
Form1.Show
Form1.Enabled = True
End Sub

