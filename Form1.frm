VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Quikwrite OCR v1.0"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Recognize After Mouse UP"
      Height          =   195
      Left            =   1680
      TabIndex        =   22
      Top             =   2400
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1920
      Width           =   855
   End
   Begin VB.PictureBox pict 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      MouseIcon       =   "Form1.frx":0442
      MousePointer    =   99  'Custom
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   15
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "CLS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Execute"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mark"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox picNum 
      BackColor       =   &H80000004&
      DrawWidth       =   2
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   16
      Top             =   720
      Width           =   855
   End
   Begin VB.PictureBox pic5 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      Picture         =   "Form1.frx":0594
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   14
      Top             =   4560
      Width           =   855
   End
   Begin VB.PictureBox pic4 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      Picture         =   "Form1.frx":30F6
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   13
      Top             =   4560
      Width           =   855
   End
   Begin VB.PictureBox pic3 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1920
      Picture         =   "Form1.frx":5C58
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   12
      Top             =   4560
      Width           =   855
   End
   Begin VB.PictureBox pic2 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      Picture         =   "Form1.frx":87BA
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   11
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Appearance      =   0  'Flat
      Caption         =   "Init Insert (Press me to do Execute!)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Picture         =   "Form1.frx":B31C
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   0
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":DCCE
      Height          =   1215
      Left            =   0
      TabIndex        =   21
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   " Writing Area"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "         1            2            3            4            5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ArrA(70, 70) As Byte ' That means the number 1 and so on down
Dim ArrB(70, 70) As Byte
Dim ArrC(70, 70) As Byte
Dim ArrD(70, 70) As Byte
Dim ArrE(70, 70) As Byte ' The number 5
Dim FLAG As Boolean
Dim Arr2(70, 70) As Byte ' The array that comes from the writing area
Private Const CharacterWidth As Long = 30
Private Const CharacterHeight As Long = 20

Private Sub command1_Click()
Dim Height As Long, Width As Long
 Height = pic.ScaleHeight / 70
 Width = pic.ScaleWidth / 70
For y = 0 To pic.ScaleHeight Step Height: For x = 0 To pic.ScaleWidth Step Width
  If pic.Point(x, y) = vbWhite Then
   ArrA(x, y) = 0
  Else
   ArrA(x, y) = 1
  End If
  If pic2.Point(x, y) = vbWhite Then
   ArrB(x, y) = 0
  Else
   ArrB(x, y) = 1
  End If
  If pic3.Point(x, y) = vbWhite Then
   ArrC(x, y) = 0
  Else
   ArrC(x, y) = 1
  End If
  If pic4.Point(x, y) = vbWhite Then
   ArrD(x, y) = 0
  Else
   ArrD(x, y) = 1
  End If
  If pic5.Point(x, y) = vbWhite Then
   ArrE(x, y) = 0
  Else
   ArrE(x, y) = 1
  End If
Next x: Next y
Debug.Print "Done"
Command5.Enabled = True
pict.Enabled = True
End Sub
Private Sub Command2_Click()
Dim MaxExtents As DWORD
Dim RealTextExtent As RECT
  MaxExtents.low = pict.ScaleHeight
  MaxExtents.high = pict.ScaleWidth
  RealTextExtent = GetTrueExtents(MaxExtents)
  If RealTextExtent.Left = -1 Or RealTextExtent.Right = -1 Or RealTextExtent.Bottom = -1 Or RealTextExtent.Top = -1 Then
    Exit Sub
  End If
  pict.Line (RealTextExtent.Left, RealTextExtent.Top)-(RealTextExtent.Right, RealTextExtent.Top), vbGreen
  pict.Line (RealTextExtent.Left, RealTextExtent.Top)-(RealTextExtent.Left, RealTextExtent.Bottom), vbYellow
  pict.Line (RealTextExtent.Right, RealTextExtent.Top)-(RealTextExtent.Right, RealTextExtent.Bottom), vbMagenta
  pict.Line (RealTextExtent.Left, RealTextExtent.Bottom)-(RealTextExtent.Right, RealTextExtent.Bottom), vbCyan
StretchBlt picNum.hdc, 0, 0, picNum.ScaleWidth, picNum.ScaleHeight, pict.hdc, RealTextExtent.Left + 1, RealTextExtent.Top + 1, (RealTextExtent.Right - RealTextExtent.Left) - 1, (RealTextExtent.Bottom - RealTextExtent.Top) - 1, vbSrcCopy
End Sub

Private Sub Command4_Click()
Text7.Text = ""
End Sub
Private Sub Command5_Click()
Dim Counter As Integer
Debug.Print ""
Counter = 0
Counter2 = 0
Counter3 = 0
Counter4 = 0
Counter5 = 0
arrchk = 0
Dim Height As Long, Width As Long
 Height = picNum.ScaleHeight / 70
 Width = picNum.ScaleWidth / 70
For y = 0 To picNum.ScaleHeight Step Height: For x = 0 To picNum.ScaleWidth Step Width
  If picNum.Point(x, y) = vbWhite Then
   Arr2(x, y) = 0
  Else
   Arr2(x, y) = 1
  End If
Next x: Next y
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For y = 0 To picNum.ScaleHeight Step Height: For x = 0 To picNum.ScaleWidth Step Width
 arrchk = arrchk + Arr2(x, y)
Next x: Next y
If arrchk = 123 Then MsgBox "Error!, Writing Area is empty!": Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Debug.Print "Insert OK"
For x = 0 To 70: For y = 0 To 70
 If ArrA(x, y) = Arr2(x, y) Then
  Counter = Counter + 1
 End If
 If ArrB(x, y) = Arr2(x, y) Then
  Counter2 = Counter2 + 1
 End If
 If ArrC(x, y) = Arr2(x, y) Then
  Counter3 = Counter3 + 1
 End If
 If ArrD(x, y) = Arr2(x, y) Then
  Counter4 = Counter4 + 1
 End If
 If ArrE(x, y) = Arr2(x, y) Then
  Counter5 = Counter5 + 1
 End If
Next y: Next x
text1(0).Text = (Counter / 5041) * 100
text1(1).Text = (Counter2 / 5041) * 100
text1(2).Text = (Counter3 / 5041) * 100
text1(3).Text = (Counter4 / 5041) * 100
text1(4).Text = (Counter5 / 5041) * 100
Dim Biggest As Long, Indx As Integer
Biggest = 0
For i = 0 To 4
 If text1(i) > Biggest Then
  Biggest = text1(i).Text
  Indx = i
 End If
Next i
text1(Indx).BackColor = vbBlack
text1(Indx).ForeColor = vbRed
Text7.Text = Text7.Text & Indx + 1
Text2.Text = Biggest
End Sub
Private Sub Command7_Click()
pict.Cls
Text2.Text = ""
For i = 0 To 4
 text1(i).BackColor = vbWhite
 text1(i).ForeColor = vbBlack
 text1(i).Text = "0"
Next i
picNum.Cls
End Sub
Private Sub Command8_Click()
pict.PaintPicture Picture1.Picture, 10, 0, , , , , 40
End Sub
Private Sub Form_Load()
Debug.Print ""
End Sub
Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
End Sub
Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
End Sub
Private Sub List3_Click()
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
Dim Height As Long, Width As Long
Height = Picture2.ScaleHeight / 70
Width = Picture2.ScaleWidth / 70
End Sub
Private Sub pict_KeyPress(KeyAscii As Integer)
pict.Cls
pict.Print Chr(KeyAscii)
End Sub

Private Sub Pict_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FLAG = True
End Sub

Private Sub Pict_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If FLAG = True Then
 pict.DrawWidth = 5
 pict.PSet (x, y)
 pict.DrawWidth = 1
End If
End Sub

Private Sub Pict_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
FLAG = False
If Check1.Value = 1 Then
 Call Command2_Click
 DoEvents
 Call Command5_Click
 DoEvents
 Call Command7_Click
End If
End Sub
