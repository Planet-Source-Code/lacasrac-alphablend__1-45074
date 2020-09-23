VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alphablend"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Felvesz1 
      Caption         =   "Create maps"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.PictureBox DST 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   3240
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   3
      Top             =   0
      Width           =   3135
   End
   Begin VB.HScrollBar v 
      Height          =   255
      LargeChange     =   10
      Left            =   3360
      Max             =   255
      SmallChange     =   10
      TabIndex        =   2
      Top             =   4440
      Width           =   2775
   End
   Begin VB.PictureBox SRC 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   120
      Picture         =   "alpha1.frx":0000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   3120
      Width           =   3000
   End
   Begin VB.PictureBox SRC2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   120
      Picture         =   "alpha1.frx":7485
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Alphablend project, created by Laca in 2003
'Kozari Laszlo, Hungary


Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Dim rgbs() As Long
Dim rgbs2() As Long
Private Sub Felvesz(Shade)

'This method is just create the 2
'picture maps by their rgb values.

ReDim rgbs(SRC.ScaleWidth - 1, SRC.ScaleHeight - 1)
ReDim rgbs2(SRC.ScaleWidth - 1, SRC.ScaleHeight - 1)

For x = 0 To SRC.ScaleWidth - 1
 For y = 0 To SRC.ScaleHeight - 1
        C = GetPixel(SRC.hdc, x, y)
        C2 = GetPixel(SRC2.hdc, x, y)
                
        rgbs(x, y) = C
        rgbs2(x, y) = C2
 Next y
Next x

'This line creates the alphablending...
Call Made(Shade)

End Sub


Private Sub Felvesz1_Click()

DST.Cls
Call Felvesz(v.Value)

End Sub

Private Sub v_Change()

Call Made(v.Value)
DST.Refresh

End Sub



Private Sub Made(Shade)
'Alphablend method
'Not too fast, but with the setpixel api _
 and with pixeldrawing its nice...


Alpha = Shade / 255
Alpha2 = (255 - Shade) / 255

bit0 = 255
bit1 = bit0 * 256
bit2 = bit1 * 256

For x = 0 To SRC.ScaleWidth - 1
 For y = 0 To SRC.ScaleHeight - 1
      SRC1 = rgbs(x, y): DST1 = rgbs2(x, y)
      col = _
            (SRC1 And bit0) * Alpha + (DST1 And bit0) * Alpha2 Or _
            (SRC1 And bit1) * Alpha + (DST1 And bit1) * Alpha2 And bit1 Or _
            (SRC1 And bit2) * Alpha + (DST1 And bit2) * Alpha2 And bit2
      SetPixel DST.hdc, x, y, col
 Next y
Next x
    
End Sub

Private Sub v_Scroll()
v_Change
End Sub





