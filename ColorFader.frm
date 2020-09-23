VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Color Percentage Fader "
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   2640
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picGetColor 
      BackColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.PictureBox picGetColor 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picShowPercent 
      Height          =   495
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton cmdDoFill 
      Caption         =   "Do Fill"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblColours 
      Alignment       =   2  'Center
      Caption         =   "Click on colour to change it"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblColours 
      Alignment       =   1  'Right Justify
      Caption         =   "End Colour"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblColours 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Colour"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'developed after seeing
'ColorFade
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=50010&lngWId=1
' felt people might want greater control over what colours you can use
Private Sub cmdDoFill_Click()

  Dim X    As Long

  For X = 1 To picShowPercent.Width
    picShowPercent.Line (X, 0)-(X, picShowPercent.Height), RGBPercentage(picGetColor(0).BackColor, picGetColor(1).BackColor, Percent(X, picShowPercent.Width))
  Next X

End Sub

Private Sub picGetColor_Click(Index As Integer)

  With cdlColor
    .ShowColor
    Select Case Index
     Case 0
      picGetColor(0).BackColor = .Color
     Case 1
      picGetColor(1).BackColor = .Color
    End Select
  End With

End Sub

':)Roja's VB Code Fixer V1.1.59 (22/11/2003 1:05:02 PM) 1 + 27 = 28 Lines Thanks Ulli for inspiration and lots of code.
