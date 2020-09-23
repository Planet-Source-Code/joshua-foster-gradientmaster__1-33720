VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GradientMaster"
   ClientHeight    =   5280
   ClientLeft      =   4830
   ClientTop       =   3945
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4080
      TabIndex        =   12
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bottom Color"
      Height          =   1575
      Left            =   240
      TabIndex        =   22
      Top             =   3480
      Width           =   1575
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "0"
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Red"
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Green"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Blue"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Middle Color"
      Height          =   1575
      Left            =   240
      TabIndex        =   18
      Top             =   1800
      Width           =   1575
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "0"
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Red"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Green"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Blue"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Top Color"
      Height          =   1575
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   1575
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "0"
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Red"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Green"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Blue"
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Image"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   2400
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5280
      Left            =   2040
      ScaleHeight     =   500
      ScaleMode       =   0  'User
      ScaleWidth      =   1
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "by Joshua Foster"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   27
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "GradientMaster"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Enter a value between 0 and 100 for each color."
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   1440
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

  Dim incR1, incG1, incB1, _
      incR2, incG2, incB2, _
      red, green, blue As Double
  
  If Text1.Text > 100 Or Text2.Text > 100 Or Text3.Text > 100 _
  Or Text4.Text > 100 Or Text5.Text > 100 Or Text6.Text > 100 _
  Or Text7.Text > 100 Or Text8.Text > 100 Or Text9.Text > 100 _
  Then
    MsgBox ("All color values must be between 0 and 100!")
    Exit Sub
  End If
  
  red = Text1.Text
  green = Text2.Text
  blue = Text3.Text
  
  incR1 = (Text4.Text - red) / 250
  incG1 = (Text5.Text - green) / 250
  incB1 = (Text6.Text - blue) / 250
  incR2 = (Text7.Text - Text4.Text) / 250
  incG2 = (Text8.Text - Text5.Text) / 250
  incB2 = (Text9.Text - Text6.Text) / 250
  
  For i = 0 To 500
    Picture1.Line (0, i)-(1, i), RGB(red * 2.55, green * 2.55, blue * 2.55)
    If i < 250 Then
      red = red + incR1
      green = green + incG1
      blue = blue + incB1
    Else
      red = red + incR2
      green = green + incG2
      blue = blue + incB2
    End If
  Next i
  
End Sub

Private Sub Command2_Click()
  
  fname = InputBox("Type the path and filename for" & _
                   " your image." & vbCrLf & vbCrLf & _
                   "( .bmp extension not necessary )")
  
  If fname = "" Then Exit Sub
  If Not Right(fname, 4) = ".bmp" Then fname = fname & ".bmp"
  If Not Dir(fname) = "" Then
    answer = MsgBox("File already exists.  Overwrite?", _
                     vbYesNo + vbExclamation)
    If answer = vbNo Then Call Command2_Click
  End If
  
  Picture1.Picture = Picture1.Image
  SavePicture Picture1.Picture, fname
  
End Sub

Private Sub Command3_Click()
  End
End Sub

Private Sub Text1_Click()
  Text1.Text = ""
End Sub

Private Sub Text1_LostFocus()
  If Text1.Text = "" Then Text1.Text = "0"
End Sub

Private Sub Text2_LostFocus()
  If Text2.Text = "" Then Text2.Text = "0"
End Sub

Private Sub Text3_LostFocus()
  If Text3.Text = "" Then Text3.Text = "0"
End Sub

Private Sub Text4_LostFocus()
  If Text4.Text = "" Then Text4.Text = "0"
End Sub

Private Sub Text5_LostFocus()
  If Text5.Text = "" Then Text5.Text = "0"
End Sub

Private Sub Text6_LostFocus()
  If Text6.Text = "" Then Text6.Text = "0"
End Sub

Private Sub Text7_LostFocus()
  If Text7.Text = "" Then Text7.Text = "0"
End Sub

Private Sub Text8_LostFocus()
  If Text8.Text = "" Then Text8.Text = "0"
End Sub

Private Sub Text9_LostFocus()
  If Text9.Text = "" Then Text9.Text = "0"
End Sub

Private Sub Text2_Click()
  Text2.Text = ""
End Sub

Private Sub Text3_Click()
  Text3.Text = ""
End Sub

Private Sub Text4_Click()
  Text4.Text = ""
End Sub

Private Sub Text5_Click()
  Text5.Text = ""
End Sub

Private Sub Text6_Click()
  Text6.Text = ""
End Sub

Private Sub Text7_Click()
  Text7.Text = ""
End Sub

Private Sub Text8_Click()
  Text8.Text = ""
End Sub

Private Sub Text9_Click()
  Text9.Text = ""
End Sub

