VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Menu"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2520
      Picture         =   "Menu.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Pic3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1320
      Picture         =   "Menu.frx":0682
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      Picture         =   "Menu.frx":0EBB
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      Picture         =   "Menu.frx":12D8
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5280
      Top             =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on this Picture"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Double
Dim j As Double
Dim k As Double
Dim l As Double
Dim m As Double
Dim a As Double

Private Sub pic1_Click()
i = 0
k = 0
l = 0
m = 0
Timer2.Enabled = True
pic2.Visible = True
Pic3.Visible = True
Pic4.Visible = True
Label1.Move 4000, 1600
Label1.Caption = "Now click one of these three."
End Sub

Private Sub pic2_Click()
    MsgBox "You have choosed Diamond Pendent."
End Sub

Private Sub Pic3_Click()
MsgBox "You have choosed Gem Ring."
End Sub

Private Sub Pic4_Click()
MsgBox "You have choosed Diamond Neckless."
End Sub

Private Sub Timer2_Timer()
i = i + 50
If i < 4000 Then
    If m < 1700 Then
        j = Math.Sqr((a * a) - (i - 2000) * (i - 2000)) + 1700 + m
        m = m + 100
    Else
        j = Math.Sqr((a * a) - (i - 2000) * (i - 2000)) + 1700
    End If
    
    pic2.Move i, j
    
    If k < 1215 Then
        Pic3.Move i, j + k
        k = k + 100
    Else
        Pic3.Move i, j + 1215
    End If
    If l < 1215 Then
        Pic4.Move i, j + l
        l = l + 100
    Else
        Pic4.Move i, j + 2430
    End If
End If
End Sub

Private Sub Form_Load()
Me.Move 0, 0
pic2.Move 120, 1680
Pic4.Move 120, 1680
Pic3.Move 120, 1680
a = 2000
End Sub
