VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3060
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   3060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FF00&
      Height          =   270
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFF00&
      Height          =   270
      Left            =   2175
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00808080&
      Height          =   270
      Left            =   1875
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00000000&
      Height          =   270
      Left            =   1575
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FFFF&
      Height          =   270
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Height          =   270
      Left            =   975
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   270
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   15
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This is just an example, you can have way more colors than these"
      Height          =   390
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   2880
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.RichTextBox1.SelColor = vbWhite
End Sub

Private Sub Command2_Click()
Form1.RichTextBox1.SelColor = vbRed
End Sub

Private Sub Command3_Click()
Form1.RichTextBox1.SelColor = vbBlue
End Sub

Private Sub Command4_Click()
Form1.RichTextBox1.SelColor = vbYellow
End Sub

Private Sub Command5_Click()
Form1.RichTextBox1.SelColor = vbBlack
End Sub

Private Sub Command6_Click()
Form1.RichTextBox1.SelColor = &H808080
End Sub

Private Sub Command7_Click()
Form1.RichTextBox1.SelColor = &HFFFF00
End Sub

Private Sub Command8_Click()
Form1.RichTextBox1.SelColor = vbGreen
End Sub
