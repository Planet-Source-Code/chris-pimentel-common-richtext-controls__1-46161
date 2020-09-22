VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1410
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2340
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   2340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   300
      Left            =   330
      TabIndex        =   2
      Top             =   510
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1395
      TabIndex        =   0
      Text            =   "10"
      Top             =   75
      Width           =   630
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Must be a valid font size  exaample  ""48 or 72"""
      Height          =   405
      Left            =   60
      TabIndex        =   4
      Top             =   930
      Width           =   2205
   End
   Begin VB.Line Line1 
      X1              =   15
      X2              =   2340
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label Label2 
      Caption         =   "10 - 72"
      Height          =   210
      Left            =   1410
      TabIndex        =   3
      Top             =   375
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Font Size ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   75
      Width           =   1245
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.RichTextBox1.SelFontSize = Form3.Text1.Text
Unload Me
End Sub

