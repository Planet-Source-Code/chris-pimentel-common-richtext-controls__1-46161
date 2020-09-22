VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9795
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "PRINT"
      Height          =   360
      Left            =   2895
      TabIndex        =   3
      Top             =   0
      Width           =   1290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Font size"
      Height          =   360
      Left            =   1395
      TabIndex        =   2
      Top             =   0
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Font color"
      Height          =   360
      Left            =   225
      TabIndex        =   1
      Top             =   0
      Width           =   1005
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6615
      Left            =   60
      TabIndex        =   0
      Top             =   435
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   11668
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":030A
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const AnInch As Long = 1440
Private Const QuarterInch As Long = 360
Private Sub Command1_Click()
Form2.Show
End Sub
Private Sub Command2_Click()
Form3.Show
End Sub
Private Sub Command4_Click()
Form4.Show
End Sub
Private Sub Command6_Click()
 PrintRTF RichTextBox1, AnInch, AnInch, AnInch, AnInch
End Sub


