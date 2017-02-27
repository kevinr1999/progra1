VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H000000C0&
   Caption         =   "Form7"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form7"
   ScaleHeight     =   7500
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "SALIR"
      Height          =   975
      Left            =   5400
      TabIndex        =   6
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CLIENTE"
      Height          =   1095
      Left            =   3480
      TabIndex        =   5
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ALQUILER"
      Height          =   1095
      Left            =   3600
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DISCO"
      Height          =   1095
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ACTOR"
      Height          =   1095
      Left            =   480
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PELICULAS"
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TIPO DE PELICULA"
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show

End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Command4_Click()
Form4.Show
End Sub

Private Sub Command5_Click()
Form5.Show
End Sub

Private Sub Command6_Click()
Form6.Show
End Sub

Private Sub Command7_Click()
End
End Sub
