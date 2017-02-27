VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF00FF&
      Caption         =   "Eliminar"
      Height          =   555
      Left            =   5040
      TabIndex        =   8
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "crear"
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante.SERVERINT\Desktop\ELIZABETH\Tienda de Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Tipo de Pelicula"
      Top             =   4200
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      DataField       =   "categoria"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      DataField       =   "Tipo"
      DataSource      =   "Data1"
      Height          =   855
      Left            =   3120
      TabIndex        =   3
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "CATEGORIA"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "TIPO"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "TIPO DE PELICULA"
      BeginProperty Font 
         Name            =   "Adobe Gothic Std B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew

End Sub

Private Sub Command2_Click()
Data1.Recordset.Update
End Sub


Private Sub Command4_Click()
Data1.Recordset.Delete
End Sub
