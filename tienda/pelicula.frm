VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   Caption         =   "Form2"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form2"
   ScaleHeight     =   7395
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "ELIMINAR"
      Height          =   615
      Left            =   3360
      TabIndex        =   8
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MODIFICAR"
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GUARDAR"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREAR"
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante.SERVERINT\Desktop\ELIZABETH\Tienda de Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pelicula"
      Top             =   4440
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      DataField       =   "cod_Actor"
      DataSource      =   "Data1"
      Height          =   975
      Left            =   3480
      TabIndex        =   4
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      DataField       =   "cod_Tipo"
      DataSource      =   "Data1"
      Height          =   855
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Cod_Actor"
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Cod_Tipo"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PELICULA"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command3_Click()
Data1.Recordset.Update
End Sub

Private Sub Command5_Click()
Data1.Recordset.Delete
End Sub
