VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H0000C0C0&
   Caption         =   "Form4"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form4"
   ScaleHeight     =   7830
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "ELIMINAR"
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MODIFICAR"
      Height          =   615
      Left            =   3240
      TabIndex        =   11
      Top             =   6240
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GUARDAR"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREAR"
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante.SERVERINT\Desktop\ELIZABETH\Tienda de Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Disco"
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   4320
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      DataField       =   "cod_pelicula"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      Top             =   3360
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      DataField       =   "Num_copias"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3360
      TabIndex        =   6
      Top             =   2400
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      DataField       =   "código"
      DataSource      =   "Data1"
      Height          =   855
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "FORMATO"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "COD_PELICULA"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "NUM_COPIAS"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "CODIGO"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "DISCO"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form4"
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
