VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form6"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form6"
   ScaleHeight     =   8220
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "ELIMINAR"
      Height          =   615
      Left            =   3840
      TabIndex        =   12
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MODIFICAR"
      Height          =   735
      Left            =   3720
      TabIndex        =   11
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GUARDAR"
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREAR"
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante.SERVERINT\Desktop\ELIZABETH\Tienda de Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cliente"
      Top             =   4680
      Width           =   4215
   End
   Begin VB.TextBox Text4 
      DataField       =   "Telefono"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   3960
      Width           =   4935
   End
   Begin VB.TextBox Text3 
      DataField       =   "Direccion"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   3240
      Width           =   5055
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2520
      TabIndex        =   6
      Top             =   2400
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      DataField       =   "Num_Membresia"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label Label5 
      Caption         =   "TELEFONO"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "DIRECCION"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "NOMBRE"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "NUM_MEMBRESIA"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "CLIENTE"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   26.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form6"
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
