VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form5"
   ClientHeight    =   9210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form5"
   ScaleHeight     =   9210
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "ELIMINAR"
      Height          =   495
      Left            =   3480
      TabIndex        =   18
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MODIFICAR"
      Height          =   495
      Left            =   3480
      TabIndex        =   17
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GUARDAR"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   8520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREAR"
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante.SERVERINT\Desktop\ELIZABETH\Tienda de Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Alquiler"
      Top             =   6720
      Width           =   4335
   End
   Begin VB.TextBox Text7 
      DataField       =   "cantidad"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2880
      TabIndex        =   14
      Top             =   5760
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      DataField       =   "Valor_Alquiler"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2880
      TabIndex        =   13
      Top             =   5040
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      DataField       =   "Fecha_Devolucion"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2760
      TabIndex        =   12
      Top             =   4320
      Width           =   4575
   End
   Begin VB.TextBox Text4 
      DataField       =   "Fecha_Alquiler"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2760
      TabIndex        =   11
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      DataField       =   "cod_Cliente"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   2760
      TabIndex        =   10
      Top             =   2760
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      DataField       =   "cod_Disco"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2760
      TabIndex        =   9
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      DataField       =   "código"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label8 
      Caption         =   "CANTIDAD"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "VALOR_ALQUILER"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "FECHA_DEVOLUCION"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "FECHA_ALQUILER"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "COD_CLIENTE"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "COD_DISCO"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "CODIGO"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "ALQUILER"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "Form5"
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
