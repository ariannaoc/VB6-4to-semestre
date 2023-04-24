VERSION 5.00
Begin VB.Form Instrucciones 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form3"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   LinkTopic       =   "Form3"
   ScaleHeight     =   4680
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton continuar 
      BackColor       =   &H0000C000&
      Caption         =   "Continuar"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   5100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "¿Como jugar?"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   3150
   End
End
Attribute VB_Name = "Instrucciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub continuar_Click()

    Juego.Show
    Instrucciones.Hide

End Sub

Private Sub Form_Load()
Label2.Caption = "El juego consiste en pulsar la mayor cantidad de recuadros azules que se posible antes de que cambien, de esta manera puedes sumar puntos. Recuerda que si pulsas una casilla que no sea azul puedes perder puntos"
End Sub


