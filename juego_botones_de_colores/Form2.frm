VERSION 5.00
Begin VB.Form Juego 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9540
   LinkTopic       =   "Form2"
   ScaleHeight     =   4635
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tiempo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Comienzo 
      BackColor       =   &H0000FFFF&
      Caption         =   "Jugar"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   360
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   8880
      Top             =   4080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   15
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   14
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   13
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   12
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   11
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   10
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   9
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   8
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   7
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   6
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   5
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   4
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   3
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   2
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Index           =   1
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "Command1"
      Height          =   615
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Puntuacion 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Tiempo_etiquta 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Turno:"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      TabIndex        =   20
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Puntos"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6840
      TabIndex        =   0
      Top             =   3120
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      Height          =   3855
      Left            =   360
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "Juego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim contador As Integer
Dim colores(8) As Variant
Dim Inicio As Boolean

Private Sub Comienzo_Click()

    Inicio = True
    contador = 30
    Timer1.Enabled = True
    tiempo.Text = contador
        
End Sub


Private Sub Command1_Click(Index As Integer)

    If Command1(Index).BackColor = colores(0) Then
        Command1(Index).Visible = False
        Puntuacion.Text = Puntuacion.Text + 1
    Else
        Puntuacion.Text = Puntuacion.Text - 1
    
    End If


End Sub

Private Sub Form_Load()
    For i = 0 To 15
        Command1(i).Caption = Empty
    
    Next
    
    tiempo.Locked = True
    tiempo.Text = Empty
    Puntuacion.Locked = True
    Puntuacion.Text = 0
    Inicio = False
    


End Sub



Private Sub Timer1_Timer()
contador = contador - 1
tiempo.Text = contador

If contador = 0 Then
    MsgBox "Puntaje: " & Puntuacion.Text, vbOKOnly, "Fin del juego"
    Timer1.Enabled = False
    Comienzo.Enabled = False
    
    
End If

If Inicio = True Then
    colores(0) = RGB(0, 153, 255)
    colores(1) = RGB(255, 0, 0)
    colores(2) = RGB(2, 255, 51)
    colores(3) = RGB(153, 102, 255)
    colores(4) = RGB(255, 153, 51)
    colores(5) = RGB(0, 150, 0)
    colores(6) = RGB(255, 155, 250)
    colores(7) = RGB(255, 204, 0)

    For i = 0 To 15
        j = Int(Rnd() * 8)
        Command1(i).BackColor = colores(j)
        Command1(i).Visible = True
     Next

    Command1(Int(Rnd * 15)).BackColor = colores(0)
         
End If

If contador Mod 5 = 0 Then
    Timer1.Interval = Timer1.Interval - 200
End If
    
End Sub
