VERSION 5.00
Begin VB.Form Principal 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Inicio 
      BackColor       =   &H000080FF&
      Caption         =   "Iniciar"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   3
      Left            =   8040
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   2
      Left            =   8040
      Top             =   960
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   1
      Left            =   6720
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   6840
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Titulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Colores"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   5835
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Arianna Olivares"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   825
      TabIndex        =   1
      Top             =   3840
      Width           =   1845
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim colores(8) As Variant


Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Inicio_Click()

    Instrucciones.Show
    Principal.Hide
    

End Sub


Private Sub Timer1_Timer()

    If Timer1.Interval = 300 Then
        Titulo.ForeColor = Rnd * RGB(255, 255, 155)
    End If
    
'    colores(0) = RGB(0, 153, 255)
 '   colores(1) = RGB(255, 0, 0)
  '  colores(2) = RGB(2, 255, 51)
   ' colores(3) = RGB(153, 102, 255)
    'colores(4) = RGB(255, 153, 51)
    'colores(5) = RGB(0, 150, 0)
    'colores(6) = RGB(255, 155, 250)
    'colores(7) = RGB(255, 204, 0)

    For i = 0 To 3
        'j = Int(Rnd() * 8)
        'Shape1(i).FillColor = colores(j)
        Shape1(i).FillColor = Rnd * RGB(255, 255, 255)
     Next

End Sub

