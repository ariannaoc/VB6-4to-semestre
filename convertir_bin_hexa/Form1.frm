VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Convertidor Bin/Hex"
   ClientHeight    =   4905
   ClientLeft      =   5730
   ClientTop       =   2970
   ClientWidth     =   7140
   FillColor       =   &H80000004&
   ForeColor       =   &H00FF8080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   7140
   Begin VB.TextBox text_hexa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox text_binario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton Convertir 
      BackColor       =   &H8000000B&
      Caption         =   "Convertir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Limpiar 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Salir 
      BackColor       =   &H8000000B&
      Caption         =   "Salir"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Boton0 
      BackColor       =   &H000000FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Boton1 
      BackColor       =   &H00C0C000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Hexadecimal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Binario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim uno As String
Dim cero As String


Private Sub Boton0_Click()
    
    cero = "0"
    text_binario.Text = text_binario.Text & cero

    Call NumDigitos

End Sub

Private Sub Boton1_Click()

    uno = "1"
    text_binario.Text = text_binario.Text & uno
    
    Call NumDigitos
    
End Sub

Private Function NumDigitos()
    
    If (Len(text_binario.Text) = 4) Then
        Boton0.Enabled = False
        Boton1.Enabled = False
    End If

End Function


Private Sub Convertir_Click()
Dim NumBinario As Integer

NumBinario = Val(text_binario.Text)

Select Case (NumBinario)
    Case 0
        text_hexa.Text = 0
    Case 1
        text_hexa.Text = 1
    Case 10
        text_hexa.Text = 2
    Case 11
        text_hexa.Text = 3
    Case 100
        text_hexa.Text = 4
    Case 101
        text_hexa.Text = 5
    Case 110
        text_hexa.Text = 6
    Case 111
        text_hexa.Text = 7
    Case 1000
        text_hexa.Text = 8
    Case 1001
        text_hexa.Text = 9
    Case 1010
        text_hexa.Text = "A"
    Case 1011
        text_hexa.Text = "B"
    Case 1100
        text_hexa.Text = "C"
    Case 1101
        text_hexa.Text = "D"
    Case 1110
        text_hexa.Text = "E"
    Case Else
        text_hexa.Text = "F"
End Select



End Sub

Private Sub Limpiar_Click()
    text_binario.Text = ""
    text_hexa.Text = ""
    
    Boton1.Enabled = True
    Boton0.Enabled = True
    
End Sub

Private Sub Salir_Click()
End
End Sub
