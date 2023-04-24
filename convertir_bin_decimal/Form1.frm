VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "Convertidor "
   ClientHeight    =   5385
   ClientLeft      =   6075
   ClientTop       =   3195
   ClientWidth     =   7320
   FillColor       =   &H00404040&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000C&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7320
   Begin VB.TextBox Digit_5 
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Digit_4 
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Digit_3 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Digit_2 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      Caption         =   "Número decimal"
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   3720
      TabIndex        =   6
      Top             =   2400
      Width           =   2655
      Begin VB.TextBox RDecimal 
         BackColor       =   &H80000004&
         Height          =   495
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.TextBox Binario 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Digit_1 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Salir"
      Height          =   615
      Left            =   4920
      MaskColor       =   &H00808080&
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convertir"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Número Binario"
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   3720
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Ingrese dígitos binarios"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Arianna A. Olivares C.        30.680.385
Dim Digit1 As Integer
Dim Digit2 As Integer
Dim Digit3 As Integer
Dim Digit4 As Integer
Dim Digit5 As Integer

'Validar cuadros de texto

Private Function Validar(obj As TextBox)
    If Not IsNumeric(obj.Text) Then
        obj.Text = ""
    Else
        Digito = Val(obj.Text)
    End If
    
    If (Digito > 1) Then
    
    MsgBox "El valor no corresponde al sistema binario ingrese un número válido (0,1)", vbCritical + vbOKOnly, "Mensaje del Sistema"

    obj.Text = ""

   End If
End Function

'Botón de Convertir

Private Sub Command1_Click()

    Dim D1 As Integer
    Dim D2 As Integer
    Dim D3 As Integer
    Dim D4 As Integer
    Dim D5 As Integer
    Dim NumDecimal As Integer
    
    Binario.Text = Digit_1 & Digit_2 & Digit_3 & Digit_4 & Digit_5

    If (Len(Binario.Text) < 5) Then
        Binario.Text = ""
        MsgBox "Todas las casillas deben contener un valor 0 o 1", vbCritical + vbOKOnly, "Mensaje del Sistema"
        
    Else

' Operacion de conversion

       D5 = Val(Digit_5) * 2 ^ 0
       D4 = Val(Digit_4) * 2 ^ 1
       D3 = Val(Digit_3) * 2 ^ 2
       D2 = Val(Digit_2) * 2 ^ 3
       D1 = Val(Digit_1) * 2 ^ 4
       
    NumDecimal = D5 + D4 + D3 + D2 + D1
    
    RDecimal.Text = NumDecimal
 End If

End Sub

'Botón de Salir

Private Sub Command2_Click()
 End
End Sub

'Cuadros de texto

Private Sub Digit_1_Change()
    Call Validar(Digit_1)
End Sub

Private Sub Digit_2_Change()
    Call Validar(Digit_2)
End Sub

Private Sub Digit_3_Change()
    Call Validar(Digit_3)
End Sub

Private Sub Digit_4_Change()
    Call Validar(Digit_4)
End Sub

Private Sub Digit_5_Change()
    Call Validar(Digit_5)
End Sub

