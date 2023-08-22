VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "clave desencriptada"
      Height          =   1695
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   7935
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Encriptar"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   6120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Desencriptar"
      Height          =   855
      Left            =   5040
      TabIndex        =   1
      Top             =   6120
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Audiowide"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   7215
   End
   Begin VB.Frame Frame1 
      Caption         =   "contraseña"
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   7935
   End
   Begin VB.Frame Frame2 
      Caption         =   "clave encriptada"
      Height          =   2055
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   7935
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   7215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub crear_randon()

End Sub


Private Sub Command1_Click()
Dim i As Integer
Dim sacar(40) As String
Dim a, b, c  As Integer
Label1.Caption = ""
c = 1
For i = 1 To Len(Label2.Caption)
 a = Val(i)
 b = a Mod 2
 If b = 0 Then
    sacar(i) = Mid(Label2.Caption, a, 1)
    sacar(i) = Asc(sacar(i))
    sacar(i) = Val(sacar(i)) - c
    sacar(i) = Chr(sacar(i))
    c = c + 1
End If
Label1.Caption = Label1.Caption & sacar(i) 'sacar(1) & sacar(2) & sacar(3) & sacar(4) & sacar(5) & sacar(6)
Next i
End Sub

Private Sub Command2_Click()
Dim clave(20) As String
Dim sacar(20) As String
Dim i, a, b, c As Integer
c = 1
Label2.Caption = ""
For i = 1 To Len(Text1.Text)
 a = Val(i)
 b = a Mod 2
 If b = 0 Then
    clave(i) = Int((9 * Rnd) + 1)
 Else
    clave(i) = Chr(Int(24 * Rnd) + 65)
 End If
 sacar(i) = UCase(Mid(Text1.Text, i, 1))
 sacar(i) = Asc(sacar(i))
 sacar(i) = Val(sacar(i)) + c
 sacar(i) = Chr(sacar(i))
 Label2.Caption = Label2.Caption & clave(i) & sacar(i)
 c = c + 1
Next i
Label2.Caption = UCase(Label2.Caption)
End Sub
