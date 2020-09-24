VERSION 5.00
Begin VB.Form calendar 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cambfecha 
      Caption         =   ">"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   10
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cambfecha 
      Caption         =   ">>"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   9
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cambfecha 
      Caption         =   "<<"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   8
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cambfecha 
      Caption         =   "<"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Una Semana"
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   6
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6 dia(s)"
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   5
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5 dia(s)"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   4
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4 dia(s)"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   3
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3 dia(s)"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2 dia(s)"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   7440
      Width           =   1095
   End
   Begin Project1.cal calendario 
      Height          =   6945
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   12250
   End
End
Attribute VB_Name = "calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mdias As Integer
Dim fecha_poner As Date




Private Sub cambfecha_Click(Index As Integer)
Select Case Index
    Case 0
        fecha = calendario(0).fecha
        fecha_poner = DateAdd("d", -1, fecha)
    Case 1
        fecha = calendario(0).fecha
        fecha_poner = DateAdd("d", -7, fecha)
    Case 2
        fecha = calendario(0).fecha
        fecha_poner = DateAdd("d", 7, fecha)
    
    Case 3
        fecha = calendario(0).fecha
        fecha_poner = DateAdd("d", 1, fecha)
End Select
pongo_cal
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0 'Mostrar 1 Dia
        mdias = 0
    Case 1 'Mostrar 2 Dias
        mdias = 1
    Case 2 'Mostrar 3 Días
        mdias = 2
    Case 3 'Mostrar 4 Días
        mdias = 3
    Case 4 'Mostrar 5 Días
        mdias = 4
    Case 5 'Mostrar 6 Días
        mdias = 5
    Case 6 'Mostrar una Semana
        mdias = 6
End Select
pongo_cal

End Sub

Private Sub Form_Load()
mdias = 6
fecha_poner = Date
pongo_cal
calendar.Width = calendario(calendario.Count - 1).Left + calendario(calendario.Count - 1).Width
End Sub
Sub pongo_cal()
If calendario.Count > 0 Then
    For n = 1 To calendario.Count - 1
        Unload calendario(n)
    Next n
End If
If mdias = 0 Then
    ancho = calendar.Width
Else
    ancho = calendar.Width / (mdias + 1)
End If
calendario(0).fecha = fecha_poner
calendario(0).poner_dia
calendario(0).Width = ancho
calendario(0).Left = 0
For n = 1 To mdias
    Load calendario(n)
    calendario(n).Width = ancho
    calendario(n).Left = (calendario(n - 1).Left + calendario(n - 1).Width)
    calendario(n).Visible = True
    fecha_siguiente = DateAdd("d", n, fecha_poner)
    fecha_siguiente = Format(fecha_siguiente, "Short Date")
    calendario(n).fecha = fecha_siguiente
    calendario(n).poner_dia
Next n

End Sub
