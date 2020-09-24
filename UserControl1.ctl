VERSION 5.00
Begin VB.UserControl cal 
   BackColor       =   &H000000FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   ScaleHeight     =   3600
   ScaleWidth      =   1755
   Begin VB.Label interval 
      Caption         =   "1"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label nota 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "hora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label hora 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "hora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Label fechas 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label dia 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "Dia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "cal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim calendario As Date
Public Property Get Day() As String
    Day = dia.Caption
End Property

Public Property Let Day(ByVal sNewValue As String)
    dia.Caption = sNewValue
End Property
Public Property Get fecha() As String
    fecha = fechas.Caption
    poner_dia
End Property

Public Property Let fecha(ByVal sNewValue As String)
    fechas.Caption = sNewValue
End Property
Public Property Get intervalo() As String
    intervalo = interval.Caption
End Property

Public Property Let intervalo(ByVal sNewValue As String)
    interval.Caption = sNewValue
    poner_horas
End Property

Private Sub fechas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For n = 0 To nota.Count - 1
    nota(n).BackColor = &HC0FFFF
    nota(n).ForeColor = &HFF&
Next n
End Sub

Private Sub nota_DblClick(Index As Integer)
If nota(Index).Caption <> "" And Len(Trim(nota(Index).Caption)) > 0 Then
    actual = nota(Index).Caption
End If
retStr = InputBox("Introduce la nota", "Notas", actual)
nota(Index).Caption = retStr
If nota(Index).Caption <> "" And Len(Trim(nota(Index).Caption)) > 0 Then
    nota(Index).BackColor = &H80FFFF
    nota(Index).ForeColor = &HC00000
Else
    nota(Index).BackColor = &HC0FFFF
    nota(Index).ForeColor = &HFF&
End If
End Sub

Private Sub nota_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'For n = 0 To nota.Count - 1
'    If n = Index Then
'        nota(Index).BackColor = &HFF8080
'        nota(Index).ForeColor = RGB(255, 255, 255)
'    Else
'        nota(n).BackColor = &HC0FFFF
'        nota(n).ForeColor = &HFF&
'    End If
'Next n
End Sub

Private Sub UserControl_Initialize()

dia.Width = UserControl.Width
fechas.Width = UserControl.Width
    poner_dia
    'fechas.Caption = Date
    poner_horas
End Sub
Sub poner_dia()
If IsDate(fechas.Caption) Then
    Select Case Weekday(fechas.Caption, vbMonday)
        Case Is = 1: dia.Caption = "Lunes":
        Case Is = 2: dia.Caption = "Martes":
        Case Is = 3: dia.Caption = "Miercoles":
        Case Is = 4: dia.Caption = "Jueves":
        Case Is = 5: dia.Caption = "Viernes":
        Case Is = 6: dia.Caption = "Sabado":
        Case Else
            dia.Caption = "Domingo":
    End Select
End If
End Sub
Sub poner_horas()
If nota.Count > 0 Then
    For i = 1 To nota.Count - 1
        Unload nota(i)
    Next i
End If
If hora.Count > 0 Then
    For i = 1 To hora.Count - 1
        Unload hora(i)
    Next i
End If

Select Case CInt(interval.Caption)
    Case 1 'Mostraremos de hora en hora
        hora(0).Caption = "00:00"
        nota(0).Caption = ""
        hora(0).Height = 255
        nota(0).Height = 255
        For j = 1 To 23
            Load hora(j)
            Load nota(j)
            hora(j).Top = (hora(j - 1).Top + hora(j).Height) + 15
            nota(j).Top = (nota(j - 1).Top + nota(j).Height) + 15
            hora(j).Left = 0
            nota(j).Left = nota(j - 1).Left
            nota(j).Caption = ""
            nota(j).Visible = True
            If j < 10 Then
                etiqueta = "0" & j & ":00"
            Else
                etiqueta = j & ":00"
            End If
            hora(j).Caption = etiqueta
            hora(j).Visible = True
        Next j
    Case 2 'Mostraremos cada 2 horas
        hora(0).Caption = "00:00"
        hora(0).Height = hora(0).Height * 2
        nota(0).Caption = ""
        nota(0).Height = nota(0).Height * 2
        anadir = 1
        For j = 1 To 12
            Load hora(j)
            Load nota(j)
            hora(j).Top = (hora(j - 1).Top + hora(j).Height) + 15
            nota(j).Top = (nota(j - 1).Top + nota(j).Height) + 15
            hora(j).Left = 0
            nota(j).Left = nota(j - 1).Left
            nota(j).Caption = ""
            nota(j).Visible = True
            If j < 10 And j + anadir < 10 Then
                etiqueta = "0" & j + anadir & ":00"
            Else
                If j + anadir = 24 Then
                    etiqueta = "23:00"
                Else
                    etiqueta = j + anadir & ":00"
                End If
            End If
            hora(j).Caption = etiqueta
            anadir = anadir + 1
            hora(j).Visible = True
        Next j
    Case 0 'Mostraremos cada 30minutos
        hora(0).Caption = "00:00"
        nota(0).Caption = ""
        hora(0).Height = hora(0).Height - 85
        nota(0).Height = nota(0).Height - 85
        nota(0).FontSize = 9
        hora(0).FontSize = 9
        anadir = 1
        veces = 1
        For j = 1 To 48
            Load hora(j)
            Load nota(j)
            hora(j).Top = (hora(j - 1).Top + hora(j).Height)
            nota(j).Top = (nota(j - 1).Top + nota(j).Height)
            hora(j).Left = 0
            nota(j).Left = nota(j - 1).Left
            nota(j).Caption = ""
            nota(j).Visible = True
            If anadir = 1 Then
                If j - veces < 10 Then
                    etiqueta = "0" & j - veces & ":30"
                Else
                    etiqueta = j - veces & ":30"
                End If
            Else
                If j - veces < 10 Then
                    etiqueta = "0" & j - veces & ":00"
                Else
                    etiqueta = j - veces & ":00"
                End If
            End If
            hora(j).Caption = etiqueta
            If anadir = 1 Then
                anadir = 0
            Else
                anadir = 1
                veces = veces + 1
            End If
            hora(j).Visible = True
        Next j
    
End Select
UserControl.Height = hora(hora.Count - 1).Top + hora(hora.Count - 1).Height
UserControl.Width = fechas.Width
        UserControl_Resize
End Sub
'Public Sub quitar_color()
''For n = 0 To nota.Count - 1
''    nota(n).BackColor = &HC0FFFF
''    nota(n).ForeColor = &HFF&
''Next n
'        UserControl.Height = hora(hora.Count - 1).Top + hora(hora.Count - 1).Height
'        UserControl_Resize
'End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For n = 0 To nota.Count - 1
    If nota(n).Caption <> "" And Len(Trim(nota(n).Caption)) > 0 Then
        nota(n).BackColor = &H80FFFF
        nota(n).ForeColor = &HC00000
    Else
        nota(n).BackColor = &HC0FFFF
        nota(n).ForeColor = &HFF&
    End If
    'nota(n).BackColor = &HC0FFFF
    'nota(n).ForeColor = &HFF&
Next n


End Sub

Private Sub UserControl_Resize()
On Error Resume Next
If UserControl.Width < 1500 Then
    UserControl.Width = 1500
End If
dia.Width = UserControl.Width
fechas.Width = UserControl.Width

For n = 0 To hora.Count - 1
    hora(n).Left = 0
    If UserControl.Width > 2000 Then
        j = (UserControl.Width / 4)
    Else
        j = (UserControl.Width / 3)
    End If
    'j = j - 150
    hora(n).Width = j
    'hora(n).Caption = hora(n).Width
Next n
For n = 0 To nota.Count - 1
    nota(n).Left = hora(0).Width
    'nota(n).Caption = nota(n).Left
    nota(n).Width = (UserControl.Width - hora(0).Width)
Next n

UserControl.Height = hora(hora.Count - 1).Top + hora(hora.Count - 1).Height
'UserControl.Width = (nota(0).Left + nota(0).Width) + 50
End Sub
