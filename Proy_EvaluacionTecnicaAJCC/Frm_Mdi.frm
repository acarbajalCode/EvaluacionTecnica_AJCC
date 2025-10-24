VERSION 5.00
Begin VB.MDIForm Frm_Mdi 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5850
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10485
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu FrmS 
      Caption         =   "Formularios"
      Begin VB.Menu FrmPregunta1 
         Caption         =   "Pregunta 1"
      End
      Begin VB.Menu FrmPregunta2 
         Caption         =   "Pregunta 2"
      End
      Begin VB.Menu FrmPregunta3 
         Caption         =   "Pregunta 3"
      End
      Begin VB.Menu FrmReniecAPI 
         Caption         =   "Consume Datos RENIEC API"
      End
   End
   Begin VB.Menu FrmSalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "Frm_Mdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub FrmPregunta1_Click()
Form1.Show

End Sub

Private Sub FrmPregunta2_Click()
Form2.Show
End Sub

Private Sub FrmPregunta3_Click()
Form3.Show
End Sub

Private Sub FrmReniecAPI_Click()
Form4.Show
End Sub

Private Sub FrmSalir_Click()
Dim respuesta As Integer
    
    ' Mostrar un cuadro de mensaje de confirmación con botones SÍ y NO
    respuesta = MsgBox("¿Está seguro que desea salir de la aplicación?", vbYesNo + vbQuestion, "Confirmar Salida")
    
    ' Chequear la respuesta del usuario
    If respuesta = vbYes Then
        ' Si el usuario selecciona "Sí", descarga el formulario actual.
        Unload Me
        
        ' (Opcional) Si Form2 es el último formulario visible,
        ' también puedes usar End para detener la aplicación completa:
        ' End
    End If
End Sub

Private Sub MDIForm_Load()
Me.Height = 10000
Me.Width = 15000
Me.Left = 0
Me.Top = 0

End Sub
