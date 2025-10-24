VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10755
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   10755
   Begin VB.TextBox txtUserId 
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Text            =   "1"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtBody 
      Height          =   615
      Left            =   1080
      TabIndex        =   7
      Text            =   "Este es el cuerpo del post"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Text            =   "Mi Título"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtRespuestaPost 
      Height          =   855
      Left            =   6600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton btnEnviarPost 
      Caption         =   "Enviar POST"
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Json:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   1200
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   5400
      X2              =   5400
      Y1              =   120
      Y2              =   2040
   End
   Begin VB.Label Label4 
      Caption         =   "UserId:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Body:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblEstado 
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Respuesta POST:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const URL_API_POST As String = "https://jsonplaceholder.typicode.com/posts"

' FUNCIÓN envía el nuevo "post" a la API
Public Function ConsultarPost(ByRef CodigoEstado As Long) As String
    
    Dim objeto_HTTP As MSXML2.XMLHTTP60
    Dim cadena_JSON As String
    
    Dim titulo As String: titulo = Trim(txtTitle.Text)
    Dim cuerpo As String: cuerpo = Trim(txtBody.Text)
    Dim id_usuario As String: id_usuario = Trim(txtUserId.Text)
    
    On Error GoTo ManejadorErrores
    
    If titulo = "" Or cuerpo = "" Or id_usuario = "" Then
        ConsultarPost = "ERROR_VALIDACION: Campos vacíos."
        Exit Function
    End If
    
    cadena_JSON = "{" & _
                  """title"": """ & titulo & """, " & _
                  """body"": """ & cuerpo & """, " & _
                  """userId"": " & id_usuario & _
                  "}"
    
    Set objeto_HTTP = New MSXML2.XMLHTTP60
    objeto_HTTP.Open "POST", URL_API_POST, False
    objeto_HTTP.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    objeto_HTTP.setRequestHeader "Accept", "application/json"
    
    objeto_HTTP.send cadena_JSON
    
    
    CodigoEstado = objeto_HTTP.Status
    
    ConsultarPost = objeto_HTTP.responseText
    
    Set objeto_HTTP = Nothing
    Exit Function
    
ManejadorErrores:
    CodigoEstado = 0
    ConsultarPost = "ERROR_EJECUCION: " & Err.Description
    Set objeto_HTTP = Nothing
End Function


Private Sub btnEnviarPost_Click()
Dim resultado_JSON As String
    Dim codigo_HTTP As Long
    
    lblEstado.Caption = "Enviando solicitud POST..."
    lblEstado.ForeColor = vbBlack
    txtRespuestaPost.Text = ""
    
    resultado_JSON = ConsultarPost(codigo_HTTP)
    
    
    If Left(resultado_JSON, 6) = "ERROR_" Then
        lblEstado.Caption = resultado_JSON
        lblEstado.ForeColor = vbRed
        
    ElseIf codigo_HTTP = 201 Then
        
        lblEstado.Caption = "ÉXITO (CÓDIGO 201): Post creado."
        lblEstado.ForeColor = vbBlue
        txtRespuestaPost.Text = resultado_JSON
        
    Else
        lblEstado.Caption = "ERROR HTTP: Código " & codigo_HTTP
        lblEstado.ForeColor = vbRed
        txtRespuestaPost.Text = resultado_JSON
        
    End If
End Sub
























