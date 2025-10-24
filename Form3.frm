VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9420
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   9420
   Begin VB.TextBox txtResultados 
      Height          =   1095
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton btnConsumirAPI 
      Caption         =   "Procesar Respuesta API"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox lstUsuarios 
      Height          =   2400
      ItemData        =   "Form3.frx":0000
      Left            =   5040
      List            =   "Form3.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   4680
      Y1              =   240
      Y2              =   2640
   End
   Begin VB.Label Label1 
      Caption         =   "Respuesta del Consumo API:"
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
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblEstado 
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const URL_API_USUARIOS As String = "https://jsonplaceholder.typicode.com/users"

Private Const COLOR_EXITO As Long = &HFF0000
Private Const COLOR_ERROR As Long = &HFF&
Private Const COLOR_NEGRO As Long = &H0&

Sub Log(ByVal mensaje As String)
    txtResultados.Text = txtResultados.Text & mensaje & vbCrLf
End Sub


Sub ConsultarUsuarios()
    
    Dim objeto_HTTP As MSXML2.XMLHTTP60
    Dim sJsonResponse As String
    Dim colUsuarios As Collection
    Dim vUser As Dictionary
    
    
    On Error GoTo ManejadorErrores
    
    txtResultados.Text = ""
    lstUsuarios.Clear
    
    lblEstado.Caption = "Cargando..."
    lblEstado.ForeColor = COLOR_NEGRO
    
    Log "Iniciando solicitud GET a la API de Usuarios..."
    
   
    Set objeto_HTTP = New MSXML2.XMLHTTP60
    
    objeto_HTTP.Open "GET", URL_API_USUARIOS, False
    objeto_HTTP.send
    
    If objeto_HTTP.Status = 200 Then
        sJsonResponse = objeto_HTTP.responseText
        
        Log "Respuesta recibida (" & Len(sJsonResponse) & " bytes). Parseando..."
        
        Set colUsuarios = JsonConverter.ParseJson(sJsonResponse)
        
        Log String$(50, "-")
        Log "Usuarios encontrados: " & colUsuarios.Count & ". Mostrando en ListBox..."
        Log String$(50, "-")
        
        lstUsuarios.AddItem "Nombre                  | Email"
        lstUsuarios.AddItem String$(50, "-")
        
        For Each vUser In colUsuarios
            
            Dim sName As String
            Dim sEmail As String
            
            sName = vUser("name")
            sEmail = vUser("email")
            
            lstUsuarios.AddItem Left(sName & Space(25), 25) & "| " & sEmail
            
        Next vUser
        
        lblEstado.Caption = "Consumo OK"
        lblEstado.ForeColor = COLOR_EXITO
        
    Else
        Log "ERROR HTTP: Código " & objeto_HTTP.Status & ". Fallo al obtener usuarios."
        
        lblEstado.Caption = "ERROR HTTP"
        lblEstado.ForeColor = COLOR_ERROR
    End If

    Set objeto_HTTP = Nothing
    Exit Sub
    
ManejadorErrores:
    
    Log String$(50, "#")
    Log "ERROR CRÍTICO EN TIEMPO DE EJECUCIÓN:"
    Log "Descripción: " & Err.Description
    
    lblEstado.Caption = "ERROR CRÍTICO"
    lblEstado.ForeColor = COLOR_ERROR
    
    Set objeto_HTTP = Nothing
End Sub


Private Sub btnConsumirAPI_Click()
    ConsultarUsuarios
End Sub




