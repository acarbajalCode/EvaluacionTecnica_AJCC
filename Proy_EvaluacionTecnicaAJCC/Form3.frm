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

' Constante: URL para obtener el array de 10 usuarios
Private Const URL_API_USUARIOS As String = "https://jsonplaceholder.typicode.com/users"

Private Const COLOR_EXITO As Long = &HFF0000    ' Azul
Private Const COLOR_ERROR As Long = &HFF&         ' Rojo
Private Const COLOR_NEGRO As Long = &H0&

Sub Log(ByVal mensaje As String)
    ' Agrega el mensaje al texto existente seguido de un salto de línea (vbCrLf)
    txtResultados.Text = txtResultados.Text & mensaje & vbCrLf
End Sub

' a. Consumir la API y procesar la respuesta
Sub ConsultarUsuarios()
    
    ' --- Manejo de Variables y Objetos ---
    Dim objeto_HTTP As MSXML2.XMLHTTP60
    Dim sJsonResponse As String
    Dim colUsuarios As Collection
    Dim vUser As Dictionary
    
    ' --- Inicialización de la Interfaz ---
    On Error GoTo ManejadorErrores
    
    txtResultados.Text = ""         ' Limpia el log
    lstUsuarios.Clear               ' Limpia el ListBox para los usuarios
    
    lblEstado.Caption = "Cargando..."
    lblEstado.ForeColor = COLOR_NEGRO
    
    Log "Iniciando solicitud GET a la API de Usuarios..."
    
    ' --- PASO 1: Realizar una solicitud GET ---
    Set objeto_HTTP = New MSXML2.XMLHTTP60
    
    objeto_HTTP.Open "GET", URL_API_USUARIOS, False ' Solicitud GET Sincrónica
    objeto_HTTP.send
    
    ' --- Procesamiento de la Respuesta HTTP ---
    If objeto_HTTP.Status = 200 Then
        sJsonResponse = objeto_HTTP.responseText
        
        Log "Respuesta recibida (" & Len(sJsonResponse) & " bytes). Parseando..."
        
        ' --- PASO 2: Utilizar una librería de parseo de JSON ---
        Set colUsuarios = JsonConverter.ParseJson(sJsonResponse)
        
        Log String$(50, "-")
        Log "Usuarios encontrados: " & colUsuarios.Count & ". Mostrando en ListBox..."
        Log String$(50, "-")
        
        ' --- Encabezado Limpio y Concatenado EN EL LISTBOX ---
        lstUsuarios.AddItem "Nombre                  | Email"
        lstUsuarios.AddItem String$(50, "-")
        
        ' --- PASO 3 & 4: Iterar, Concatenar y Agregar a lstUsuarios ---
        For Each vUser In colUsuarios
            
            Dim sName As String
            Dim sEmail As String
            
            sName = vUser("name")
            sEmail = vUser("email")
            
            ' AÑADIR AL LISTBOX: Concatenar name y email con formato tabular
            lstUsuarios.AddItem Left(sName & Space(25), 25) & "| " & sEmail
            
        Next vUser
        
        ' --- Indicador de Éxito ---
        lblEstado.Caption = "Consumo OK"
        lblEstado.ForeColor = COLOR_EXITO
        
    Else
        ' Caso de error HTTP (código 4xx, 5xx)
        Log "ERROR HTTP: Código " & objeto_HTTP.Status & ". Fallo al obtener usuarios."
        
        ' --- Indicador de Error ---
        lblEstado.Caption = "ERROR HTTP"
        lblEstado.ForeColor = COLOR_ERROR
    End If

    ' --- Limpieza ---
    Set objeto_HTTP = Nothing
    Exit Sub
    
' --- Bloque de CATCH (Manejo de Errores Críticos en Tiempo de Ejecución) ---
ManejadorErrores:
    
    Log String$(50, "#")
    Log "ERROR CRÍTICO EN TIEMPO DE EJECUCIÓN:"
    Log "Descripción: " & Err.Description
    
    ' --- Indicador de Error ---
    lblEstado.Caption = "ERROR CRÍTICO"
    lblEstado.ForeColor = COLOR_ERROR
    
    ' --- Limpieza ---
    Set objeto_HTTP = Nothing
End Sub


Private Sub btnConsumirAPI_Click()
ConsultarUsuarios
End Sub




