VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   11295
   Begin VB.CommandButton btnConsultaDNI 
      Caption         =   "Consultar Información"
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton btnGenerarToken 
      BackColor       =   &H8000000B&
      Caption         =   "Obtener Token"
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtDNI 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin MSComctlLib.ListView lsvHistorial 
      Height          =   2415
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblEstado 
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label lblToken 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese DNI a Consultar:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Constantes del Servicio
Private Const URL_BASE_API As String = "https://miapi.cloud/v1/dni/"

' Variable a nivel de módulo para almacenar el Bearer Token
Private Token_Autorizacion As String

' Constantes de color
Private Const COLOR_EXITO As Long = &HFF0000
Private Const COLOR_ERROR As Long = &HFF&

' Variable global para el número de consulta (Historial)
Private Historial_Contador As Long




Private Sub btnConsultaDNI_Click()
ConsultarDNI
End Sub

Private Sub btnGenerarToken_Click()
' El Token real de la imagen es: eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjc2ViVyyx2KfjojM0jmQslmV4ccI6MTg2MTg5NDU5MHI0.XAT_LIZj6b64p2SZdWGfoBuhkdD5TWr2i-qIGLL4ANa4
    ' Usamos este valor como simulación
    Token_Autorizacion = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VyX2lkIjo0MjQsImV4cCI6MTc2MTg5NDU5MH0.XAT_UZI6b64p2SZdWGfoBuhkdD5TWr2i-qIGLL4ANa4"
    
    lblToken.Caption = "TOKEN OBTENIDO"
    lblToken.ForeColor = vbBlue
    
    MsgBox "Token de autorización simulado y cargado.", vbInformation
End Sub

Public Sub ConsultarDNI()
    
    Dim objeto_HTTP As MSXML2.XMLHTTP60
    Dim DNI_Consulta As String
    Dim UrlCompleta As String
    Dim respuesta_JSON As String
    Dim codigo_HTTP As Long
    
    Dim parsedJSON As Dictionary ' Requiere JsonConverter.bas y Scripting Runtime
    Dim datosUsuario As Dictionary
    
    On Error GoTo ManejadorErrores
    
    ' 1. VALIDACIÓN INICIAL ESTRICTA
    DNI_Consulta = Trim(txtDNI.Text)
    
    If Token_Autorizacion = "" Then
        lblEstado.Caption = "ERROR: Token de autorización faltante. Presione 'Obtener Token'."
        lblEstado.ForeColor = COLOR_ERROR
        Exit Sub
    End If
    
    ' --- VALIDACIÓN NUMÉRICA Y LONGITUD (8 DÍGITOS) ---
    If Not IsNumeric(DNI_Consulta) Then
        lblEstado.Caption = "ERROR: El DNI debe ser solo numérico."
        lblEstado.ForeColor = COLOR_ERROR
        Exit Sub
    End If
    
    If Len(DNI_Consulta) <> 8 Then
        lblEstado.Caption = "ERROR: El DNI debe tener exactamente 8 dígitos."
        lblEstado.ForeColor = COLOR_ERROR
        Exit Sub
    End If
    ' ----------------------------------------------------

    ' 2. SOLICITUD GET
    UrlCompleta = URL_BASE_API & DNI_Consulta
    lblEstado.Caption = "Consultando DNI " & DNI_Consulta & "..."
    lblEstado.ForeColor = vbBlack
    
    Set objeto_HTTP = New MSXML2.XMLHTTP60
    objeto_HTTP.Open "GET", UrlCompleta, False
    
    ' Cabecera de Autorización y Contenido
    objeto_HTTP.setRequestHeader "Authorization", "Bearer " & Token_Autorizacion
    objeto_HTTP.setRequestHeader "Content-Type", "application/json"
    
    objeto_HTTP.send
    
    codigo_HTTP = objeto_HTTP.Status
    respuesta_JSON = objeto_HTTP.responseText
    
    ' 3. MANEJO DE RESPUESTA Y PARSEO
    If codigo_HTTP = 200 Then
        
        Set parsedJSON = JsonConverter.ParseJson(respuesta_JSON)
        
        If parsedJSON("success") = True Then
            ' ÉXITO LÓGICO: DNI ENCONTRADO
            Set datosUsuario = parsedJSON("datos")
            
            Dim nombre_completo As String
            nombre_completo = datosUsuario("nombres") & " " & datosUsuario("ape_paterno") & " " & datosUsuario("ape_materno")
            
            ' Agrega el registro al historial
            AgregarRegistroHistorial DNI_Consulta, nombre_completo, "EXITO"
            
            lblEstado.Caption = "ÉXITO: DNI encontrado y agregado al historial."
            lblEstado.ForeColor = COLOR_EXITO
            
        Else
            ' FALLO LÓGICO: DNI no encontrado, pero la API respondió 200.
            AgregarRegistroHistorial DNI_Consulta, "NO ENCONTRADO EN API", "FALLO"
            lblEstado.Caption = "FALLO: DNI no encontrado o error lógico de API."
            lblEstado.ForeColor = COLOR_ERROR
        End If
        
    ElseIf codigo_HTTP = 404 Then
        
        ' *** LÓGICA DE EXCLUSIÓN: Muestra el mensaje CLARO pero NO lo guarda en el historial ***
        lblEstado.Caption = "ERROR 404: DNI NO SE ENCONTRÓ EN LA BD DE RENIEC O NO EXISTE."
        lblEstado.ForeColor = COLOR_ERROR
        ' NO SE LLAMA A AgregarRegistroHistorial
        
    Else
        ' Fallo HTTP General (401 Unauthorized, 500 Server Error, etc.)
        AgregarRegistroHistorial DNI_Consulta, "ERROR HTTP " & codigo_HTTP, "ERROR"
        lblEstado.Caption = "ERROR HTTP: Código " & codigo_HTTP & " - Revisar Token o Conexión."
        lblEstado.ForeColor = COLOR_ERROR
    End If
    
FinalizarProceso:
    Set objeto_HTTP = Nothing
    Exit Sub
    
ManejadorErrores:
    ' Error fatal VB6
    AgregarRegistroHistorial DNI_Consulta, "ERROR FATAL VB6", "ERROR"
    lblEstado.Caption = "ERROR CRÍTICO VB6: " & Err.Description
    lblEstado.ForeColor = COLOR_ERROR
    Resume FinalizarProceso
End Sub


Private Sub AgregarRegistroHistorial(ByVal dni As String, ByVal nombre As String, ByVal estado As String)
    
    Dim nuevo_item As ListItem
    
    ' Inicialización de Columnas (SOLO si es la primera consulta)
    If lsvHistorial.ColumnHeaders.Count = 0 Then
        lsvHistorial.View = 3 ' Asegura Vista de Reporte
        lsvHistorial.ColumnHeaders.Add , , "N°", 400
        lsvHistorial.ColumnHeaders.Add , , "DNI", 900
        lsvHistorial.ColumnHeaders.Add , , "Nombre Completo", 3500
        lsvHistorial.ColumnHeaders.Add , , "Estado", 1000
    End If
    
    Historial_Contador = Historial_Contador + 1
    
    ' Crear nueva fila
    Set nuevo_item = lsvHistorial.ListItems.Add(, , Historial_Contador)
    
    ' Añadir SubItems
    nuevo_item.SubItems(1) = dni
    nuevo_item.SubItems(2) = nombre
    nuevo_item.SubItems(3) = estado
    
    ' Colorear la fila según el estado (Éxito/Error)
    If estado = "EXITO" Then
        nuevo_item.ForeColor = vbBlue
    Else
        nuevo_item.ForeColor = vbRed
    End If
    
End Sub
