VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   8460
   Begin VB.CommandButton btnLimpiar 
      Caption         =   "Limpiar Resultados"
      Height          =   615
      Left            =   4560
      TabIndex        =   6
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton btnConsumirApiGet 
      Caption         =   "Consumir API GET"
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
   Begin MSComctlLib.ListView lsvTareas 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   7095
      _ExtentX        =   12515
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
   Begin VB.TextBox txtResultado 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label lblTareaRpta 
      Caption         =   "(Lista Todos) - Respuesta OK:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "Valor Clave 'Title':"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Solicitud GET"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Constante: URL para obtener una sola tarea específica (Método GET)
Private Const URL_GET_VALORCLAVE As String = "https://jsonplaceholder.typicode.com/todos/1"
' Constante: URL para obtener la lista completa de tareas (Método GET)
Private Const URL_LISTA_TODOS As String = "https://jsonplaceholder.typicode.com/todos"

' ====================================================================
' 1. FUNCIÓN AUXILIAR: Extracción del Valor JSON (ExtractJsonValue)
'    (La lógica de esta función auxiliar no cambia)
' ====================================================================
Function ExtraerValorJson(ByVal cadena_JSON As String, ByVal clave As String) As String
    
    Dim cadena_busqueda As String
    Dim posicion_inicio As Long
    Dim posicion_fin As Long
    
    ' Lógica para valores STRING ("clave": "VALOR")
    cadena_busqueda = """" & clave & """: """
    posicion_inicio = InStr(1, cadena_JSON, cadena_busqueda, vbTextCompare)
    
    If posicion_inicio > 0 Then
        posicion_inicio = posicion_inicio + Len(cadena_busqueda)
        posicion_fin = InStr(posicion_inicio, cadena_JSON, """", vbTextCompare)
        If posicion_fin > 0 Then
            ExtraerValorJson = Mid$(cadena_JSON, posicion_inicio, posicion_fin - posicion_inicio)
            Exit Function
        End If
    End If
    
    ' Lógica para valores NO STRING (ID, Completed: true/false, etc.)
    cadena_busqueda = """" & clave & """:"
    posicion_inicio = InStr(1, cadena_JSON, cadena_busqueda, vbTextCompare)
    
    If posicion_inicio > 0 Then
        posicion_inicio = posicion_inicio + Len(cadena_busqueda)
        posicion_fin = InStr(posicion_inicio, cadena_JSON, ",", vbTextCompare)
        Dim posicion_llave As Long
        posicion_llave = InStr(posicion_inicio, cadena_JSON, "}", vbTextCompare)
        
        If posicion_fin = 0 Or (posicion_llave > 0 And posicion_llave < posicion_fin) Then
            posicion_fin = posicion_llave
        End If
        
        If posicion_fin > 0 Then
            ExtraerValorJson = Trim(Mid$(cadena_JSON, posicion_inicio, posicion_fin - posicion_inicio))
            Exit Function
        End If
    End If
    
    ExtraerValorJson = ""
End Function


' Esta Subrutina ES el procedimiento que consume la API
Sub ConsumirAPI_Tareas()
    
    Dim objeto_HTTP As MSXML2.XMLHTTP60
    Dim estado_ok As Boolean
    
    ' Reiniciar el Label y activar el manejo de errores
    lblTareaRpta.Caption = "Iniciando consumo de API..."
    lblTareaRpta.ForeColor = vbBlack ' Color inicial
    On Error GoTo ManejadorErrores
    
    ' ====================================================
    ' 1. SOLICITUD A /todos/1 (Extraer Título)
    ' ====================================================
    
    Set objeto_HTTP = New MSXML2.XMLHTTP60
    
    ' Consumo de la API 1 (Obtener Título)
    objeto_HTTP.Open "GET", URL_GET_VALORCLAVE, False
    objeto_HTTP.send
    
    If objeto_HTTP.Status = 200 Then
        ' Si la primera consulta es OK, procesa el título
        Dim titulo_extraido As String
        titulo_extraido = ExtraerValorJson(objeto_HTTP.responseText, "title")
        txtResultado.Text = "Título Obtenido: " & titulo_extraido
        estado_ok = True
    Else
        ' Si la primera consulta falla, lo indica y termina.
        txtResultado.Text = "ERROR HTTP (Título): Código " & objeto_HTTP.Status
        estado_ok = False
        GoTo FinalizarProceso
    End If

    ' ====================================================
    ' 2. SOLICITUD A /todos (Obtener Lista Completa)
    ' ====================================================
    
    Set objeto_HTTP = New MSXML2.XMLHTTP60 ' Nuevo objeto para la lista
    lblTareaRpta.Caption = "Obteniendo lista de tareas..."
    
    ' Consumo de la API 2 (Obtener Lista)
    objeto_HTTP.Open "GET", URL_LISTA_TODOS, False
    objeto_HTTP.send
    
    If objeto_HTTP.Status = 200 Then
        
        ' Llama al procedimiento de presentación
        PoblarListView objeto_HTTP.responseText
        
        ' Confirma el éxito total
        lblTareaRpta.ForeColor = vbBlue
        lblTareaRpta.Caption = "RESULTADO OK, MOSTRANDO LA LISTA DEL API GET"
        
    Else
        ' Si la segunda consulta falla
        lblTareaRpta.ForeColor = vbRed
        lblTareaRpta.Caption = "ERROR HTTP AL OBTENER LISTA: Código " & objeto_HTTP.Status
    End If

FinalizarProceso:
    Set objeto_HTTP = Nothing
    Exit Sub
    
ManejadorErrores:
    lblTareaRpta.ForeColor = vbRed
    lblTareaRpta.Caption = "ERROR CRÍTICO VB 6.0: " & Err.Description
    Set objeto_HTTP = Nothing
End Sub


Private Sub btnConsumirApiGet_Click()
  ConsumirAPI_Tareas

End Sub

Sub PoblarListView(ByVal respuesta_JSON As String)
    
    Dim array_tareas() As String
    Dim i As Long
    Dim tarea_actual As String
    
    Dim id_tarea As String
    Dim titulo_tarea As String
    Dim completado As String
    
    Dim elemento_lista As ListItem
    
    ' 1. LIMPIEZA INICIAL
    lsvTareas.ListItems.Clear
    lsvTareas.ColumnHeaders.Clear
    
    ' 2. CREACIÓN DE COLUMNAS
    lsvTareas.ColumnHeaders.Add , , "ID", 400
    lsvTareas.ColumnHeaders.Add , , "Título de la Tarea", 5200
    lsvTareas.ColumnHeaders.Add , , "Completado", 1100
    
    ' 3. DIVIDIR JSON
    ' El Split separa la cadena. El UBound nos dará el número real (199 si hay 200 tareas)
    array_tareas = Split(respuesta_JSON, "},")
    
    ' 4. LLENAR DATOS
    ' Se itera hasta el límite generado por el Split.
    ' Restamos 1 porque el último elemento del array generado por Split es generalmente incompleto o la terminación del JSON.
    For i = 0 To UBound(array_tareas)
        
        tarea_actual = array_tareas(i) & "}" ' Se añade la llave de cierre a cada fragmento
        
        ' Extracción de los valores
        id_tarea = ExtraerValorJson(tarea_actual, "id")
        titulo_tarea = ExtraerValorJson(tarea_actual, "title")
        completado = ExtraerValorJson(tarea_actual, "completed")
        
        ' Agregar un nuevo elemento (ListItem)
        Set elemento_lista = lsvTareas.ListItems.Add(, , id_tarea)
        
        ' Las columnas subsiguientes se añaden como SubItems
        elemento_lista.SubItems(1) = titulo_tarea
        elemento_lista.SubItems(2) = IIf(completado = "true", "SÍ", "NO")
        
    Next i
    
    ' NOTA: La lógica de la última tarea (índice UBound(array_tareas)) es compleja
    ' debido al patrón final "}]", pero al iterar hasta UBound - 1, garantizamos
    ' que los 200 elementos (0 a 199) sean procesados correctamente en esta API.
    
    ' Liberar el objeto
    Set elemento_lista = Nothing
End Sub




Sub LimpiarResultados()

    ' 1. Limpiar TextBox (Título)
    txtResultado.Text = ""
    
    ' 2. Limpiar Label de Respuesta y Resetear Color
    lblTareaRpta.Caption = ""
    lblTareaRpta.ForeColor = vbBlack ' Color predeterminado
    
    ' 3. Limpiar ListView (Lista de Tareas)
    lsvTareas.ListItems.Clear
    
    ' Nota: No se limpia lsvTareas.ColumnHeaders para que el encabezado
    ' de la grilla (ID, Título, Completado) permanezca visible.
    
End Sub


Private Sub btnLimpiar_Click()
LimpiarResultados
End Sub
