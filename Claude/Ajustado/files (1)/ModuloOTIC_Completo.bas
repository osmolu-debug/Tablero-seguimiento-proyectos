Attribute VB_Name = "ModuloOTIC"
'============================================================
' PORTAFOLIO OTIC — Macro de Exportacion a JSON v3
' Variables declaradas fuera del loop para evitar desbordamiento
'============================================================
Sub ExportarDashboard()
    Dim wsData As Worksheet
    Dim wsTareas As Worksheet
    Dim jsonStr As String
    Dim filePath As String
    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim firstProj As Boolean

    ' Variables de proyecto (declaradas fuera del loop)
    Dim pid As String
    Dim pNombre As String, pSGC As String, pModalidad As String
    Dim pObjeto As String, pEstructurador As String, pAbogado As String
    Dim pEvaluadores As String, pSupervisor As String, pApoyo As String
    Dim pEstado As String
    Dim pValor As Double
    Dim fAbogado As String, fMesa As String, fRadic As String
    Dim dAbogado As String, dMesa As String, dRadic As String

    ' Variables de tareas (declaradas fuera del loop)
    Dim tareasJson As String, tareaArr As String
    Dim firstTarea As Boolean
    Dim tNombre As String, tFase As String, tInicio As String
    Dim tFin As String, tEstado As String
    Dim tDias As String, tAvance As String
    Dim progGen As String

    filePath = ThisWorkbook.Path
    If Right(filePath, 1) <> Chr(92) Then filePath = filePath & Chr(92)
    filePath = filePath & "datos_otic.json"
    On Error GoTo ErrorHandler

    Set wsData = ThisWorkbook.Sheets("Data Original")
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    jsonStr = "{" & vbNewLine
    jsonStr = jsonStr & """fecha_exportacion"": """ & Format(Now(), "YYYY-MM-DD HH:MM:SS") & """," & vbNewLine
    jsonStr = jsonStr & """proyectos"": [" & vbNewLine

    firstProj = True

    For i = 2 To lastRow
        Dim cellVal As String
        cellVal = Trim(CStr(wsData.Cells(i, 1).Value))
        If cellVal <> "" And Left(cellVal, 1) = "P" And IsNumeric(Mid(cellVal, 2)) Then

            pid = Trim(CStr(wsData.Cells(i, 1).Value))

            If Not firstProj Then
                jsonStr = jsonStr & "," & vbNewLine
            End If
            firstProj = False

            ' Campos de texto
            pNombre      = LimpiarTexto(CStr(wsData.Cells(i, 3).Value))
            pSGC         = LimpiarTexto(CStr(wsData.Cells(i, 4).Value))
            pModalidad   = LimpiarTexto(CStr(wsData.Cells(i, 5).Value))
            pObjeto      = LimpiarTexto(CStr(wsData.Cells(i, 6).Value))
            pEstructurador = LimpiarTexto(CStr(wsData.Cells(i, 20).Value))
            pAbogado     = LimpiarTexto(CStr(wsData.Cells(i, 21).Value))
            pEvaluadores = LimpiarTexto(CStr(wsData.Cells(i, 22).Value))
            pSupervisor  = LimpiarTexto(CStr(wsData.Cells(i, 23).Value))
            pApoyo       = LimpiarTexto(CStr(wsData.Cells(i, 24).Value))
            pEstado      = LimpiarTexto(CStr(wsData.Cells(i, 2).Value))

            ' Valor como Double (puede ser mayor a 2 billones)
            pValor = 0
            If IsNumeric(wsData.Cells(i, 9).Value) Then
                pValor = CDbl(wsData.Cells(i, 9).Value)
            End If

            ' Fechas y dias como String para evitar overflow
            fAbogado = FormatFecha(wsData.Cells(i, 10).Value)
            dAbogado = ValorStr(wsData.Cells(i, 11).Value)

            fMesa    = FormatFecha(wsData.Cells(i, 13).Value)
            dMesa    = ValorStr(wsData.Cells(i, 14).Value)

            fRadic   = FormatFecha(wsData.Cells(i, 16).Value)
            dRadic   = ValorStr(wsData.Cells(i, 17).Value)

            ' Construir JSON del proyecto
            jsonStr = jsonStr & "  {" & vbNewLine
            jsonStr = jsonStr & "    ""id"": """ & pid & """," & vbNewLine
            jsonStr = jsonStr & "    ""estado"": """ & pEstado & """," & vbNewLine
            jsonStr = jsonStr & "    ""nombre"": """ & pNombre & """," & vbNewLine
            jsonStr = jsonStr & "    ""sgc"": """ & pSGC & """," & vbNewLine
            jsonStr = jsonStr & "    ""modalidad"": """ & pModalidad & """," & vbNewLine
            jsonStr = jsonStr & "    ""objeto"": """ & pObjeto & """," & vbNewLine
            jsonStr = jsonStr & "    ""valor"": " & Format(pValor, "0") & "," & vbNewLine
            jsonStr = jsonStr & "    ""supervisor"": """ & pSupervisor & """," & vbNewLine
            jsonStr = jsonStr & "    ""estructurador"": """ & pEstructurador & """," & vbNewLine
            jsonStr = jsonStr & "    ""abogado"": """ & pAbogado & """," & vbNewLine
            jsonStr = jsonStr & "    ""evaluadores"": """ & pEvaluadores & """," & vbNewLine
            jsonStr = jsonStr & "    ""fecha_abogado"": """ & fAbogado & """," & vbNewLine
            jsonStr = jsonStr & "    ""dias_abogado"": " & dAbogado & "," & vbNewLine
            jsonStr = jsonStr & "    ""fecha_mesa"": """ & fMesa & """," & vbNewLine
            jsonStr = jsonStr & "    ""dias_mesa"": " & dMesa & "," & vbNewLine
            jsonStr = jsonStr & "    ""fecha_radicacion"": """ & fRadic & """," & vbNewLine
            jsonStr = jsonStr & "    ""dias_radicacion"": " & dRadic & "," & vbNewLine

            ' Tareas del proyecto (hoja Pxx)
            tareasJson = "[]"
            tareaArr   = ""

            Set wsTareas = Nothing
            On Error Resume Next
            Set wsTareas = ThisWorkbook.Sheets(pid)
            On Error GoTo ErrorHandler

            If Not wsTareas Is Nothing Then
                tareaArr  = "["
                firstTarea = True

                For j = 4 To 18
                    tNombre = Trim(CStr(wsTareas.Cells(j, 4).Value))
                    If tNombre <> "" Then
                        If Not firstTarea Then tareaArr = tareaArr & ","
                        firstTarea = False

                        tFase   = LimpiarTexto(CStr(wsTareas.Cells(j, 3).Value))
                        tInicio = FormatFecha(wsTareas.Cells(j, 5).Value)
                        tFin    = FormatFecha(wsTareas.Cells(j, 6).Value)
                        tDias   = ValorStr(wsTareas.Cells(j, 7).Value)
                        tAvance = ValorStr(wsTareas.Cells(j, 8).Value)
                        tEstado = LimpiarTexto(CStr(wsTareas.Cells(j, 9).Value))

                        tareaArr = tareaArr & vbNewLine & "      {"
                        tareaArr = tareaArr & """num"": " & (j - 3) & ","
                        tareaArr = tareaArr & """fase"": """ & tFase & ""","
                        tareaArr = tareaArr & """nombre"": """ & LimpiarTexto(tNombre) & ""","
                        tareaArr = tareaArr & """inicio"": """ & tInicio & ""","
                        tareaArr = tareaArr & """fin"": """ & tFin & ""","
                        tareaArr = tareaArr & """dias"": " & tDias & ","
                        tareaArr = tareaArr & """avance"": " & tAvance & ","
                        tareaArr = tareaArr & """estado"": """ & tEstado & """"
                        tareaArr = tareaArr & "}"
                    End If
                Next j
                tareasJson = tareaArr & "]"

                ' Progreso general desde celda H19
                progGen = "0"
                On Error Resume Next
                If IsNumeric(wsTareas.Cells(19, 8).Value) Then
                    progGen = CStr(CLng(wsTareas.Cells(19, 8).Value * 100))
                End If
                On Error GoTo ErrorHandler

                jsonStr = jsonStr & "    ""progreso_general"": " & progGen & "," & vbNewLine
            Else
                jsonStr = jsonStr & "    ""progreso_general"": 0," & vbNewLine
            End If

            jsonStr = jsonStr & "    ""tareas"": " & tareasJson & vbNewLine
            jsonStr = jsonStr & "  }"

            Set wsTareas = Nothing
        End If
    Next i

    jsonStr = jsonStr & vbNewLine & "]" & vbNewLine & "}"

    ' Escribir archivo
    Dim fileNum As Integer
    fileNum = FreeFile()
    If ThisWorkbook.Path = "" Then
        MsgBox "Guarde el archivo Excel primero antes de exportar.", vbExclamation
        Exit Sub
    End If
    Open filePath For Output As #fileNum
    Print #fileNum, jsonStr
    Close #fileNum

    MsgBox "Dashboard exportado correctamente!" & vbNewLine & vbNewLine & _
           "Archivo: datos_otic.json" & vbNewLine & _
           "Ubicacion: " & filePath & vbNewLine & vbNewLine & _
           "Abra dashboard_otic.html en Chrome/Edge y cargue el JSON.", _
           vbInformation, "Exportacion Exitosa"
    Exit Sub

ErrorHandler:
    MsgBox "Error al exportar: " & Err.Description & vbNewLine & _
           "En fila: " & i & ", columna: " & j, vbCritical, "Error de Exportacion"
End Sub

' ============================================================
' FUNCIONES AUXILIARES
' ============================================================
Function LimpiarTexto(texto As String) As String
    Dim r As String
    r = texto
    r = Replace(r, Chr(10), ", ")
    r = Replace(r, Chr(13), "")
    r = Replace(r, """", "'")
    r = Replace(r, "\", "/")
    LimpiarTexto = Trim(r)
End Function
Function FormatFecha(valor As Variant) As String
    If IsDate(valor) And Not IsEmpty(valor) And Not IsNull(valor) Then
        FormatFecha = Format(CDate(valor), "YYYY-MM-DD")
    Else
        FormatFecha = ""
    End If
End Function
Function ValorLong(valor As Variant) As Long
    If IsNumeric(valor) And Not IsEmpty(valor) Then
        ValorLong = CLng(valor)
    Else
        ValorLong = 0
    End If
End Function

Function ValorStr(valor As Variant) As String
    ' Convierte un valor numerico a String seguro (evita overflow)
    If IsEmpty(valor) Or IsNull(valor) Then
        ValorStr = "0"
    ElseIf IsNumeric(valor) Then
        On Error Resume Next
        ValorStr = CStr(CLng(CDbl(valor)))
        If Err.Number <> 0 Then ValorStr = "0"
        On Error GoTo 0
    Else
        ValorStr = "0"
    End If
End Function



' ============================================================
' MACRO: ImportarTareasDesdeJSON  v3 — robusta y simple
' ============================================================
Sub ImportarTareasDesdeJSON()

    ' --- Seleccionar archivo JSON via dialogo ---
    Dim filePath As String
    filePath = Application.GetOpenFilename( _
        FileFilter:="Archivos JSON (*.json), *.json", _
        Title:="Seleccionar archivo JSON exportado desde el dashboard")
    
    If filePath = "False" Or filePath = "" Then
        MsgBox "Operacion cancelada.", vbInformation
        Exit Sub
    End If
    
    ' --- Leer el archivo linea por linea ---
    Dim fileNum As Integer
    Dim jsonRaw As String
    Dim linea As String
    
    On Error GoTo ErrLeer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, linea
        jsonRaw = jsonRaw & linea & Chr(10)
    Loop
    Close #fileNum
    On Error GoTo 0
    
    If Len(jsonRaw) < 20 Then
        MsgBox "El archivo JSON esta vacio o no es valido.", vbCritical
        Exit Sub
    End If
    
    ' --- Procesar cada hoja de proyecto ---
    Dim ws As Worksheet
    Dim pid As String
    Dim updated As Integer
    updated = 0
    
    For Each ws In ThisWorkbook.Worksheets
        pid = ws.Name
        
        ' Solo hojas P1, P2, P3 ... P99
        If Left(pid, 1) = "P" And IsNumeric(Mid(pid, 2)) Then
            
            ' Buscar el bloque de este proyecto en el JSON
            Dim buscarId As String
            buscarId = """id"": """ & pid & """"
            
            Dim posId As Long
            posId = InStr(1, jsonRaw, buscarId)
            
            If posId > 0 Then
                ' Encontrar donde empieza el array de tareas
                Dim posTareas As Long
                posTareas = InStr(posId, jsonRaw, """tareas"":")
                
                ' Asegurar que estamos dentro del bloque de este proyecto
                Dim posNextId As Long
                posNextId = InStr(posId + 10, jsonRaw, """id"": ""P")
                If posNextId = 0 Then posNextId = Len(jsonRaw)
                
                If posTareas > 0 And posTareas < posNextId Then
                    
                    ' Encontrar el cierre del array ]
                    Dim posFinArr As Long
                    posFinArr = InStr(posTareas, jsonRaw, "]")
                    If posFinArr = 0 Then posFinArr = Len(jsonRaw)
                    
                    Dim bloqueTareas As String
                    bloqueTareas = Mid(jsonRaw, posTareas, posFinArr - posTareas + 1)
                    
                    ' Limpiar filas de tareas (4 a 18)
                    Dim r As Integer
                    For r = 4 To 18
                        ws.Cells(r, 3).ClearContents
                        ws.Cells(r, 4).ClearContents
                        ws.Cells(r, 5).ClearContents
                        ws.Cells(r, 6).ClearContents
                        ws.Cells(r, 8).ClearContents
                    Next r
                    
                    ' Escribir tareas
                    Call EscribirTareas(ws, bloqueTareas)
                    updated = updated + 1
                    
                End If
            End If
        End If
    Next ws
    
    If updated = 0 Then
        MsgBox "No se actualizaron hojas." & Chr(10) & _
               "Verifique que el JSON sea el exportado desde el dashboard.", _
               vbExclamation, "Sin cambios"
    Else
        MsgBox "Importacion exitosa." & Chr(10) & Chr(10) & _
               updated & " hojas de proyecto actualizadas." & Chr(10) & _
               "Ejecute ExportarDashboard para regenerar el JSON.", _
               vbInformation, "OK - Importacion completada"
    End If
    Exit Sub

ErrLeer:
    Close #fileNum
    MsgBox "Error al leer el archivo: " & Err.Description, vbCritical
End Sub

' --- Escribe las tareas del bloque JSON en la hoja ---
Private Sub EscribirTareas(ws As Worksheet, bloque As String)
    Dim fila As Integer
    Dim pos As Long
    fila = 4
    pos = 1

    Do While fila <= 18
        ' Buscar inicio del siguiente objeto tarea
        Dim pObj As Long
        pObj = InStr(pos, bloque, "{")
        If pObj = 0 Then Exit Do

        ' Buscar cierre del objeto
        Dim pFin As Long
        pFin = InStr(pObj, bloque, "}")
        If pFin = 0 Then Exit Do

        Dim obj As String
        obj = Mid(bloque, pObj, pFin - pObj + 1)

        ' Extraer campos del objeto
        Dim vFase   As String: vFase   = ExtraerStr(obj, "fase")
        Dim vNombre As String: vNombre = ExtraerStr(obj, "nombre")
        Dim vInicio As String: vInicio = ExtraerStr(obj, "inicio")
        Dim vFin    As String: vFin    = ExtraerStr(obj, "fin")
        Dim vAvance As Long:   vAvance = ExtraerNum(obj, "avance")

        If Len(Trim(vNombre)) > 0 Then
            ws.Cells(fila, 3).Value = vFase
            ws.Cells(fila, 4).Value = vNombre
            If Len(vInicio) = 10 Then
                On Error Resume Next
                ws.Cells(fila, 5).Value = CDate(vInicio)
                ws.Cells(fila, 5).NumberFormat = "DD/MM/YYYY"
                On Error GoTo 0
            End If
            If Len(vFin) = 10 Then
                On Error Resume Next
                ws.Cells(fila, 6).Value = CDate(vFin)
                ws.Cells(fila, 6).NumberFormat = "DD/MM/YYYY"
                On Error GoTo 0
            End If
            ws.Cells(fila, 8).Value = vAvance
            fila = fila + 1
        End If

        pos = pFin + 1
    Loop
End Sub

' --- Extrae valor string: "campo": "valor" ---
Private Function ExtraerStr(obj As String, campo As String) As String
    Dim buscar As String
    buscar = """" & campo & """: """
    Dim p As Long
    p = InStr(1, obj, buscar)
    If p = 0 Then ExtraerStr = "": Exit Function
    p = p + Len(buscar)
    Dim q As Long
    q = InStr(p, obj, """")
    If q = 0 Then ExtraerStr = "": Exit Function
    ExtraerStr = Mid(obj, p, q - p)
End Function

' --- Extrae valor numerico: "campo": 42 ---
Private Function ExtraerNum(obj As String, campo As String) As Long
    Dim buscar As String
    buscar = """" & campo & """: "
    Dim p As Long
    p = InStr(1, obj, buscar)
    If p = 0 Then ExtraerNum = 0: Exit Function
    p = p + Len(buscar)
    Dim resultado As String
    Dim k As Long
    For k = p To p + 15
        If k > Len(obj) Then Exit For
        Dim ch As String
        ch = Mid(obj, k, 1)
        If ch = "," Or ch = "}" Or ch = Chr(10) Or ch = Chr(13) Or ch = " " Then
            resultado = Trim(Mid(obj, p, k - p))
            Exit For
        End If
    Next k
    If IsNumeric(resultado) Then
        ExtraerNum = CLng(resultado)
    Else
        ExtraerNum = 0
    End If
End Function
