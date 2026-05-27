
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$base = "c:\Users\OSWALDO\OneDrive\Documentos\Trabajo IAs\Tablero Bitacora"
$hoy = "27 de mayo de 2026"
$at  = '@'

# Helpers
function H1($s,$t)  { $s.Style=$s.Document.Styles.Item(-2);  $s.TypeText($t); $s.TypeParagraph() }
function H2($s,$t)  { $s.Style=$s.Document.Styles.Item(-3);  $s.TypeText($t); $s.TypeParagraph() }
function H3($s,$t)  { $s.Style=$s.Document.Styles.Item(-4);  $s.TypeText($t); $s.TypeParagraph() }
function NL($s,$t)  { $s.Style=$s.Document.Styles.Item(-1);  $s.TypeText($t); $s.TypeParagraph() }
function BL($s,$t)  { $s.Style=$s.Document.Styles.Item(-13); $s.TypeText($t); $s.TypeParagraph() }
function NM($s,$t)  { $s.Style=$s.Document.Styles.Item(-14); $s.TypeText($t); $s.TypeParagraph() }
function Skip($s)   { $s.Style=$s.Document.Styles.Item(-1);  $s.TypeParagraph() }

function MakeTable($doc, $sel, $rows, $cols) {
    $sel.Style = $doc.Styles.Item(-1)
    $tbl = $doc.Tables.Add($sel.Range, $rows, $cols)
    try { $tbl.Style = $doc.Styles.Item("Tabla con cuadricula") } catch {
        try { $tbl.Style = $doc.Styles.Item("Table Grid") } catch {}
    }
    $tbl.Borders.InsideLineStyle  = 1
    $tbl.Borders.OutsideLineStyle = 1
    return $tbl
}

function TCell($tbl, $r, $c, $txt, [bool]$bold=$false, [string]$bg="") {
    $cell = $tbl.Cell($r,$c)
    $cell.Range.Text = $txt
    $cell.Range.Font.Bold = $bold
    $cell.Range.Font.Size = 10
    if ($bg -ne "") {
        try { $cell.Shading.BackgroundPatternColor = [int]("0x$bg") } catch {}
    }
}

# ═══════════════════════════════════════════════════════════════
# MANUAL TÉCNICO
# ═══════════════════════════════════════════════════════════════
$docT = $word.Documents.Add()
$s = $word.Selection

# --- Portada ---
$s.Style = $docT.Styles.Item(-63)
$s.TypeText("Manual Técnico")
$s.TypeParagraph()
$s.Style = $docT.Styles.Item(-1)
$s.Font.Size = 16; $s.Font.Bold = $true
$s.TypeText("Tablero Bitácora OTIC")
$s.TypeParagraph()
$s.Font.Size = 12; $s.Font.Bold = $false
$s.TypeText("Secretaría Distrital de Movilidad")
$s.TypeParagraph()
$s.TypeText("Versión: v2026.05.27   |   Fecha: $hoy")
$s.TypeParagraph()
$s.InsertBreak(7)  # wdPageBreak

# --- 1. Descripción General ---
H1 $s "1. Descripción General"
NL $s "El Tablero Bitácora OTIC es un dashboard web de seguimiento contractual para la Oficina TIC (OTIC) de la Secretaría Distrital de Movilidad de Bogotá. Permite gestionar contratos Precontractuales y Poscontractuales, registrar novedades en bitácora con control de concurrencia, y generar informes de jefatura."
Skip $s
NL $s "Plataforma: Google Apps Script (backend) + HTML/CSS/JavaScript (frontend) + Google Sheets (base de datos)."

# --- 2. Arquitectura ---
H1 $s "2. Arquitectura del Sistema"
NL $s "El sistema tiene tres capas:"
BL $s "Capa de presentación: Index.html servido por Apps Script (doGet). Contiene toda la interfaz, lógica de negocio en JavaScript y estilos CSS."
BL $s "Capa de lógica: Codigo.gs ejecutado en Google Apps Script. Maneja autenticación, permisos, lectura y escritura al Sheet."
BL $s "Capa de datos: Google Sheets con tres hojas (Precontractual, Poscontractual, Bitácora)."
Skip $s

$t1 = MakeTable $docT $s 3 2
TCell $t1 1 1 "Componente" $true "C8D8D5"
TCell $t1 1 2 "Descripción" $true "C8D8D5"
TCell $t1 2 1 "Index.html"
TCell $t1 2 2 "Frontend completo: UI, filtros, tablas, gráficas, modales, lógica JavaScript"
TCell $t1 3 1 "Codigo.gs"
TCell $t1 3 2 "Backend Apps Script: autenticación, roles, CRUD sobre Google Sheets"
$s.MoveDown(5, 1)  # mover después de la tabla
$s.TypeParagraph()

# --- 3. Archivos ---
H1 $s "3. Archivos del Proyecto"
$t2 = MakeTable $docT $s 6 3
TCell $t2 1 1 "Archivo" $true "C8D8D5"
TCell $t2 1 2 "Versión" $true "C8D8D5"
TCell $t2 1 3 "Descripción" $true "C8D8D5"
TCell $t2 2 1 "Index.html"; TCell $t2 2 2 "v2026.05.27"; TCell $t2 2 3 "Dashboard principal — interfaz web completa"
TCell $t2 3 1 "Codigo.gs"; TCell $t2 3 2 "v2026.05.27"; TCell $t2 3 3 "Backend Google Apps Script"
TCell $t2 4 1 "Manual_Tecnico_OTIC_v23.docx"; TCell $t2 4 2 "v23"; TCell $t2 4 3 "Este documento"
TCell $t2 5 1 "Manual_Usuario_OTIC_v23.docx"; TCell $t2 5 2 "v23"; TCell $t2 5 3 "Manual para usuarios finales"
TCell $t2 6 1 "OTIC_Seguimiento_Bitacora 2026.xlsx"; TCell $t2 6 2 "—"; TCell $t2 6 3 "Plantilla Excel de referencia (copia local)"
$s.MoveDown(5,1); $s.TypeParagraph()

# --- 4. Google Sheets ---
H1 $s "4. Estructura del Google Sheets"
NL $s "ID del Spreadsheet: 15J-EubhRIHff8-2KPgTIKnasVhtj_0QPaSrNVZgRH-8"
Skip $s

H2 $s "4.1 Hoja Precontractual"
NL $s "Filas 1–3: encabezados. Datos desde fila 4."
$t3 = MakeTable $docT $s 9 2
TCell $t3 1 1 "Columna" $true "C8D8D5"; TCell $t3 1 2 "Descripción" $true "C8D8D5"
TCell $t3 2 1 "ID"; TCell $t3 2 2 "Identificador interno de fila"
TCell $t3 3 1 "ID Proyecto (Línea PAA)"; TCell $t3 3 2 "Código SGC del proyecto (ej: SGC-366)"
TCell $t3 4 1 "Nombre"; TCell $t3 4 2 "Nombre abreviado del contrato"
TCell $t3 5 1 "Supervisor / Estructurador / Abogado"; TCell $t3 5 2 "Responsables del proceso"
TCell $t3 6 1 "Fechas hitos (Abogado/Mesa/Radicación)"; TCell $t3 6 2 "Fecha plan, fecha real y estado de cada hito"
TCell $t3 7 1 "CDP / Valor CDP"; TCell $t3 7 2 "Número de CDP y apropiación en pesos"
TCell $t3 8 1 "Semáforo / % Avance"; TCell $t3 8 2 "Estado calculado: rd, yl, gn, pu"
TCell $t3 9 1 "Observaciones / Línea PAA"; TCell $t3 9 2 "Notas generales y código de línea presupuestal"
$s.MoveDown(5,1); $s.TypeParagraph()

H2 $s "4.2 Hoja Poscontractual"
NL $s "Filas 1–3: encabezados. Datos desde fila 4."
$t4 = MakeTable $docT $s 8 2
TCell $t4 1 1 "Columna" $true "C8D8D5"; TCell $t4 1 2 "Descripción" $true "C8D8D5"
TCell $t4 2 1 "Año / # Contrato"; TCell $t4 2 2 "Año y número del contrato"
TCell $t4 3 1 "Contratista"; TCell $t4 3 2 "Nombre del contratista"
TCell $t4 4 1 "Objeto"; TCell $t4 4 2 "Descripción del objeto del contrato"
TCell $t4 5 1 "Abogado / Asignación"; TCell $t4 5 2 "Abogado OTIC responsable y tipo de asignación"
TCell $t4 6 1 "F. Terminación / F. Pierde Competencia"; TCell $t4 6 2 "Fechas clave de liquidación"
TCell $t4 7 1 "Urgencia"; TCell $t4 7 2 "VENCIDO / Crítico / Alerta / En Tiempo / Cerrado / Pend. SECOP"
TCell $t4 8 1 "Estado SECOP / Estado Actual"; TCell $t4 8 2 "Estado en plataforma y estado interno de gestión"
$s.MoveDown(5,1); $s.TypeParagraph()

H2 $s "4.3 Hoja Bitácora"
NL $s "Filas 1–2: encabezados (BITACORA_HEADER_ROWS = 2). Datos desde fila 3. La fila 3 siempre contiene el registro más reciente."
$t5 = MakeTable $docT $s 8 2
TCell $t5 1 1 "Columna" $true "C8D8D5"; TCell $t5 1 2 "Descripción" $true "C8D8D5"
TCell $t5 2 1 "Módulo"; TCell $t5 2 2 "Precontractual o Poscontractual"
TCell $t5 3 1 "Contrato"; TCell $t5 3 2 "ID del contrato o proyecto"
TCell $t5 4 1 "Nombre"; TCell $t5 4 2 "Nombre del contrato"
TCell $t5 5 1 "Fecha"; TCell $t5 5 2 "Formato texto: D Mmm YYYY (ej: 22 May 2026)"
TCell $t5 6 1 "Usuario"; TCell $t5 6 2 "Correo del usuario que registró la nota"
TCell $t5 7 1 "Novedad"; TCell $t5 7 2 "Texto de la novedad registrada"
TCell $t5 8 1 "Compromiso"; TCell $t5 8 2 "Compromiso asociado (opcional)"
$s.MoveDown(5,1); $s.TypeParagraph()

# --- 5. Roles ---
H1 $s "5. Roles y Permisos"
NL $s "Los roles se definen en el objeto ROLES del archivo Codigo.gs. Correos no registrados reciben rol readonly por defecto."
Skip $s
$t6 = MakeTable $docT $s 5 4
TCell $t6 1 1 "Rol" $true "C8D8D5"; TCell $t6 1 2 "Descripción" $true "C8D8D5"
TCell $t6 1 3 "Escribe Pre/Pos" $true "C8D8D5"; TCell $t6 1 4 "Edita notas ajenas" $true "C8D8D5"
TCell $t6 2 1 "admin"; TCell $t6 2 2 "Administrador total"; TCell $t6 2 3 "Ambas"; TCell $t6 2 4 "Sí"
TCell $t6 3 1 "pre"; TCell $t6 3 2 "Gestor Precontractual"; TCell $t6 3 3 "Solo Pre"; TCell $t6 3 4 "No"
TCell $t6 4 1 "pos"; TCell $t6 4 2 "Gestor Poscontractual"; TCell $t6 4 3 "Solo Pos"; TCell $t6 4 4 "No"
TCell $t6 5 1 "readonly"; TCell $t6 5 2 "Solo lectura + notas propias"; TCell $t6 5 3 "Ninguna"; TCell $t6 5 4 "No"
$s.MoveDown(5,1); $s.TypeParagraph()

# --- 6. Bitácora técnica ---
H1 $s "6. Módulo Bitácora — Detalles Técnicos"
H2 $s "6.1 Flujo de Guardado (appendBitOne)"
NL $s "Cada nota se persiste de forma individual e inmediata. No depende del botón Guardar principal."
NM $s "Usuario escribe novedad y hace clic en Guardar dentro del contrato."
NM $s "addBitaPre() o addBitaPos() construye el objeto entry con módulo, contrato, fecha, usuario, novedad."
NM $s "persistBitaAppend() llama a _gsCall('appendBitOne', ...) que usa google.script.run."
NM $s "appendBitOneFromClient() en el servidor obtiene el email del usuario activo (Session.getActiveUser)."
NM $s "Se adquiere LockService.getDocumentLock() con timeout de 15 segundos para evitar escrituras simultáneas."
NM $s "Se inserta una nueva fila en la posición 3 de la hoja Bitácora (más reciente siempre arriba)."
NM $s "La columna Fecha se formatea como texto (@) para evitar que Sheets lo convierta a fecha automáticamente."
NM $s "Solo si el servidor retorna ok:true, la UI actualiza la vista y muestra el toast de confirmación."
Skip $s

H2 $s "6.2 Búsqueda para Editar/Eliminar (findBitRowIndex_)"
NL $s "Para editar o eliminar una nota, el servidor busca la fila exacta por coincidencia de cinco campos: módulo + contrato + fecha + usuario + novedad. Si dos notas tienen exactamente el mismo contenido, se toma la primera encontrada."
Skip $s

H2 $s "6.3 Permisos de Edición"
BL $s "Los botones Editar y Eliminar solo aparecen en notas del DÍA ACTUAL y del PROPIO USUARIO (a menos que sea admin)."
BL $s "El servidor valida el propietario antes de ejecutar cualquier modificación, independientemente de lo que envíe el cliente."

# --- 7. Funciones principales ---
H1 $s "7. Referencia de Funciones Principales (Codigo.gs)"
$t7 = MakeTable $docT $s 9 3
TCell $t7 1 1 "Función" $true "C8D8D5"; TCell $t7 1 2 "Tipo" $true "C8D8D5"; TCell $t7 1 3 "Descripción" $true "C8D8D5"
TCell $t7 2 1 "doGet(e)"; TCell $t7 2 2 "Entry point"; TCell $t7 2 3 "Sirve el HTML del dashboard o responde acciones GET (read, whoami)"
TCell $t7 3 1 "doPost(e)"; TCell $t7 3 2 "Entry point"; TCell $t7 3 3 "Procesa acciones POST: appendBitOne, editBitOne, deleteBitOne, savePre, savePos, saveAll"
TCell $t7 4 1 "whoAmI()"; TCell $t7 4 2 "Auth"; TCell $t7 4 3 "Retorna el email y rol del usuario activo"
TCell $t7 5 1 "readAll()"; TCell $t7 5 2 "Lectura"; TCell $t7 5 3 "Lee las tres hojas y retorna todo el contenido al cliente"
TCell $t7 6 1 "appendBitOne(entry, email)"; TCell $t7 6 2 "Bitácora"; TCell $t7 6 3 "Inserta una nota en la fila 3 de Bitácora con LockService"
TCell $t7 7 1 "editBitOne(match, update, email, role)"; TCell $t7 7 2 "Bitácora"; TCell $t7 7 3 "Edita una nota existente validando propiedad"
TCell $t7 8 1 "deleteBitOne(match, email, role)"; TCell $t7 8 2 "Bitácora"; TCell $t7 8 3 "Elimina una fila de bitácora validando propiedad"
TCell $t7 9 1 "savePre / savePos / saveAll"; TCell $t7 9 2 "Escritura"; TCell $t7 9 3 "Guarda filas de Pre o Pos (NO toca la bitácora)"
$s.MoveDown(5,1); $s.TypeParagraph()

H2 $s "7.1 Actualización puntual de campo (updatePreField)"
NL $s "Función introducida en v2026.05.27. Permite escribir una sola celda del Sheet Precontractual sin necesidad del ciclo completo de carga y guardado masivo (clearContents + setValues). Esto elimina el riesgo de que un guardado masivo posterior pise los cambios directos al Sheet."
Skip $s
$t7b = MakeTable $docT $s 5 2
TCell $t7b 1 1 "Parámetro / Detalle" $true "C8D8D5"; TCell $t7b 1 2 "Descripción" $true "C8D8D5"
TCell $t7b 2 1 "rowId"; TCell $t7b 2 2 "Identificador del proyecto (ej: P28). Busca la fila donde columna A = rowId."
TCell $t7b 3 1 "field"; TCell $t7b 3 2 "Nombre del campo JS (ej: 'estructurador', 'nombre', 'supervisor'). Validado contra PRE_FIELD_COL."
TCell $t7b 4 1 "value"; TCell $t7b 4 2 "Valor nuevo a escribir en la celda."
TCell $t7b 5 1 "PRE_FIELD_COL"; TCell $t7b 5 2 "Mapa que traduce campo JS a número de columna del Sheet: nombre=3, supervisor=4, cargo=5, abogado=6, estructurador=7, fPlanAbg=8, fPlanMesa=12, fPlanRad=16, valor=22, cdp=23, fase=24, modalidad=28, obs=29, fTermContrato=30, lineaPaa=31."
$s.MoveDown(5,1); $s.TypeParagraph()
NL $s "Función wrapper expuesta al cliente: updatePreFieldFromClient(rowId, field, value). Valida que el rol sea admin o pre antes de delegar a updatePreField."

# --- 8. Despliegue ---
H1 $s "8. Guía de Despliegue en Google Apps Script"
NM $s "Abrir script.google.com con la cuenta del dominio movilidadbogota.gov.co."
NM $s "Abrir el proyecto existente o crear uno nuevo."
NM $s "Pegar el contenido de Codigo.gs en el editor. Verificar que SPREADSHEET_ID apunte al Sheet correcto."
NM $s "Crear o actualizar el archivo HTML llamado exactamente 'index' y pegar el contenido de Index.html."
NM $s "Ir a Implementar > Nueva implementación > Tipo: Aplicación web."
NM $s "Configurar: Ejecutar como la cuenta propietaria. Acceso: Todos en movilidadbogota.gov.co."
NM $s "Copiar la URL de implementación y distribuirla a los usuarios."
NM $s "IMPORTANTE: ante cualquier cambio de código, crear NUEVA implementación (no editar la existente) para que los usuarios vean los cambios al recargar."
Skip $s
H2 $s "8.1 Agregar un nuevo usuario"
NL $s "En Codigo.gs, agregar el correo del usuario al objeto ROLES con su rol correspondiente y redesplegar:"
$s.Style = $docT.Styles.Item(-1)
$s.Font.Name = "Courier New"; $s.Font.Size = 9
$s.TypeText("'nuevousuario${at}movilidadbogota.gov.co' : 'readonly',")
$s.TypeParagraph()
$s.Font.Name = "Calibri"; $s.Font.Size = 11

# --- 9. Changelog ---
H1 $s "9. Historial de Cambios"
$t8 = MakeTable $docT $s 7 3
TCell $t8 1 1 "Versión" $true "C8D8D5"; TCell $t8 1 2 "Fecha" $true "C8D8D5"; TCell $t8 1 3 "Cambios" $true "C8D8D5"
TCell $t8 2 1 "v2026.05.27"; TCell $t8 2 2 "27/05/2026"; TCell $t8 2 3 "Nuevo: updatePreField — escritura puntual celda a celda (evita sobreescritura con guardado masivo). Botón Editar por fila (admin/pre): formulario con 14 campos editables (nombre, supervisor, cargo, abogado, estructurador, fechas plan, valor, CDP, fase, modalidad, línea PAA, f. terminación)."
TCell $t8 3 1 "v2026.05.22"; TCell $t8 3 2 "22/05/2026"; TCell $t8 3 3 "Corrección bug pérdida de notas; responsive; nuevo formato informe; corrección encabezados; mostrar todos los contratos Pos en informe"
TCell $t8 4 1 "v2026.05.20"; TCell $t8 4 2 "20/05/2026"; TCell $t8 4 3 "Bitácora fila por fila (appendBitOne) con LockService — eliminado saveBit masivo"
TCell $t8 5 1 "v2026.05.20-d"; TCell $t8 5 2 "20/05/2026"; TCell $t8 5 3 "Correcciones módulo Análisis IA + renombrado carpeta"
TCell $t8 6 1 "v6.3"; TCell $t8 6 2 "—"; TCell $t8 6 3 "Pruebas completas + manual supervisor"
TCell $t8 7 1 "v1.0"; TCell $t8 7 2 "—"; TCell $t8 7 3 "Primera versión del tablero"
$s.MoveDown(5,1); $s.TypeParagraph()

$pathT = "$base\Manual_Tecnico_OTIC_v23.docx"
$docT.SaveAs2($pathT)
$docT.Close()
Write-Output "Manual Técnico creado: $pathT"

# ═══════════════════════════════════════════════════════════════
# MANUAL DE USUARIO
# ═══════════════════════════════════════════════════════════════
$docU = $word.Documents.Add()
$s = $word.Selection

# Portada
$s.Style = $docU.Styles.Item(-63)
$s.TypeText("Manual de Usuario")
$s.TypeParagraph()
$s.Style = $docU.Styles.Item(-1)
$s.Font.Size = 16; $s.Font.Bold = $true
$s.TypeText("Tablero Bitácora OTIC")
$s.TypeParagraph()
$s.Font.Size = 12; $s.Font.Bold = $false
$s.TypeText("Secretaría Distrital de Movilidad")
$s.TypeParagraph()
$s.TypeText("Versión: v2026.05.27   |   Fecha: $hoy")
$s.TypeParagraph()
$s.InsertBreak(7)

# --- 1. Introducción ---
H1 $s "1. Introducción"
NL $s "El Tablero Bitácora OTIC es la herramienta oficial de seguimiento contractual de la OTIC — Secretaría Distrital de Movilidad. Permite consultar el estado de todos los contratos activos, registrar novedades, revisar semáforos de alerta y generar informes para jefatura."
Skip $s
NL $s "El tablero funciona directamente en el navegador web — no requiere instalar ningún programa."

# --- 2. Acceso ---
H1 $s "2. Cómo Acceder"
NM $s "Abrir el navegador web (Chrome recomendado) con la cuenta del dominio ${at}movilidadbogota.gov.co."
NM $s "Ingresar a la URL del tablero proporcionada por el administrador."
NM $s "El sistema detecta automáticamente el correo y asigna el rol correspondiente."
NM $s "Los datos se cargan automáticamente desde Google Sheets al abrir el tablero."
Skip $s
NL $s "Si aparece el mensaje 'No se pudo conectar a Google Sheets', recargar la página (F5). Si persiste, verificar que se esté usando la cuenta de la organización."

# --- 3. Interfaz ---
H1 $s "3. Interfaz General"
H2 $s "3.1 Barra superior"
$t9 = MakeTable $docU $s 6 2
TCell $t9 1 1 "Elemento" $true "C8D8D5"; TCell $t9 1 2 "Función" $true "C8D8D5"
TCell $t9 2 1 "Precontractual (N)"; TCell $t9 2 2 "Cambia al módulo de contratos en proceso de contratación"
TCell $t9 3 1 "Poscontractual (N)"; TCell $t9 3 2 "Cambia al módulo de contratos en liquidación"
TCell $t9 4 1 "Google Sheets | nombre usuario"; TCell $t9 4 2 "Indica que la conexión al Sheet está activa y muestra el usuario logueado"
TCell $t9 5 1 "Cargar"; TCell $t9 5 2 "Permite cargar un Excel local o reconectarse a Google Sheets"
TCell $t9 6 1 "Guardar"; TCell $t9 6 2 "Guarda los cambios de Pre/Poscontractual en el Sheet (las notas de bitácora ya se guardan al instante)"
$s.MoveDown(5,1); $s.TypeParagraph()

H2 $s "3.2 Pestañas de navegación"
$t10 = MakeTable $docU $s 6 2
TCell $t10 1 1 "Pestaña" $true "C8D8D5"; TCell $t10 1 2 "Contenido" $true "C8D8D5"
TCell $t10 2 1 "Resumen"; TCell $t10 2 2 "Tabla principal de todos los contratos con semáforo y bitácora"
TCell $t10 3 1 "Semáforos"; TCell $t10 3 2 "Vista resumida del estado de cada contrato"
TCell $t10 4 1 "Plan de Trabajo"; TCell $t10 4 2 "Tareas y actividades por proyecto con fechas y avance"
TCell $t10 5 1 "Informe"; TCell $t10 5 2 "Informe de jefatura imprimible con formato de documento oficial"
TCell $t10 6 1 "Equipo"; TCell $t10 6 2 "Carga de trabajo por persona del equipo OTIC"
$s.MoveDown(5,1); $s.TypeParagraph()

# --- 4. Módulo Precontractual ---
H1 $s "4. Módulo Precontractual"
NL $s "Muestra todos los contratos activos en proceso de contratación. Los KPIs superiores muestran la cantidad de contratos por estado de semáforo."
Skip $s

H2 $s "4.1 Semáforo de Radicación"
$t11 = MakeTable $docU $s 5 2
TCell $t11 1 1 "Semáforo" $true "C8D8D5"; TCell $t11 1 2 "Condición" $true "C8D8D5"
TCell $t11 2 1 "Crítico (rojo)"; TCell $t11 2 2 "5 días o menos para radicar (o ya vencido)"
TCell $t11 3 1 "Alerta (amarillo)"; TCell $t11 3 2 "Entre 6 y 20 días para la radicación"
TCell $t11 4 1 "En Tiempo (verde)"; TCell $t11 4 2 "Más de 20 días para radicar"
TCell $t11 5 1 "Radicado (morado)"; TCell $t11 5 2 "Contrato radicado a la DC — hito cumplido"
$s.MoveDown(5,1); $s.TypeParagraph()

H2 $s "4.2 Ver y expandir la Bitácora"
NL $s "Al hacer clic sobre una fila de la tabla, se expande la bitácora del contrato mostrando todas las novedades registradas."

H2 $s "4.3 Filtros disponibles"
BL $s "Estructurador: filtra por la persona responsable de estructurar el contrato."
BL $s "Estado: filtra por color de semáforo (Crítico, Alerta, En Tiempo, Radicado)."
BL $s "Mes Mesa: filtra por el mes de la solicitud de mesa de trabajo."
BL $s "Mes Radicación: filtra por el mes planeado de radicación."

H2 $s "4.4 Editar datos del proyecto (solo admin y pre)"
NL $s "Los usuarios con rol admin (Edgar Romero, Oswaldo Montero, Roger González) y pre (Santiago Vargas) pueden editar los datos de cada proyecto directamente desde el tablero, sin necesidad de modificar el archivo Sheet manualmente."
Skip $s
NL $s "¿Por qué es importante usar esta función y no editar el Sheet directamente?"
NL $s "Cuando alguien edita una celda del Sheet directamente, esa edición puede ser sobreescrita si otro usuario guarda desde el tablero. Con el botón Editar del tablero, el cambio se persiste de forma inmediata celda a celda, garantizando que no sea borrado por guardados posteriores."
Skip $s
NM $s "Pasar el cursor sobre la fila del proyecto en la pestaña Resumen."
NM $s "Hacer clic en el botón 'Editar' (azul claro) que aparece en la columna de acciones al final de la fila."
NM $s "Se abre un formulario con los siguientes campos agrupados:"
BL $s "Nombre del proyecto."
BL $s "Responsables: Supervisor, Apoyo supervisión, Abogado, Estructurador."
BL $s "Fechas plan: F. Plan Abogado, F. Plan Mesa de trabajo, F. Plan Radicación, F. Terminación contrato."
BL $s "Información: Valor, CDP, Fase, Modalidad de selección, Línea PAA."
NM $s "Modificar solo los campos que necesita cambiar. Los campos que queden igual no generan ningún cambio."
NM $s "Hacer clic en 'Guardar cambios'. Aparece la confirmación: 'X campos actualizados en Sheets'."
Skip $s
NL $s "Nota: si no hay ningún cambio respecto al valor actual, el sistema muestra 'Sin cambios para guardar' sin escribir nada."

# --- 5. Módulo Poscontractual ---
H1 $s "5. Módulo Poscontractual (Liquidaciones)"
NL $s "Muestra contratos en etapa de liquidación. Se accede con el botón 'Poscontractual' en la barra superior."
Skip $s
H2 $s "5.1 Urgencia de liquidación"
$t12 = MakeTable $docU $s 7 2
TCell $t12 1 1 "Urgencia" $true "C8D8D5"; TCell $t12 1 2 "Condición" $true "C8D8D5"
TCell $t12 2 1 "VENCIDO"; TCell $t12 2 2 "Ya perdió vigencia"
TCell $t12 3 1 "Crítico"; TCell $t12 3 2 "Menos de 60 días hábiles para vencer"
TCell $t12 4 1 "Alerta"; TCell $t12 4 2 "Entre 60 y 90 días hábiles"
TCell $t12 5 1 "En Tiempo"; TCell $t12 5 2 "Más de 90 días hábiles"
TCell $t12 6 1 "Pend. SECOP"; TCell $t12 6 2 "Con acta de liquidación pendiente en SECOP"
TCell $t12 7 1 "Cerrado"; TCell $t12 7 2 "Liquidación completada"
$s.MoveDown(5,1); $s.TypeParagraph()

# --- 6. Bitácora ---
H1 $s "6. Cómo Usar la Bitácora"
NL $s "La bitácora permite registrar novedades, compromisos y observaciones sobre cada contrato. Cada nota se guarda automáticamente en el momento de crearla — no hay que esperar al botón Guardar principal."
Skip $s

H2 $s "6.1 Registrar una nueva novedad"
NM $s "En la pestaña Resumen o en Poscontractual, hacer clic sobre la fila del contrato para expandir su bitácora."
NM $s "En el campo 'Nueva novedad...', escribir el texto de la novedad."
NM $s "Hacer clic en el botón 'Guardar' junto al campo de texto."
NM $s "Aparecerá el mensaje verde: 'Novedad guardada en Sheets por: [usuario]'."
NM $s "La nota queda registrada con la fecha de hoy y el nombre del usuario logueado."
Skip $s

H2 $s "6.2 Editar o eliminar una nota"
NL $s "Solo es posible editar o eliminar notas registradas HOY y que sean propias. Los administradores pueden editar cualquier nota."
NM $s "Expandir la bitácora del contrato."
NM $s "Hacer clic en el botón Editar (lapiz) para modificar el texto, o en el botón Eliminar (papelera) para borrar."
NM $s "Confirmar la acción en el diálogo que aparece."
Skip $s

H2 $s "6.3 Notas importantes"
BL $s "Si aparece el mensaje 'Conectando a Google Sheets — espera un momento', esperar 5 segundos y volver a intentar."
BL $s "Las notas no se pueden modificar al día siguiente (solo el mismo día que se registraron)."
BL $s "El campo 'Registrado por' muestra el usuario que quedará como autor de la nota."

# --- 7. Informe ---
H1 $s "7. Módulo Informe"
NL $s "Genera un informe formal de jefatura con dos secciones principales, en formato de tabla oficial."
Skip $s

H2 $s "7.1 Sección 4 — Seguimiento Procesos (Precontractual)"
NL $s "Muestra una tabla con todos los contratos precontractuales incluyendo: Línea PAA, Nombre abreviado, Fecha de radicación, CDP, Valor CDP y Estado (última nota de bitácora)."

H2 $s "7.2 Sección 5 — Seguimiento de Liquidaciones (Poscontractual)"
NL $s "Muestra todos los contratos poscontractuales en tabla con: N°, Año, # Contrato, Contratista, Objeto, Fecha pierde competencia y Estado."
NL $s "El subtítulo 'Liquidaciones y/o cierres con pérdida de competencia al DD/MM/YYYY' se calcula automáticamente a 60 días desde la fecha actual."

H2 $s "7.3 Configurar el informe"
NM $s "Hacer clic en el botón Configurar (engranaje) en la parte superior del informe."
NM $s "Seleccionar los contratos específicos que desea incluir en el informe."
NM $s "Hacer clic en Generar. El informe mostrará solo los contratos seleccionados."
NM $s "Para volver a ver todos los contratos, volver a Configurar y seleccionar todos."

H2 $s "7.4 Imprimir o exportar a PDF"
NM $s "Hacer clic en el botón 'Imprimir / PDF' en la parte superior del informe."
NM $s "En el diálogo de impresión del navegador, seleccionar 'Guardar como PDF' o la impresora deseada."
NM $s "Se recomienda orientación horizontal (paisaje) para mejor visualización de las tablas."

# --- 8. Roles ---
H1 $s "8. Roles de Usuario"
$t13 = MakeTable $docU $s 5 3
TCell $t13 1 1 "Rol" $true "C8D8D5"; TCell $t13 1 2 "Puede hacer" $true "C8D8D5"; TCell $t13 1 3 "No puede hacer" $true "C8D8D5"
TCell $t13 2 1 "admin (Edgar, Oswaldo, Roger)"; TCell $t13 2 2 "Todo: leer, editar datos de proyectos, notas de cualquier usuario, formulario Editar proyecto"; TCell $t13 2 3 "—"
TCell $t13 3 1 "pre (Santiago)"; TCell $t13 3 2 "Editar contratos precontractuales, registrar notas propias, formulario Editar proyecto"; TCell $t13 3 3 "Editar datos poscontractuales ni notas ajenas"
TCell $t13 4 1 "pos (Paola)"; TCell $t13 4 2 "Editar contratos poscontractuales, registrar y editar cualquier nota de bitácora"; TCell $t13 4 3 "Editar datos precontractuales ni usar formulario Editar proyecto"
TCell $t13 5 1 "readonly (todos los demás)"; TCell $t13 5 2 "Ver todo el tablero, registrar y editar sus propias notas del día"; TCell $t13 5 3 "Editar datos contractuales, notas de otros usuarios, formulario Editar proyecto"
$s.MoveDown(5,1); $s.TypeParagraph()

# --- 9. Preguntas frecuentes ---
H1 $s "9. Preguntas Frecuentes"
$t14 = MakeTable $docU $s 9 2
TCell $t14 1 1 "Situación" $true "C8D8D5"; TCell $t14 1 2 "Solución" $true "C8D8D5"
TCell $t14 2 1 "El tablero aparece vacío al abrir"; TCell $t14 2 2 "Recargar la página (F5). Verificar que se usa cuenta ${at}movilidadbogota.gov.co"
TCell $t14 3 1 "Aparece 'No se pudo conectar'"; TCell $t14 3 2 "Recargar la página. Si persiste, contactar al administrador"
TCell $t14 4 1 "No aparecen los botones de editar/eliminar en una nota"; TCell $t14 4 2 "Solo aparecen en notas registradas HOY y por el propio usuario. Las notas de otros días o de otros usuarios no se pueden modificar"
TCell $t14 5 1 "Guardé una nota pero no aparece"; TCell $t14 5 2 "Verificar que apareció el mensaje verde de confirmación. Si salió un mensaje rojo, intentar de nuevo"
TCell $t14 6 1 "El botón Guardar no hace nada visible"; TCell $t14 6 2 "Si el rol es readonly, las notas ya se guardan al instante. El botón Guardar principal es solo para administradores y gestores"
TCell $t14 7 1 "Quiero aparecer como autor con mi nombre real, no con el correo"; TCell $t14 7 2 "El sistema muestra la parte del correo antes del ${at}. Contactar al administrador para ajustar el nombre en el sistema"
TCell $t14 8 1 "Cambié un dato en el Sheet (nombre, supervisor, etc.) y al día siguiente volvió al valor anterior"; TCell $t14 8 2 "Ocurre porque el botón Guardar del tablero sobreescribe todo el Sheet con los datos que cargó al inicio de esa sesión. Solución: usar siempre el botón 'Editar' del tablero (admin/pre) para cambiar estos datos — se guarda celda a celda de forma segura."
TCell $t14 9 1 "No veo el botón 'Editar' en las filas de la tabla"; TCell $t14 9 2 "El botón solo es visible para roles admin (Edgar, Oswaldo, Roger) y pre (Santiago). Pasar el cursor sobre la fila para que aparezca."
$s.MoveDown(5,1); $s.TypeParagraph()

$pathU = "$base\Manual_Usuario_OTIC_v23.docx"
$docU.SaveAs2($pathU)
$docU.Close()
Write-Output "Manual de Usuario creado: $pathU"

$word.Quit()
Write-Output "Listo."
