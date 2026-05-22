# Tablero Bitácora OTIC — Dashboard de Seguimiento
**Secretaría Distrital de Movilidad — OTIC**  
Versión: `v2026.05.22` | Plataforma: Google Apps Script + HTML

---

## ¿Qué es este proyecto?

Dashboard web de seguimiento contractual para la Oficina TIC (OTIC) de la Secretaría Distrital de Movilidad de Bogotá. Permite hacer seguimiento a contratos **Precontractuales** y **Poscontractuales**, registrar novedades en bitácora, y generar informes de jefatura.

Funciona como una página web hospedada en **Google Apps Script**, conectada a un **Google Sheets** que actúa como base de datos.

---

## Archivos del proyecto

| Archivo | Descripción |
|---|---|
| `Index.html` | Dashboard completo — interfaz web, lógica JavaScript, estilos CSS |
| `Codigo.gs` | Backend en Google Apps Script — conexión a Sheets, control de roles, bitácora |
| `Manual_Tecnico_OTIC_v21.docx` | Manual técnico del sistema |
| `Manual_Usuario_OTIC_v21.docx` | Manual de uso para usuarios finales |
| `OTIC_Seguimiento_Bitacora 2026 (10).xlsx` | Plantilla Excel de referencia (copia local del Sheet) |

---

## Arquitectura

```
Navegador del usuario
      │
      ▼
Index.html  ←──── sirve desde Apps Script (doGet)
      │
      ▼ google.script.run / fetch
Codigo.gs   ←──── procesa acciones (doPost / doGet)
      │
      ▼
Google Sheets (SPREADSHEET_ID: 15J-EubhRIHff8-2KPgTIKnasVhtj_0QPaSrNVZgRH-8)
  ├── Hoja: Precontractual
  ├── Hoja: Poscontractual
  └── Hoja: Bitácora
```

---

## Estructura del Google Sheets

### Hoja: Precontractual
Filas 1–3: encabezados. Datos desde fila 4.  
Columnas principales: ID, ID Proyecto, Nombre, Supervisor, Cargo, Abogado, Estructurador, Fechas hitos (Abogado / Mesa / Radicación), % Avance, Semáforo, Valor CDP, CDP, Fase, Año, Modalidad, Observaciones, F. Terminación, Línea PAA.

### Hoja: Poscontractual
Filas 1–3: encabezados. Datos desde fila 4.  
Columnas principales: Año, # Contrato, Contratista, Objeto, Abogado, Asignación, Estado SECOP, Estado Actual, Meses, F. Terminación, F. Pierde Competencia, F. Cierre, F. Póliza, Urgencia, Valor, Observaciones, F. Adición.

### Hoja: Bitácora
Filas 1–2: encabezados (`BITACORA_HEADER_ROWS = 2`). Datos desde fila 3.  
Columnas: Módulo, Contrato, Nombre, Fecha, Usuario, Novedad, Compromiso.  
**Nota:** Los registros más recientes siempre quedan en la fila 3 (se inserta hacia abajo).

---

## Roles de usuario

Los roles se definen en el objeto `ROLES` del archivo `Codigo.gs`:

| Rol | Descripción | Permisos |
|---|---|---|
| `admin` | Administrador total | Lee y escribe Pre, Pos y Bitácora; edita notas de cualquier usuario |
| `pre` | Gestor Precontractual | Escribe Precontractual; lee Poscontractual; registra notas propias |
| `pos` | Gestor Poscontractual | Escribe Poscontractual; lee Precontractual; registra notas propias |
| `readonly` | Solo lectura | Ve todo; puede registrar, editar y eliminar **sus propias notas** del día |

> Los correos no registrados en `ROLES` reciben rol `readonly` por defecto.

---

## Funcionamiento de la Bitácora

La bitácora es el módulo más crítico del sistema. Cada nota se persiste **de forma individual e inmediata** al servidor, sin esperar al botón "Guardar".

### Flujo de guardado de una nota

```
Usuario escribe novedad → clic "Guardar"
      │
      ▼
addBitaPre() / addBitaPos()  [Index.html]
      │
      ▼
persistBitaAppend(entry, callback)
      │
      ▼
_gsCall('appendBitOne', payload, cb)
      │
      ▼
google.script.run.appendBitOneFromClient(entry)  [Codigo.gs]
      │
      ▼
appendBitOne(entry, email)
  ├── Valida que novedad no esté vacía
  ├── Obtiene LockService (timeout 15s) → previene escrituras simultáneas
  ├── Inserta fila en posición 3 de la hoja Bitácora (más reciente arriba)
  ├── Formato de fecha como texto (@) para evitar conversión automática
  └── Retorna { ok: true, ... }
      │
      ▼
UI actualiza en memoria solo si servidor confirma OK
```

### Editar / Eliminar nota

- Solo aparecen los botones ✏️ y 🗑️ en **notas del día actual** y **del propio usuario** (o admin para cualquier nota).
- El servidor valida el propietario antes de permitir la operación.
- La búsqueda de la nota en el Sheet usa coincidencia exacta por: módulo + contrato + fecha + usuario + novedad.

---

## Módulos del Dashboard

### 📊 Resumen (Precontractual)
Vista principal con tabla de todos los contratos activos. Incluye semáforo de radicación, hitos, y bitácora expandible por contrato.

**Semáforo de radicación:**
| Color | Condición |
|---|---|
| 🔴 Crítico | ≤ 5 días para radicar (o vencido) |
| 🟡 Alerta | 6 – 20 días |
| 🟢 En Tiempo | > 20 días |
| 🏆 Radicado | Hito cumplido |

### 🚦 Semáforos
Resumen visual por estado con detalle de proyectos.

### 📋 Plan de Trabajo
Plan de tareas por proyecto, con fases, fechas y porcentaje de avance.

### 📑 Informe
Informe formal de jefatura con dos secciones:

**4. Seguimiento Procesos (Precontractual)**  
Tabla con columnas: LÍNEA PAA, ABREVIADO, FECHA RADICACIÓN A CONTRATACIÓN, CDP, VALOR CDP, ESTADO  
El campo ESTADO muestra la última nota de bitácora del contrato (o el campo Observaciones si no hay notas).

**5. Seguimiento de Liquidaciones (Poscontractual)**  
Tabla con todos los contratos poscontractuales. Columnas: N°, AÑO, # DE CONTRATO, CONTRATISTA, OBJETO, FECHA PIERDE COMPETENCIA, ESTADO  
Con subtítulo dinámico: *"Liquidaciones y/o cierres con pérdida de competencia al DD/MM/YYYY"*  
Al usar ⚙️ Configurar se pueden seleccionar contratos específicos para el informe.

### 👥 Equipo
Carga de trabajo por persona (Supervisor, Estructurador, Apoyo).

### Poscontractual (módulo)
Gestión de contratos en fase de liquidación con semáforo por urgencia de vencimiento.

---

## Despliegue en Google Apps Script

1. Abrir [script.google.com](https://script.google.com) con la cuenta de la organización.
2. Crear nuevo proyecto o abrir el existente.
3. Pegar el contenido de `Codigo.gs` en el editor de scripts.
4. Crear un archivo HTML llamado `index` y pegar el contenido de `Index.html`.
5. Verificar que `SPREADSHEET_ID` en `Codigo.gs` apunte al Sheet correcto.
6. **Publicar → Implementar como aplicación web:**
   - Ejecutar como: **Yo** (la cuenta propietaria)
   - Quién tiene acceso: **Todos en [dominio]** o los usuarios requeridos
7. Copiar la URL de implementación y distribuirla a los usuarios.
8. Ante cualquier cambio en el código, crear una **nueva implementación** (no editar la existente) para que los usuarios vean los cambios.

---

## Correcciones recientes (v2026.05.22)

### Bug crítico resuelto — Pérdida de notas en bitácora
**Problema:** Con el sistema anterior (`saveBit` masivo), cuando dos usuarios guardaban al mismo tiempo, el último en guardar sobreescribía las notas del otro.  
**Solución:** La bitácora ahora se persiste **fila por fila** con `LockService` para manejo concurrente seguro.

### Otras correcciones
| # | Corrección |
|---|---|
| 1 | Error de conexión ahora muestra mensaje visible en lugar de fallar silenciosamente |
| 2 | Botones ✏️ y 🗑️ solo aparecen en notas propias (no en notas de otros usuarios) |
| 3 | Si la conexión no está lista al guardar, el sistema avisa en lugar de simular un guardado falso |

### Mejoras de interfaz
- **Responsive:** El tablero se adapta a pantallas ≤1200px (tablet) y ≤768px (móvil).
- **Informe:** Secciones 4 y 5 rediseñadas como tablas formales estilo documento Word.

---

## Variables de configuración importantes (`Codigo.gs`)

```javascript
var SPREADSHEET_ID = '15J-EubhRIHff8-2KPgTIKnasVhtj_0QPaSrNVZgRH-8';
var BITACORA_HEADER_ROWS = 2; // Filas 1-2 son encabezado; datos desde fila 3
var ROLES = { 'correo@dominio.gov.co': 'admin', ... };
```

Para agregar un usuario nuevo, añadir su correo al objeto `ROLES` con el rol correspondiente y redesplegar.

---

## Notas de operación

- El botón **💾 Guardar** solo guarda cambios de Pre/Poscontractual. La bitácora se guarda automáticamente al hacer clic en "Guardar" dentro de cada contrato.
- Los usuarios con rol `readonly` **no necesitan** presionar el botón Guardar — sus notas ya quedaron persistidas al momento de registrarlas.
- Si el tablero muestra "Sin datos" al cargar, recargar la página. Si persiste, verificar permisos de la cuenta en el Sheet.
- El informe se puede imprimir con **🖨️ Imprimir / PDF** directo desde el navegador.
