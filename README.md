# SGAM – Sistema de Gestión de Asistencias Médicas
### Versión 1.2.0

Aplicación de escritorio monolítica portable en Python que automatiza la conciliación de asistencias de médicos e internos hospitalarios, cruzando una plantilla maestra Excel contra el reporte del lector biométrico de huella digital.

---

## Estructura del Proyecto

```
SGAM/
├── main.py                       ← Punto de entrada y verificación de dependencias
├── ingestion.py                  ← Carga, validación y normalización de archivos
├── core.py                       ← Motor de cruce de datos y algoritmo semáforo
├── ui.py                         ← Interfaz gráfica (CustomTkinter)
├── export.py                     ← Exportación Excel institucional (individual y maestro)
├── utils.py                      ← Estadísticas avanzadas, rankings y tendencias
├── generar_datos_ejemplo.py      ← Genera archivos de prueba con datos ficticios
├── build_exe.py                  ← Empaqueta la app en ejecutable .exe (PyInstaller)
├── requirements.txt              ← Dependencias Python
├── assets/
│   ├── logo_hospital.png         ← Logo institucional (PNG, opcional)
│   └── icono.ico                 ← Ícono de la aplicación (.exe)
├── data/
│   ├── Estructura_Maestra_Hospital2.xlsx
│   └── Reporte_Scanner_Enero2025.xlsx
└── output/                       ← Reportes generados automáticamente
```

---

## Instalación Rápida

```bash
# 1. Clonar o descomprimir el proyecto
cd SGAM/

# 2. Crear entorno virtual (recomendado)
python -m venv venv
venv\Scripts\activate          # Windows
source venv/bin/activate       # Linux / macOS

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Generar datos de prueba (opcional, para verificar que todo funciona)
python generar_datos_ejemplo.py

# 5. Ejecutar
python main.py
```

---

## Archivos de Entrada Requeridos

### 1. Plantilla Maestra — `Estructura_Maestra_Hospital2.xlsx`

| Hoja | Columnas obligatorias | Columnas opcionales |
|------|-----------------------|---------------------|
| `1_Reglas_ID` | Regla de Negocio, Descripción | — |
| `2_Catalogo_Personal` | ID_Biometrico_SIRA, Nombre_Completo, Tipo, Grado, Universidad, Estatus, Vigencia | Especialidad_Base¹, Periodo_Ingreso², Foto_Ruta³ |
| `3_Registro_Incidencias` | ID_Institucional, Tipo_Incidencia, Fecha_Inicio, Fecha_Fin, Destino_o_Servicio, Notas_Motivo | — |
| `4_Rol_Guardias` | Fecha_Guardia, ID_Institucional, Servicio_Cubierto, TIPO | — |

> ¹ Especialidad_Base puede estar vacía para Internos.  
> ² Periodo_Ingreso acepta `A` (Primavera) o `B` (Otoño).  
> ³ Foto_Ruta: ruta relativa o absoluta a JPG/PNG del médico (se muestra en el tablero).

### 2. Reporte del Escáner Biométrico

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `ID_Biometrico` | Texto | Debe coincidir con `ID_Biometrico_SIRA` del catálogo |
| `Fecha` | Fecha | Formato YYYY-MM-DD |
| `Hora_CheckIn` | Hora | Formato HH:MM:SS |
| `Hora_CheckOut` | Hora | Formato HH:MM:SS |

Formatos aceptados: `.xlsx`, `.xls`, `.csv`

---

## Reglas de Negocio

### Catálogo de Personal

| Tipo | Regla |
|------|-------|
| **Interno** | `Especialidad_Base` puede quedar vacía — no genera error |
| **Residente** | `Grado` debe seguir el patrón `R1`, `R2`, `R3`… |
| **Período de Ingreso** | `A` = Primavera · `B` = Otoño |

### Registro de Incidencias

| Tipo | Comportamiento |
|------|---------------|
| Vacaciones, Incapacidad, Permiso, Comisión | Cubren el día sin evaluar asistencia |
| **Rotación** | Requiere `Destino_o_Servicio` obligatorio; genera alerta si está vacío |
| Cualquier tipo | `Notas_Motivo` se propaga al tooltip del calendario y al Excel |

### Rol de Guardias — Tipos de Turno

| Tipo | Horario por defecto | Configurable en `1_Reglas_ID` |
|------|---------------------|-------------------------------|
| **A** — Matutino | 08:00 – 15:00 | Sí, con "Horario Guardia A: HH:MM - HH:MM" |
| **B** — Vespertino | 15:00 – 21:00 | Sí |
| **C** — Nocturno | 21:00 – 08:00 | Sí |

Para sobreescribir un horario, añadir una fila en `1_Reglas_ID`:
- **Regla de Negocio:** `Horario Guardia A`
- **Descripción:** `08:00 - 14:30`

---

## Algoritmo Semáforo — Lógica de Prioridades

Por cada empleado y cada día del mes, el sistema aplica en orden:

```
1. ¿Existe incidencia registrada?
   → Asignar color según tipo (Vacaciones, Rotación, etc.) y detener.

2. ¿No tiene turno en Rol de Guardias?
   → ⚪ No Laborable (gris claro)

3. ¿Tiene turno programado? → Evaluar escáner biométrico:
   a. Sin marca de entrada → 🟠 Falta
   b. CheckIn ≤ hora_turno + tolerancia → 🟡 Asistencia
   c. CheckIn > hora_turno + tolerancia → 🔴 Retardo
```

**Casos especiales:**
- **Doble turno el mismo día:** se acepta sin error; se usa el CheckIn más temprano para evaluar puntualidad.
- **Turno nocturno (Tipo C):** si el CheckOut es antes de las 06:00, el sistema lo propaga al día siguiente para evitar falsos "Falta".

### Paleta de colores

| Color | Estatus | Hex |
|-------|---------|-----|
| 🟡 Amarillo | Asistencia | `#FFD966` |
| 🔴 Rojo | Retardo | `#FF4B4B` |
| 🟠 Naranja | Falta | `#FF8C00` |
| ⚪ Gris | No Laborable | `#D9D9D9` |
| 🟢 Verde | Vacaciones | `#70AD47` |
| 🔵 Azul claro | Incapacidad | `#9DC3E6` |
| 💛 Amarillo pálido | Permiso | `#FFE699` |
| 🔵 Celeste | Comisión | `#BDD7EE` |
| 🟤 Ámbar | Rotación | `#F4B942` |

---

## Módulos del Sistema

### `ingestion.py`
- Lee y valida las 4 hojas del archivo maestro y el reporte del escáner.
- Normaliza tipos de datos, fechas y horas.
- Genera lista de **alertas no-bloqueantes** (rotaciones sin destino, grados no estándar, periodos inválidos).
- Parsea horarios de turno A/B/C desde `1_Reglas_ID`.

### `core.py`
- Motor principal: cruza catálogo × incidencias × guardias × escáner.
- Aplica la lógica de prioridades y asigna estatus + color por día.
- Propaga `Notas_Motivo` y `Destino_o_Servicio` de incidencias a cada registro.
- Funciones de filtrado: `filtrar_por_empleado()`, `filtrar_por_tipo()`, `filtrar_por_especialidad()`.

### `export.py`
- `exportar_reporte_empleado()` — reporte individual con calendario coloreado + espacio de firma.
- `exportar_todos()` — un archivo por cada empleado.
- `exportar_filtrado()` — igual con filtros por Tipo o Especialidad.
- `exportar_maestro_consolidado()` — **hoja única compacta**: una fila por médico, columnas 1–31 coloreadas, orientación landscape, optimizado para impresión.

### `utils.py`
- `estadisticas_por_empleado()` — métricas individuales completas.
- `ranking_asistencia()` — top N empleados por % de asistencia.
- `resumen_por_tipo()` — agregado Interno / Residente / Médico Adscrito.
- `resumen_por_especialidad()` — agregado por especialidad.
- `calcular_tendencia_semanal()` — distribución semana por semana.
- `dias_criticos()` — días con mayor número de faltas + retardos.
- `exportar_estadisticas_excel()` — Excel con 5 hojas de análisis.

### `ui.py`
- Sidebar con secciones: Archivos de Entrada, Procesamiento, Exportación, Estadísticas.
- Barra de búsqueda + ComboBox de empleados.
- **Tablero Individual:** foto del médico, nombre, Tipo · Especialidad, calendario mensual coloreado + leyenda.
- **Panel de Analítica:** gráfica de pastel con filtros dinámicos por Tipo y Especialidad + botón Vista Individual.
- **Ventana de Estadísticas Avanzadas:** 5 pestañas con tablas interactivas y mini gráfica de tendencia.
- Barra de estado inferior con contadores en tiempo real.
- Alertas de validación mostradas al cargar la plantilla maestra.

---

## Nombres de Archivos Generados

```
# Reporte individual:
SGAM_Reporte_[Nombre_Empleado]_[Mes][Año].xlsx

# Maestro consolidado:
SGAM_Maestro_[Mes][Año].xlsx
SGAM_Maestro_Residente_[Mes][Año].xlsx     ← con filtro de tipo
SGAM_Maestro_Urgencias_[Mes][Año].xlsx     ← con filtro de especialidad

# Estadísticas avanzadas:
SGAM_Estadisticas_[Mes][Año].xlsx
```

---

## Generar Ejecutable `.exe`

```bash
pip install pyinstaller
python build_exe.py
# Genera: dist/SGAM.exe  (portable, sin instalar Python)
```

---

## Requerimientos del Sistema

| Requisito | Mínimo |
|-----------|--------|
| Python | 3.11+ |
| Sistema operativo | Windows 10/11 (UI gráfica) |
| RAM | 4 GB recomendado |
| Resolución | 1280 × 768 mínimo |
| Dependencias | pandas · openpyxl · customtkinter · matplotlib · Pillow · pyinstaller |

---

## Historial de Versiones

| Versión | Cambios principales |
|---------|---------------------|
| **v1.2.0** | Módulo `utils.py` con estadísticas avanzadas · Ventana de estadísticas en UI · ROTACION como estatus propio · Notas de incidencia propagadas · Reporte maestro en hoja única compacta · Horarios A/B/C configurables · Barra de estado con contadores en vivo |
| **v1.1.0** | Dobles turnos · Turnos nocturnos · Perfil con Tipo y Especialidad · Exportación con filtros · Reporte maestro multi-hoja |
| **v1.0.0** | Versión inicial: carga, cruce, semáforo, calendario, exportación individual |

---

## Mejoras Futuras

1. **Base de datos SQLite** — persistencia entre sesiones y búsqueda de históricos
2. **Módulo de login** — control de acceso RH vs consulta general
3. **Exportación PDF** — reportes con firma digital integrada
4. **Notificaciones por correo** — envío automático al generar reportes (smtplib)
5. **Vigilancia de directorio** — importación automática al detectar nuevo archivo del escáner (watchdog)
6. **Comparativo mensual** — gráficas de tendencia entre meses
7. **Módulo de auditoría** — trazabilidad de cambios con usuario y timestamp
8. **Soporte multi-hospital** — separación por unidad o departamento
