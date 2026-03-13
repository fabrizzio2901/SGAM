# SGAM – Sistema de Gestión de Asistencias Médicas

## Descripción
Aplicación de escritorio monolítica portable que automatiza la conciliación de asistencias de médicos e internos hospitalarios, cruzando una plantilla maestra Excel con reportes del lector biométrico de huella digital.

---

## Estructura del Proyecto

```
SGAM/
├── main.py                      ← Punto de entrada
├── ingestion.py                 ← Carga y validación de archivos
├── core.py                      ← Motor de cruce de datos
├── ui.py                        ← Interfaz gráfica (CustomTkinter)
├── export.py                    ← Exportación Excel institucional
├── generar_datos_ejemplo.py     ← Genera datos ficticios para pruebas
├── build_exe.py                 ← Empaqueta en ejecutable .exe
├── requirements.txt             ← Dependencias Python
├── assets/
│   ├── logo_hospital.png        ← Logo institucional (PNG)
│   └── icono.ico                ← Ícono de la aplicación
├── data/
│   ├── Estructura_Maestra_Hospital2.xlsx
│   └── Reporte_Scanner_Ejemplo.xlsx
└── output/                      ← Reportes generados
```

---

## Instalación

```bash
# 1. Crear entorno virtual (recomendado)
python -m venv venv
venv\Scripts\activate         # Windows
source venv/bin/activate      # Linux/Mac

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Generar datos de prueba (opcional)
python generar_datos_ejemplo.py

# 4. Ejecutar la aplicación
python main.py
```

---

## Archivos Requeridos

### Archivo Maestro: `Estructura_Maestra_Hospital2.xlsx`

| Hoja | Columnas clave |
|------|---------------|
| `1_Reglas_ID` | Regla de Negocio, Descripción |
| `2_Catalogo_Personal` | ID_Biometrico_SIRA, Nombre_Completo, Tipo, Especialidad_Base, Estatus |
| `3_Registro_Incidencias` | ID_Institucional, Tipo_Incidencia, Fecha_Inicio, Fecha_Fin |
| `4_Rol_Guardias` | Fecha_Guardia, ID_Institucional, Servicio_Cubierto, TIPO |

### Reporte Escáner Biométrico

| Columna | Descripción |
|---------|-------------|
| `ID_Biometrico` | Debe coincidir con `ID_Biometrico_SIRA` del catálogo |
| `Fecha` | Fecha del registro (YYYY-MM-DD) |
| `Hora_CheckIn` | Hora de entrada (HH:MM:SS) |
| `Hora_CheckOut` | Hora de salida (HH:MM:SS) |

---

## Lógica de Semáforo

| Color | Estatus | Condición |
|-------|---------|-----------|
| 🟡 Amarillo | Asistencia | Check-in dentro de tolerancia |
| 🔴 Rojo | Retardo | Check-in fuera de tolerancia |
| 🟠 Naranja | Falta | Sin registro biométrico en día laborable |
| ⚪ Gris | No Laborable | Sin turno en Rol de Guardias |
| 🟢 Verde | Vacaciones | Cubierto por incidencia vacacional |
| 🔵 Azul | Incapacidad | Cubierto por incidencia médica |
| 💛 Amarillo claro | Permiso | Cubierto por permiso |
| 🔵 Celeste | Comisión | Cubierto por comisión |

### Prioridad del algoritmo:
1. **Incidencia** → Asignar color y detener evaluación
2. **Sin turno** en Rol_Guardias → No Laborable (gris)
3. **Con turno** → Evaluar escáner:
   - Sin marca → Falta
   - CheckIn ≤ hora_esperada + tolerancia → Asistencia
   - CheckIn > hora_esperada + tolerancia → Retardo

---

## Generar Ejecutable (.exe)

```bash
# Instalar PyInstaller
pip install pyinstaller

# Ejecutar script de build
python build_exe.py

# El .exe se genera en: dist/SGAM.exe
```

---

## Nombre de Archivos Exportados

```
SGAM_Reporte_[NombreEmpleado]_[Mes][Año].xlsx
```

Ejemplo: `SGAM_Reporte_Dra._Laura_Mendoza_Torres_Enero2025.xlsx`

---

## Requerimientos del Sistema

- Python 3.11+
- Windows 10/11 (interfaz gráfica con CustomTkinter)
- 4 GB RAM mínimo recomendado
- Resolución mínima: 1280 × 768

---

## Mejoras Futuras Sugeridas

1. **Base de datos SQLite** para persistencia entre sesiones
2. **Dashboard web** con Flask/FastAPI + exportación PDF
3. **Notificaciones por correo** al generar reportes (smtplib)
4. **Importación automática** desde directorio vigilado (watchdog)
5. **Histórico de meses** con comparativos de tendencias
6. **Control de acceso** (login RH vs consulta general)
7. **Firma digital** integrada en reportes PDF
8. **Módulo de auditoría** de cambios con trazabilidad
