"""
SGAM - Sistema de Gestión de Asistencias Médicas
Módulo: ui.py
Responsabilidad: Interfaz gráfica con CustomTkinter + calendario interactivo.
"""

import threading
import calendar
from datetime import date
from pathlib import Path
from typing import Optional

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from ingestion import cargar_archivo_maestro, cargar_reporte_scanner, extraer_reglas
from core import procesar_asistencias, calcular_resumen, filtrar_por_empleado, COLOR_MAP
from export import exportar_reporte_empleado, exportar_todos


# ──────────────────────────────────────────────
# Configuración visual
# ──────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

APP_TITLE   = "SGAM – Sistema de Gestión de Asistencias Médicas"
APP_VERSION = "v1.0.0"
WIN_WIDTH   = 1280
WIN_HEIGHT  = 800

COLOR_SIDEBAR = "#1F3864"
COLOR_ACCENT  = "#2E75B6"
COLOR_BG      = "#F0F4F8"
COLOR_WHITE   = "#FFFFFF"
COLOR_TEXT    = "#1F1F1F"
COLOR_SUCCESS = "#217346"
COLOR_WARN    = "#C55A11"
COLOR_ERROR   = "#C00000"


# ──────────────────────────────────────────────
# Widget: Calendario Mensual
# ──────────────────────────────────────────────
class CalendarioWidget(ctk.CTkFrame):
    """
    Widget reutilizable que muestra un calendario mensual
    con celdas coloreadas según el estatus de asistencia.
    """

    DIAS_HEADER = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]

    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color=COLOR_WHITE, corner_radius=10, **kwargs)
        self._construir_estructura()

    def _construir_estructura(self):
        self.lbl_titulo = ctk.CTkLabel(
            self, text="Sin datos cargados",
            font=ctk.CTkFont(family="Calibri", size=14, weight="bold"),
            text_color=COLOR_SIDEBAR
        )
        self.lbl_titulo.pack(pady=(12, 6))

        self.grid_frame = ctk.CTkFrame(self, fg_color=COLOR_WHITE)
        self.grid_frame.pack(padx=16, pady=(0, 12))

        for j, dia_nombre in enumerate(self.DIAS_HEADER):
            lbl = ctk.CTkLabel(
                self.grid_frame, text=dia_nombre, width=52, height=28,
                font=ctk.CTkFont(size=10, weight="bold"),
                fg_color=COLOR_SIDEBAR, text_color=COLOR_WHITE,
                corner_radius=4
            )
            lbl.grid(row=0, column=j, padx=2, pady=2)

    def cargar_mes(self, df_empleado: pd.DataFrame, mes: int, anio: int,
                    nombre_empleado: str):
        """Renderiza el calendario para el empleado y mes indicados."""
        # Limpiar celdas anteriores (mantener headers)
        for widget in self.grid_frame.winfo_children():
            info = widget.grid_info()
            if int(info.get("row", 0)) > 0:
                widget.destroy()

        self.lbl_titulo.configure(
            text=f"{nombre_empleado}  |  "
                 f"{self._nombre_mes(mes)} {anio}"
        )

        # Construir mapa {dia: info_estatus}
        mapa_dia = {}
        for _, row in df_empleado.iterrows():
            d = pd.to_datetime(row["Fecha"]).day
            mapa_dia[d] = {
                "label":   row.get("Label", ""),
                "color":   f"#{row.get('Color_Hex', 'FFFFFF')}",
                "notas":   row.get("Notas", ""),
                "estatus": row.get("Estatus", ""),
            }

        primer_dia_semana = calendar.monthrange(anio, mes)[0]  # 0=Lun
        dias_del_mes = calendar.monthrange(anio, mes)[1]

        fila = 1
        col  = primer_dia_semana

        for d in range(1, dias_del_mes + 1):
            info = mapa_dia.get(d, {
                "label": "",
                "color": "#FFFFFF",
                "notas": "",
                "estatus": "SIN_DATOS",
            })

            texto_celda = str(d)
            color_bg    = info["color"]
            tooltip_txt = info["notas"] or info["label"]

            celda = ctk.CTkLabel(
                self.grid_frame,
                text=texto_celda,
                width=52, height=42,
                font=ctk.CTkFont(size=10, weight="bold"),
                fg_color=color_bg,
                text_color="#1F1F1F",
                corner_radius=6,
            )
            celda.grid(row=fila, column=col, padx=2, pady=2)

            # Tooltip al pasar el cursor
            self._bind_tooltip(celda, tooltip_txt)

            col += 1
            if col > 6:
                col = 0
                fila += 1

    @staticmethod
    def _nombre_mes(mes: int) -> str:
        nombres = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                   "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        return nombres[mes] if 1 <= mes <= 12 else str(mes)

    @staticmethod
    def _bind_tooltip(widget, texto: str):
        """Añade mini tooltip al widget."""
        if not texto:
            return
        tip = None

        def mostrar(event):
            nonlocal tip
            tip = tk.Toplevel(widget)
            tip.overrideredirect(True)
            tip.geometry(f"+{event.x_root + 14}+{event.y_root + 14}")
            lbl = tk.Label(tip, text=texto, bg="#FFFFE0",
                           font=("Calibri", 9), relief="solid", borderwidth=1,
                           padx=6, pady=4)
            lbl.pack()

        def ocultar(event):
            nonlocal tip
            if tip:
                tip.destroy()
                tip = None

        widget.bind("<Enter>", mostrar)
        widget.bind("<Leave>", ocultar)


# ──────────────────────────────────────────────
# Ventana Principal
# ──────────────────────────────────────────────
class SGAMApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        # Estado de la aplicación
        self._ruta_maestro:  Optional[str] = None
        self._ruta_scanner:  Optional[str] = None
        self._ruta_logo:     Optional[str] = None
        self._datos:         Optional[dict] = None
        self._df_resultado:  Optional[pd.DataFrame] = None
        self._reglas:        Optional[dict] = None

        self._configurar_ventana()
        self._construir_ui()

    # ── Configuración de ventana ─────────────────────────────────────────
    def _configurar_ventana(self):
        self.title(f"{APP_TITLE}  {APP_VERSION}")
        self.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}")
        self.minsize(1100, 680)
        self.configure(fg_color=COLOR_BG)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

    # ── Construcción de UI ───────────────────────────────────────────────
    def _construir_ui(self):
        self._construir_sidebar()
        self._construir_panel_central()
        self._construir_barra_estado()

    def _construir_sidebar(self):
        """Panel lateral izquierdo con controles."""
        self.sidebar = ctk.CTkFrame(
            self, width=240, corner_radius=0,
            fg_color=COLOR_SIDEBAR
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew", rowspan=2)
        self.sidebar.grid_propagate(False)

        # Logo / Título
        ctk.CTkLabel(
            self.sidebar, text="⚕ SGAM",
            font=ctk.CTkFont(family="Calibri", size=22, weight="bold"),
            text_color=COLOR_WHITE
        ).pack(pady=(24, 2))

        ctk.CTkLabel(
            self.sidebar, text="Gestión de Asistencias\nMédicas",
            font=ctk.CTkFont(size=11),
            text_color="#A8C8E8"
        ).pack(pady=(0, 24))

        # Separador
        ctk.CTkFrame(self.sidebar, height=1, fg_color="#3A5C8C").pack(
            fill="x", padx=16, pady=8
        )

        # ── Sección: Archivos ────────────────────────────────────────────
        self._sidebar_seccion("📂 ARCHIVOS DE ENTRADA")

        self.btn_maestro = self._sidebar_boton(
            "Cargar Plantilla Maestra",
            comando=self._cargar_maestro,
            icono="📋"
        )
        self.lbl_maestro = self._sidebar_estado("Sin cargar")

        self.btn_scanner = self._sidebar_boton(
            "Cargar Reporte Escáner",
            comando=self._cargar_scanner,
            icono="🖱"
        )
        self.lbl_scanner = self._sidebar_estado("Sin cargar")

        self.btn_logo = self._sidebar_boton(
            "Logo Institucional (PNG)",
            comando=self._cargar_logo,
            icono="🖼"
        )
        self.lbl_logo = self._sidebar_estado("Opcional")

        # Separador
        ctk.CTkFrame(self.sidebar, height=1, fg_color="#3A5C8C").pack(
            fill="x", padx=16, pady=12
        )

        # ── Sección: Procesamiento ───────────────────────────────────────
        self._sidebar_seccion("⚙ PROCESAMIENTO")

        self.btn_procesar = self._sidebar_boton(
            "Procesar Asistencias",
            comando=self._procesar,
            icono="▶",
            color_primario="#217346",
            color_hover="#1A5C37",
        )

        # Separador
        ctk.CTkFrame(self.sidebar, height=1, fg_color="#3A5C8C").pack(
            fill="x", padx=16, pady=12
        )

        # ── Sección: Exportación ─────────────────────────────────────────
        self._sidebar_seccion("📤 EXPORTACIÓN")

        self.btn_exportar_uno = self._sidebar_boton(
            "Exportar Empleado Actual",
            comando=self._exportar_empleado_actual,
            icono="📄"
        )
        self.btn_exportar_todos = self._sidebar_boton(
            "Exportar Todos",
            comando=self._exportar_todos,
            icono="📦"
        )

        # Barra de progreso
        self.progress_bar = ctk.CTkProgressBar(self.sidebar, width=200)
        self.progress_bar.pack(padx=20, pady=(8, 2))
        self.progress_bar.set(0)
        self.lbl_progreso = ctk.CTkLabel(
            self.sidebar, text="",
            font=ctk.CTkFont(size=9), text_color="#A8C8E8"
        )
        self.lbl_progreso.pack()

    def _sidebar_seccion(self, texto: str):
        ctk.CTkLabel(
            self.sidebar, text=texto,
            font=ctk.CTkFont(size=9, weight="bold"),
            text_color="#7EB3D8"
        ).pack(anchor="w", padx=18, pady=(8, 2))

    def _sidebar_boton(self, texto: str, comando, icono: str = "",
                        color_primario: str = COLOR_ACCENT,
                        color_hover: str = "#1F5C8A") -> ctk.CTkButton:
        btn = ctk.CTkButton(
            self.sidebar,
            text=f"{icono}  {texto}",
            command=comando,
            fg_color=color_primario,
            hover_color=color_hover,
            text_color=COLOR_WHITE,
            font=ctk.CTkFont(size=11),
            height=36, width=200,
            corner_radius=8,
        )
        btn.pack(padx=20, pady=4)
        return btn

    def _sidebar_estado(self, texto: str) -> ctk.CTkLabel:
        lbl = ctk.CTkLabel(
            self.sidebar, text=f"  {texto}",
            font=ctk.CTkFont(size=9, slant="italic"),
            text_color="#90B8D8",
            anchor="w"
        )
        lbl.pack(padx=20, fill="x")
        return lbl

    # ── Panel Central ────────────────────────────────────────────────────
    def _construir_panel_central(self):
        self.panel_central = ctk.CTkFrame(
            self, corner_radius=0, fg_color=COLOR_BG
        )
        self.panel_central.grid(row=0, column=1, sticky="nsew", padx=0)
        self.panel_central.grid_columnconfigure(0, weight=3)
        self.panel_central.grid_columnconfigure(1, weight=2)
        self.panel_central.grid_rowconfigure(1, weight=1)

        # ── Barra de búsqueda ────────────────────────────────────────────
        barra_busq = ctk.CTkFrame(
            self.panel_central, height=56,
            fg_color=COLOR_WHITE, corner_radius=0
        )
        barra_busq.grid(row=0, column=0, columnspan=2, sticky="ew")

        ctk.CTkLabel(
            barra_busq, text="🔍 Buscar empleado:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=COLOR_TEXT
        ).pack(side="left", padx=(16, 8), pady=14)

        self.entry_buscar = ctk.CTkEntry(
            barra_busq, placeholder_text="ID Biométrico o Nombre...",
            width=340, height=34,
            font=ctk.CTkFont(size=11)
        )
        self.entry_buscar.pack(side="left", padx=4, pady=10)
        self.entry_buscar.bind("<Return>", lambda e: self._buscar())

        ctk.CTkButton(
            barra_busq, text="Buscar", command=self._buscar,
            width=90, height=34, corner_radius=8,
            fg_color=COLOR_ACCENT
        ).pack(side="left", padx=6)

        # ComboBox de empleados
        ctk.CTkLabel(
            barra_busq, text="   |   Empleado:",
            font=ctk.CTkFont(size=11), text_color="#595959"
        ).pack(side="left", padx=(12, 4))

        self.combo_empleados = ctk.CTkComboBox(
            barra_busq, values=["(sin datos)"],
            width=260, height=34,
            command=self._cambiar_empleado_combo
        )
        self.combo_empleados.pack(side="left", padx=4)

        # ── Columna izquierda: Calendario ─────────────────────────────────
        frame_cal = ctk.CTkFrame(
            self.panel_central, fg_color=COLOR_BG, corner_radius=0
        )
        frame_cal.grid(row=1, column=0, sticky="nsew", padx=(12, 6), pady=12)

        ctk.CTkLabel(
            frame_cal, text="📅 Calendario de Asistencia",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLOR_SIDEBAR
        ).pack(anchor="w", pady=(0, 8))

        self.calendario = CalendarioWidget(frame_cal)
        self.calendario.pack(fill="both", expand=True)

        # ── Columna derecha: Gráfica + Stats ─────────────────────────────
        frame_der = ctk.CTkFrame(
            self.panel_central, fg_color=COLOR_BG, corner_radius=0
        )
        frame_der.grid(row=1, column=1, sticky="nsew", padx=(6, 12), pady=12)

        ctk.CTkLabel(
            frame_der, text="📊 Resumen del Período",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLOR_SIDEBAR
        ).pack(anchor="w", pady=(0, 8))

        self.frame_grafica = ctk.CTkFrame(frame_der, fg_color=COLOR_WHITE, corner_radius=10)
        self.frame_grafica.pack(fill="both", expand=True)

        # Placeholder inicial
        self._grafica_placeholder()

    def _construir_barra_estado(self):
        self.barra_estado = ctk.CTkFrame(
            self, height=28, corner_radius=0,
            fg_color="#D6E4F0"
        )
        self.barra_estado.grid(row=1, column=1, sticky="ew")

        self.lbl_estado = ctk.CTkLabel(
            self.barra_estado,
            text="Listo. Cargue los archivos para comenzar.",
            font=ctk.CTkFont(size=10),
            text_color="#1F3864"
        )
        self.lbl_estado.pack(side="left", padx=16)

    # ── Acciones de los botones ──────────────────────────────────────────

    def _cargar_maestro(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar Plantilla Maestra",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not ruta:
            return
        try:
            self._datos = cargar_archivo_maestro(ruta)
            self._reglas = extraer_reglas(self._datos["reglas"])
            self._ruta_maestro = ruta
            nombre = Path(ruta).name
            self.lbl_maestro.configure(
                text=f"  ✅ {nombre[:28]}...",
                text_color="#70AD47"
            )
            self._set_estado(f"Plantilla maestra cargada: {nombre}", "ok")
        except Exception as e:
            messagebox.showerror("Error al cargar plantilla", str(e))
            self._set_estado(f"Error: {e}", "error")

    def _cargar_scanner(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar Reporte del Escáner Biométrico",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")]
        )
        if not ruta:
            return
        try:
            df = cargar_reporte_scanner(ruta)
            if self._datos is None:
                self._datos = {}
            self._datos["scanner"] = df
            self._ruta_scanner = ruta
            nombre = Path(ruta).name
            n_registros = len(df)
            self.lbl_scanner.configure(
                text=f"  ✅ {n_registros} registros",
                text_color="#70AD47"
            )
            self._set_estado(f"Escáner cargado: {n_registros} registros", "ok")
        except Exception as e:
            messagebox.showerror("Error al cargar escáner", str(e))
            self._set_estado(f"Error: {e}", "error")

    def _cargar_logo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar Logo Institucional",
            filetypes=[("PNG Image", "*.png")]
        )
        if ruta:
            self._ruta_logo = ruta
            self.lbl_logo.configure(
                text=f"  ✅ {Path(ruta).name[:28]}",
                text_color="#70AD47"
            )

    def _procesar(self):
        if not self._datos:
            messagebox.showwarning("Datos faltantes", "Cargue la Plantilla Maestra primero.")
            return
        if "scanner" not in (self._datos or {}):
            messagebox.showwarning("Datos faltantes", "Cargue el Reporte del Escáner.")
            return

        self._set_estado("⏳ Procesando asistencias...", "ok")
        self.btn_procesar.configure(state="disabled")

        def _tarea():
            try:
                df = procesar_asistencias(self._datos, self._reglas or {})
                self._df_resultado = df
                self.after(0, self._post_procesamiento)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error de procesamiento", str(e)))
                self.after(0, lambda: self._set_estado(f"Error: {e}", "error"))
            finally:
                self.after(0, lambda: self.btn_procesar.configure(state="normal"))

        threading.Thread(target=_tarea, daemon=True).start()

    def _post_procesamiento(self):
        """Se ejecuta en el hilo principal tras procesar."""
        df = self._df_resultado
        resumen = calcular_resumen(df)
        n_emp = resumen["empleados"]

        # Actualizar combo de empleados
        nombres = sorted(df["Nombre_Completo"].dropna().unique().tolist())
        self.combo_empleados.configure(values=nombres)
        if nombres:
            self.combo_empleados.set(nombres[0])
            self._mostrar_empleado(nombres[0])

        # Mostrar gráfica
        self._mostrar_grafica(resumen)

        self._set_estado(
            f"✅ Procesamiento completo — {n_emp} empleados | "
            f"{resumen['total_dias_laborables']} días laborables evaluados",
            "ok"
        )

    def _buscar(self):
        if self._df_resultado is None:
            return
        termino = self.entry_buscar.get().strip()
        if not termino:
            return
        df_filtrado = filtrar_por_empleado(self._df_resultado, termino)
        if df_filtrado.empty:
            messagebox.showinfo("Sin resultados", f"No se encontraron registros para '{termino}'")
            return
        nombre = df_filtrado.iloc[0]["Nombre_Completo"]
        self._mostrar_empleado(nombre)
        self.combo_empleados.set(nombre)

    def _cambiar_empleado_combo(self, seleccion: str):
        self._mostrar_empleado(seleccion)

    def _mostrar_empleado(self, nombre: str):
        if self._df_resultado is None:
            return
        df_emp = self._df_resultado[
            self._df_resultado["Nombre_Completo"] == nombre
        ].copy()

        if df_emp.empty:
            return

        fechas = pd.to_datetime(df_emp["Fecha"])
        mes  = fechas.dt.month.mode()[0]
        anio = fechas.dt.year.mode()[0]

        df_mes = df_emp[
            (fechas.dt.month == mes) & (fechas.dt.year == anio)
        ]
        self.calendario.cargar_mes(df_mes, mes, anio, nombre)

    def _exportar_empleado_actual(self):
        if self._df_resultado is None:
            messagebox.showwarning("Sin datos", "Primero procese las asistencias.")
            return
        nombre = self.combo_empleados.get()
        if not nombre or nombre == "(sin datos)":
            messagebox.showwarning("Sin empleado", "Seleccione un empleado del desplegable.")
            return

        directorio = filedialog.askdirectory(title="Carpeta de destino para el reporte")
        if not directorio:
            return

        df_emp = self._df_resultado[
            self._df_resultado["Nombre_Completo"] == nombre
        ].copy()

        try:
            ruta = exportar_reporte_empleado(df_emp, directorio, self._ruta_logo)
            messagebox.showinfo(
                "Exportación exitosa",
                f"Reporte generado:\n{ruta}"
            )
            self._set_estado(f"✅ Exportado: {Path(ruta).name}", "ok")
        except Exception as e:
            messagebox.showerror("Error de exportación", str(e))

    def _exportar_todos(self):
        if self._df_resultado is None:
            messagebox.showwarning("Sin datos", "Primero procese las asistencias.")
            return
        directorio = filedialog.askdirectory(title="Carpeta de destino para los reportes")
        if not directorio:
            return

        self.progress_bar.set(0)
        self.btn_exportar_todos.configure(state="disabled")

        def _callback(progreso: float, nombre: str):
            self.after(0, lambda: self.progress_bar.set(progreso))
            self.after(0, lambda: self.lbl_progreso.configure(text=nombre[:32]))

        def _tarea():
            try:
                archivos = exportar_todos(
                    self._df_resultado, directorio,
                    self._ruta_logo, callback=_callback
                )
                self.after(0, lambda: messagebox.showinfo(
                    "Exportación completa",
                    f"Se generaron {len(archivos)} archivos en:\n{directorio}"
                ))
                self.after(0, lambda: self._set_estado(
                    f"✅ {len(archivos)} reportes exportados", "ok"
                ))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally:
                self.after(0, lambda: self.btn_exportar_todos.configure(state="normal"))
                self.after(0, lambda: self.progress_bar.set(0))
                self.after(0, lambda: self.lbl_progreso.configure(text=""))

        threading.Thread(target=_tarea, daemon=True).start()

    # ── Gráfica de pastel ─────────────────────────────────────────────────
    def _grafica_placeholder(self):
        """Muestra un mensaje en lugar de gráfica cuando no hay datos."""
        for w in self.frame_grafica.winfo_children():
            w.destroy()
        ctk.CTkLabel(
            self.frame_grafica,
            text="📊\n\nCargue y procese los\narchivos para ver\nel resumen estadístico.",
            font=ctk.CTkFont(size=12),
            text_color="#A0A0A0",
            justify="center"
        ).pack(expand=True)

    def _mostrar_grafica(self, resumen: dict):
        """Renderiza la gráfica de pastel de resumen."""
        for w in self.frame_grafica.winfo_children():
            w.destroy()

        conteos = resumen.get("conteos", {})
        if not conteos:
            self._grafica_placeholder()
            return

        etiquetas = []
        valores   = []
        colores   = []

        for estatus, cantidad in conteos.items():
            info = COLOR_MAP.get(estatus, COLOR_MAP["SIN_DATOS"])
            hex_color = f"#{info['hex']}"
            etiquetas.append(f"{info['label']}\n({cantidad})")
            valores.append(cantidad)
            colores.append(hex_color)

        fig, ax = plt.subplots(figsize=(4.2, 3.8), dpi=80)
        fig.patch.set_facecolor("#FFFFFF")
        ax.set_facecolor("#FFFFFF")

        wedges, texts, autotexts = ax.pie(
            valores, labels=etiquetas, colors=colores,
            autopct="%1.1f%%", startangle=140,
            textprops={"fontsize": 7.5},
            wedgeprops={"linewidth": 1.5, "edgecolor": "white"}
        )
        for t in autotexts:
            t.set_fontsize(7)
            t.set_color("#2F2F2F")

        ax.set_title(
            f"Distribución de Asistencia\n"
            f"{resumen.get('empleados', 0)} empleados · {resumen['total_dias_laborables']} días",
            fontsize=9, color="#1F3864", fontweight="bold"
        )
        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.frame_grafica)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=8)
        plt.close(fig)

    # ── Estado ────────────────────────────────────────────────────────────
    def _set_estado(self, mensaje: str, tipo: str = "ok"):
        colores = {
            "ok":    "#1F3864",
            "error": COLOR_ERROR,
            "warn":  COLOR_WARN,
        }
        self.lbl_estado.configure(
            text=mensaje,
            text_color=colores.get(tipo, "#1F3864")
        )


# ──────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────
def iniciar_app():
    app = SGAMApp()
    app.mainloop()
