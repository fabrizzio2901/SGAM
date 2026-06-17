"""
SGAM - Sistema de Gestión de Asistencias Médicas
Módulo: ui.py
Responsabilidad: Interfaz gráfica con CustomTkinter.

Mejoras v1.3:
  - Ampliación de información del personal en el Tablero Individual.
  - Renderizado visual de porcentajes y métricas de desempeño.
  - Protección de memoria fig.clf() en gráficas.
  - Integración del Estatus "Festivo" a las leyendas.
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

from ingestion import (cargar_plantilla, extraer_reglas,
                        cargar_reporte_scanner, cargar_reporte_scanner_hospital,
                        cargar_dos_scanners)
from core import (
    procesar_asistencias, calcular_resumen, filtrar_por_empleado, COLOR_MAP,
    obtener_tipos_unicos, obtener_especialidades_unicas,
)
from export import (
    exportar_reporte_empleado, exportar_todos,
    exportar_filtrado, exportar_maestro_consolidado,
)
from utils import (
    estadisticas_por_empleado, ranking_asistencia,
    resumen_por_tipo, resumen_por_especialidad,
    dias_criticos, calcular_tendencia_semanal,
    exportar_estadisticas_excel,
)


# ──────────────────────────────────────────────
# Configuración visual global
# ──────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

APP_TITLE   = "SGAM – Sistema de Gestión de Asistencias Médicas"
APP_VERSION = "v1.3.0 Alan Jimenez & Fabrizzio Ramirez"
WIN_W, WIN_H = 1380, 860

SIDEBAR_W    = 245
COLOR_SIDEBAR = "#1F3864"
COLOR_ACCENT  = "#2E75B6"
COLOR_BG      = "#F0F4F8"
COLOR_WHITE   = "#FFFFFF"
COLOR_TEXT    = "#1F1F1F"
COLOR_OK      = "#217346"
COLOR_WARN    = "#C55A11"
COLOR_ERR     = "#C00000"
COLOR_PURPLE  = "#7030A0"


# ══════════════════════════════════════════════
# Widget: Calendario Mensual
# ══════════════════════════════════════════════
class CalendarioWidget(ctk.CTkFrame):
    """
    Muestra un calendario mensual coloreado con ampliación de datos del usuario.
    """
    DIAS_HDR = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
    # [MEJORA] Integración del color Festivo
    LEYENDA  = [
        ("#FFD966", "Asistencia"), ("#FF4B4B", "Falta"),
        ("#FF8C00", "Retardo"),    ("#D9D9D9", "No Laborable"),
        ("#70AD47", "Vacaciones"), ("#ADD8E6", "Festivo"), 
        ("#9DC3E6", "Incapacidad"),("#FFE699", "Permiso"), 
        ("#BDD7EE", "Comisión"),   ("#F4B942", "Rotación"),
    ]

    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color=COLOR_WHITE, corner_radius=10, **kwargs)
        self._foto_label: Optional[ctk.CTkLabel] = None
        self._construir()

    def _construir(self):
        # ── Foto + Título ──────────────────────────────
        top = ctk.CTkFrame(self, fg_color=COLOR_WHITE)
        top.pack(fill="x", padx=12, pady=(10, 2))

        self.foto_frame = ctk.CTkFrame(top, width=58, height=58, fg_color="#E8EEF4", corner_radius=8)
        self.foto_frame.pack(side="left", padx=(0, 10))
        self.foto_frame.pack_propagate(False)
        self._lbl_foto_ico = ctk.CTkLabel(self.foto_frame, text="👤", font=ctk.CTkFont(size=26), text_color="#AAAAAA")
        self._lbl_foto_ico.place(relx=0.5, rely=0.5, anchor="center")

        # Contenedor de Títulos
        self.titulo_frame = ctk.CTkFrame(top, fg_color=COLOR_WHITE)
        self.titulo_frame.pack(side="left", fill="x", expand=True)

        self.lbl_nombre = ctk.CTkLabel(self.titulo_frame, text="Sin datos cargados", font=ctk.CTkFont(family="Calibri", size=14, weight="bold"), text_color=COLOR_SIDEBAR, anchor="w")
        self.lbl_nombre.pack(anchor="w")

        # [MEJORA REQUERIMIENTO 1] Contenedor para desplegar campos extendidos 
        self.info_frame = ctk.CTkFrame(self.titulo_frame, fg_color="transparent")
        self.info_frame.pack(fill="x", pady=2)

        # ── Grid del calendario ──────────────────────────────────────────
        self.grid_frame = ctk.CTkFrame(self, fg_color=COLOR_WHITE)
        self.grid_frame.pack(padx=12, pady=(4, 2))

        for j, dia in enumerate(self.DIAS_HDR):
            lbl = ctk.CTkLabel(self.grid_frame, text=dia, width=50, height=26, font=ctk.CTkFont(size=9, weight="bold"), fg_color=COLOR_SIDEBAR, text_color=COLOR_WHITE, corner_radius=4)
            lbl.grid(row=0, column=j, padx=2, pady=2)

        # ── Leyenda de colores ──────────────────────────────
        ley_frame = ctk.CTkFrame(self, fg_color="#EEF2F7", corner_radius=8)
        ley_frame.pack(fill="x", padx=12, pady=(6, 12))

        ctk.CTkLabel(ley_frame, text="Leyenda:", font=ctk.CTkFont(size=9, weight="bold"), text_color="#1F3864").pack(anchor="w", padx=10, pady=(6, 2))

        # Render dinámico en base a la lista extendida
        fila1, fila2 = self.LEYENDA[:5], self.LEYENDA[5:]
        for fila_items in [fila1, fila2]:
            row = ctk.CTkFrame(ley_frame, fg_color="#EEF2F7")
            row.pack(fill="x", padx=6, pady=2)
            for hex_c, label in fila_items:
                chip = tk.Canvas(row, width=15, height=15, bg=hex_c, highlightthickness=1, highlightbackground="#AAAAAA")
                chip.pack(side="left", padx=(4, 2), pady=4)
                ctk.CTkLabel(row, text=label, font=ctk.CTkFont(size=9), text_color="#2F2F2F").pack(side="left", padx=(0, 10))
        tk.Frame(ley_frame, height=4, bg="#EEF2F7").pack()

    # ── API pública ───────────────────────────────────────────────────────
    def cargar_mes(self, df_emp: pd.DataFrame, mes: int, anio: int, nombre: str, 
                   id_emp: str = "", tipo: str = "", esp: str = "", univ: str = "", 
                   subesp: str = "", altaesp: str = "", periodo: str = "", 
                   turno_tipo: str = "", foto_ruta: str = ""):
        
        for w in self.grid_frame.winfo_children():
            if int(w.grid_info().get("row", 0)) > 0:
                w.destroy()

        self.lbl_nombre.configure(text=f"{nombre} | {self._mes_nombre(mes)} {anio}")
        
        # [MEJORA REQUERIMIENTO 1] Renderizado de todos los campos nuevos y opcionales
        for w in self.info_frame.winfo_children():
            w.destroy()
            
        campos_a_mostrar = [
            ("ID", id_emp),
            ("Tipo", tipo),
            ("Especialidad", esp),
            ("Universidad", univ),
            ("Sub especialidad", subesp),
            ("Alta especialidad", altaesp),
            ("Ingreso", periodo),
            ("Turno", turno_tipo)
        ]
        
        # Grid para aprovechar el espacio (2 columnas)
        fila_idx, col_idx = 0, 0
        for label_text, valor in campos_a_mostrar:
            # Filtramos estrictamente los nulos para no saturar la UI
            if valor and valor not in ["N/A", "nan"]:
                row_container = ctk.CTkFrame(self.info_frame, fg_color="transparent")
                row_container.grid(row=fila_idx, column=col_idx, sticky="w", padx=(0, 15), pady=1)
                
                ctk.CTkLabel(row_container, text=f"{label_text}: ", font=ctk.CTkFont(size=10, weight="bold"), text_color=COLOR_SIDEBAR).pack(side="left")
                ctk.CTkLabel(row_container, text=f"{valor}", font=ctk.CTkFont(size=10), text_color="#595959").pack(side="left")
                
                col_idx += 1
                if col_idx > 1:  # Máximo 2 columnas
                    col_idx = 0
                    fila_idx += 1

        self._cargar_foto(foto_ruta)

        mapa: dict[int, dict] = {}
        for _, row in df_emp.iterrows():
            d = pd.to_datetime(row["Fecha"]).day
            mapa[d] = {
                "color":  f"#{row.get('Color_Hex', 'FFFFFF')}",
                "label":  row.get("Label", ""),
                "notas":  row.get("Notas", ""),
                "turno":  row.get("Turno_Tipo", ""),
            }

        primer_dia = calendar.monthrange(anio, mes)[0]
        dias_mes   = calendar.monthrange(anio, mes)[1]
        fila, col  = 1, primer_dia

        for d in range(1, dias_mes + 1):
            info    = mapa.get(d, {"color": "#FFFFFF", "label": "", "notas": "", "turno": ""})
            tooltip = "\n".join(filter(None, [info["turno"], info["notas"] or info["label"]]))

            celda = ctk.CTkLabel(
                self.grid_frame, text=str(d), width=50, height=40,
                font=ctk.CTkFont(size=10, weight="bold"),
                fg_color=info["color"], text_color="#1F1F1F", corner_radius=5,
            )
            celda.grid(row=fila, column=col, padx=2, pady=2)
            self._tooltip(celda, tooltip)

            col += 1
            if col > 6:
                col = 0
                fila += 1

    def _cargar_foto(self, foto_ruta: str):
        for w in self.foto_frame.winfo_children():
            w.destroy()

        ruta = Path(foto_ruta) if foto_ruta else None
        if ruta and ruta.exists():
            try:
                from PIL import Image as PILImage
                img = PILImage.open(str(ruta))
                img = img.resize((54, 54))
                ctk_img = ctk.CTkImage(light_image=img, size=(54, 54))
                lbl = ctk.CTkLabel(self.foto_frame, image=ctk_img, text="")
                lbl.place(relx=0.5, rely=0.5, anchor="center")
                lbl.image = ctk_img   
                return
            except Exception:
                pass   

        lbl = ctk.CTkLabel(self.foto_frame, text="👤", font=ctk.CTkFont(size=26), text_color="#AAAAAA")
        lbl.place(relx=0.5, rely=0.5, anchor="center")

    @staticmethod
    def _mes_nombre(mes: int) -> str:
        ns = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
              "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        return ns[mes] if 1 <= mes <= 12 else str(mes)

    @staticmethod
    def _tooltip(widget, texto: str):
        if not texto: return
        tip = None
        def show(e):
            nonlocal tip
            tip = tk.Toplevel(widget)
            tip.overrideredirect(True)
            tip.geometry(f"+{e.x_root + 14}+{e.y_root + 14}")
            tk.Label(tip, text=texto, bg="#FFFFE0", font=("Calibri", 9), relief="solid", bd=1, padx=6, pady=4).pack()
        def hide(e):
            nonlocal tip
            if tip: tip.destroy(); tip = None
        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)


# ══════════════════════════════════════════════
# Widget: Panel de Analítica con Filtros
# ══════════════════════════════════════════════
class PanelAnalitica(ctk.CTkFrame):
    """
    Panel derecho con gráfica de pastel, porcentajes en texto + filtros dinámicos.
    """

    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color=COLOR_BG, corner_radius=0, **kwargs)
        self._df: Optional[pd.DataFrame] = None
        self._nombre_actual: str = ""
        self._construir()

    def _construir(self):
        hdr = ctk.CTkFrame(self, fg_color=COLOR_WHITE, corner_radius=8)
        hdr.pack(fill="x", padx=0, pady=(0, 8))

        ctk.CTkLabel(hdr, text="📊 Analítica de Asistencia", font=ctk.CTkFont(size=13, weight="bold"), text_color=COLOR_SIDEBAR).pack(anchor="w", padx=12, pady=(10, 2))

        frow = ctk.CTkFrame(hdr, fg_color=COLOR_WHITE)
        frow.pack(fill="x", padx=12, pady=(4, 10))

        ctk.CTkLabel(frow, text="Tipo:", font=ctk.CTkFont(size=10), text_color="#595959").pack(side="left", padx=(0, 4))
        self.combo_tipo = ctk.CTkComboBox(frow, values=["(Todos)"], width=130, height=28, command=self._on_tipo_cambiado)
        self.combo_tipo.set("(Todos)")
        self.combo_tipo.pack(side="left", padx=(0, 8))

        self.lbl_esp = ctk.CTkLabel(frow, text="Especialidad:", font=ctk.CTkFont(size=10), text_color="#595959")
        self.combo_esp = ctk.CTkComboBox(frow, values=["(Todas)"], width=140, height=28, command=self._actualizar_grafica_filtros)
        self.combo_esp.set("(Todas)")
        self._esp_visible = False

        self.btn_individual = ctk.CTkButton(frow, text="👤 Vista Individual", width=130, height=28, fg_color=COLOR_ACCENT, hover_color="#1F5C8A", corner_radius=6, command=self._toggle_vista_individual)
        self.btn_individual.pack(side="left", padx=(8, 0))
        self._modo_individual = False

        # Contenedor padre de la información visual
        self.frame_main = ctk.CTkFrame(self, fg_color=COLOR_WHITE, corner_radius=8)
        self.frame_main.pack(fill="both", expand=True)
        
        # Subcontenedores (Textos de Métrica y Gráfica)
        self.frame_metricas = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        self.frame_metricas.pack(side="left", fill="y", padx=20, pady=20)
        self.frame_chart = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        self.frame_chart.pack(side="right", fill="both", expand=True)

        self._placeholder()

    def set_datos(self, df: pd.DataFrame, tipos: list[str], especialidades: list[str]):
        self._df             = df
        self._especialidades = especialidades

        TIPOS_OBSOLETOS = {"médico adscrito", "medico adscrito"}
        tipos_validos   = [t for t in tipos if t.lower() not in TIPOS_OBSOLETOS]

        self.combo_tipo.configure(values=["(Todos)"] + tipos_validos)
        self.combo_tipo.set("(Todos)")
        self.combo_esp.configure(values=["(Todas)"] + especialidades)
        self.combo_esp.set("(Todas)")

        self._ocultar_esp()
        self._modo_individual = False
        self.btn_individual.configure(text="👤 Vista Individual", fg_color=COLOR_ACCENT)
        self._dibujar_grafica(df)

    def set_empleado_actual(self, nombre: str):
        self._nombre_actual = nombre
        if self._modo_individual:
            self._dibujar_individual()

    def _ocultar_esp(self):
        if self._esp_visible:
            self.lbl_esp.pack_forget()
            self.combo_esp.pack_forget()
            self._esp_visible = False

    def _mostrar_esp(self):
        if not self._esp_visible:
            self.lbl_esp.pack(side="left", padx=(8, 4), before=self.btn_individual)
            self.combo_esp.pack(side="left", padx=(0, 8), before=self.btn_individual)
            self._esp_visible = True

    def _on_tipo_cambiado(self, seleccion: str):
        es_residente = seleccion.lower() == "residente"
        if es_residente: self._mostrar_esp()
        else:
            self._ocultar_esp()
            self.combo_esp.set("(Todas)")
        self._actualizar_grafica_filtros()

    def _actualizar_grafica_filtros(self, _=None):
        if self._df is None: return
        if self._modo_individual:
            self._dibujar_individual()
            return

        df_f     = self._df.copy()
        sel_tipo = self.combo_tipo.get()
        sel_esp  = self.combo_esp.get()

        if sel_tipo != "(Todos)": df_f = df_f[df_f["Tipo"].str.lower().str.contains(sel_tipo.lower(), na=False)]
        if sel_esp != "(Todas)":  df_f = df_f[df_f["Especialidad"].str.lower().str.contains(sel_esp.lower(), na=False)]

        if df_f.empty:
            self._placeholder("Sin datos para\nlos filtros seleccionados.")
            return
        self._dibujar_grafica(df_f)

    def _toggle_vista_individual(self):
        self._modo_individual = not self._modo_individual
        if self._modo_individual:
            self.btn_individual.configure(text="🌐 Vista General", fg_color="#7030A0")
            self._dibujar_individual()
        else:
            self.btn_individual.configure(text="👤 Vista Individual", fg_color=COLOR_ACCENT)
            self._actualizar_grafica_filtros()

    def _dibujar_individual(self):
        if self._df is None or not self._nombre_actual:
            self._placeholder("Seleccione un empleado\nen la lista superior.")
            return
        df_emp = self._df[self._df["Nombre_Completo"] == self._nombre_actual]
        if df_emp.empty:
            self._placeholder("Sin datos para este empleado.")
            return
        primera = df_emp.iloc[0]
        subtitulo = " · ".join(filter(None, [str(primera.get("Tipo", "")), str(primera.get("Especialidad", ""))]))
        self._dibujar_grafica(df_emp, titulo_extra=f"{self._nombre_actual}\n{subtitulo}")

    def _dibujar_grafica(self, df: pd.DataFrame, titulo_extra: str = ""):
        """Renderiza gráfica de pastel y calcula los porcentajes reales de rendimiento."""
        for w in self.frame_chart.winfo_children(): w.destroy()
        for w in self.frame_metricas.winfo_children(): w.destroy()

        lab = df[df["Estatus"] != "NO_LABORABLE"]
        conteos = lab["Estatus"].value_counts().to_dict()
        if not conteos:
            self._placeholder()
            return
            
        # [MEJORA REQUERIMIENTO 4] Lógica estricta de evaluación y render en texto
        asist = conteos.get("ASISTENCIA", 0)
        ret   = conteos.get("RETARDO", 0)
        fal   = conteos.get("FALTA", 0)
        dias_evaluados = asist + ret + fal
        
        pct_asist = round((asist / dias_evaluados)*100, 1) if dias_evaluados else 0.0
        pct_ret   = round((ret / dias_evaluados)*100, 1) if dias_evaluados else 0.0
        pct_fal   = round((fal / dias_evaluados)*100, 1) if dias_evaluados else 0.0
        
        n_emp  = df["ID"].nunique()
        titulo_seccion = titulo_extra or f"Resumen de Desempeño\n({n_emp} empleados)"

        ctk.CTkLabel(self.frame_metricas, text=titulo_seccion, font=ctk.CTkFont(size=12, weight="bold"), text_color=COLOR_SIDEBAR).pack(anchor="w", pady=(10, 15))
        
        texto_stats = (
            f"Días Puntuables: {dias_evaluados}\n\n"
            f"✔️ Asistencias: {asist} ({pct_asist}%)\n\n"
            f"⚠️ Retardos: {ret} ({pct_ret}%)\n\n"
            f"❌ Faltas: {fal} ({pct_fal}%)"
        )
        ctk.CTkLabel(self.frame_metricas, text=texto_stats, justify="left", font=ctk.CTkFont(size=12, weight="bold"), text_color="#1F1F1F").pack(anchor="w")

        # Configuración de la gráfica de pastel con todos los estatus (incluye vacaciones, etc.)
        etiquetas, valores, colores = [], [], []
        for est, cant in conteos.items():
            info = COLOR_MAP.get(est, COLOR_MAP["SIN_DATOS"])
            etiquetas.append(f"{info['label']}\n({cant})")
            valores.append(cant)
            colores.append(f"#{info['hex']}")

        from matplotlib.figure import Figure
        # [MEJORA: PROTECCIÓN DE MEMORIA]
        fig = Figure(figsize=(3.4, 3.0), dpi=80)
        fig.clf() # Limpiamos el caché interno de Figure
        fig.patch.set_facecolor("#FFFFFF")
        ax = fig.add_subplot(111)
        ax.set_facecolor("#FFFFFF")
        
        wedges, texts, autotexts = ax.pie(
            valores, labels=etiquetas, colors=colores,
            autopct="%1.1f%%", startangle=140,
            textprops={"fontsize": 7.5},
            wedgeprops={"linewidth": 1.5, "edgecolor": "white"}
        )
        for t in autotexts:
            t.set_fontsize(7); t.set_color("#2F2F2F")
        ax.set_title("Distribución General del Mes", fontsize=9, color="#1F3864", fontweight="bold", pad=10)
        
        fig.tight_layout()

        self.canvas_pie = FigureCanvasTkAgg(fig, master=self.frame_chart)
        self.canvas_pie.draw()
        self.canvas_pie.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=8)

    def _placeholder(self, mensaje: str = "Cargue y procese los\narchivos para ver\nel análisis."):
        for w in self.frame_chart.winfo_children(): w.destroy()
        for w in self.frame_metricas.winfo_children(): w.destroy()
        ctk.CTkLabel(self.frame_chart, text=f"📊\n\n{mensaje}", font=ctk.CTkFont(size=11), text_color="#A0A0A0", justify="center").pack(expand=True)


# ══════════════════════════════════════════════
# Ventana Principal
# ══════════════════════════════════════════════
class SGAMApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self._ruta_plantilla:  Optional[str] = None   
        self._ruta_scanner:    Optional[str] = None   
        self._df_scanner:      None                    
        self._ruta_logo:       Optional[str] = None
        self._datos:           Optional[dict] = None
        self._reglas:          Optional[dict] = None
        self._df:              Optional[pd.DataFrame] = None

        self._configurar_ventana()
        self._construir_ui()

    def _configurar_ventana(self):
        self.title(f"{APP_TITLE}  {APP_VERSION}")
        self.geometry(f"{WIN_W}x{WIN_H}")
        self.minsize(1150, 700)
        self.configure(fg_color=COLOR_BG)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

    def _construir_ui(self):
        self._sidebar()
        self._panel_central()
        self._barra_estado()

    def _sidebar(self):
        sb_outer = ctk.CTkFrame(self, width=SIDEBAR_W, corner_radius=0, fg_color=COLOR_SIDEBAR)
        sb_outer.grid(row=0, column=0, sticky="nsew", rowspan=2)
        sb_outer.grid_propagate(False)
        sb_outer.pack_propagate(False)

        hdr = ctk.CTkFrame(sb_outer, fg_color=COLOR_SIDEBAR, corner_radius=0)
        hdr.pack(fill="x")
        ctk.CTkLabel(hdr, text="⚕ SGAM", font=ctk.CTkFont(size=20, weight="bold"), text_color=COLOR_WHITE).pack(pady=(16, 1))
        ctk.CTkLabel(hdr, text="Gestión de Asistencias Médicas", font=ctk.CTkFont(size=9), text_color="#A8C8E8").pack(pady=(0, 10))
        ctk.CTkFrame(sb_outer, height=1, fg_color="#3A5C8C").pack(fill="x")

        canvas   = tk.Canvas(sb_outer, bg=COLOR_SIDEBAR, highlightthickness=0, bd=0, relief="flat")
        scrollbar = tk.Scrollbar(sb_outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        sb = tk.Frame(canvas, bg=COLOR_SIDEBAR)
        win_id = canvas.create_window((0, 0), window=sb, anchor="nw")
        self._sb = sb  

        def _on_resize(e):
            canvas.itemconfig(win_id, width=e.width)
        canvas.bind("<Configure>", _on_resize)

        def _on_frame(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        sb.bind("<Configure>", _on_frame)

        def _on_wheel(e):
            canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_wheel)

        self._sep()
        self._sec("📂 ARCHIVOS DE ENTRADA")
        self._btn("Plantilla Maestra", self._cargar_plantilla, "📋")
        self.lbl_plantilla = self._lbl_est("Sin cargar")
        self._btn("Reporte Escáner", self._cargar_scanner, "🖱")
        self.lbl_scanner = self._lbl_est("Sin cargar")
        self._btn("Logo Institucional (PNG)", self._cargar_logo, "🖼")
        self.lbl_logo = self._lbl_est("Opcional")
        self._sep()

        self._sec("⚙ PROCESAMIENTO")
        self.btn_proc = self._btn("Procesar Asistencias", self._procesar, "▶", color="#217346", hover="#1A5C37")
        self._sep()

        self._sec("📤 EXPORTACIÓN")
        self.btn_exp1 = self._btn("Exportar Empleado Actual",     self._exp_actual,   "📄")
        self.btn_exp2 = self._btn("Exportar Todos (individual)",  self._exp_todos,    "📦")
        self.btn_exp3 = self._btn("Exportar con Filtros…",        self._exp_filtrado, "🔽")
        self.btn_exp4 = self._btn("Exportar Maestro (1 archivo)", self._exp_maestro, "📊", color=COLOR_PURPLE, hover="#5C2480")
        self._sep()

        self._sec("📈 ESTADÍSTICAS")
        self.btn_stats     = self._btn("Ver Estadísticas",     self._ver_estadisticas, "🔢", color="#B8450A", hover="#8C3208")
        self.btn_exp_stats = self._btn("Exportar Estadísticas", self._exp_estadisticas, "📑", color="#8C3208",  hover="#6B2506")
        self._sep()

        fr_prog = tk.Frame(sb, bg=COLOR_SIDEBAR)
        fr_prog.pack(fill="x", padx=12, pady=(4, 8))
        self.prog = ctk.CTkProgressBar(fr_prog)
        self.prog.pack(fill="x", pady=(2, 2))
        self.prog.set(0)
        self.lbl_prog = ctk.CTkLabel(fr_prog, text="", font=ctk.CTkFont(size=8), text_color="#A8C8E8")
        self.lbl_prog.pack()

    def _sep(self): tk.Frame(self._sb, height=1, bg="#3A5C8C").pack(fill="x", padx=12, pady=6)
    def _sec(self, texto): tk.Label(self._sb, text=texto, font=("Calibri", 8, "bold"), fg="#7EB3D8", bg=COLOR_SIDEBAR, anchor="w").pack(anchor="w", padx=16, pady=(4, 1))
    
    def _btn(self, texto, cmd, ico="", color=COLOR_ACCENT, hover="#1F5C8A"):
        b = ctk.CTkButton(self._sb, text=f"{ico}  {texto}", command=cmd, fg_color=color, hover_color=hover, text_color=COLOR_WHITE, font=ctk.CTkFont(size=10), height=32, corner_radius=7, anchor="w")
        b.pack(padx=12, pady=2, fill="x")
        return b

    def _lbl_est(self, texto):
        lbl = ctk.CTkLabel(self._sb, text=f"  {texto}", font=ctk.CTkFont(size=8, slant="italic"), text_color="#90B8D8", anchor="w")
        lbl.pack(padx=16, pady=(0, 2), fill="x")
        return lbl

    def _panel_central(self):
        pc = ctk.CTkFrame(self, corner_radius=0, fg_color=COLOR_BG)
        pc.grid(row=0, column=1, sticky="nsew")
        pc.grid_columnconfigure(0, weight=3)
        pc.grid_columnconfigure(1, weight=2)
        pc.grid_rowconfigure(1, weight=1)

        bb = ctk.CTkFrame(pc, height=58, fg_color=COLOR_WHITE, corner_radius=0)
        bb.grid(row=0, column=0, columnspan=2, sticky="ew")

        ctk.CTkLabel(bb, text="🔍", font=ctk.CTkFont(size=13, weight="bold"), text_color=COLOR_TEXT).pack(side="left", padx=(14, 4), pady=14)

        self.entry_bus = ctk.CTkEntry(bb, placeholder_text="ID Biométrico o Nombre…", width=300, height=34, font=ctk.CTkFont(size=11))
        self.entry_bus.pack(side="left", padx=4)
        self.entry_bus.bind("<Return>", lambda e: self._buscar())

        ctk.CTkButton(bb, text="Buscar", command=self._buscar, width=88, height=34, corner_radius=8, fg_color=COLOR_ACCENT).pack(side="left", padx=4)

        ctk.CTkLabel(bb, text="|", text_color="#CCCCCC").pack(side="left", padx=8)
        ctk.CTkLabel(bb, text="Empleado:", font=ctk.CTkFont(size=11), text_color="#595959").pack(side="left", padx=(0, 4))

        self.combo_emp = ctk.CTkComboBox(bb, values=["(sin datos)"], width=280, height=34, command=self._cambiar_empleado)
        self.combo_emp.pack(side="left", padx=4)

        fc = ctk.CTkFrame(pc, fg_color=COLOR_BG, corner_radius=0)
        fc.grid(row=1, column=0, sticky="nsew", padx=(10, 5), pady=10)

        ctk.CTkLabel(fc, text="📅 Tablero Individual", font=ctk.CTkFont(size=13, weight="bold"), text_color=COLOR_SIDEBAR).pack(anchor="w", pady=(0, 6))

        self.calendario = CalendarioWidget(fc)
        self.calendario.pack(fill="both", expand=True)

        self.analitica = PanelAnalitica(pc)
        self.analitica.grid(row=1, column=1, sticky="nsew", padx=(5, 10), pady=10)

    def _barra_estado(self):
        ba = ctk.CTkFrame(self, height=26, corner_radius=0, fg_color="#D6E4F0")
        ba.grid(row=1, column=1, sticky="ew")

        self.lbl_estado = ctk.CTkLabel(ba, text="Listo. Cargue los archivos para comenzar.", font=ctk.CTkFont(size=10), text_color="#1F3864")
        self.lbl_estado.pack(side="left", padx=14)
        self.lbl_contadores = ctk.CTkLabel(ba, text="", font=ctk.CTkFont(size=9), text_color="#5A7A9A")
        self.lbl_contadores.pack(side="right", padx=14)
        ctk.CTkLabel(ba, text=f"SGAM {APP_VERSION}", font=ctk.CTkFont(size=9), text_color="#8AACC8").pack(side="right", padx=(0, 6))

    def _cargar_plantilla(self):
        ruta = filedialog.askopenfilename(title="Seleccionar Plantilla Maestra", filetypes=[("Excel", "*.xlsx *.xls")])
        if not ruta: return
        try:
            datos = cargar_plantilla(ruta)
            self._ruta_plantilla = ruta
            cat    = datos["catalogo"]
            n_tot  = len(cat[cat["_activo"]])
            n_int  = len(cat[cat["_activo"] & (cat["Tipo de personal"] == "Interno")])
            n_res  = len(cat[cat["_activo"] & (cat["Tipo de personal"] == "Residente")])
            nombre = Path(ruta).name
            self.lbl_plantilla.configure(text=f"  ✅ {nombre}  ({n_int} int · {n_res} res)", text_color="#70AD47")
            self._set_estado(f"Plantilla cargada: {n_tot} activos ({n_int} internos, {n_res} residentes)", "ok")
            self._mostrar_alertas(datos.get("alertas", []))
        except Exception as e:
            messagebox.showerror("Error al cargar plantilla", str(e))
            self._set_estado(f"Error: {e}", "error")

    def _mostrar_alertas(self, alertas: list):
        if alertas:
            messagebox.showwarning("Advertencias de validación", "Se detectaron los siguientes avisos:\n\n" + "\n".join(f"• {a}" for a in alertas[:15]))

    def _cargar_scanner(self):
        rutas = filedialog.askopenfilenames(title="Reporte(s) Escáner — selecciona uno o dos archivos", filetypes=[("Excel", "*.xls *.xlsx"), ("CSV", "*.csv")])
        if not rutas: return
        try:
            from pathlib import Path as _P
            def _auto_load(ruta):
                ext = _P(ruta).suffix.lower()
                if ext == ".xls": return cargar_reporte_scanner_hospital(ruta)
                import pandas as _pd
                xl = _pd.ExcelFile(ruta)
                if "Reporte de Excepciones" in xl.sheet_names: return cargar_reporte_scanner_hospital(ruta)
                return cargar_reporte_scanner(ruta)

            if len(rutas) == 1:
                df = _auto_load(rutas[0])
                detalle = _P(rutas[0]).name
            else:
                df = cargar_dos_scanners(rutas[0], rutas[1])
                detalle = f"{_P(rutas[0]).name} + {_P(rutas[1]).name}"

            self._ruta_scanner = rutas[0]   
            self._df_scanner   = df         
            self.lbl_scanner.configure(text=f"  ✅ {len(df)} registros  ({len(rutas)} archivo{'s' if len(rutas)>1 else ''})", text_color="#70AD47")
            self._set_estado(f"Escáner: {len(df)} registros — {detalle}", "ok")
        except Exception as e:
            messagebox.showerror("Error al cargar escáner", str(e))
            self._set_estado(f"Error escáner: {e}", "error")

    def _cargar_logo(self):
        ruta = filedialog.askopenfilename(title="Logo Institucional", filetypes=[("PNG", "*.png")])
        if ruta:
            self._ruta_logo = ruta
            self.lbl_logo.configure(text=f"  ✅ {Path(ruta).name[:26]}", text_color="#70AD47")

    def _procesar(self):
        if not self._ruta_plantilla:
            messagebox.showwarning("Datos faltantes", "Cargue la Plantilla Maestra primero.")
            return
        if not hasattr(self, '_df_scanner') or self._df_scanner is None:
            messagebox.showwarning("Datos faltantes", "Cargue el Reporte del Escáner.")
            return

        self._set_estado("⏳ Procesando…", "ok")
        self.btn_proc.configure(state="disabled")

        def _run():
            try:
                import pandas as pd
                datos = cargar_plantilla(self._ruta_plantilla)
                datos["scanner"] = self._df_scanner
                reglas = extraer_reglas(pd.DataFrame())
                self._datos  = datos
                self._reglas = reglas
                df = procesar_asistencias(datos, reglas)
                self._df = df
                self.after(0, self._post_proceso)
            except Exception as e:
                msg = str(e)
                self.after(0, lambda: messagebox.showerror("Error al procesar", msg))
                self.after(0, lambda: self._set_estado(f"Error: {msg[:80]}", "error"))
            finally:
                self.after(0, lambda: self.btn_proc.configure(state="normal"))

        threading.Thread(target=_run, daemon=True).start()

    def _post_proceso(self):
        df = self._df
        resumen = calcular_resumen(df)
        nombres = sorted(df["Nombre_Completo"].dropna().unique().tolist())

        self.combo_emp.configure(values=nombres)
        if nombres:
            self.combo_emp.set(nombres[0])
            self._mostrar_empleado(nombres[0])

        self.analitica.set_datos(df, obtener_tipos_unicos(df), obtener_especialidades_unicas(df))

        c = resumen["conteos"]
        self.lbl_contadores.configure(
            text=(
                f"👥 {resumen['empleados']} empleados  |  "
                f"📅 {resumen['total_dias_laborables']} días lab.  |  "
                f"🟡 {c.get('ASISTENCIA', 0)} asist.  "
                f"🔴 {c.get('RETARDO', 0)} retardos  "
                f"🟠 {c.get('FALTA', 0)} faltas"
            )
        )

        self._set_estado(f"✅ Procesado — {resumen['empleados']} empleados | {resumen['total_dias_laborables']} días laborables", "ok")

    def _buscar(self):
        if self._df is None: return
        termino = self.entry_bus.get().strip()
        if not termino: return
        df_f = filtrar_por_empleado(self._df, termino)
        if df_f.empty:
            messagebox.showinfo("Sin resultados", f"No se encontraron registros para '{termino}'")
            return
        nombre = df_f.iloc[0]["Nombre_Completo"]
        self._mostrar_empleado(nombre)
        self.combo_emp.set(nombre)

    def _cambiar_empleado(self, nombre: str):
        self._mostrar_empleado(nombre)

    def _mostrar_empleado(self, nombre: str):
        if self._df is None: return
        df_emp = self._df[self._df["Nombre_Completo"] == nombre].copy()
        if df_emp.empty: return

        # [MEJORA REQUERIMIENTO 1] Extracción completa de variables para la inyección visual
        primera  = df_emp.iloc[0]
        id_emp   = str(primera.get("ID", ""))
        tipo     = str(primera.get("Tipo", ""))
        esp      = str(primera.get("Especialidad", ""))
        univ     = str(primera.get("Universidad", "N/A"))
        subesp   = str(primera.get("Subespecialidad", "N/A"))
        altaesp  = str(primera.get("Alta_Especialidad", "N/A"))
        periodo  = str(primera.get("Periodo_Ingreso", ""))
        foto     = str(primera.get("Foto_Ruta", "")).strip()

        fechas = pd.to_datetime(df_emp["Fecha"])
        mes  = int(fechas.dt.month.mode()[0])
        anio = int(fechas.dt.year.mode()[0])

        df_mes = df_emp[(fechas.dt.month == mes) & (fechas.dt.year == anio)]

        turnos_mes = df_mes["Guardia_Tipo"].replace("", pd.NA).dropna()
        turno_tipo = str(turnos_mes.mode().iloc[0]) if len(turnos_mes) > 0 else ""

        self.calendario.cargar_mes(
            df_mes, mes, anio, nombre,
            id_emp=id_emp, tipo=tipo, esp=esp, univ=univ, 
            subesp=subesp, altaesp=altaesp, periodo=periodo, 
            turno_tipo=turno_tipo, foto_ruta=foto
        )
        self.analitica.set_empleado_actual(nombre)

    def _exp_actual(self):
        if self._df is None:
            messagebox.showwarning("Sin datos", "Procese las asistencias primero.")
            return
        nombre = self.combo_emp.get()
        if not nombre or nombre == "(sin datos)":
            messagebox.showwarning("Sin empleado", "Seleccione un empleado.")
            return
        directorio = filedialog.askdirectory(title="Carpeta de destino")
        if not directorio: return
        df_emp = self._df[self._df["Nombre_Completo"] == nombre].copy()
        try:
            ruta = exportar_reporte_empleado(df_emp, directorio, self._ruta_logo)
            messagebox.showinfo("Exportado", f"Reporte generado:\n{ruta}")
            self._set_estado(f"✅ Exportado: {Path(ruta).name}", "ok")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _exp_todos(self): self._exportar_con_opciones(maestro=False, filtrado=False)
    def _exp_filtrado(self): self._exportar_con_opciones(maestro=False, filtrado=True)
    def _exp_maestro(self): self._exportar_con_opciones(maestro=True, filtrado=True)

    def _exportar_con_opciones(self, maestro: bool, filtrado: bool):
        if self._df is None:
            messagebox.showwarning("Sin datos", "Procese las asistencias primero.")
            return

        filtros = self._dialogo_filtros() if filtrado else {"tipo": None, "especialidad": None}
        if filtros is None: return

        directorio = filedialog.askdirectory(title="Carpeta de destino")
        if not directorio: return

        self.prog.set(0)

        def _cb(progreso, nombre):
            self.after(0, lambda: self.prog.set(progreso))
            self.after(0, lambda: self.lbl_prog.configure(text=nombre[:32]))

        def _run():
            try:
                if maestro:
                    ruta = exportar_maestro_consolidado(self._df, directorio, ruta_logo=self._ruta_logo, filtro_tipo=filtros["tipo"], filtro_especialidad=filtros["especialidad"], callback=_cb)
                    self.after(0, lambda: messagebox.showinfo("Maestro generado", f"Archivo:\n{ruta}"))
                    self.after(0, lambda: self._set_estado(f"✅ Maestro: {Path(ruta).name}", "ok"))
                else:
                    archivos = exportar_filtrado(self._df, directorio, filtro_tipo=filtros["tipo"], filtro_especialidad=filtros["especialidad"], ruta_logo=self._ruta_logo, callback=_cb)
                    self.after(0, lambda: messagebox.showinfo("Exportación completa", f"Se generaron {len(archivos)} archivos en:\n{directorio}"))
                    self.after(0, lambda: self._set_estado(f"✅ {len(archivos)} archivos exportados", "ok"))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally:
                self.after(0, lambda: self.prog.set(0))
                self.after(0, lambda: self.lbl_prog.configure(text=""))

        threading.Thread(target=_run, daemon=True).start()

    def _dialogo_filtros(self) -> Optional[dict]:
        if self._df is None: return None

        tipos  = ["(Todos)"] + obtener_tipos_unicos(self._df)
        esps   = ["(Todas)"] + obtener_especialidades_unicas(self._df)
        result = {}
        cancel = [False]

        win = ctk.CTkToplevel(self)
        win.title("Filtros de exportación")
        win.geometry("440x330")
        win.resizable(False, False)
        win.grab_set()
        win.configure(fg_color=COLOR_BG)

        ctk.CTkLabel(win, text="Filtros de Exportación", font=ctk.CTkFont(size=13, weight="bold"), text_color=COLOR_SIDEBAR).pack(pady=(18, 6))
        ctk.CTkLabel(win, text="Deje en '(Todos/Todas)' para exportar sin restricción.", font=ctk.CTkFont(size=10), text_color="#595959").pack(pady=(0, 14))

        for lbl_txt, valores, key in [("Tipo de Personal:", tipos, "tipo"), ("Especialidad:", esps, "especialidad")]:
            fr = ctk.CTkFrame(win, fg_color=COLOR_WHITE, corner_radius=8)
            fr.pack(fill="x", padx=24, pady=4)
            ctk.CTkLabel(fr, text=lbl_txt, font=ctk.CTkFont(size=10, weight="bold"), text_color=COLOR_SIDEBAR).pack(anchor="w", padx=12, pady=(8, 2))
            cb = ctk.CTkComboBox(fr, values=valores, height=32)
            cb.set(valores[0])
            cb.pack(padx=12, pady=(0, 10))
            result[key] = cb   

        fr_btns = ctk.CTkFrame(win, fg_color=COLOR_BG)
        fr_btns.pack(fill="x", padx=24, pady=(8, 16))
        fr_btns.grid_columnconfigure(0, weight=1)
        fr_btns.grid_columnconfigure(1, weight=1)

        def ok():
            for key, cb in result.items():
                sel = cb.get()
                result[key] = None if sel.startswith("(") else sel
            win.destroy()

        def no():
            cancel[0] = True
            win.destroy()

        ctk.CTkButton(fr_btns, text="✅  Continuar", command=ok, fg_color=COLOR_OK, hover_color="#1A5C37", height=38, corner_radius=8, font=ctk.CTkFont(size=11, weight="bold")).grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ctk.CTkButton(fr_btns, text="Cancelar", command=no, fg_color="#888888", hover_color="#666666", height=38, corner_radius=8, font=ctk.CTkFont(size=11)).grid(row=0, column=1, sticky="ew", padx=(5, 0))

        win.wait_window()
        return None if cancel[0] else result

    def _ver_estadisticas(self):
        if self._df is None:
            messagebox.showwarning("Sin datos", "Procese las asistencias primero.")
            return

        win = ctk.CTkToplevel(self)
        win.title("SGAM – Estadísticas Avanzadas")
        win.geometry("980x640")
        win.configure(fg_color=COLOR_BG)
        win.grab_set()

        hdr = ctk.CTkFrame(win, fg_color="#1F3864", corner_radius=0)
        hdr.pack(fill="x")
        ctk.CTkLabel(hdr, text="📈  Estadísticas Avanzadas del Período", font=ctk.CTkFont(size=14, weight="bold"), text_color="white").pack(side="left", padx=18, pady=12)

        tabs = ctk.CTkTabview(win, fg_color=COLOR_WHITE, corner_radius=8)
        tabs.pack(fill="both", expand=True, padx=12, pady=10)

        for nombre_tab in ["Individual", "Por Tipo", "Por Especialidad", "Tendencia Semanal", "Días Críticos"]:
            tabs.add(nombre_tab)

        # [MEJORA REQUERIMIENTO 4] Reemplazamos visualmente Días Lab por Días Evaluados
        df_ind = estadisticas_por_empleado(self._df)
        self._tabla_stats(tabs.tab("Individual"), df_ind, [
            ("Nombre", 220, "w"), ("Tipo", 120, "center"),
            ("Especialidad", 130, "center"), ("Grado", 55, "center"),
            ("Periodo", 60, "center"), ("Días Eval.", 70, "center"),
            ("Asist.", 55, "center"), ("Retardos", 65, "center"),
            ("Faltas", 55, "center"), ("% Asist.", 65, "center"),
            ("% Falta", 60, "center"), ("% Presencia", 80, "center"),
        ], [
            "Nombre_Completo", "Tipo", "Especialidad", "Grado",
            "Periodo_Ingreso", "dias_evaluados", "asistencias", "retardos",
            "faltas", "pct_asistencia", "pct_falta", "pct_presencia",
        ])

        df_tipo = resumen_por_tipo(self._df)
        self._tabla_stats(tabs.tab("Por Tipo"), df_tipo, [(c, 110, "center") for c in df_tipo.columns], list(df_tipo.columns))

        df_esp = resumen_por_especialidad(self._df)
        self._tabla_stats(tabs.tab("Por Especialidad"), df_esp, [(c, 120 if i == 0 else 95, "w" if i == 0 else "center") for i, c in enumerate(df_esp.columns)], list(df_esp.columns))

        df_sem = calcular_tendencia_semanal(self._df)
        self._tabla_stats(tabs.tab("Tendencia Semanal"), df_sem, [(c, 120, "center") for c in df_sem.columns], list(df_sem.columns))
        self._grafica_tendencia(tabs.tab("Tendencia Semanal"), df_sem)

        df_crit = dias_criticos(self._df, top_n=15)
        self._tabla_stats(tabs.tab("Días Críticos"), df_crit, [(c, 130 if i == 0 else 100, "center") for i, c in enumerate(df_crit.columns)], list(df_crit.columns))

        pie = ctk.CTkFrame(win, fg_color=COLOR_BG)
        pie.pack(fill="x", padx=16, pady=(4, 12))
        ctk.CTkButton(
            pie, text="📑  Exportar estas estadísticas a Excel",
            command=lambda: self._exp_estadisticas(win),
            fg_color="#8C3208", hover_color="#6B2506",
            height=36, corner_radius=8, font=ctk.CTkFont(size=11),
        ).pack(fill="x")

    def _tabla_stats(self, parent, df: pd.DataFrame, columnas_cfg: list, col_keys: list):
        if df.empty:
            ctk.CTkLabel(parent, text="Sin datos disponibles.", text_color="#888888").pack(expand=True)
            return

        outer = ctk.CTkFrame(parent, fg_color=COLOR_WHITE)
        outer.pack(fill="both", expand=True, padx=8, pady=8)

        canvas   = tk.Canvas(outer, bg=COLOR_WHITE, highlightthickness=0)
        scroll_y = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        scroll_x = tk.Scrollbar(outer, orient="horizontal", command=canvas.xview)
        inner    = tk.Frame(canvas, bg=COLOR_WHITE)

        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.pack(side="right",  fill="y")
        scroll_x.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)

        COL_HDR = "#1F3864"
        COL_ALT = "#F2F7FC"
        F_HDR   = ("Calibri", 9, "bold")
        F_BODY  = ("Calibri", 9)

        for j, (lbl, ancho, anchor) in enumerate(columnas_cfg):
            tk.Label(inner, text=lbl, bg=COL_HDR, fg="white", font=F_HDR, width=ancho // 7, anchor="center", relief="flat", pady=5, padx=4).grid(row=0, column=j, sticky="ew", padx=1, pady=1)

        for i, (_, row) in enumerate(df.iterrows()):
            bg = COL_ALT if i % 2 == 0 else "white"
            for j, (_, ancho, anchor) in enumerate(columnas_cfg):
                key = col_keys[j] if j < len(col_keys) else ""
                val = row.get(key, "") if key else ""
                texto = f"{val:.1f}%" if isinstance(val, float) else str(val)
                fg_color = "black"
                if isinstance(val, float):
                    if "falta" in key.lower() and val > 25: fg_color = "#C00000"
                    elif "asistencia" in key.lower() or "presencia" in key.lower():
                        fg_color = "#217346" if val >= 80 else ("#C55A11" if val >= 60 else "#C00000")
                tk.Label(inner, text=texto, bg=bg, fg=fg_color, font=F_BODY, anchor=anchor, width=ancho // 7, pady=3, padx=4).grid(row=i + 1, column=j, sticky="ew", padx=1, pady=0)

    def _grafica_tendencia(self, parent, df_sem: pd.DataFrame):
        if df_sem.empty: return
        try:
            from matplotlib.figure import Figure
            fig = Figure(figsize=(3.8, 2.4), dpi=80)
            fig.clf()
            fig.patch.set_facecolor("white")
            ax = fig.add_subplot(111)
            
            x     = range(len(df_sem))
            ancho = 0.28
            ax.bar([i - ancho for i in x], df_sem["Asistencias"], width=ancho, label="Asistencias", color="#FFD966", edgecolor="white")
            ax.bar(x, df_sem["Retardos"], width=ancho, label="Retardos", color="#FF4B4B", edgecolor="white")
            ax.bar([i + ancho for i in x], df_sem["Faltas"], width=ancho, label="Faltas", color="#FF8C00", edgecolor="white")
            ax.set_xticks(list(x))
            ax.set_xticklabels(df_sem["Semana"].tolist(), fontsize=7)
            ax.legend(fontsize=7, loc="upper right")
            ax.set_title("Tendencia por semana", fontsize=8, color="#1F3864")
            ax.tick_params(labelsize=7)
            
            fig.tight_layout()
            
            self.canvas_tendencia = FigureCanvasTkAgg(fig, master=parent)
            self.canvas_tendencia.draw()
            self.canvas_tendencia.get_tk_widget().pack(side="bottom", fill="x", padx=8, pady=(4, 8))
        except Exception as e:
            print("Error renderizando tendencia:", e)

    def _exp_estadisticas(self, parent_win=None):
        if self._df is None:
            messagebox.showwarning("Sin datos", "Procese las asistencias primero.")
            return
        directorio = filedialog.askdirectory(title="Carpeta de destino")
        if not directorio: return
        try:
            ruta = exportar_estadisticas_excel(self._df, directorio, self._ruta_logo)
            messagebox.showinfo("Estadísticas exportadas", f"Archivo generado:\n{ruta}")
            self._set_estado(f"✅ Estadísticas: {Path(ruta).name}", "ok")
            if parent_win:
                parent_win.grab_release()
                parent_win.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _set_estado(self, msg: str, tipo: str = "ok"):
        colores = {"ok": "#1F3864", "error": COLOR_ERR, "warn": COLOR_WARN}
        self.lbl_estado.configure(text=msg, text_color=colores.get(tipo, "#1F3864"))


# ──────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────
def iniciar_app():
    app = SGAMApp()
    app.mainloop()