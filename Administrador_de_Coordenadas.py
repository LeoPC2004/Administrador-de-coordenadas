# -*- coding: utf-8 -*-
"""
PROSEGUR - Administrador de Coordenadas RPA
===========================================

Herramienta de escritorio para recalibrar coordenadas de scripts RPA
(basados en PyAutoGUI) sin tocar el código fuente.

Dependencias:
    pip install pyautogui keyboard

Ejecución:
    python Administrador_de_Coordenadas.py

Empaquetado:
    pyinstaller --onefile --windowed Administrador_de_Coordenadas.py

Logo:
    Ya viene embebido. Para cambiarlo, colocar "logo_prosegur.png"
    (alto recomendado: 64 px) junto a este script / el .exe.
"""

import getpass
import os
import re
import shutil
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog, ttk

import keyboard
import pyautogui

# ==========================================================
# CONFIGURACIÓN
# ==========================================================

# Ruta por defecto del RPA. Si no existe, la app pedirá seleccionarlo.
RUTA_RPA = r"C:\Users\LeonardoMiguelPachas\Downloads\LeoPC\Prueba python\Admin de Cord\RPA PRUEBA.py"

TECLA_CAPTURA = "8"
TECLA_CANCELAR = "esc"
ARCHIVO_LOGO = "logo_prosegur.png"

# Paleta corporativa PROSEGUR: amarillo (igual al del logo) y negro
AMARILLO = "#F9D119"        # amarillo del logo
AMARILLO_OSC = "#E2B800"    # hover / acentos
AMARILLO_CLARO = "#FFF6CC"  # fondos suaves de paneles
NEGRO = "#1A1A1A"           # texto y botones principales
GRIS_OSC = "#3A3A3A"        # hover de botones negros
BLANCO = "#FFFFFF"
GRIS = "#5B6573"
VERDE = "#1F8A4C"
ROJO = "#B3261E"
AMBAR = "#B26A00"

# Logo embebido (PNG en base64) para que funcione también dentro del .exe.
# Si existe un archivo "logo_prosegur.png" junto a la app, se usa ese.
LOGO_B64 = (
    "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"
)

PATRON_COORD = re.compile(
    r'^([A-Z0-9_]+)\s*=\s*\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)',
    re.MULTILINE,
)


# ==========================================================
# LÓGICA DE ARCHIVO (lectura, backup, actualización, reporte)
# ==========================================================

def directorio_app():
    """Carpeta base de la app (funciona también con PyInstaller)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def leer_archivo(ruta):
    """Devuelve el contenido del archivo respetando saltos de línea."""
    with open(ruta, "r", encoding="utf-8", newline="") as f:
        return f.read()


def obtener_coordenadas(contenido):
    """Extrae todas las variables NOMBRE = (x, y) del contenido."""
    return [
        {
            "nombre": m.group(1),
            "x": int(m.group(2)),
            "y": int(m.group(3)),
            "texto_original": m.group(0),
        }
        for m in PATRON_COORD.finditer(contenido)
    ]


def crear_backup(ruta_rpa):
    """Crea RPA_BACKUP_YYYYMMDD_HHMMSS.py junto al archivo original."""
    fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta = os.path.dirname(ruta_rpa)
    base = os.path.splitext(os.path.basename(ruta_rpa))[0]
    destino = os.path.join(carpeta, f"{base}_BACKUP_{fecha}.py")

    contador = 2
    while os.path.exists(destino):  # evita sobrescribir si hay 2 saves/segundo
        destino = os.path.join(carpeta, f"{base}_BACKUP_{fecha}_{contador}.py")
        contador += 1

    shutil.copy2(ruta_rpa, destino)
    return destino


def actualizar_coordenada(ruta_rpa, nombre, nx, ny):
    """
    Reemplaza la coordenada en el archivo conservando la anterior comentada:

        # [2026-11-16 10:15:00] MENU_CONTRATOS = (270, 270)
        MENU_CONTRATOS = (355, 270)
    """
    contenido = leer_archivo(ruta_rpa)
    patron = re.compile(
        rf'^{re.escape(nombre)}\s*=\s*\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)',
        re.MULTILINE,
    )
    m = patron.search(contenido)
    if not m:
        raise ValueError(
            f"No se encontró '{nombre}' en el archivo. "
            "Puede haber sido modificado externamente; actualice la lista."
        )

    salto = "\r\n" if "\r\n" in contenido else "\n"
    sello = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    reemplazo = f"# [{sello}] {m.group(0)}{salto}{nombre} = ({nx}, {ny})"
    nuevo = contenido[:m.start()] + reemplazo + contenido[m.end():]

    # Escritura segura: archivo temporal + reemplazo atómico
    temporal = ruta_rpa + ".tmp"
    with open(temporal, "w", encoding="utf-8", newline="") as f:
        f.write(nuevo)
    os.replace(temporal, ruta_rpa)

    # Verificación posterior
    for c in obtener_coordenadas(leer_archivo(ruta_rpa)):
        if c["nombre"] == nombre and (c["x"], c["y"]) == (nx, ny):
            return
    raise RuntimeError("La verificación posterior al guardado falló.")


def generar_reporte(ruta_rpa, usuario, cambios, backups):
    """Genera el TXT resumen junto al RPA y devuelve su ruta."""
    ahora = datetime.now()
    nombre = f"REPORTE_COORDENADAS_{ahora.strftime('%Y%m%d_%H%M%S')}.txt"
    ruta = os.path.join(os.path.dirname(ruta_rpa), nombre)

    lineas = [
        "=" * 70,
        "PROSEGUR - REPORTE DE CALIBRACION DE COORDENADAS RPA",
        "=" * 70,
        f"Usuario : {usuario}",
        f"Fecha   : {ahora.strftime('%Y-%m-%d %H:%M:%S')}",
        f"Archivo : {ruta_rpa}",
        "",
        f"CAMBIOS REALIZADOS ({len(cambios)})",
        "-" * 70,
    ]
    if cambios:
        for i, c in enumerate(cambios, 1):
            lineas.append(
                f"{i:02d}. [{c['hora']}] {c['nombre']}: "
                f"({c['antes'][0]}, {c['antes'][1]}) -> "
                f"({c['nuevo'][0]}, {c['nuevo'][1]})"
            )
    else:
        lineas.append("Sin cambios en esta sesion.")

    if backups:
        lineas += ["", "BACKUPS CREADOS", "-" * 70] + backups

    lineas += ["", "COORDENADAS FINALES", "-" * 70]
    try:
        for c in obtener_coordenadas(leer_archivo(ruta_rpa)):
            lineas.append(f"{c['nombre']} = ({c['x']}, {c['y']})")
    except OSError as e:
        lineas.append(f"No se pudo leer el archivo: {e}")
    lineas.append("=" * 70)

    with open(ruta, "w", encoding="utf-8") as f:
        f.write("\n".join(lineas))
    return ruta


# ==========================================================
# OVERLAY FLOTANTE DE CAPTURA
# ==========================================================

class OverlayCaptura(tk.Toplevel):
    """Mensaje flotante siempre visible mientras se espera la tecla 8."""

    def __init__(self, master, nombre, on_cancel):
        super().__init__(master)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=NEGRO, highlightbackground=AMARILLO, highlightthickness=3)

        tk.Label(
            self, text=f"Capturando: {nombre}", bg=NEGRO, fg=AMARILLO,
            font=("Segoe UI", 9, "bold"),
        ).pack(padx=24, pady=(12, 0))

        tk.Label(
            self,
            text=f"Ubique el mouse en la posición deseada\ny presione la tecla {TECLA_CAPTURA}",
            bg=NEGRO, fg=BLANCO, font=("Segoe UI", 13, "bold"), justify="center",
        ).pack(padx=24, pady=(6, 4))

        self.lbl_pos = tk.Label(
            self, text="Mouse: (0, 0)", bg=NEGRO, fg=AMARILLO,
            font=("Consolas", 12, "bold"),
        )
        self.lbl_pos.pack(pady=(0, 6))

        tk.Button(
            self, text=f"Cancelar captura ({TECLA_CANCELAR.upper()})",
            command=on_cancel, bg=AMARILLO, fg=NEGRO, relief="flat",
            font=("Segoe UI", 9, "bold"), cursor="hand2", padx=12, pady=3,
        ).pack(pady=(0, 12))

        # Centrado arriba de la pantalla
        self.update_idletasks()
        ancho = self.winfo_reqwidth()
        self.geometry(f"+{(self.winfo_screenwidth() - ancho) // 2}+30")

    def actualizar_posicion(self, x, y):
        self.lbl_pos.config(text=f"Mouse: ({x}, {y})")


# ==========================================================
# APLICACIÓN PRINCIPAL
# ==========================================================

class AdministradorCoordenadas:

    def __init__(self, root):
        self.root = root
        self.ruta_rpa = RUTA_RPA
        self.usuario = ""
        self.coordenadas = []
        self.cambios = []          # historial de la sesión
        self.backups = []
        self.pendiente = None      # {"nombre", "antes", "nuevo"}
        self.capturando = False
        self.overlay = None
        self.logo_img = None

        self._configurar_ventana()
        self._configurar_estilos()

        # Usuario
        self.root.withdraw()
        if not self._pedir_usuario():
            self.root.destroy()
            return
        if not self._resolver_ruta_rpa():
            self.root.destroy()
            return

        self._construir_interfaz()
        self.root.deiconify()
        self.cargar_coordenadas()
        self._estado(f"Sesión iniciada por {self.usuario}.")

    # ------------------------------------------------------
    # Inicio
    # ------------------------------------------------------

    def _configurar_ventana(self):
        self.root.title("PROSEGUR - Administrador de Coordenadas RPA")
        self.root.geometry("1120x740")
        self.root.minsize(1000, 660)
        self.root.configure(bg=BLANCO)
        self.root.protocol("WM_DELETE_WINDOW", self.salir)

    def _configurar_estilos(self):
        s = ttk.Style()
        s.theme_use("clam")

        s.configure("Treeview", rowheight=28, font=("Segoe UI", 10),
                    background=BLANCO, fieldbackground=BLANCO, borderwidth=0)
        s.configure("Treeview.Heading", background=NEGRO, foreground=AMARILLO,
                    font=("Segoe UI", 10, "bold"), relief="flat", padding=6)
        s.map("Treeview.Heading", background=[("active", GRIS_OSC)])
        s.map("Treeview", background=[("selected", AMARILLO)],
              foreground=[("selected", NEGRO)])

        def boton(nombre, fondo, hover, fg=BLANCO):
            s.configure(nombre, background=fondo, foreground=fg, borderwidth=0,
                        font=("Segoe UI", 10, "bold"), padding=(14, 10))
            s.map(nombre, background=[("disabled", "#D9D9D9"), ("active", hover)],
                  foreground=[("disabled", "#9A9A9A")])

        boton("Primario.TButton", NEGRO, GRIS_OSC, fg=AMARILLO)
        boton("Exito.TButton", VERDE, "#17703D")
        boton("Peligro.TButton", ROJO, "#8E1F19")
        boton("Neutro.TButton", AMARILLO, AMARILLO_OSC, fg=NEGRO)

    def _pedir_usuario(self):
        try:
            sugerido = getpass.getuser()
        except Exception:
            sugerido = ""
        nombre = simpledialog.askstring(
            "PROSEGUR - Administrador de Coordenadas RPA",
            "Ingrese su nombre (quedará registrado en el reporte):",
            initialvalue=sugerido, parent=self.root,
        )
        if not nombre or not nombre.strip():
            return False
        self.usuario = nombre.strip()
        return True

    def _resolver_ruta_rpa(self):
        """Si la ruta configurada no existe, permite elegir el archivo."""
        while not os.path.isfile(self.ruta_rpa):
            messagebox.showwarning(
                "Archivo RPA no encontrado",
                f"No existe el archivo configurado:\n\n{self.ruta_rpa}\n\n"
                "Seleccione el archivo del RPA a calibrar.",
            )
            ruta = filedialog.askopenfilename(
                title="Seleccione el archivo del RPA",
                filetypes=[("Python", "*.py"), ("Todos", "*.*")],
            )
            if not ruta:
                return False
            self.ruta_rpa = ruta
        return True

    # ------------------------------------------------------
    # Interfaz
    # ------------------------------------------------------

    def _construir_interfaz(self):
        self._construir_cabecera()
        self._construir_barra_archivo()

        cuerpo = tk.Frame(self.root, bg=BLANCO)
        cuerpo.pack(fill="both", expand=True, padx=16, pady=(8, 0))
        cuerpo.columnconfigure(0, weight=3)
        cuerpo.columnconfigure(1, weight=2, minsize=340)
        cuerpo.rowconfigure(0, weight=3)
        cuerpo.rowconfigure(1, weight=2)

        self._construir_tabla(cuerpo)
        self._construir_panel_detalle(cuerpo)
        self._construir_historial(cuerpo)

        self.lbl_estado = tk.Label(
            self.root, text="", anchor="w", bg=AMARILLO_CLARO, fg=NEGRO,
            font=("Segoe UI", 9), padx=16, pady=6,
        )
        self.lbl_estado.pack(fill="x", side="bottom")

        self._actualizar_botones()

    def _construir_cabecera(self):
        # Cabecera amarilla: el fondo del logo coincide con el de la barra
        cab = tk.Frame(self.root, bg=AMARILLO)
        cab.pack(fill="x")

        self._cargar_logo()
        if self.logo_img:
            tk.Label(cab, image=self.logo_img, bg=AMARILLO).pack(
                side="left", padx=(20, 0), pady=8)

        titulos = tk.Frame(cab, bg=AMARILLO)
        titulos.pack(side="left", padx=24)
        tk.Label(titulos, text="PROSEGUR", bg=AMARILLO, fg=NEGRO,
                 font=("Segoe UI", 22, "bold")).pack(anchor="w")
        tk.Label(titulos, text="Administrador de Coordenadas RPA", bg=AMARILLO,
                 fg=GRIS_OSC, font=("Segoe UI", 13)).pack(anchor="w")

        tk.Label(cab, text=f"Usuario: {self.usuario}", bg=AMARILLO, fg=NEGRO,
                 font=("Segoe UI", 10, "bold")).pack(side="right", padx=24)

        # Línea negra bajo la cabecera
        tk.Frame(self.root, bg=NEGRO, height=4).pack(fill="x")

    def _cargar_logo(self):
        """Carga el logo externo si existe; si no, el embebido."""
        ruta_logo = os.path.join(directorio_app(), ARCHIVO_LOGO)
        try:
            if os.path.isfile(ruta_logo):
                self.logo_img = tk.PhotoImage(file=ruta_logo)
            else:
                self.logo_img = tk.PhotoImage(data=LOGO_B64)
        except tk.TclError:
            self.logo_img = None

    def _construir_barra_archivo(self):
        barra = tk.Frame(self.root, bg=AMARILLO_CLARO)
        barra.pack(fill="x")
        tk.Label(barra, text="Archivo RPA:", bg=AMARILLO_CLARO, fg=NEGRO,
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=(16, 6), pady=6)
        self.lbl_ruta = tk.Label(barra, text=self.ruta_rpa, bg=AMARILLO_CLARO, fg=GRIS,
                                 font=("Consolas", 9), anchor="w")
        self.lbl_ruta.pack(side="left", fill="x", expand=True)
        ttk.Button(barra, text="Cambiar archivo…", style="Neutro.TButton",
                   command=self.cambiar_archivo).pack(side="right", padx=12, pady=3)

    def _construir_tabla(self, padre):
        marco = tk.Frame(padre, bg=BLANCO)
        marco.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        marco.rowconfigure(1, weight=1)
        marco.columnconfigure(0, weight=1)

        encabezado = tk.Frame(marco, bg=BLANCO)
        encabezado.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        tk.Label(encabezado, text="Coordenadas del RPA", bg=BLANCO, fg=NEGRO,
                 font=("Segoe UI", 12, "bold")).pack(side="left")

        self.var_filtro = tk.StringVar()
        self.var_filtro.trace_add("write", lambda *_: self._poblar_tabla())
        tk.Entry(encabezado, textvariable=self.var_filtro, width=24,
                 font=("Segoe UI", 10), relief="solid", bd=1).pack(side="right")
        tk.Label(encabezado, text="Buscar:", bg=BLANCO, fg=GRIS,
                 font=("Segoe UI", 9)).pack(side="right", padx=6)

        cols = ("nombre", "x", "y")
        self.tabla = ttk.Treeview(marco, columns=cols, show="headings",
                                  selectmode="browse")
        self.tabla.heading("nombre", text="Nombre")
        self.tabla.heading("x", text="X actual")
        self.tabla.heading("y", text="Y actual")
        self.tabla.column("nombre", width=300, anchor="w")
        self.tabla.column("x", width=90, anchor="center")
        self.tabla.column("y", width=90, anchor="center")
        self.tabla.tag_configure("par", background="#FAF7E6")
        self.tabla.tag_configure("pendiente", background="#FFE08A")
        self.tabla.grid(row=1, column=0, sticky="nsew")

        sb = ttk.Scrollbar(marco, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=sb.set)
        sb.grid(row=1, column=1, sticky="ns")

        self.tabla.bind("<<TreeviewSelect>>", self._al_seleccionar)
        self.tabla.bind("<Double-1>", lambda e: self.iniciar_captura())

    def _construir_panel_detalle(self, padre):
        panel = tk.Frame(padre, bg=AMARILLO_CLARO)
        panel.grid(row=0, column=1, sticky="nsew")
        panel.columnconfigure(0, weight=1)

        tk.Label(panel, text="Coordenada seleccionada", bg=AMARILLO_CLARO, fg=NEGRO,
                 font=("Segoe UI", 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(12, 2))

        self.lbl_nombre = tk.Label(panel, text="— Seleccione una fila —",
                                   bg=AMARILLO_CLARO, fg=NEGRO, anchor="w",
                                   font=("Segoe UI", 14, "bold"), wraplength=300,
                                   justify="left")
        self.lbl_nombre.grid(row=1, column=0, sticky="ew", padx=14)

        self.lbl_actual = tk.Label(panel, text="", bg=AMARILLO_CLARO, fg=GRIS,
                                   anchor="w", font=("Consolas", 12))
        self.lbl_actual.grid(row=2, column=0, sticky="ew", padx=14, pady=(2, 8))

        # Comparación ANTES / NUEVO
        comp = tk.Frame(panel, bg=BLANCO, highlightbackground="#D8C98A",
                        highlightthickness=1)
        comp.grid(row=3, column=0, sticky="ew", padx=14, pady=4)
        tk.Label(comp, text="ANTES:", bg=BLANCO, fg=GRIS,
                 font=("Segoe UI", 10, "bold")).grid(row=0, column=0, padx=10, pady=(8, 2), sticky="w")
        self.lbl_antes = tk.Label(comp, text="—", bg=BLANCO, fg=GRIS,
                                  font=("Consolas", 13, "bold"))
        self.lbl_antes.grid(row=0, column=1, sticky="w")
        tk.Label(comp, text="NUEVO:", bg=BLANCO, fg=NEGRO,
                 font=("Segoe UI", 10, "bold")).grid(row=1, column=0, padx=10, pady=(2, 8), sticky="w")
        self.lbl_nuevo = tk.Label(comp, text="—", bg=BLANCO, fg=NEGRO,
                                  font=("Consolas", 13, "bold"))
        self.lbl_nuevo.grid(row=1, column=1, sticky="w")

        self.lbl_instruccion = tk.Label(
            panel, bg=AMARILLO_CLARO, fg=GRIS, wraplength=300, justify="left",
            font=("Segoe UI", 9),
            text="1) Seleccione una fila  2) Capturar Nueva Coordenada  "
                 f"3) Mueva el mouse y presione {TECLA_CAPTURA}  4) Confirme.",
        )
        self.lbl_instruccion.grid(row=4, column=0, sticky="ew", padx=14, pady=(8, 4))

        # Botones de acción
        self.btn_capturar = ttk.Button(panel, text="Capturar Nueva Coordenada",
                                       style="Primario.TButton",
                                       command=self.iniciar_captura)
        self.btn_guardar = ttk.Button(panel, text="Guardar Cambio",
                                      style="Exito.TButton",
                                      command=self.guardar_cambio)
        self.btn_cancelar = ttk.Button(panel, text="Cancelar Captura",
                                       style="Peligro.TButton",
                                       command=self.cancelar_captura)
        self.btn_capturar.grid(row=5, column=0, sticky="ew", padx=14, pady=(8, 4))
        self.btn_guardar.grid(row=6, column=0, sticky="ew", padx=14, pady=4)
        self.btn_cancelar.grid(row=7, column=0, sticky="ew", padx=14, pady=4)

        fila = tk.Frame(panel, bg=AMARILLO_CLARO)
        fila.grid(row=8, column=0, sticky="ew", padx=14, pady=(8, 12))
        for i in range(3):
            fila.columnconfigure(i, weight=1)
        ttk.Button(fila, text="Actualizar Lista", style="Neutro.TButton",
                   command=self.cargar_coordenadas).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ttk.Button(fila, text="Generar Reporte", style="Neutro.TButton",
                   command=self.reporte_manual).grid(row=0, column=1, sticky="ew", padx=3)
        ttk.Button(fila, text="Salir", style="Neutro.TButton",
                   command=self.salir).grid(row=0, column=2, sticky="ew", padx=(3, 0))

    def _construir_historial(self, padre):
        marco = tk.Frame(padre, bg=BLANCO)
        marco.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(10, 10))
        marco.rowconfigure(1, weight=1)
        marco.columnconfigure(0, weight=1)

        tk.Label(marco, text="Últimos cambios realizados", bg=BLANCO, fg=NEGRO,
                 font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))

        cols = ("hora", "nombre", "antes", "nuevo")
        self.tabla_hist = ttk.Treeview(marco, columns=cols, show="headings", height=5)
        for col, texto, ancho in (("hora", "Fecha y hora", 160),
                                  ("nombre", "Coordenada", 300),
                                  ("antes", "Antes", 130),
                                  ("nuevo", "Nuevo", 130)):
            self.tabla_hist.heading(col, text=texto)
            self.tabla_hist.column(col, width=ancho, anchor="center" if col != "nombre" else "w")
        self.tabla_hist.grid(row=1, column=0, sticky="nsew")
        sb = ttk.Scrollbar(marco, orient="vertical", command=self.tabla_hist.yview)
        self.tabla_hist.configure(yscrollcommand=sb.set)
        sb.grid(row=1, column=1, sticky="ns")

    # ------------------------------------------------------
    # Utilidades de UI
    # ------------------------------------------------------

    def _estado(self, texto):
        self.lbl_estado.config(text=texto)

    def _nombre_seleccionado(self):
        sel = self.tabla.selection()
        if not sel:
            return None
        return self.coordenadas[int(sel[0])]

    def _actualizar_botones(self):
        sel = self._nombre_seleccionado() is not None
        hay_pend = self.pendiente is not None
        self.btn_capturar.config(
            state="normal" if sel and not hay_pend and not self.capturando else "disabled")
        self.btn_guardar.config(state="normal" if hay_pend else "disabled")
        self.btn_cancelar.config(
            state="normal" if hay_pend or self.capturando else "disabled")

    def _poblar_tabla(self):
        """Rellena la tabla aplicando el filtro de búsqueda."""
        previo = self._nombre_seleccionado()
        previo = previo["nombre"] if previo else None
        filtro = self.var_filtro.get().strip().upper()

        self.tabla.delete(*self.tabla.get_children())
        for i, c in enumerate(self.coordenadas):
            if filtro and filtro not in c["nombre"]:
                continue
            tags = []
            if self.pendiente and self.pendiente["nombre"] == c["nombre"]:
                tags.append("pendiente")
            elif i % 2:
                tags.append("par")
            self.tabla.insert("", "end", iid=str(i),
                              values=(c["nombre"], c["x"], c["y"]), tags=tags)
            if c["nombre"] == previo:
                self.tabla.selection_set(str(i))
        self._actualizar_botones()

    def _refrescar_detalle(self):
        c = self._nombre_seleccionado()
        if self.pendiente:
            p = self.pendiente
            self.lbl_nombre.config(text=p["nombre"])
            self.lbl_actual.config(text=f"Valor actual: ({p['antes'][0]}, {p['antes'][1]})")
            self.lbl_antes.config(text=f"({p['antes'][0]}, {p['antes'][1]})")
            self.lbl_nuevo.config(text=f"({p['nuevo'][0]}, {p['nuevo'][1]})", fg=AMBAR)
            self.lbl_instruccion.config(
                text="Cambio PENDIENTE: presione «Guardar Cambio» para aplicarlo "
                     "o «Cancelar Captura» para descartarlo.", fg=AMBAR)
        elif c:
            self.lbl_nombre.config(text=c["nombre"])
            self.lbl_actual.config(text=f"Valor actual: ({c['x']}, {c['y']})")
            self.lbl_antes.config(text="—")
            self.lbl_nuevo.config(text="—", fg=NEGRO)
            self.lbl_instruccion.config(
                text="Presione «Capturar Nueva Coordenada» (o doble click en la fila).",
                fg=GRIS)
        else:
            self.lbl_nombre.config(text="— Seleccione una fila —")
            self.lbl_actual.config(text="")
            self.lbl_antes.config(text="—")
            self.lbl_nuevo.config(text="—", fg=NEGRO)
        self._actualizar_botones()

    def _al_seleccionar(self, _evento=None):
        # Con un cambio pendiente no se permite cambiar de fila
        if self.pendiente:
            c = self._nombre_seleccionado()
            if c and c["nombre"] != self.pendiente["nombre"]:
                self._estado("Hay un cambio pendiente: guárdelo o cancélelo primero.")
                for i, item in enumerate(self.coordenadas):
                    if item["nombre"] == self.pendiente["nombre"]:
                        self.tabla.selection_set(str(i))
                        break
        self._refrescar_detalle()

    # ------------------------------------------------------
    # Carga de datos
    # ------------------------------------------------------

    def cargar_coordenadas(self):
        if self.pendiente and self.winfo_busy_pendiente():
            return
        try:
            contenido = leer_archivo(self.ruta_rpa)
        except OSError as e:
            messagebox.showerror("Error de lectura", f"No se pudo leer el archivo:\n{e}")
            return
        self.coordenadas = obtener_coordenadas(contenido)
        self._poblar_tabla()
        self._refrescar_detalle()
        self._estado(f"{len(self.coordenadas)} coordenadas cargadas.")
        if not self.coordenadas:
            messagebox.showinfo(
                "Sin coordenadas",
                "No se encontraron variables con formato NOMBRE = (x, y) en el archivo.")

    def winfo_busy_pendiente(self):
        """Impide recargar la lista con un cambio pendiente sin resolver."""
        messagebox.showinfo("Cambio pendiente",
                            "Guarde o cancele el cambio pendiente antes de actualizar la lista.")
        return True

    def cambiar_archivo(self):
        if self.pendiente or self.capturando:
            messagebox.showinfo("Cambio pendiente",
                                "Guarde o cancele el cambio pendiente antes de cambiar de archivo.")
            return
        ruta = filedialog.askopenfilename(
            title="Seleccione el archivo del RPA",
            filetypes=[("Python", "*.py"), ("Todos", "*.*")])
        if ruta:
            self.ruta_rpa = ruta
            self.lbl_ruta.config(text=ruta)
            self.cargar_coordenadas()

    # ------------------------------------------------------
    # Captura
    # ------------------------------------------------------

    def iniciar_captura(self):
        c = self._nombre_seleccionado()
        if not c:
            messagebox.showwarning("Aviso", "Seleccione primero una coordenada de la lista.")
            return
        if self.pendiente or self.capturando:
            return

        self.capturando = True
        self._captura_actual = c
        self._actualizar_botones()

        self.root.withdraw()  # oculta la ventana para poder ir a Siebel
        self.overlay = OverlayCaptura(self.root, c["nombre"], self.cancelar_captura)
        # Pequeña espera para no capturar una tecla ya presionada
        self.root.after(400, self._sondear_tecla)

    def _sondear_tecla(self):
        """Revisa cada 30 ms la tecla de captura sin bloquear la interfaz."""
        if not self.capturando:
            return
        try:
            if keyboard.is_pressed(TECLA_CANCELAR):
                self.cancelar_captura()
                return
            x, y = pyautogui.position()
            if self.overlay:
                self.overlay.actualizar_posicion(x, y)
            if keyboard.is_pressed(TECLA_CAPTURA):
                self._finalizar_captura(x, y)
                return
        except Exception as e:
            self._cerrar_overlay()
            self.capturando = False
            self.root.deiconify()
            messagebox.showerror("Error de captura", f"No se pudo leer el teclado/mouse:\n{e}")
            self._actualizar_botones()
            return
        self.root.after(30, self._sondear_tecla)

    def _cerrar_overlay(self):
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None

    def _restaurar_ventana(self):
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(200, lambda: self.root.attributes("-topmost", False))
        self.root.focus_force()

    def _finalizar_captura(self, x, y):
        c = self._captura_actual
        self.capturando = False
        self._cerrar_overlay()
        self._restaurar_ventana()

        self.pendiente = {
            "nombre": c["nombre"],
            "antes": (c["x"], c["y"]),
            "nuevo": (x, y),
        }
        self._poblar_tabla()
        self._refrescar_detalle()
        self._estado(f"Capturado ({x}, {y}) para {c['nombre']}.")

        if (x, y) == (c["x"], c["y"]):
            messagebox.showinfo("Sin diferencias",
                                "La nueva coordenada es igual a la actual.", parent=self.root)
            self.pendiente = None
            self._poblar_tabla()
            self._refrescar_detalle()
            return

        if messagebox.askyesno(
            "Confirmar cambio",
            f"{c['nombre']}\n\n"
            f"ANTES:  ({c['x']}, {c['y']})\n"
            f"NUEVO:  ({x}, {y})\n\n"
            "¿Desea guardar el cambio?",
            parent=self.root,
        ):
            self.guardar_cambio(confirmado=True)
        else:
            self._estado("Cambio pendiente: use «Guardar Cambio» o «Cancelar Captura».")

    def cancelar_captura(self):
        """Cancela la captura en curso o descarta el cambio pendiente."""
        if self.capturando:
            self.capturando = False
            self._cerrar_overlay()
            self._restaurar_ventana()
            self._estado("Captura cancelada. No se modificó nada.")
        elif self.pendiente:
            if not messagebox.askyesno("Descartar cambio",
                                       "¿Desea descartar el cambio pendiente?",
                                       parent=self.root):
                return
            self.pendiente = None
            self._estado("Cambio descartado. No se modificó nada.")
            self._poblar_tabla()
        self._refrescar_detalle()

    # ------------------------------------------------------
    # Guardado
    # ------------------------------------------------------

    def guardar_cambio(self, confirmado=False):
        """Guarda el cambio pendiente. Devuelve True si se guardó."""
        p = self.pendiente
        if not p:
            messagebox.showinfo("Aviso", "No hay ningún cambio pendiente.")
            return False

        if not confirmado and not messagebox.askyesno(
            "Confirmar cambio",
            f"{p['nombre']}\n\nANTES:  ({p['antes'][0]}, {p['antes'][1]})\n"
            f"NUEVO:  ({p['nuevo'][0]}, {p['nuevo'][1]})\n\n¿Desea guardar el cambio?",
            parent=self.root,
        ):
            return False

        try:
            ruta_backup = crear_backup(self.ruta_rpa)
            self.backups.append(ruta_backup)
            actualizar_coordenada(self.ruta_rpa, p["nombre"], *p["nuevo"])
        except Exception as e:
            messagebox.showerror(
                "Error al guardar",
                f"No se pudo actualizar el archivo:\n{e}\n\n"
                "El archivo original se conserva y existe un backup.",
                parent=self.root)
            return False

        registro = {
            "hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "nombre": p["nombre"],
            "antes": p["antes"],
            "nuevo": p["nuevo"],
        }
        self.cambios.append(registro)
        self.tabla_hist.insert(
            "", 0, values=(registro["hora"], registro["nombre"],
                           f"({p['antes'][0]}, {p['antes'][1]})",
                           f"({p['nuevo'][0]}, {p['nuevo'][1]})"))

        self.pendiente = None
        self.coordenadas = obtener_coordenadas(leer_archivo(self.ruta_rpa))
        self._poblar_tabla()
        self._refrescar_detalle()
        self._estado(f"Guardado: {p['nombre']} = ({p['nuevo'][0]}, {p['nuevo'][1]}). "
                     f"Backup: {os.path.basename(ruta_backup)}")
        return True

    # ------------------------------------------------------
    # Reporte y salida
    # ------------------------------------------------------

    def reporte_manual(self):
        try:
            ruta = generar_reporte(self.ruta_rpa, self.usuario, self.cambios, self.backups)
        except OSError as e:
            messagebox.showerror("Error", f"No se pudo generar el reporte:\n{e}")
            return
        self._estado(f"Reporte generado: {ruta}")
        messagebox.showinfo("Reporte generado", f"Reporte guardado en:\n\n{ruta}")

    def salir(self):
        if self.capturando:
            self.cancelar_captura()

        if not messagebox.askyesno("Salir", "¿Desea salir?", parent=self.root):
            return

        if self.pendiente:
            r = messagebox.askyesnocancel(
                "Cambios pendientes",
                f"Hay un cambio pendiente en {self.pendiente['nombre']}.\n\n"
                "¿Desea guardar antes de salir?",
                parent=self.root)
            if r is None:
                return                      # Cancelar: no salir
            if r and not self.guardar_cambio(confirmado=True):
                return                      # falló el guardado: no salir
            self.pendiente = None

        ruta = None
        try:
            ruta = generar_reporte(self.ruta_rpa, self.usuario, self.cambios, self.backups)
        except OSError as e:
            messagebox.showwarning("Reporte", f"No se pudo generar el reporte:\n{e}")

        if ruta:
            messagebox.showinfo("Reporte generado", f"Reporte final guardado en:\n\n{ruta}")
        self.root.destroy()


# ==========================================================
# MAIN
# ==========================================================

def main():
    root = tk.Tk()
    AdministradorCoordenadas(root)
    try:
        root.mainloop()
    except tk.TclError:
        pass  # la ventana se destruyó durante el inicio (usuario canceló)


if __name__ == "__main__":
    main()
