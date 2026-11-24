import os
import re
import shutil
import time
import unicodedata
from datetime import datetime
from getpass import getpass # futura limpieza
from pathlib import Path

import pandas as pd
import pyautogui
import ctypes
import tkinter as tk # futura limpieza
import pyperclip

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

import threading

# ============================================================
# VALIDACION DE PANTALLA
# ============================================================

def validar_pantalla():

    ancho = ctypes.windll.user32.GetSystemMetrics(0)
    alto = ctypes.windll.user32.GetSystemMetrics(1)

    root = tk.Tk()

    escala = root.winfo_fpixels("1i") / 96

    root.destroy()

    escala_porcentaje = round(
        escala * 100
    )

    if (
        ancho != 1920
        or alto != 1200
        or escala_porcentaje != 125
    ):

        mensaje = (
            "\nERROR DE CONFIGURACION\n\n"
            f"Resolucion detectada: {ancho}x{alto}\n"
            f"Escala detectada: {escala_porcentaje}%\n\n"
            "Configuracion requerida:\n"
            "Resolucion: 1920x1200\n"
            "Escala: 125%\n\n"
            "Corrige la configuracion y vuelve a ejecutar el RPA."
        )

        ctypes.windll.user32.MessageBoxW(
            0,
            mensaje,
            "RPA BLOQUEADO",
            16
        )

        raise SystemExit(
            mensaje
        )

    print(
        "\n✅ Resolucion validada: 1920x1200"
    )

    print(
        "✅ Escala validada: 125%"
    )

# ============================================================
# CONFIGURACION GENERAL
# ============================================================

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 1

validar_pantalla()

TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

ESPERA_EXPORTACION = 120

FECHA_CIERRE = input(
    "Ingresa la fecha de cierre (DD/MM/AAAA): "
).strip()

if not FECHA_CIERRE:
    raise ValueError("No se ingreso fecha de cierre.")

FECHA_CIERRE_FILTRO = f">={FECHA_CIERRE}"


# ============================================================
# COORDENADAS GENERALES DE SIEBEL
# ============================================================
# ============================================================
# Coordenadas SIEBEL
# ============================================================



HERRAMIENTAS = (508, 117)
SIEBEL = (508, 207)
COORDENADA_1 = (1058, 691)
COORDENADA_2 = (1212, 690)
LOGIN = (1354, 680)

FLECHA_MODULOS = (1806, 235)

MENU = (270, 270)
MENU_CONTRATOS = (274, 264)
MENU_ACTIVIDADES = (355, 270)

SIGUIENTE = (1685, 845)
GUARDAR_SIEBEL = (1032, 599)


# ============================================================
# CONSULTA DE FACTURAS
# ============================================================

MODULO_FACTURAS = (1761, 231)
CONSULTA_FACTURAS = (424, 267)
NUMERO_FACTURA = (84, 328)
EXPORTAR_FACTURAS = (309, 593)


# ============================================================
# CONSULTA DE CONTRATOS
# ============================================================

MODULO_CONTRATOS = (590, 236)
MIS_CONTRATOS = (101, 269)
TODOS_CONTRATOS = (99, 302)

CONSULTA_CONTRATOS = (502, 270)
NUMERO_CONTRATO = (491, 319)
EXPORTAR_CONTRATOS = (310, 617)


# ============================================================
# ACTIVIDADES ARGENTINA
# ============================================================

CONSULTA_ACTIVIDADES_ARG = (1774, 435)
CONSULTA_ACTIVIDADES = (512, 269)
CAMPO_TIPO = (423, 320)
CAMPO_CATEGORIA = (540, 322)
CAMPO_ESTADO = (658, 321)

FLECHA_DESPLAZAMIENTO = (1872, 608)
CAMPO_REGION = (841, 322)
CAMPO_FECHA_CIERRE = (673, 318)
EXPORTAR_ACTIVIDADES = (391, 614)

# ============================================================
# FUNCIONES GENERALES
# ============================================================

def buscar_carpeta_descargas():

    home = Path.home()

    opciones = [
        home / "Downloads",
        home / "Descargas",
        home / "OneDrive" / "Downloads",
        home / "OneDrive" / "Descargas",
    ]

    for carpeta_onedrive in home.glob("OneDrive*"):

        opciones.append(
            carpeta_onedrive / "Downloads"
        )

        opciones.append(
            carpeta_onedrive / "Descargas"
        )

    for carpeta in opciones:

        if carpeta.exists():

            return carpeta

    raise FileNotFoundError(
        "No se encontro la carpeta Descargas."
    )


def limpiar_excel(valor):

    if valor is None:

        return ""

    try:

        if pd.isna(valor):

            return ""

    except (TypeError, ValueError):

        pass

    texto = str(valor)

    texto = texto.replace("\ufeff", "")
    texto = texto.replace("ÿþ", "")
    texto = texto.replace("þÿ", "")
    texto = texto.replace("\xa0", " ")
    texto = texto.replace("\u200b", "")

    texto = re.sub(
        r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]",
        "",
        texto
    )

    return texto.strip()


def limpiar_codigo(valor):

    texto = limpiar_excel(valor)

    texto = re.sub(
        r"\s+",
        "",
        texto
    )

    valores_ignorados = {
        "",
        "0",
        "0.0",
        "nan",
        "none",
        "total",
        "suma",
        "recuento",
    }

    if texto.lower() in valores_ignorados:

        return ""

    if re.fullmatch(
        r"\d+\.0",
        texto
    ):

        texto = texto[:-2]

    return texto


def normalizar(valor):

    texto = limpiar_excel(valor).lower()

    texto = texto.replace("nº", "n")
    texto = texto.replace("n°", "n")
    texto = texto.replace("nro.", "n")
    texto = texto.replace("nro", "n")

    texto = unicodedata.normalize(
        "NFKD",
        texto
    )

    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )

    texto = re.sub(
        r"[^a-z0-9]",
        "",
        texto
    )

    return texto


def buscar_columna(
    dataframe,
    candidatos
):

    mapa_columnas = {
        normalizar(columna): columna
        for columna in dataframe.columns
    }

    for candidato in candidatos:

        clave = normalizar(
            candidato
        )

        if clave in mapa_columnas:

            return mapa_columnas[
                clave
            ]

    for candidato in candidatos:

        clave = normalizar(
            candidato
        )

        for (
            columna_normalizada,
            columna_original
        ) in mapa_columnas.items():

            if (
                clave in columna_normalizada
                or columna_normalizada in clave
            ):

                return columna_original

    return None


def detectar_encoding_csv(ruta):

    primeros_bytes = ruta.read_bytes()[:4]

    if primeros_bytes.startswith(
        (
            b"\xff\xfe",
            b"\xfe\xff",
        )
    ):

        return "utf-16"

    if primeros_bytes.startswith(
        b"\xef\xbb\xbf"
    ):

        return "utf-8-sig"

    for encoding in [
        "utf-8",
        "cp1252",
        "latin-1",
    ]:

        try:

            ruta.read_text(
                encoding=encoding
            )

            return encoding

        except UnicodeDecodeError:

            continue

    return "cp1252"


def detectar_separador_csv(
    ruta,
    encoding
):

    contenido = ruta.read_text(
        encoding=encoding,
        errors="replace"
    )

    muestra = contenido[:20000]

    if "\t" in muestra:

        return "\t"

    if ";" in muestra:

        return ";"

    if "," in muestra:

        return ","

    if "|" in muestra:

        return "|"

    raise ValueError(
        "No se pudo detectar el separador "
        f"de {ruta.name}."
    )


def leer_csv_siebel(ruta):

    encoding = detectar_encoding_csv(
        ruta
    )

    separador = detectar_separador_csv(
        ruta,
        encoding
    )

    print(
        f"\nLeyendo {ruta.name}: "
        f"encoding={encoding}, "
        f"separador={repr(separador)}"
    )

    dataframe = pd.read_csv(
        ruta,
        sep=separador,
        dtype=str,
        encoding=encoding,
        engine="python",
        keep_default_na=False,
        on_bad_lines="skip"
    )

    dataframe.columns = [
        limpiar_excel(columna)
        for columna in dataframe.columns
    ]

    nombres_usados = {}
    nombres_nuevos = []

    for columna in dataframe.columns:

        nombre_base = (
            columna
            or "Columna"
        )

        nombres_usados[nombre_base] = (
            nombres_usados.get(
                nombre_base,
                0
            )
            + 1
        )

        if nombres_usados[nombre_base] == 1:

            nombres_nuevos.append(
                nombre_base
            )

        else:

            nombres_nuevos.append(
                f"{nombre_base}_"
                f"{nombres_usados[nombre_base]}"
            )

    dataframe.columns = nombres_nuevos

    for columna in dataframe.columns:

        dataframe[columna] = (
            dataframe[columna]
            .apply(limpiar_excel)
        )

    columnas_con_datos = [
        columna
        for columna in dataframe.columns
        if (
            dataframe[columna]
            .astype(str)
            .str.strip()
            .ne("")
            .any()
        )
    ]

    dataframe = dataframe[
        columnas_con_datos
    ].copy()

    dataframe.reset_index(
        drop=True,
        inplace=True
    )

    print(
        f"Registros leidos: "
        f"{len(dataframe)}"
    )

    print(
        "Columnas detectadas:"
    )

    for columna in dataframe.columns:

        print(
            f" - {columna}"
        )

    return dataframe


def guardar_tabla(
    dataframe,
    ruta,
    nombre_tabla,
    columnas_texto
):

    dataframe = dataframe.copy()

    dataframe.columns = [
        limpiar_excel(columna)
        for columna in dataframe.columns
    ]

    for columna in dataframe.columns:

        dataframe[columna] = (
            dataframe[columna]
            .apply(limpiar_excel)
        )

    for columna in columnas_texto:

        if columna in dataframe.columns:

            dataframe[columna] = (
                dataframe[columna]
                .apply(limpiar_codigo)
            )

    if ruta.exists():

        ruta.unlink()

    dataframe.to_excel(
        ruta,
        index=False,
        engine="openpyxl"
    )

    libro = load_workbook(
        ruta
    )

    hoja = libro.active

    hoja.title = "Datos"

    hoja.freeze_panes = "A2"

    hoja.auto_filter.ref = (
        hoja.dimensions
    )

    encabezados = {
        celda.value: celda.column
        for celda in hoja[1]
    }

    for columna in columnas_texto:

        if columna in encabezados:

            numero_columna = (
                encabezados[columna]
            )

            for fila in range(
                2,
                hoja.max_row + 1
            ):

                celda = hoja.cell(
                    row=fila,
                    column=numero_columna
                )

                celda.number_format = "@"

                if celda.value is not None:

                    celda.value = str(
                        celda.value
                    )

    if hoja.max_row >= 2:

        referencia = (
            f"A1:"
            f"{hoja.cell(hoja.max_row, hoja.max_column).coordinate}"
        )

        tabla = Table(
            displayName=nombre_tabla,
            ref=referencia
        )

        tabla.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showRowStripes=True,
            showFirstColumn=False,
            showLastColumn=False,
            showColumnStripes=False
        )

        hoja.add_table(
            tabla
        )

    for columna_celdas in hoja.columns:

        letra = (
            columna_celdas[0]
            .column_letter
        )

        ancho = max(
            (
                len(
                    str(celda.value)
                )
                for celda in columna_celdas
                if celda.value is not None
            ),
            default=0
        )

        hoja.column_dimensions[
            letra
        ].width = min(
            ancho + 2,
            45
        )

    libro.save(
        ruta
    )


def mostrar_click(
    x,
    y,
    duracion=0.8
):

    pass

def abrir_edge(
    maximizar=False
):
    

    print(
        "\nAbriendo Microsoft Edge..."
    )

    os.system(
        "start microsoft-edge:"
    )

    print(
        "Esperando 6 segundos..."
    )

    time.sleep(6)

    if maximizar:

        pyautogui.hotkey(
            "win",
            "up"
        )

        time.sleep(2)


# ============================================================
# ============================================================

def login_siebel():

    pyautogui.click(
        *HERRAMIENTAS
    )

    time.sleep(3.5)

    pyautogui.click(
        *SIEBEL
    )

    time.sleep(3.5)

    # Primer click
    pyautogui.click(
        *COORDENADA_1
    )

    time.sleep(3.5)

    # Segundo click
    pyautogui.click(
        *COORDENADA_2
    )

    time.sleep(3.5)

    # Ingresar
    pyautogui.click(
        *LOGIN
    )

    time.sleep(15)

# ============================================================
# ============================================================

def esperar_archivo(
    ruta,
    limite=ESPERA_EXPORTACION
):

    fecha_limite = (
        time.time()
        + limite
    )

    while time.time() < fecha_limite:

        if (
            ruta.exists()
            and ruta.stat().st_size > 0
        ):

            tamano_inicial = (
                ruta.stat().st_size
            )

            time.sleep(2)

            tamano_final = (
                ruta.stat().st_size
            )

            if tamano_inicial == tamano_final:

                return

        time.sleep(1)

    raise FileNotFoundError(
        f"No se detecto el archivo: "
        f"{ruta}"
    )


def exportar(
    ruta,
    coordenada_exportar,
    coordenada_menu=None
):

    if coordenada_menu is None:

        coordenada_menu = MENU

    print(
        "\nIniciando exportacion:"
    )

    print(
        ruta.name
    )

    pyautogui.click(
        *coordenada_menu
    )

    time.sleep(7)

    pyautogui.click(
        *coordenada_exportar
    )

    time.sleep(7)

    pyautogui.click(
        *SIGUIENTE
    )

    time.sleep(16)

    pyautogui.click(
        *GUARDAR_SIEBEL
    )

    time.sleep(8)

    pyautogui.hotkey(
        "alt",
        "n"
    )

    time.sleep(2)

    pyautogui.hotkey(
        "ctrl",
        "a"
    )

    time.sleep(1)

    pyperclip.copy(
        str(ruta)
    )

    pyautogui.hotkey(
        "ctrl",
        "v"
    )

    time.sleep(3)

    pyautogui.press(
        "enter"
    )

    print(
        f"Esperando exportacion: "
        f"{ruta}"
    )

    esperar_archivo(
        ruta
    )

    print(
        "Exportacion completada."
    )

    time.sleep(3)

    pyautogui.hotkey(
        "alt",
        "f4"
    )

    time.sleep(3)


def escribir_campo(
    coordenada,
    texto
):

    pyautogui.click(
        *coordenada
    )

    time.sleep(2)

    pyautogui.hotkey(
        "ctrl",
        "a"
    )

    pyperclip.copy(
        texto
    )

    pyautogui.hotkey(
        "ctrl",
        "v"
    )

    time.sleep(2)


def localizar_configuracion():

    patrones = [
        "ConfiguracionFacturacion*.xlsx",
        "ConfiguraciónFacturacion*.xlsx",
        "Configuracion Facturacion*.xlsx",
        "Configuración Facturación*.xlsx",
        "ConfiguracionValidar*.xlsx",
        "ConfiguraciónValidar*.xlsx",
    ]

    archivos_encontrados = []

    for patron in patrones:

        archivos_encontrados.extend(
            CARPETA_RPA1.glob(
                patron
            )
        )

    archivos_encontrados = [
        archivo
        for archivo in archivos_encontrados
        if (
            archivo.is_file()
            and not archivo.name.startswith("~$")
            and not str(archivo).startswith(
                str(CARPETA)
            )
        )
    ]

    if not archivos_encontrados:

        raise FileNotFoundError(
            "No se encontro "
            "ConfiguracionFacturacion.xlsx "
            "o ConfiguracionValidar.xlsx "
            "en Descargas/RPA1."
        )

    return max(
        archivos_encontrados,
        key=lambda archivo:
        archivo.stat().st_mtime
    )


def buscar_hoja(
    libro,
    nombre_buscado
):

    nombre_normalizado = (
        normalizar(
            nombre_buscado
        )
    )

    for nombre_real in libro.sheetnames:

        if (
            normalizar(nombre_real)
            == nombre_normalizado
        ):

            return libro[
                nombre_real
            ]

    return None


def buscar_columna_hoja(
    hoja,
    candidatos
):

    mapa_columnas = {
        normalizar(celda.value):
        celda.column
        for celda in hoja[1]
        if celda.value is not None
    }

    for candidato in candidatos:

        clave = normalizar(
            candidato
        )

        if clave in mapa_columnas:

            return mapa_columnas[
                clave
            ]

    for candidato in candidatos:

        clave = normalizar(
            candidato
        )

        for (
            columna_normalizada,
            numero_columna
        ) in mapa_columnas.items():

            if (
                clave in columna_normalizada
                or columna_normalizada in clave
            ):

                return numero_columna

    return None


def mostrar_cabeceras(
    hoja
):

    print(
        f"\nCabeceras disponibles "
        f"en la hoja {hoja.title}:"
    )

    for celda in hoja[1]:

        if celda.value is not None:

            print(
                f" - {repr(celda.value)} "
                f"=> {normalizar(celda.value)}"
            )


def ajustar_tablas(
    hoja,
    ultima_fila
):

    for tabla in hoja.tables.values():

        referencia_actual = (
            tabla.ref
        )

        celda_inicial = (
            referencia_actual
            .split(":")[0]
        )

        celda_final = (
            referencia_actual
            .split(":")[1]
        )

        ultima_columna = re.sub(
            r"\d+",
            "",
            celda_final
        )

        tabla.ref = (
            f"{celda_inicial}:"
            f"{ultima_columna}"
            f"{ultima_fila}"
        )


# ============================================================
# ACTUALIZAR ARCHIVO DE CONFIGURACION
#
# FUENTE:
# NC_Generadas.xlsx
# columna Nº Factura
#
# DESTINO:
# ConfiguracionFacturacion / ConfiguracionValidar
# hoja Notadecredito
# columna FACTURA RELACIONADA
# ============================================================

def actualizar_configuracion(
    ruta_nc_generadas,
    ruta_gestion_contrato
):

    ruta_configuracion_original = (
        localizar_configuracion()
    )

    ruta_configuracion_final = (
        CARPETA
        / ruta_configuracion_original.name
    )

    print(
        "\nConfiguracion original:"
    )

    print(
        ruta_configuracion_original
    )

    if (
        ruta_configuracion_original.resolve()
        != ruta_configuracion_final.resolve()
    ):

        try:

            shutil.copy2(
                ruta_configuracion_original,
                ruta_configuracion_final
            )

        except PermissionError as error:

            raise PermissionError(
                "No se pudo copiar el archivo "
                "de configuracion. "
                "Cierra el archivo en Excel."
            ) from error

    print(
        "\nConfiguracion de trabajo:"
    )

    print(
        ruta_configuracion_final
    )

    # ========================================================
    # LEER NC_GENERADAS.XLSX
    # ========================================================

    df_nc_generadas = pd.read_excel(
        ruta_nc_generadas,
        sheet_name=0,
        dtype=str,
        engine="openpyxl",
        keep_default_na=False
    )

    df_nc_generadas.columns = [
        limpiar_excel(columna)
        for columna
        in df_nc_generadas.columns
    ]

    for columna in df_nc_generadas.columns:

        df_nc_generadas[columna] = (
            df_nc_generadas[columna]
            .apply(limpiar_excel)
        )

    columna_factura_nc = buscar_columna(
        df_nc_generadas,
        [
            "Nº Factura",
            "N° Factura",
            "N Factura",
            "Numero Factura",
            "Número Factura",
        ]
    )

    if not columna_factura_nc:

        raise ValueError(
            "No se encontro la columna "
            "Nº Factura en NC_Generadas.xlsx."
        )

    # ========================================================
    # PREPARAR TODAS LAS FILAS DE Nº FACTURA
    #
    # IMPORTANTE:
    # NO FILTRAR VACIOS.
    # NO ELIMINAR DUPLICADOS.
    # UNA POSICION POR CADA FILA DEL OUTPUT 3.
    # ========================================================

    facturas_relacionadas = []

    for factura in df_nc_generadas[
        columna_factura_nc
    ]:

        facturas_relacionadas.append(
            limpiar_codigo(
                factura
            )
        )

    if (
        len(facturas_relacionadas)
        != len(df_nc_generadas)
    ):

        raise ValueError(
            "Error preparando Nº Factura. "
            f"NC_Generadas={len(df_nc_generadas)}, "
            f"preparadas={len(facturas_relacionadas)}"
        )

    if not any(
        facturas_relacionadas
    ):

        raise ValueError(
            "La columna Nº Factura "
            "de NC_Generadas.xlsx "
            "no contiene valores validos."
        )

    print(
        f"\nFilas leidas desde "
        f"NC_Generadas.xlsx: "
        f"{len(df_nc_generadas)}"
    )

    print(
        f"Filas preparadas para "
        f"FACTURA RELACIONADA: "
        f"{len(facturas_relacionadas)}"
    )

    # ========================================================
    # LEER GESTIONCONTRATO.XLSX
    # ========================================================

    df_gestion_archivo = pd.read_excel(
        ruta_gestion_contrato,
        sheet_name=0,
        dtype=str,
        engine="openpyxl",
        keep_default_na=False
    )

    df_gestion_archivo.columns = [
        limpiar_excel(columna)
        for columna
        in df_gestion_archivo.columns
    ]

    for columna in df_gestion_archivo.columns:

        df_gestion_archivo[columna] = (
            df_gestion_archivo[columna]
            .apply(limpiar_excel)
        )

    columna_contrato_gestion = (
        buscar_columna(
            df_gestion_archivo,
            [
                "N° Contrato",
                "Nº Contrato",
                "N Contrato",
                "Numero Contrato",
                "Número Contrato",
            ]
        )
    )

    columna_cliente_gestion = (
        buscar_columna(
            df_gestion_archivo,
            [
                "Código Cliente",
                "Codigo Cliente",
                "Número Cliente",
                "Numero Cliente",
            ]
        )
    )

    if not columna_contrato_gestion:

        raise ValueError(
            "No se encontro N° Contrato "
            "en GestionContrato.xlsx."
        )

    if not columna_cliente_gestion:

        raise ValueError(
            "No se encontro Código Cliente "
            "en GestionContrato.xlsx."
        )

    filas_gestion = []

    for contrato, cliente in zip(
        df_gestion_archivo[
            columna_contrato_gestion
        ],
        df_gestion_archivo[
            columna_cliente_gestion
        ]
    ):

        filas_gestion.append(
            (
                limpiar_codigo(
                    contrato
                ),
                limpiar_codigo(
                    cliente
                ),
            )
        )

    print(
        f"Filas leidas desde "
        f"GestionContrato.xlsx: "
        f"{len(df_gestion_archivo)}"
    )

    print(
        f"Filas preparadas para "
        f"GestionContrato: "
        f"{len(filas_gestion)}"
    )

    # ========================================================
    # ABRIR CONFIGURACION
    # ========================================================

    try:

        libro_configuracion = (
            load_workbook(
                ruta_configuracion_final
            )
        )

    except PermissionError as error:

        raise PermissionError(
            "No se pudo abrir el archivo "
            "de configuracion. "
            "Cierra el archivo en Excel."
        ) from error

    hoja_nota_credito = buscar_hoja(
        libro_configuracion,
        "Notadecredito"
    )

    hoja_gestion_contrato = buscar_hoja(
        libro_configuracion,
        "GestionContrato"
    )

    if hoja_nota_credito is None:

        raise ValueError(
            "No se encontro la hoja "
            "Notadecredito."
        )

    if hoja_gestion_contrato is None:

        raise ValueError(
            "No se encontro la hoja "
            "GestionContrato."
        )

    # ========================================================
    # BUSCAR COLUMNAS DESTINO
    # ========================================================

    columna_factura_relacionada = (
        buscar_columna_hoja(
            hoja_nota_credito,
            [
                "FACTURA RELACIONADA"
            ]
        )
    )

    columna_contrato_destino = (
        buscar_columna_hoja(
            hoja_gestion_contrato,
            [
                "N° Contrato",
                "Nº Contrato",
                "N Contrato",
                "Numero Contrato",
                "Número Contrato",
            ]
        )
    )

    columna_cliente_destino = (
        buscar_columna_hoja(
            hoja_gestion_contrato,
            [
                "Código Cliente",
                "Codigo Cliente",
                "Número Cliente",
                "Numero Cliente",
            ]
        )
    )

    if not columna_factura_relacionada:

        mostrar_cabeceras(
            hoja_nota_credito
        )

        raise ValueError(
            "No se encontro "
            "FACTURA RELACIONADA "
            "en la hoja Notadecredito."
        )

    if not columna_contrato_destino:

        mostrar_cabeceras(
            hoja_gestion_contrato
        )

        raise ValueError(
            "No se encontro N° Contrato "
            "en la hoja GestionContrato."
        )

    if not columna_cliente_destino:

        mostrar_cabeceras(
            hoja_gestion_contrato
        )

        raise ValueError(
            "No se encontro Código Cliente "
            "en la hoja GestionContrato."
        )

    # ========================================================
    # PEGAR TODAS LAS FACTURAS
    #
    # SI HAY 779 FILAS:
    # PEGA DESDE FILA 2 HASTA FILA 780
    # ========================================================

    limite_limpieza_facturas = max(
        hoja_nota_credito.max_row,
        len(facturas_relacionadas) + 1
    )

    for numero_fila in range(
        2,
        limite_limpieza_facturas + 1
    ):

        hoja_nota_credito.cell(
            row=numero_fila,
            column=columna_factura_relacionada
        ).value = None

    for numero_fila, factura in enumerate(
        facturas_relacionadas,
        start=2
    ):

        celda = (
            hoja_nota_credito.cell(
                row=numero_fila,
                column=columna_factura_relacionada
            )
        )

        celda.value = factura

        celda.number_format = "@"

    ultima_fila_facturas = (
        len(facturas_relacionadas)
        + 1
    )

    ajustar_tablas(
        hoja_nota_credito,
        ultima_fila_facturas
    )

    print(
        f"\nFACTURA RELACIONADA "
        f"escrita desde fila 2 "
        f"hasta fila "
        f"{ultima_fila_facturas}."
    )

    print(
        f"Posiciones escritas en "
        f"Notadecredito: "
        f"{len(facturas_relacionadas)}"
    )

    # ========================================================
    # PEGAR GESTION CONTRATO
    # ========================================================

    limite_limpieza_gestion = max(
        hoja_gestion_contrato.max_row,
        len(filas_gestion) + 1
    )

    for numero_fila in range(
        2,
        limite_limpieza_gestion + 1
    ):

        hoja_gestion_contrato.cell(
            row=numero_fila,
            column=columna_contrato_destino
        ).value = None

        hoja_gestion_contrato.cell(
            row=numero_fila,
            column=columna_cliente_destino
        ).value = None

    for numero_fila, (
        contrato,
        cliente
    ) in enumerate(
        filas_gestion,
        start=2
    ):

        celda_contrato = (
            hoja_gestion_contrato.cell(
                row=numero_fila,
                column=columna_contrato_destino
            )
        )

        celda_cliente = (
            hoja_gestion_contrato.cell(
                row=numero_fila,
                column=columna_cliente_destino
            )
        )

        celda_contrato.value = contrato
        celda_contrato.number_format = "@"

        celda_cliente.value = cliente
        celda_cliente.number_format = "@"

    ultima_fila_gestion = max(
        2,
        len(filas_gestion) + 1
    )

    ajustar_tablas(
        hoja_gestion_contrato,
        ultima_fila_gestion
    )

    # ========================================================
    # GUARDAR
    # ========================================================

    try:

        libro_configuracion.save(
            ruta_configuracion_final
        )

    except PermissionError as error:

        raise PermissionError(
            "No se pudo guardar el archivo "
            "de configuracion. "
            "Cierra el archivo en Excel."
        ) from error

    # ========================================================
    # VALIDAR TODAS LAS FACTURAS
    # ========================================================

    libro_validacion_final = (
        load_workbook(
            ruta_configuracion_final,
            data_only=False
        )
    )

    hoja_nc_validacion = buscar_hoja(
        libro_validacion_final,
        "Notadecredito"
    )

    hoja_gc_validacion = buscar_hoja(
        libro_validacion_final,
        "GestionContrato"
    )

    facturas_guardadas = []

    for numero_fila in range(
        2,
        len(facturas_relacionadas) + 2
    ):

        valor_guardado = (
            hoja_nc_validacion.cell(
                row=numero_fila,
                column=columna_factura_relacionada
            ).value
        )

        facturas_guardadas.append(
            limpiar_codigo(
                valor_guardado
            )
        )

    if (
        len(facturas_guardadas)
        != len(facturas_relacionadas)
    ):

        raise ValueError(
            "Cantidad incorrecta de "
            "facturas guardadas. "
            f"Guardadas={len(facturas_guardadas)}, "
            f"Esperadas={len(facturas_relacionadas)}"
        )

    for numero_fila, (
        esperado,
        guardado
    ) in enumerate(
        zip(
            facturas_relacionadas,
            facturas_guardadas
        ),
        start=2
    ):

        if esperado != guardado:

            raise ValueError(
                f"Error en FACTURA RELACIONADA "
                f"fila {numero_fila}: "
                f"esperado={repr(esperado)}, "
                f"guardado={repr(guardado)}"
            )

    # ========================================================
    # VALIDAR TODAS LAS FILAS GESTIONCONTRATO
    # ========================================================

    filas_gestion_guardadas = []

    for numero_fila in range(
        2,
        len(filas_gestion) + 2
    ):

        contrato_guardado = (
            limpiar_codigo(
                hoja_gc_validacion.cell(
                    row=numero_fila,
                    column=columna_contrato_destino
                ).value
            )
        )

        cliente_guardado = (
            limpiar_codigo(
                hoja_gc_validacion.cell(
                    row=numero_fila,
                    column=columna_cliente_destino
                ).value
            )
        )

        filas_gestion_guardadas.append(
            (
                contrato_guardado,
                cliente_guardado,
            )
        )

    for numero_fila, (
        esperado,
        guardado
    ) in enumerate(
        zip(
            filas_gestion,
            filas_gestion_guardadas
        ),
        start=2
    ):

        if esperado != guardado:

            raise ValueError(
                f"Error en GestionContrato "
                f"fila {numero_fila}: "
                f"esperado={repr(esperado)}, "
                f"guardado={repr(guardado)}"
            )

    print(
        "\n=============================================="
    )

    print(
        "CRUCE FINALIZADO CORRECTAMENTE"
    )

    print(
        "=============================================="
    )

    print(
        "\nArchivo actualizado:"
    )

    print(
        ruta_configuracion_final
    )

    print(
        f"Filas copiadas a "
        f"FACTURA RELACIONADA: "
        f"{len(facturas_guardadas)}"
    )

    print(
        f"Filas copiadas a hoja "
        f"GestionContrato: "
        f"{len(filas_gestion_guardadas)}"
    )

    return (
        ruta_configuracion_final,
        len(facturas_guardadas),
        len(filas_gestion_guardadas)
    )


# ============================================================
# PREPARAR RUTAS DEL PROCESO
# ============================================================

DESCARGAS = (
    buscar_carpeta_descargas()
)

CARPETA_RPA1 = (
    DESCARGAS
    / "RPA1"
)

ARCHIVO_SF = (
    CARPETA_RPA1
    / "Solicitudes Salesforce.xlsx"
)

if not ARCHIVO_SF.exists():

    raise FileNotFoundError(
        "No se encontro el archivo exacto "
        "'Solicitudes Salesforce.xlsx' "
        f"en {CARPETA_RPA1}"
    )


CARPETA = (
    DESCARGAS
    / f"CASOS NOTAS DE CREDITOS_{TIMESTAMP}"
)

CARPETA.mkdir(
    parents=True,
    exist_ok=False
)


VALIDACION = (
    CARPETA
    / "ArchivoValidar.xlsx"
)


CSV_FACTURAS = (
    CARPETA
    / f"output_1_{TIMESTAMP}.csv"
)


CONSULTA_SIEBEL = (
    CARPETA
    / "ConsultaSiebel.xlsx"
)


CSV_CONTRATOS = (
    CARPETA
    / f"output_2_{TIMESTAMP}.csv"
)


GESTION_CONTRATO = (
    CARPETA
    / "GestionContrato.xlsx"
)


CSV_ACTIVIDADES = (
    CARPETA
    / f"output_3_{TIMESTAMP}.csv"
)


NC_GENERADAS = (
    CARPETA
    / "NC_Generadas.xlsx"
)


print(
    "\n=============================================="
)

print(
    "INICIO DEL PROCESO"
)

print(
    "=============================================="
)

print(
    "\nDescargas:"
)

print(
    DESCARGAS
)

print(
    "\nReporte Salesforce:"
)

print(
    ARCHIVO_SF.name
)

print(
    "\nCarpeta del proceso:"
)

print(
    CARPETA
)


# ============================================================
# PASO 1: PROCESAR SALESFORCE
# ============================================================

raw = pd.read_excel(
    ARCHIVO_SF,
    sheet_name=0,
    header=None,
    dtype=str,
    engine="openpyxl",
    keep_default_na=False
)


fila_header = next(
    (
        indice
        for indice in range(
            len(raw)
        )
        if "Id. del caso" in (
            raw.iloc[indice]
            .fillna("")
            .astype(str)
            .str.strip()
            .tolist()
        )
    ),
    None
)


if fila_header is None:

    raise ValueError(
        "No se encontro la cabecera "
        "'Id. del caso'."
    )


# ============================================================
# IMPORTANTE:
# LEER TODO COMO TEXTO PARA NO PERDER CEROS INICIALES
# ============================================================

df = pd.read_excel(
    ARCHIVO_SF,
    sheet_name=0,
    header=fila_header,
    dtype=str,
    engine="openpyxl",
    keep_default_na=False
)


df = df.dropna(
    axis=1,
    how="all"
)


if "Núm. Factura" not in df.columns:

    raise ValueError(
        "No se encontro la columna "
        "'Núm. Factura'."
    )


# ============================================================
# LIMPIAR Y VALIDAR FACTURAS
# ============================================================

df["NumFactura_Limpio"] = (
    df["Núm. Factura"]
    .apply(limpiar_codigo)
)


df["Largo_NumFactura"] = (
    df["NumFactura_Limpio"]
    .apply(
        lambda factura:
        len(factura)
        if factura
        else ""
    )
)


def generar_observacion(
    factura,
    largo
):

    if not factura:

        return ""

    if largo != 14:

        return (
            f"REVISAR: factura tiene "
            f"{largo} caracteres; "
            f"debe tener 14"
        )

    return ""


df["Observacion_NumFactura"] = (
    df.apply(
        lambda fila:
        generar_observacion(
            fila[
                "NumFactura_Limpio"
            ],
            fila[
                "Largo_NumFactura"
            ]
        ),
        axis=1
    )
)


# ============================================================
# OR DE FACTURAS
# ============================================================

facturas_validas = (
    df.loc[
        df[
            "Largo_NumFactura"
        ] == 14,
        "NumFactura_Limpio"
    ]
    .tolist()
)


facturas_validas = [
    factura
    for factura
    in facturas_validas
    if factura
]


if not facturas_validas:

    raise ValueError(
        "No existen facturas validas "
        "de 14 caracteres."
    )


OR_FACTURAS = (
    " OR ".join(
        facturas_validas
    )
)


cantidad_observadas = (
    df[
        "Observacion_NumFactura"
    ]
    .ne("")
    .sum()
)


# ============================================================
# EXCEL DE VALIDACION
# ============================================================

df.to_excel(
    VALIDACION,
    index=False,
    engine="openpyxl"
)


libro_validacion = (
    load_workbook(
        VALIDACION
    )
)


hoja_validacion = (
    libro_validacion.active
)


hoja_validacion.title = (
    "Validacion"
)


hoja_validacion.freeze_panes = (
    "A2"
)


hoja_validacion.auto_filter.ref = (
    hoja_validacion.dimensions
)


encabezados_validacion = {
    celda.value:
    celda.column
    for celda
    in hoja_validacion[1]
}


# ============================================================
# CONSERVAR COLUMNAS CLAVE COMO TEXTO
# PARA NO PERDER CEROS INICIALES
# ============================================================

columnas_texto_archivo_validar = [
    "Id. del caso",
    "Número del caso",
    "Contrato de servicio: Nombre",
    "Contrato de servicio: Número de contrato",
    "Código Cliente",
    "Núm. Factura",
    "NumFactura_Limpio",
]


for nombre_columna in columnas_texto_archivo_validar:

    if nombre_columna in encabezados_validacion:

        numero_columna = (
            encabezados_validacion[
                nombre_columna
            ]
        )

        for fila in range(
            2,
            hoja_validacion.max_row + 1
        ):

            celda = (
                hoja_validacion.cell(
                    row=fila,
                    column=numero_columna
                )
            )

            celda.number_format = "@"

            if celda.value is not None:

                celda.value = str(
                    celda.value
                )


columna_observacion = (
    encabezados_validacion[
        "Observacion_NumFactura"
    ]
)


relleno_observacion = (
    PatternFill(
        fill_type="solid",
        fgColor="FFF2CC"
    )
)


fuente_observacion = (
    Font(
        color="9C0006",
        bold=True
    )
)


for fila in range(
    2,
    hoja_validacion.max_row + 1
):

    celda = (
        hoja_validacion.cell(
            row=fila,
            column=columna_observacion
        )
    )

    if celda.value:

        celda.fill = (
            relleno_observacion
        )

        celda.font = (
            fuente_observacion
        )


libro_validacion.save(
    VALIDACION
)


print(
    "\nArchivoValidar generado:"
)

print(
    VALIDACION
)

print(
    f"Facturas validas: "
    f"{len(facturas_validas)}"
)

print(
    f"Facturas observadas: "
    f"{cantidad_observadas}"
)


# ============================================================
# PASO 2
# CONSULTA DE FACTURAS
# OUTPUT 1
# ============================================================

abrir_edge(
    maximizar=False
)


login_siebel()

pyautogui.click(
    *MODULO_FACTURAS
)


time.sleep(6)


pyautogui.click(
    *CONSULTA_FACTURAS
)


time.sleep(4)


escribir_campo(
    NUMERO_FACTURA,
    OR_FACTURAS
)


pyautogui.press(
    "enter"
)


time.sleep(13)


exportar(
    CSV_FACTURAS,
    EXPORTAR_FACTURAS
)


# ============================================================
# CREAR CONSULTASIEBEL.XLSX
# ============================================================

df_consulta = (
    leer_csv_siebel(
        CSV_FACTURAS
    )
)


columna_contrato_1 = (
    buscar_columna(
        df_consulta,
        [
            "Numero Contrato",
            "Número Contrato",
            "N° Contrato",
            "Nº Contrato",
            "N Contrato",
        ]
    )
)


columna_cliente_1 = (
    buscar_columna(
        df_consulta,
        [
            "Numero Cliente",
            "Número Cliente",
            "Código Cliente",
            "Codigo Cliente",
        ]
    )
)


if (
    not columna_contrato_1
    or not columna_cliente_1
):

    raise ValueError(
        "No se encontro Numero Contrato "
        "o Numero Cliente en output 1."
    )


guardar_tabla(
    df_consulta,
    CONSULTA_SIEBEL,
    "TablaConsultaSiebel",
    [
        columna_contrato_1,
        columna_cliente_1,
    ]
)


# ============================================================
# OR DE CONTRATOS
# ============================================================

contratos_validos = [
    contrato
    for contrato in (
        df_consulta[
            columna_contrato_1
        ]
        .apply(limpiar_codigo)
    )
    if contrato
]


if not contratos_validos:

    raise ValueError(
        "No se encontraron contratos "
        "para la segunda consulta."
    )


OR_CONTRATOS = (
    " OR ".join(
        contratos_validos
    )
)


print(
    "\nConsultaSiebel generado:"
)

print(
    CONSULTA_SIEBEL
)

print(
    f"Contratos preparados: "
    f"{len(contratos_validos)}"
)


# ============================================================
# PASO 3
# CONSULTA DE CONTRATOS
# OUTPUT 2
# ============================================================

abrir_edge(
    maximizar=True
)


login_siebel()


pyautogui.click(
    *MODULO_CONTRATOS
)


time.sleep(6)


pyautogui.click(
    *MIS_CONTRATOS
)


time.sleep(3)


pyautogui.click(
    *TODOS_CONTRATOS
)


time.sleep(6)


pyautogui.click(
    *CONSULTA_CONTRATOS
)


time.sleep(4)


escribir_campo(
    NUMERO_CONTRATO,
    OR_CONTRATOS
)


pyautogui.press(
    "enter"
)


time.sleep(14)


exportar(
    CSV_CONTRATOS,
    EXPORTAR_CONTRATOS,
    MENU_CONTRATOS
)


# ============================================================
# CREAR GESTIONCONTRATO.XLSX
# ============================================================

df_gestion = (
    leer_csv_siebel(
        CSV_CONTRATOS
    )
)


columna_contrato_2 = (
    buscar_columna(
        df_gestion,
        [
            "N° Contrato",
            "Nº Contrato",
            "N Contrato",
            "Numero Contrato",
            "Número Contrato",
        ]
    )
)


columna_cliente_2 = (
    buscar_columna(
        df_gestion,
        [
            "Código Cliente",
            "Codigo Cliente",
            "Numero Cliente",
            "Número Cliente",
        ]
    )
)


if (
    not columna_contrato_2
    or not columna_cliente_2
):

    raise ValueError(
        "No se encontro N° Contrato "
        "o Código Cliente en output 2."
    )


guardar_tabla(
    df_gestion,
    GESTION_CONTRATO,
    "TablaGestionContrato",
    [
        columna_contrato_2,
        columna_cliente_2,
    ]
)


print(
    "\nGestionContrato generado:"
)

print(
    GESTION_CONTRATO
)

print(
    f"Filas guardadas en "
    f"GestionContrato.xlsx: "
    f"{len(df_gestion)}"
)


# ============================================================
# PASO 4
# ACTIVIDADES ARGENTINA
# OUTPUT 3
# ============================================================

abrir_edge(
    maximizar=True
)


login_siebel()


pyautogui.click(
    *FLECHA_MODULOS
)


time.sleep(3)


pyautogui.click(
    *CONSULTA_ACTIVIDADES_ARG
)


time.sleep(6)


pyautogui.click(
    *CONSULTA_ACTIVIDADES
)


time.sleep(4)


escribir_campo(
    CAMPO_TIPO,
    "NOTA DE CREDITO"
)


escribir_campo(
    CAMPO_CATEGORIA,
    "FACTURACIÓN"
)


escribir_campo(
    CAMPO_ESTADO,
    "CERRADA"
)


for _ in range(8):

    pyautogui.click(
        *FLECHA_DESPLAZAMIENTO
    )

    time.sleep(0.25)


escribir_campo(
    CAMPO_REGION,
    "FACTURACION LIMA"
)


for _ in range(20):

    pyautogui.click(
        *FLECHA_DESPLAZAMIENTO
    )

    time.sleep(0.25)


escribir_campo(
    CAMPO_FECHA_CIERRE,
    FECHA_CIERRE_FILTRO
)


pyautogui.press(
    "enter"
)


print(
    "\nEsperando 1 minuto "
    "para los resultados "
    "de actividades..."
)


time.sleep(60)


exportar(
    CSV_ACTIVIDADES,
    EXPORTAR_ACTIVIDADES,
    MENU_ACTIVIDADES
)


# ============================================================
# CREAR NC_GENERADAS.XLSX
# DESDE OUTPUT 3
# ============================================================

df_nc_generadas = (
    leer_csv_siebel(
        CSV_ACTIVIDADES
    )
)


columna_factura_nc = (
    buscar_columna(
        df_nc_generadas,
        [
            "Nº Factura",
            "N° Factura",
            "N Factura",
            "Numero Factura",
            "Número Factura",
        ]
    )
)


if not columna_factura_nc:

    print(
        "\nColumnas encontradas "
        "en output 3:"
    )

    for columna in (
        df_nc_generadas.columns
    ):

        print(
            f" - {repr(columna)} "
            f"=> {normalizar(columna)}"
        )

    raise ValueError(
        "No se encontro la columna "
        "Nº Factura en output 3."
    )


guardar_tabla(
    df_nc_generadas,
    NC_GENERADAS,
    "TablaNCGeneradas",
    [
        columna_factura_nc
    ]
)


print(
    "\nNC_Generadas.xlsx generado:"
)

print(
    NC_GENERADAS
)

print(
    f"Filas guardadas en "
    f"NC_Generadas.xlsx: "
    f"{len(df_nc_generadas)}"
)


# ============================================================
# CRUCE FINAL
# ============================================================

(
    CONFIGURACION_ACTUALIZADA,
    CANTIDAD_FACTURAS_COPIADAS,
    CANTIDAD_FILAS_GESTION_COPIADAS

) = actualizar_configuracion(
    NC_GENERADAS,
    GESTION_CONTRATO
)


# ============================================================
# RESUMEN FINAL
# ============================================================

print(
    "\n=============================================="
)

print(
    "             PROCESO FINALIZADO"
)

print(
    "=============================================="
)

print(
    "\nCarpeta:"
)

print(
    CARPETA
)

print(
    "\nArchivos principales:"
)

print(
    f"- {VALIDACION.name}"
)

print(
    f"- {CONSULTA_SIEBEL.name}"
)

print(
    f"- {GESTION_CONTRATO.name}"
)

print(
    f"- {CSV_ACTIVIDADES.name}"
)

print(
    f"- {NC_GENERADAS.name}"
)

print(
    f"- {CONFIGURACION_ACTUALIZADA.name}"
)

print(
    f"\nFacturas consultadas: "
    f"{len(facturas_validas)}"
)

print(
    f"Facturas observadas: "
    f"{cantidad_observadas}"
)

print(
    f"Contratos consultados: "
    f"{len(contratos_validos)}"
)

print(
    f"Filas guardadas en "
    f"NC_Generadas: "
    f"{len(df_nc_generadas)}"
)

print(
    f"Facturas copiadas a "
    f"Notadecredito: "
    f"{CANTIDAD_FACTURAS_COPIADAS}"
)

print(
    f"Filas copiadas a "
    f"GestionContrato: "
    f"{CANTIDAD_FILAS_GESTION_COPIADAS}"
)

print(
    "=============================================="
)
