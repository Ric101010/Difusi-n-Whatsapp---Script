"""
WhatsApp Difusión Pro - Modo Ventana Única (Sin Recargar) & Pestañas
====================================================================
Aplicación de escritorio en PySide6 para el envío automatizado de mensajes
personalizados de WhatsApp desde Excel o CSV.

Modo Principal (Ventana Única de Opera - Sin Recargar):
- Todo el proceso ocurre DENTRO DE LA MISMA VENTANA de WhatsApp Web ya abierta.
- Cero recargas de página: no vuelve a cargar la barra verde ni reconectar WebSockets.
- Cero pestañas nuevas: busca el contacto con el motor interno de WhatsApp y lo abre al instante.
- Reutiliza tu sesión activa existente: ¡NO pide código QR!
- Pegado fiel de emojis, saltos de línea y acentos vía portapapeles.
"""

from dataclasses import dataclass, field
from datetime import datetime
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import logging
import random
import re
import sys
import time
import urllib.parse
import webbrowser

import pandas as pd
import pyautogui
import pyperclip

from PySide6.QtCore import QObject, QThread, Qt, Signal, Slot
from PySide6.QtGui import QColor, QFont, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


# Configuración de seguridad para PyAutoGUI
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.02


# ==============================================================================
# MODELOS DE DATOS Y CONFIGURACIÓN
# ==============================================================================

@dataclass
class Contacto:
    """Representa a un contacto cargado desde el archivo Excel o CSV."""
    nombre: str
    telefono_original: str
    telefono_normalizado: str
    datos_extra: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ResultadoEnvio:
    """Resultado del intento de envío a un contacto."""
    telefono: str
    nombre: str
    estado: str          # "ENVIADO" | "FALLIDO"
    detalle: str
    timestamp: str


@dataclass
class WhatsAppConfig:
    """Configuración para la campaña de envíos."""
    report_path: Path = Path("reporte_envios.csv")
    modo_ejecucion: str = "ventana_unica"   # "ventana_unica" o "pestana_nueva"
    tiempo_busqueda_contacto: float = 0.8   # Segundos de espera para que WhatsApp filtre el número (optimizado)
    delay_entre_contactos: float = 0.6      # Pausa rápida entre un contacto y el siguiente
    tiempo_cuenta_regresiva: int = 2        # Segundos de aviso para poner WhatsApp al frente


# ==============================================================================
# NORMALIZADOR DE TELÉFONOS Y CARGADOR DE DATOS
# ==============================================================================

class PhoneNormalizer:
    """Normaliza y valida números de teléfono para WhatsApp Web."""

    @staticmethod
    def normalizar(raw_phone: Any) -> Optional[str]:
        if raw_phone is None or pd.isna(raw_phone):
            return None

        phone_str = str(raw_phone).strip()
        phone_str = re.sub(r"\.0+$", "", phone_str)
        digits_only = re.sub(r"\D", "", phone_str)

        # Si el usuario ingresa 10 dígitos comenzando en 3 (móvil Colombia sin prefijo), añadir 57 automáticamente
        if len(digits_only) == 10 and digits_only.startswith("3"):
            digits_only = "57" + digits_only

        if len(digits_only) < 8 or len(digits_only) > 16:
            return None

        return digits_only


class DataLoader:
    """Carga y procesa archivos Excel y CSV."""

    COLUMNAS_NOMBRE = ["nombre", "nombres", "name", "contacto", "destinatario", "cliente"]
    COLUMNAS_TELEFONO = ["telefono", "teléfono", "celular", "movil", "móvil", "phone", "whatsapp", "numero", "número"]

    @classmethod
    def cargar_archivo(cls, ruta_archivo: Path) -> Tuple[List[Contacto], List[ResultadoEnvio], str]:
        if not ruta_archivo.exists():
            raise FileNotFoundError(f"El archivo no existe: {ruta_archivo}")

        suffix = ruta_archivo.suffix.lower()
        if suffix in [".xlsx", ".xls"]:
            df = pd.read_excel(ruta_archivo, dtype=str)
        elif suffix == ".csv":
            df = pd.read_csv(ruta_archivo, dtype=str)
        else:
            raise ValueError("Formato no soportado. Debe ser .xlsx, .xls o .csv")

        cols_map = {str(col).strip().lower(): col for col in df.columns}

        col_nombre_real: Optional[str] = None
        for candidata in cls.COLUMNAS_NOMBRE:
            if candidata in cols_map:
                col_nombre_real = cols_map[candidata]
                break

        col_tel_real: Optional[str] = None
        for candidata in cls.COLUMNAS_TELEFONO:
            if candidata in cols_map:
                col_tel_real = cols_map[candidata]
                break

        if not col_nombre_real or not col_tel_real:
            columnas_disponibles = ", ".join(list(df.columns))
            raise ValueError(
                f"No se encontraron las columnas necesarias en el archivo.\n"
                f"Columnas detectadas: [{columnas_disponibles}]\n\n"
                f"El archivo debe incluir una columna de nombre ('nombre') y una de teléfono ('telefono')."
            )

        contactos_validos: List[Contacto] = []
        descartes: List[ResultadoEnvio] = []

        for idx, row in df.iterrows():
            raw_nombre = row[col_nombre_real]
            raw_tel = row[col_tel_real]

            nombre = str(raw_nombre).strip() if pd.notna(raw_nombre) else f"Contacto #{idx+1}"
            tel_limpio = PhoneNormalizer.normalizar(raw_tel)

            datos_extra = {
                str(col): (str(row[col]).strip() if pd.notna(row[col]) else "")
                for col in df.columns
            }

            if tel_limpio:
                contactos_validos.append(
                    Contacto(
                        nombre=nombre,
                        telefono_original=str(raw_tel),
                        telefono_normalizado=tel_limpio,
                        datos_extra=datos_extra
                    )
                )
            else:
                descartes.append(
                    ResultadoEnvio(
                        telefono=str(raw_tel),
                        nombre=nombre,
                        estado="FALLIDO",
                        detalle="Número de teléfono inválido o longitud incorrecta",
                        timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    )
                )

        info_resumen = f"Cargados {len(contactos_validos)} contactos válidos ({len(descartes)} descartados)"
        return contactos_validos, descartes, info_resumen


# ==============================================================================
# HILO DE EJECUCIÓN (QThread)
# ==============================================================================

class EnvioWorker(QThread):
    """Ejecuta la difusión dentro de la misma ventana sin recargar la página."""

    sig_log = Signal(str, str)             # mensaje, nivel ('info', 'success', 'warning', 'error')
    sig_progreso = Signal(int, int)        # actual, total
    sig_finalizado = Signal(dict)          # métricas finales
    sig_estado_contacto = Signal(str, str, str) # nombre, telefono, estado

    def __init__(
        self,
        config: WhatsAppConfig,
        contactos: List[Contacto],
        descartes_previos: List[ResultadoEnvio],
        cuerpo_mensaje: str,
        parent: Optional[QObject] = None
    ):
        super().__init__(parent)
        self.config = config
        self.contactos = contactos
        self.resultados: List[ResultadoEnvio] = list(descartes_previos)
        self.cuerpo_mensaje = cuerpo_mensaje
        self._detener_solicitado = False

    def solicitar_detencion(self) -> None:
        self._detener_solicitado = True
        self.sig_log.emit("⏹ Solicitud de detención recibida. Finalizando contacto actual...", "warning")

    def _construir_mensaje_para(self, contacto: Contacto) -> str:
        nombre_contacto = contacto.nombre.strip()

        if "{nombre}" in self.cuerpo_mensaje:
            mensaje = self.cuerpo_mensaje.replace("{nombre}", nombre_contacto)
        else:
            saludo = f"¡Hola {nombre_contacto}! 👋"
            if self.cuerpo_mensaje.strip():
                mensaje = f"{saludo}\n\n{self.cuerpo_mensaje.strip()}"
            else:
                mensaje = saludo

        for clave, valor in contacto.datos_extra.items():
            token = f"{{{clave}}}"
            if token in mensaje:
                mensaje = mensaje.replace(token, str(valor))

        return mensaje

    def run(self) -> None:
        total = len(self.contactos)
        self.sig_progreso.emit(0, total)

        if self.config.modo_ejecucion == "ventana_unica":
            self._ejecutar_modo_ventana_unica(total)
        else:
            self._ejecutar_modo_pestanas(total)

        self._guardar_reporte()

        enviados = sum(1 for r in self.resultados if r.estado == "ENVIADO")
        fallidos = sum(1 for r in self.resultados if r.estado == "FALLIDO")

        metricas = {
            "total_procesados": len(self.resultados),
            "enviados": enviados,
            "fallidos": fallidos,
            "ruta_reporte": str(self.config.report_path.resolve())
        }
        self.sig_finalizado.emit(metricas)

    def _ejecutar_modo_ventana_unica(self, total: int) -> None:
        """
        Ejecuta los envíos dentro de la MISMA VENTANA activa de WhatsApp Web,
        utilizando la búsqueda interna para saltar de chat a chat SIN RECARGAR LA PÁGINA.
        """
        self.sig_log.emit("🎯 Modo: VENTANA ÚNICA (Búsqueda interna sin recargas)", "info")
        self.sig_log.emit("⚠️ Asegúrate de tener visible tu ventana de WhatsApp Web en pantalla.", "warning")

        # Cuenta regresiva para dar tiempo al usuario de enfocarse en WhatsApp
        for c in range(self.config.tiempo_cuenta_regresiva, 0, -1):
            if self._detener_solicitado:
                return
            self.sig_log.emit(f"⏳ Iniciando en {c} segundo(s)...", "info")
            time.sleep(1.0)

        self.sig_log.emit("🚀 Iniciando despacho dentro de la misma ventana...", "success")

        for i, contacto in enumerate(self.contactos, start=1):
            if self._detener_solicitado:
                self.sig_log.emit(f"Campaña detenida por el usuario antes del contacto {i}/{total}.", "warning")
                break

            self.sig_progreso.emit(i, total)
            mensaje = self._construir_mensaje_para(contacto)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            self.sig_log.emit(
                f"[{i}/{total}] Localizando a: {contacto.nombre} (+{contacto.telefono_normalizado})...",
                "info"
            )

            try:
                # 1. Limpiar cualquier búsqueda o diálogo previo
                pyautogui.press("escape")
                time.sleep(0.08)

                # 2. Abrir 'Nuevo chat' / Buscador interno en WhatsApp Web
                pyautogui.hotkey("ctrl", "alt", "n")
                time.sleep(0.18)

                # 3. Copiar el número telefónico y pegarlo en el buscador interno
                numero_a_buscar = f"+{contacto.telefono_normalizado}"
                pyperclip.copy(numero_a_buscar)
                pyautogui.hotkey("ctrl", "v")

                # 4. Esperar a que WhatsApp busque y resuelva el contacto
                t_busqueda = max(0.4, float(self.config.tiempo_busqueda_contacto))
                time.sleep(t_busqueda)

                if self._detener_solicitado:
                    break

                # 5. Seleccionar el resultado y abrir la conversación
                pyautogui.press("down")
                time.sleep(0.1)
                pyautogui.press("enter")
                time.sleep(0.15)
                pyautogui.press("enter")

                # Espera breve para que el panel del chat tome el foco en la caja de texto
                time.sleep(0.35)

                if self._detener_solicitado:
                    break

                # 6. Copiar y pegar el mensaje personalizado con formato íntegro
                pyperclip.copy(mensaje)
                pyautogui.hotkey("ctrl", "v")
                time.sleep(0.15)

                # 7. Presionar Enter para enviar el mensaje
                pyautogui.press("enter")
                time.sleep(0.35)

                # 8. Limpiar foco para dejar WhatsApp listo para el siguiente contacto
                pyautogui.press("escape")
                time.sleep(0.08)

                self.sig_log.emit(
                    f"[{i}/{total}] ✔ ENVIADO a {contacto.nombre} (+{contacto.telefono_normalizado})",
                    "success"
                )
                self.sig_estado_contacto.emit(contacto.nombre, contacto.telefono_normalizado, "ENVIADO")

                self.resultados.append(
                    ResultadoEnvio(
                        telefono=contacto.telefono_normalizado,
                        nombre=contacto.nombre,
                        estado="ENVIADO",
                        detalle="Enviado dentro de la misma ventana sin recargar",
                        timestamp=timestamp
                    )
                )

            except Exception as e:
                self.sig_log.emit(f"[{i}/{total}] ✖ Error al enviar a {contacto.nombre}: {e}", "error")
                self.sig_estado_contacto.emit(contacto.nombre, contacto.telefono_normalizado, "FALLIDO")
                self.resultados.append(
                    ResultadoEnvio(
                        telefono=contacto.telefono_normalizado,
                        nombre=contacto.nombre,
                        estado="FALLIDO",
                        detalle=f"Error en interacción: {str(e)[:80]}",
                        timestamp=timestamp
                    )
                )

            # Pausa de seguridad antes del siguiente contacto
            if i < total and not self._detener_solicitado:
                pausa = max(0.2, float(self.config.delay_entre_contactos))
                self.sig_log.emit(f"⚡ Siguiente contacto en {pausa:.1f}s...", "info")
                t_p = 0.0
                while t_p < pausa:
                    if self._detener_solicitado:
                        break
                    time.sleep(0.1)
                    t_p += 0.1

    def _ejecutar_modo_pestanas(self, total: int) -> None:
        """Modo alternativo: Pestaña nueva por contacto."""
        for i, contacto in enumerate(self.contactos, start=1):
            if self._detener_solicitado:
                break
            self.sig_progreso.emit(i, total)
            msg = self._construir_mensaje_para(contacto)
            encoded = urllib.parse.quote(msg)
            webbrowser.open(f"https://web.whatsapp.com/send?phone={contacto.telefono_normalizado}&text={encoded}")
            time.sleep(5.0)
            pyautogui.press("enter")
            time.sleep(1.0)
            pyautogui.hotkey("ctrl", "w")
            self.sig_log.emit(f"[{i}/{total}] ✔ Enviado a {contacto.nombre}", "success")
            time.sleep(1.5)

    def _guardar_reporte(self) -> None:
        if not self.resultados:
            return
        try:
            df = pd.DataFrame([
                {
                    "telefono": r.telefono,
                    "nombre": r.nombre,
                    "estado": r.estado,
                    "detalle": r.detalle,
                    "timestamp": r.timestamp
                }
                for r in self.resultados
            ])
            df.to_csv(self.config.report_path, index=False, encoding="utf-8-sig")
            self.sig_log.emit(f"📊 Reporte guardado en: {self.config.report_path.resolve()}", "success")
        except Exception as e:
            self.sig_log.emit(f"Error guardando reporte: {e}", "error")


# ==============================================================================
# INTERFAZ GRÁFICA MODERNA CON PYSIDE6
# ==============================================================================

ESTILO_QSS = """
QMainWindow {
    background-color: #0f172a;
}
QWidget {
    color: #f8fafc;
    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
    font-size: 13px;
}
QGroupBox {
    border: 1px solid #334155;
    border-radius: 8px;
    margin-top: 10px;
    padding-top: 14px;
    font-weight: bold;
    background-color: #1e293b;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 8px;
    color: #38bdf8;
}
QPushButton {
    background-color: #2563eb;
    color: #ffffff;
    border: none;
    padding: 8px 16px;
    border-radius: 6px;
    font-weight: 600;
}
QPushButton:hover {
    background-color: #1d4ed8;
}
QPushButton:pressed {
    background-color: #1e40af;
}
QPushButton:disabled {
    background-color: #475569;
    color: #94a3b8;
}
QPushButton#btn_enviar {
    background-color: #16a34a;
    font-size: 14px;
    padding: 11px 24px;
    border-radius: 8px;
}
QPushButton#btn_enviar:hover {
    background-color: #15803d;
}
QPushButton#btn_enviar:disabled {
    background-color: #334155;
    color: #64748b;
}
QPushButton#btn_detener {
    background-color: #dc2626;
    font-size: 14px;
    padding: 11px 24px;
    border-radius: 8px;
}
QPushButton#btn_detener:hover {
    background-color: #b91c1c;
}
QPushButton#btn_detener:disabled {
    background-color: #334155;
    color: #64748b;
}
QTextEdit {
    background-color: #0f172a;
    border: 1px solid #334155;
    border-radius: 6px;
    padding: 8px;
    color: #f1f5f9;
}
QTextEdit:focus {
    border: 1px solid #38bdf8;
}
QTableWidget {
    background-color: #0f172a;
    border: 1px solid #334155;
    border-radius: 6px;
    gridline-color: #1e293b;
    color: #e2e8f0;
}
QHeaderView::section {
    background-color: #1e293b;
    color: #94a3b8;
    padding: 5px;
    border: 1px solid #334155;
    font-weight: bold;
}
QProgressBar {
    border: 1px solid #334155;
    border-radius: 6px;
    text-align: center;
    background-color: #0f172a;
    color: #ffffff;
    font-weight: bold;
    height: 22px;
}
QProgressBar::chunk {
    background-color: #22c55e;
    border-radius: 5px;
}
QDoubleSpinBox {
    background-color: #0f172a;
    border: 1px solid #334155;
    border-radius: 6px;
    padding: 4px 8px;
    color: #22c55e;
    font-weight: bold;
}
QRadioButton {
    color: #f8fafc;
    font-weight: 500;
}
QRadioButton::indicator:checked {
    background-color: #22c55e;
    border: 2px solid #ffffff;
    border-radius: 6px;
}
"""


class MainWindow(QMainWindow):
    """Ventana principal de la aplicación."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("WhatsApp Difusión Pro")
        self.resize(1060, 810)
        self.setMinimumSize(880, 670)
        self.setStyleSheet(ESTILO_QSS)

        self.contactos_cargados: List[Contacto] = []
        self.descartes_iniciales: List[ResultadoEnvio] = []
        self.ruta_archivo_actual: Optional[Path] = None
        self.worker: Optional[EnvioWorker] = None

        self._crear_interfaz()
        self._cargar_archivo_por_defecto()

    def _crear_interfaz(self) -> None:
        contenedor_central = QWidget()
        self.setCentralWidget(contenedor_central)
        layout_principal = QVBoxLayout(contenedor_central)
        layout_principal.setContentsMargins(18, 12, 18, 12)
        layout_principal.setSpacing(10)

        # Encabezado con estado
        layout_encabezado = QHBoxLayout()
        lbl_titulo = QLabel("💬 WhatsApp Difusión Pro")
        lbl_titulo.setStyleSheet("font-size: 20px; font-weight: bold; color: #22c55e;")

        badge_modo = QLabel("⚡ Modo: Ventana Única (Sin Recargas)")
        badge_modo.setStyleSheet(
            "background-color: #1e3a8a; color: #93c5fd; padding: 4px 10px; "
            "border-radius: 12px; font-weight: 600; font-size: 12px;"
        )

        layout_encabezado.addWidget(lbl_titulo)
        layout_encabezado.addSpacing(10)
        layout_encabezado.addWidget(badge_modo)
        layout_encabezado.addStretch()
        layout_principal.addLayout(layout_encabezado)

        splitter = QSplitter(Qt.Vertical)

        # Panel Superior: Entrada y configuración
        panel_superior = QWidget()
        layout_sup = QVBoxLayout(panel_superior)
        layout_sup.setContentsMargins(0, 0, 0, 0)
        layout_sup.setSpacing(10)

        # 1. Grupo: Archivo y Parámetros
        grupo_archivo = QGroupBox("1. Archivo de contactos y configuración de envío")
        layout_archivo_v = QVBoxLayout(grupo_archivo)

        layout_archivo_h = QHBoxLayout()
        self.btn_seleccionar_archivo = QPushButton("📂 Seleccionar Excel (.xlsx, .xls) o CSV...")
        self.btn_seleccionar_archivo.clicked.connect(self._abrir_dialogo_archivo)
        self.lbl_archivo_info = QLabel("Ningún archivo seleccionado")
        self.lbl_archivo_info.setStyleSheet("color: #94a3b8; font-style: italic;")

        layout_archivo_h.addWidget(self.btn_seleccionar_archivo)
        layout_archivo_h.addWidget(self.lbl_archivo_info, 1)
        layout_archivo_v.addLayout(layout_archivo_h)

        # Configuración de tiempos
        layout_tiempos = QHBoxLayout()

        lbl_modo_info = QLabel("Modo: <b>Ventana Única</b> (busca y envía sin abrir pestañas ni recargar)")
        lbl_modo_info.setStyleSheet("color: #22c55e; font-size: 12px;")

        lbl_espera_busqueda = QLabel("⏱️ Tiempo de búsqueda:")
        lbl_espera_busqueda.setStyleSheet("color: #38bdf8;")
        self.spin_busqueda = QDoubleSpinBox()
        self.spin_busqueda.setRange(0.4, 6.0)
        self.spin_busqueda.setValue(0.8)
        self.spin_busqueda.setSuffix(" seg")
        self.spin_busqueda.setToolTip("Tiempo de espera para que WhatsApp resuelva el contacto en la búsqueda.")

        lbl_pausa = QLabel("⚡ Pausa entre contactos:")
        lbl_pausa.setStyleSheet("color: #38bdf8;")
        self.spin_pausa = QDoubleSpinBox()
        self.spin_pausa.setRange(0.2, 10.0)
        self.spin_pausa.setValue(0.6)
        self.spin_pausa.setSuffix(" seg")
        self.spin_pausa.setToolTip("Pausa antes de saltar al siguiente contacto.")

        layout_tiempos.addWidget(lbl_modo_info)
        layout_tiempos.addStretch()
        layout_tiempos.addWidget(lbl_espera_busqueda)
        layout_tiempos.addWidget(self.spin_busqueda)
        layout_tiempos.addSpacing(10)
        layout_tiempos.addWidget(lbl_pausa)
        layout_tiempos.addWidget(self.spin_pausa)

        layout_archivo_v.addLayout(layout_tiempos)
        layout_sup.addWidget(grupo_archivo)

        # 2. Grupo: Redacción y Vista Previa
        grupo_mensaje = QGroupBox("2. Redacción del mensaje (Saludo automático con el nombre)")
        layout_msg_grid = QHBoxLayout(grupo_mensaje)

        col_input = QVBoxLayout()
        lbl_nota_saludo = QLabel(
            "ℹ️ El saludo inicial <b>¡Hola [Nombre]! 👋</b> se concatena automáticamente al principio de cada mensaje."
        )
        lbl_nota_saludo.setStyleSheet("color: #38bdf8; font-size: 12px;")
        lbl_nota_saludo.setWordWrap(True)

        self.txt_mensaje = QTextEdit()
        self.txt_mensaje.setPlaceholderText(
            "Escribe aquí el cuerpo del mensaje que recibirá cada contacto..."
        )
        self.txt_mensaje.setText("Este es un mensaje de prueba. No contestes.")
        self.txt_mensaje.textChanged.connect(self._actualizar_vista_previa)

        col_input.addWidget(lbl_nota_saludo)
        col_input.addWidget(self.txt_mensaje)
        layout_msg_grid.addLayout(col_input, 3)

        col_preview = QVBoxLayout()
        lbl_preview_titulo = QLabel("👁️ Vista Previa en Tiempo Real:")
        lbl_preview_titulo.setStyleSheet("font-weight: bold; color: #a7f3d0;")

        self.txt_preview = QTextEdit()
        self.txt_preview.setReadOnly(True)
        self.txt_preview.setStyleSheet("background-color: #0b1320; color: #22c55e; border: 1px dashed #22c55e;")

        col_preview.addWidget(lbl_preview_titulo)
        col_preview.addWidget(self.txt_preview)
        layout_msg_grid.addLayout(col_preview, 2)

        layout_sup.addWidget(grupo_mensaje)
        splitter.addWidget(panel_superior)

        # Panel Inferior: Tabla, Consola y Ejecución
        panel_inferior = QWidget()
        layout_inf = QVBoxLayout(panel_inferior)
        layout_inf.setContentsMargins(0, 0, 0, 0)
        layout_inf.setSpacing(10)

        splitter_inf = QSplitter(Qt.Horizontal)

        grupo_tabla = QGroupBox("📋 Contactos detectados")
        layout_tabla = QVBoxLayout(grupo_tabla)
        self.tabla_contactos = QTableWidget()
        self.tabla_contactos.setColumnCount(3)
        self.tabla_contactos.setHorizontalHeaderLabels(["Nombre", "Teléfono", "Estado"])
        self.tabla_contactos.horizontalHeader().setStretchLastSection(True)
        layout_tabla.addWidget(self.tabla_contactos)
        splitter_inf.addWidget(grupo_tabla)

        grupo_consola = QGroupBox("🖥️ Registro de ejecución en vivo")
        layout_consola = QVBoxLayout(grupo_consola)
        self.txt_consola = QTextEdit()
        self.txt_consola.setReadOnly(True)
        layout_consola.addWidget(self.txt_consola)
        splitter_inf.addWidget(grupo_consola)

        splitter_inf.setSizes([450, 600])
        layout_inf.addWidget(splitter_inf)

        # Barra de Progreso
        self.barra_progreso = QProgressBar()
        self.barra_progreso.setValue(0)
        self.barra_progreso.setFormat("Esperando inicio... (0/0)")
        layout_inf.addWidget(self.barra_progreso)

        # Botones de Acción
        layout_acciones = QHBoxLayout()

        self.btn_enviar = QPushButton("🚀 Iniciar Envíos")
        self.btn_enviar.setObjectName("btn_enviar")
        self.btn_enviar.clicked.connect(self._iniciar_envios)

        self.btn_detener = QPushButton("⏹ Detener")
        self.btn_detener.setObjectName("btn_detener")
        self.btn_detener.setEnabled(False)
        self.btn_detener.clicked.connect(self._detener_envios)

        layout_acciones.addWidget(self.btn_enviar, 2)
        layout_acciones.addWidget(self.btn_detener, 1)
        layout_inf.addLayout(layout_acciones)

        splitter.addWidget(panel_inferior)
        splitter.setSizes([350, 420])
        layout_principal.addWidget(splitter)

    # --------------------------------------------------------------------------
    # GESTIÓN DE ARCHIVOS Y VISTA PREVIA
    # --------------------------------------------------------------------------

    def _cargar_archivo_por_defecto(self) -> None:
        for archivo in [Path("contactos.csv"), Path("contactos_ejemplo.xlsx")]:
            if archivo.exists():
                self._procesar_archivo(archivo)
                break

    def _abrir_dialogo_archivo(self) -> None:
        archivo, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo de contactos",
            "",
            "Archivos de Excel y CSV (*.xlsx *.xls *.csv);;Excel (*.xlsx *.xls);;CSV (*.csv)"
        )
        if archivo:
            self._procesar_archivo(Path(archivo))

    def _procesar_archivo(self, ruta: Path) -> None:
        try:
            contactos, descartes, resumen = DataLoader.cargar_archivo(ruta)
            self.contactos_cargados = contactos
            self.descartes_iniciales = descartes
            self.ruta_archivo_actual = ruta

            self.lbl_archivo_info.setText(f"✔ {ruta.name} — {resumen}")
            self.lbl_archivo_info.setStyleSheet("color: #22c55e; font-weight: 500;")

            self._actualizar_tabla_contactos()
            self._actualizar_vista_previa()

            self._agregar_log_consola(f"Archivo cargado: {ruta.name} ({len(contactos)} contactos listos)", "success")
        except Exception as e:
            QMessageBox.critical(self, "Error al cargar archivo", str(e))
            self.lbl_archivo_info.setText("Error al cargar el archivo")
            self.lbl_archivo_info.setStyleSheet("color: #ef4444;")

    def _actualizar_tabla_contactos(self) -> None:
        self.tabla_contactos.setRowCount(0)
        for row, c in enumerate(self.contactos_cargados):
            self.tabla_contactos.insertRow(row)
            self.tabla_contactos.setItem(row, 0, QTableWidgetItem(c.nombre))
            self.tabla_contactos.setItem(row, 1, QTableWidgetItem(f"+{c.telefono_normalizado}"))
            item_estado = QTableWidgetItem("Pendiente")
            item_estado.setForeground(QColor("#94a3b8"))
            self.tabla_contactos.setItem(row, 2, item_estado)

    def _actualizar_vista_previa(self) -> None:
        cuerpo = self.txt_mensaje.toPlainText().strip()
        if not self.contactos_cargados:
            primer_nombre = "NombreEjemplo"
        else:
            primer_nombre = self.contactos_cargados[0].nombre

        if "{nombre}" in cuerpo:
            preview = cuerpo.replace("{nombre}", primer_nombre)
        else:
            saludo = f"¡Hola {primer_nombre}! 👋"
            if cuerpo:
                preview = f"{saludo}\n\n{cuerpo}"
            else:
                preview = saludo

        self.txt_preview.setText(preview)

    # --------------------------------------------------------------------------
    # CONTROL DE EJECUCIÓN
    # --------------------------------------------------------------------------

    def _iniciar_envios(self) -> None:
        if not self.contactos_cargados:
            QMessageBox.warning(
                self,
                "Sin contactos",
                "Por favor, selecciona primero un archivo Excel o CSV con contactos válidos."
            )
            return

        cuerpo = self.txt_mensaje.toPlainText().strip()
        if not cuerpo:
            respuesta = QMessageBox.question(
                self,
                "Mensaje vacío",
                "El cuerpo del mensaje está vacío. Solo se enviará el saludo inicial. ¿Deseas continuar?",
                QMessageBox.Yes | QMessageBox.No
            )
            if respuesta == QMessageBox.No:
                return

        total = len(self.contactos_cargados)
        pausa_actual = self.spin_pausa.value()
        tiempo_busqueda = self.spin_busqueda.value()

        confirmacion = QMessageBox.question(
            self,
            "Confirmar inicio de envíos (Ventana Única)",
            f"Se iniciará el despacho hacia {total} contactos dentro de tu ventana actual de WhatsApp Web.\n\n"
            f"• NO se abrirán pestañas nuevas.\n"
            f"• NO se recargará la página.\n"
            f"• Tiempo de búsqueda: {tiempo_busqueda}s.\n"
            f"• Pausa entre contactos: {pausa_actual}s.\n\n"
            f"👉 Al presionar 'Yes', tendrás 3 segundos para asegurarte de que tu ventana de WhatsApp Web esté visible en pantalla.\n\n"
            f"¿Deseas comenzar?",
            QMessageBox.Yes | QMessageBox.No
        )
        if confirmacion != QMessageBox.Yes:
            return

        # Ajuste de controles
        self.btn_enviar.setEnabled(False)
        self.btn_detener.setEnabled(True)
        self.btn_seleccionar_archivo.setEnabled(False)
        self.txt_mensaje.setReadOnly(True)
        self.spin_pausa.setEnabled(False)
        self.spin_busqueda.setEnabled(False)
        self.barra_progreso.setMaximum(total)
        self.barra_progreso.setValue(0)
        self.barra_progreso.setFormat(f"Iniciando... (0/{total})")

        # Iniciar Worker en Modo Ventana Única
        config = WhatsAppConfig(
            tiempo_busqueda_contacto=tiempo_busqueda,
            delay_entre_contactos=pausa_actual,
            tiempo_cuenta_regresiva=3,
            modo_ejecucion="ventana_unica"
        )
        self.worker = EnvioWorker(
            config=config,
            contactos=self.contactos_cargados,
            descartes_previos=self.descartes_iniciales,
            cuerpo_mensaje=cuerpo,
            parent=self
        )
        self.worker.sig_log.connect(self._agregar_log_consola)
        self.worker.sig_progreso.connect(self._actualizar_progreso)
        self.worker.sig_estado_contacto.connect(self._actualizar_estado_tabla)
        self.worker.sig_finalizado.connect(self._campana_finalizada)
        self.worker.start()

    def _detener_envios(self) -> None:
        if self.worker and self.worker.isRunning():
            self.btn_detener.setEnabled(False)
            self.worker.solicitar_detencion()

    @Slot(str, str)
    def _agregar_log_consola(self, mensaje: str, nivel: str) -> None:
        hora = datetime.now().strftime("%H:%M:%S")
        color = "#e2e8f0"
        if nivel == "success":
            color = "#22c55e"
        elif nivel == "warning":
            color = "#f59e0b"
        elif nivel == "error":
            color = "#ef4444"
        elif nivel == "info":
            color = "#38bdf8"

        html = f"<span style='color: #64748b;'>[{hora}]</span> <span style='color: {color};'>{mensaje}</span>"
        self.txt_consola.append(html)
        self.txt_consola.moveCursor(QTextCursor.End)

    @Slot(int, int)
    def _actualizar_progreso(self, actual: int, total: int) -> None:
        self.barra_progreso.setValue(actual)
        self.barra_progreso.setFormat(f"Contacto {actual} de {total} ({int((actual/total)*100)}%)")

    @Slot(str, str, str)
    def _actualizar_estado_tabla(self, nombre: str, telefono: str, estado: str) -> None:
        for row in range(self.tabla_contactos.rowCount()):
            item_tel = self.tabla_contactos.item(row, 1)
            if item_tel and telefono in item_tel.text():
                item_estado = QTableWidgetItem(estado)
                if estado == "ENVIADO":
                    item_estado.setForeground(QColor("#22c55e"))
                else:
                    item_estado.setForeground(QColor("#ef4444"))
                self.tabla_contactos.setItem(row, 2, item_estado)
                break

    @Slot(dict)
    def _campana_finalizada(self, metricas: dict) -> None:
        self.btn_enviar.setEnabled(True)
        self.btn_detener.setEnabled(False)
        self.btn_seleccionar_archivo.setEnabled(True)
        self.txt_mensaje.setReadOnly(False)
        self.spin_pausa.setEnabled(True)
        self.spin_busqueda.setEnabled(True)

        enviados = metricas.get("enviados", 0)
        fallidos = metricas.get("fallidos", 0)
        total = metricas.get("total_procesados", 0)
        ruta = metricas.get("ruta_reporte", "reporte_envios.csv")

        self.barra_progreso.setFormat(f"Completado: {enviados}/{total} enviados")

        QMessageBox.information(
            self,
            "Campaña Finalizada",
            f"El proceso de difusión ha finalizado.\n\n"
            f"• Modo: Ventana Única (Sin recargas)\n"
            f"• Total procesados: {total}\n"
            f"• Enviados exitosos: {enviados}\n"
            f"• Fallidos / Error: {fallidos}\n\n"
            f"Reporte de auditoría en:\n{ruta}"
        )


# ==============================================================================
# PUNTO DE ENTRADA PRINCIPAL
# ==============================================================================

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName("WhatsApp Difusión Pro")

    fuente = QFont("Segoe UI", 10)
    app.setFont(fuente)

    ventana = MainWindow()
    ventana.show()

    sys.exit(app.exec())
