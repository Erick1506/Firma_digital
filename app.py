import sys
import os
import fitz  # PyMuPDF
import tempfile
import traceback
import unicodedata
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog,
    QLabel, QLineEdit, QFormLayout, QTextEdit, QHBoxLayout,
    QInputDialog, QMessageBox, QProgressBar
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from pyhanko.sign import signers
from pyhanko.sign.fields import SigFieldSpec, append_signature_field, MDPPerm
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.serialization import pkcs12
from asn1crypto import x509
from pyhanko_certvalidator.registry import SimpleCertificateStore
import docx2pdf
from asn1crypto.keys import PrivateKeyInfo
import subprocess
import shutil
import platform

# Valor por defecto (se sobrescribirá según selección de rol)
TARGET_TEXT = "WOLFANG ALBERTO LATORRE MARTINEZ"  # - Coordinador (fallback)

# =====================================
# CONFIGURACIÓN DE ROLES (NO MODIFICAR)
# =====================================
roles_config = {
    "Director": {
        "nombre": "GERARDO ARTURO MEDINA ROSAS",
        "posibles_nombres": [
            "GERARDO ARTURO MEDINA ROSAS",
            "Gerardo Arturo Medina Rosas",
            "gerardo arturo medina rosas",
            "G. A. Medina Rosas",
            "Gerardo A. Medina",
            "Medina Rosas Gerardo Arturo"
        ],
        "logo": os.path.join("firma", "logo_director.png"),
        "metadata": signers.PdfSignatureMetadata(
            field_name="FirmaDigital",
            name="Gerardo Arturo Medina Rosas",
            reason="Documento firmado digitalmente",
            location="Bogotá, Colombia",
            contact_info="gerardo.medina@sena.edu.co", # Editar --
            certify=True, # TRUE (CERTIFICADO) || FALSE (VALIDADO)  (ROLES DIFERENTES)
            docmdp_permissions=MDPPerm.NO_CHANGES
        )
    },
    "Coordinador": {
        "nombre": "WOLFANG ALBERTO LATORRE MARTINEZ",
        "posibles_nombres": [
            "WOLFANG ALBERTO LATORRE MARTINEZ",
            "WOLFANG ALBERTO LATORRE MARTÍNEZ",
            "Wolfang Alberto Latorre Martinez",
            "Wolfang Alberto Latorre Martínez",
            "wolfang alberto latorre martinez",
            "wolfang alberto latorre martínez",
            "WOLFANG A. LATORRE",
            "Wolfang Latorre",
            "W. A. LATORRE MARTINEZ",
        ],
        "logo": os.path.join("firma", "logo_coordinador.png"),
        "metadata": signers.PdfSignatureMetadata(
            field_name="FirmaDigital",
            name="Wolfang Alberto Latorre Martinez",
            reason="Documento firmado digitalmente",
            location="Bogotá, Colombia",
            contact_info="walatorrem@sena.edu.co",
            certify=False, # TRUE (CERTIFICADO) || FALSE (VALIDADO)  (ROLES DIFERENTES)
            docmdp_permissions=MDPPerm.NO_CHANGES
        )
    }
}

# variables globales que se llenan tras elegir rol
SELECTED_METADATA = None
SELECTED_POSIBLES_NOMBRES = None
SELECTED_LOGO_REL = None  # ruta relativa dentro del proyecto (ej: "firma/logo_coordinador.png")
SELECTED_ROLE_NAME = None


# =====================================
# FUNCIÓN PARA SELECCIONAR ROL (AL INICIO)
# =====================================
def seleccionar_rol():
    """
    Muestra un diálogo para elegir rol (Director / Coordinador).
    Rellena las variables globales SELECTED_... según la elección.
    Debe llamarse después de crear QApplication.
    """
    global SELECTED_METADATA, SELECTED_POSIBLES_NOMBRES, SELECTED_LOGO_REL, SELECTED_ROLE_NAME, TARGET_TEXT

    opciones = list(roles_config.keys())
    rol, ok = QInputDialog.getItem(None, "Seleccionar Rol", "Elija el rol:", opciones, 0, False)
    if not ok or not rol:
        QMessageBox.critical(None, "Rol no seleccionado", "Debe seleccionar un rol para continuar. Saliendo.")
        sys.exit(1)

    cfg = roles_config[rol]
    SELECTED_METADATA = cfg["metadata"]
    SELECTED_POSIBLES_NOMBRES = cfg["posibles_nombres"]
    SELECTED_LOGO_REL = cfg["logo"]
    SELECTED_ROLE_NAME = cfg["nombre"]
    TARGET_TEXT = cfg["nombre"]  # sobrescribir TARGET_TEXT para compatibilidad


# =====================================
# FUNCIÓN: convertir DOC/DOCX -> PDF 
# =====================================
def convert_word_to_pdf(input_path, prefer_tempdir=True):
    """
    Intenta convertir input_path (.docx/.doc) a PDF.
    1) Usa LibreOffice (soffice) si está disponible.
    2) Si falla, intenta docx2pdf (Word+COM).
    Devuelve la ruta absoluta del PDF generado o lanza Exception si no pudo.
    """
    if not os.path.exists(input_path):
        raise Exception(f"Archivo de entrada no existe: {input_path}")

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    # preferir tempdir para no escribir en carpetas con permisos especiales
    outdir = tempfile.gettempdir() if prefer_tempdir else os.path.dirname(input_path)
    out_pdf = os.path.join(outdir, base_name + ".pdf")

    # 1) Buscar soffice en rutas comunes
    soffice_candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/snap/bin/libreoffice",
        shutil.which("soffice") or "",
        shutil.which("libreoffice") or ""
    ]
    soffice_cmd = None
    for p in soffice_candidates:
        if p and os.path.exists(p):
            soffice_cmd = p
            break

    if not soffice_cmd:
        # intentar los comandos detectados por shutil.which
        for cmd in ("soffice", "libreoffice"):
            path_cmd = shutil.which(cmd)
            if path_cmd:
                soffice_cmd = path_cmd
                break

    # Intentar LibreOffice si se encontró
    if soffice_cmd:
        try:
            # Asegurar que outdir existe
            os.makedirs(outdir, exist_ok=True)
            proc = subprocess.run([
                soffice_cmd,
                "--headless", "--convert-to", "pdf", input_path,
                "--outdir", outdir
            ], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)

            # Verificar salida
            if os.path.exists(out_pdf):
                return out_pdf
            else:
                # si no existe, capturar salida para el error
                stdout = proc.stdout.decode(errors="ignore") if proc.stdout else ""
                stderr = proc.stderr.decode(errors="ignore") if proc.stderr else ""
                # intentar otras variantes (por ejemplo en algunas instalaciones el parámetro --outdir no funciona igual)
                # pero no complicamos más: pasamos a fallback
                # registrar mensaje en excepción
                raise Exception(f"LibreOffice falló a convertir. stdout: {stdout}; stderr: {stderr}")
        except Exception as e:
            # no terminar aquí: intentaremos docx2pdf como fallback
            pass

    # 2) Fallback: intentar docx2pdf -> usa Word + COM en Windows
    try:
        # docx2pdf.accepts both (input, output) signature in modern versions
        # intentamos generar en outdir
        os.makedirs(outdir, exist_ok=True)
        docx2pdf.convert(input_path, out_pdf)
        if os.path.exists(out_pdf):
            return out_pdf
        # en algunos entornos docx2pdf.convert(input) crea en la misma carpeta de origen; comprobarlo
        alt = input_path.replace(".docx", ".pdf").replace(".doc", ".pdf")
        if os.path.exists(alt):
            return alt
        raise Exception("docx2pdf no produjo el PDF esperado.")
    except Exception as ex:
        # nada funcionó
        raise Exception(f"No se pudo convertir {input_path} a PDF: {ex}")


# =====================================
# HILO DE FIRMA (estructura original preservada)
# =====================================
class SignThread(QThread):
    progress = pyqtSignal(int)
    message = pyqtSignal(str, bool)
    finished_process = pyqtSignal()  # Señal para indicar que finalizó el proceso

    def __init__(self, files, cert_path, cert_pass, logo_path, output_dir, signer):
        super().__init__()
        self.files = files
        self.cert_path = cert_path
        self.cert_pass = cert_pass
        self.logo_path = logo_path
        self.output_dir = output_dir
        self.signer = signer

    def run(self):
        total = len(self.files)
        for idx, file in enumerate(self.files):
            filename = os.path.basename(file)
            output_file = os.path.join(
                self.output_dir,
                filename.replace(".docx", ".pdf").replace(".doc", ".pdf")
            )

            try:
                # Conversión Word -> PDF
                if file.endswith((".docx", ".doc")):
                    self.message.emit(f"Convirtiendo {filename} a PDF...", False)

                    # Usar la función robusta de conversión (LibreOffice preferido, docx2pdf fallback)
                    try:
                        temp_pdf = convert_word_to_pdf(file, prefer_tempdir=True)
                    except Exception as conv_err:
                        # Forzar mensaje claro y re-lanzar para manejo unificado
                        raise Exception(f"Fallo conversión Word->PDF: {conv_err}")

                    if not temp_pdf or not os.path.exists(temp_pdf):
                        raise Exception("No se generó el PDF tras conversión.")
                    file = temp_pdf
                    self.message.emit(f"Archivo convertido: {file}", False)

                # Normalizar PDF
                self.message.emit(f"Normalizando PDF {filename}...", False)
                doc = fitz.open(file)
                temp_normalized = os.path.join(tempfile.gettempdir(), filename + "_normalized.pdf")
                doc.save(temp_normalized)
                doc.close()
                file = temp_normalized
                self.message.emit(f"PDF normalizado: {file}", False)

                # Insertar logo dinámico en la última página
                if self.logo_path:
                    self.message.emit(f"Inserción de logo en {filename}...", False)
                    self.insert_logo_dynamic(file)
                    self.message.emit("Logo insertado (o no) según detección.", False)

                # Firmar PDF
                self.message.emit(f"Firmando {filename}...", False)
                with open(file, "rb") as inf:
                    w = IncrementalPdfFileWriter(inf)
                    append_signature_field(w, SigFieldSpec(sig_field_name="FirmaDigital"))

                    # Usar metadata según rol seleccionado
                    metadata = SELECTED_METADATA if SELECTED_METADATA is not None else signers.PdfSignatureMetadata(
                        field_name="FirmaDigital",
                        name="Wolfang Alberto Latorre Martinez",
                        reason="Documento firmado digitalmente",
                        location="Bogotá, Colombia",
                        contact_info="walatorrem@sena.edu.co",
                        certify=False,  # firma de Aprobación (False) || Certificación (True)  (Roles diferentes)
                        docmdp_permissions=MDPPerm.NO_CHANGES
                    )

                    # signers.sign_pdf devuelve un file-like; manejar excepción si es None
                    pdf_signed = signers.sign_pdf(w, metadata, signer=self.signer, new_field_spec=None)
                    if pdf_signed is None:
                        raise Exception("La función de firma devolvió None (signers.sign_pdf falló).")

                    with open(output_file, "wb") as outf:
                        data = pdf_signed.read()
                        if data is None:
                            raise Exception("El resultado de la firma está vacío (None).")
                        outf.write(data)

                    self.message.emit(f"Documento firmado correctamente: {output_file}", False)

            except Exception as e:
                self.message.emit(f"Error firmando {filename}: {e}", True)
                traceback.print_exc()

            self.progress.emit(int((idx + 1) / total * 100))

        # Emitir señal de finalización
        self.finished_process.emit()

    # Método que coloca el logo dinámicamente, buscando variantes del nombre ignorando tildes y mayúsculas
    def insert_logo_dynamic(self, pdf_path):
        """
        Abre pdf_path, busca en la última página alguno de los nombres de la lista 
        (ignorando tildes y mayúsculas) y coloca el logo a la derecha y por encima 
        del nombre (posición vertical ajustada dinámicamente).
        Si no encuentra, emite mensaje (no hay fallback).
        """
        doc = fitz.open(pdf_path)
        page = doc[-1]  # Última página

        # Obtener lista de nombres desde selección; si no está, usar lista por defecto (Coordinador)
        posibles_nombres = SELECTED_POSIBLES_NOMBRES or [
            "WOLFANG ALBERTO LATORRE MARTINEZ",
            "WOLFANG ALBERTO LATORRE MARTÍNEZ",
            "Wolfang Alberto Latorre Martinez",
            "Wolfang Alberto Latorre Martínez",
            "wolfang alberto latorre martinez",
            "wolfang alberto latorre martínez",
            "WOLFANG A. LATORRE",
            "Wolfang Latorre",
            "W. A. LATORRE MARTINEZ",
        ]

        # Normalizar función: quita tildes y pasa a minúsculas
        def normalize_text(t: str) -> str:
            if t is None:
                return ""
            # descomponer y eliminar marcas diacríticas
            t_norm = unicodedata.normalize("NFD", t)
            t_no_accents = "".join(ch for ch in t_norm if unicodedata.category(ch) != "Mn")
            return t_no_accents.lower().strip()

        # Normalizar la lista de nombres una sola vez
        nombres_norm = [normalize_text(n) for n in posibles_nombres]
        logo_inserted = False

        # Recorrer los spans (más preciso que bloques) para encontrar la mejor posición
        # get_text("dict") devuelve estructura con páginas, bloques, lines, spans con bbox
        try:
            page_dict = page.get_text("dict")
        except Exception:
            # Fallback más simple si get_text("dict") falla (raro)
            page_dict = {"blocks": page.get_text("blocks")}

        # Iterar bloques -> líneas -> spans
        for block in page_dict.get("blocks", []):
            if logo_inserted:
                break
            # algunos bloques (imágenes) no tienen 'lines'
            lines = block.get("lines", [])
            for line in lines:
                if logo_inserted:
                    break
                for span in line.get("spans", []):
                    raw_text = span.get("text", "")
                    if not raw_text:
                        continue

                    text_norm = normalize_text(raw_text)

                    # si cualquier nombre normalizado está contenido en el span normalizado -> coincidencia
                    matched = None
                    for nn in nombres_norm:
                        if nn and nn in text_norm:
                            matched = nn
                            break

                    if matched:
                        # coordenadas del span: bbox = (x0, y0, x1, y1)
                        x0, y0, x1, y1 = span.get("bbox", (0, 0, 0, 0))

                        # ==============================
                        # CONFIGURACIÓN DEL LOGO (dinámica)
                        # ==============================
                        # Tamaño deseado del logo (ajustable)
                        logo_width = 167
                        logo_height = 62

                        # Separación horizontal y vertical respecto al texto detectado
                        gap_right = 10  # espacio entre el final del texto (x1) y el borde derecho del logo
                        gap_above = 5   # espacio entre el texto y la parte inferior del logo

                        # lo ideal es que el logo quede alineado a la derecha del texto detectado:
                        # se calcula right = x1 - gap_right
                        right = x1 - gap_right
                        left = right - logo_width

                        # Si left sale negativo (logo grande), se ajusta al borde izquierdo
                        if left < 0:
                            left = 5
                            right = left + logo_width

                        # Posicion vertical: se coloca el logo un poco por encima de la caja del span
                        # se da la margen gap_above (es decir, el borde inferior del logo queda gap_above por encima de y0)
                        bottom = y0 - gap_above
                        top = bottom - logo_height

                        # Si top < 0 (logo saldría fuera arriba), se empuja abajo hasta 5
                        if top < 0:
                            top = 5
                            bottom = top + logo_height

                        # se asegura que el rect esté dentro del tamaño de la página
                        page_rect = page.rect  # tiene width/height implícitos

                        if right > page_rect.width:
                            # si derecha excede, se dezplaza el logo para que quede dentro
                            right = page_rect.width - 5
                            left = right - logo_width
                            if left < 0:
                                left = 5
                                right = left + logo_width

                        if bottom > page_rect.height:
                            # si por alguna razón bottom excede, lo limitamos
                            bottom = page_rect.height - 5
                            top = bottom - logo_height
                            if top < 0:
                                top = 5
                                bottom = top + logo_height

                        # Construir rect y validar (fitz exige finito y no vacío)
                        rect = fitz.Rect(left, top, right, bottom)

                        # Insertar imagen (logo) en ese rect
                        try:
                            page.insert_image(rect, filename=self.logo_path)
                            logo_inserted = True
                            # Emitir mensaje a UI
                            self.message.emit(f"Logo insertado en última página sobre texto: '{raw_text[:60]}'", False)
                        except Exception as ie:
                            # Si algo falla al insertar, reportarlo
                            self.message.emit(f"Error insertando logo: {ie}", True)
                        break
                # fin spans
            # fin lines
        # fin bloques

        if not logo_inserted:
            # Si no lo encuentra, se muestra un aviso en interfaz
            self.message.emit(f"⚠ No se encontró ningún nombre en {os.path.basename(pdf_path)}; no se insertó logo.", True)

        # Guardar cambios en el PDF (incremental)
        try:
            doc.save(pdf_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
        except Exception:
            # si falla incremental, intentar guardar completo (más seguro)
            try:
                doc.save(pdf_path)
            except Exception as e_save:
                self.message.emit(f"Error guardando PDF tras insertar logo: {e_save}", True)
        finally:
            doc.close()


# =====================================
# APP PRINCIPAL (estructura original preservada)
# =====================================
class FirmaDigitalApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Firma Digital Masiva - Certificación de Documentos By: Erick.Z")
        self.resize(600, 600)

        # Ruta base del logo por defecto 
        default_logo = os.path.join(os.getcwd(), "firma", "logo.png")
        # Si el rol seleccionado puso una ruta relativa, construir la ruta absoluta y usarla si existe
        if SELECTED_LOGO_REL:
            candidate = os.path.join(os.getcwd(), SELECTED_LOGO_REL)
            self.logo_path = candidate if os.path.exists(candidate) else (default_logo if os.path.exists(default_logo) else None)
        else:
            self.logo_path = default_logo if os.path.exists(default_logo) else None

        self.cert_path = None
        self.cert_password = None
        self.files_to_sign = []
        self.output_dir = None  # Guardar carpeta de salida

        layout = QVBoxLayout()

        # Botón de Instrucciones 

        btn_help = QPushButton("📘 Instrucciones de uso")
        btn_help.clicked.connect(self.show_instructions)
        layout.addWidget(btn_help)

        btn_files = QPushButton("Seleccionar archivos (PDF/Word)")
        btn_files.clicked.connect(self.load_files)
        layout.addWidget(btn_files)

        self.files_display = QTextEdit()
        self.files_display.setReadOnly(True)
        layout.addWidget(self.files_display)

        btn_cert = QPushButton("Seleccionar certificado (.pfx)")
        btn_cert.clicked.connect(self.load_cert)
        layout.addWidget(btn_cert)

        self.cert_display = QLineEdit()
        self.cert_display.setReadOnly(True)
        layout.addWidget(self.cert_display)

        form_layout = QFormLayout()
        self.cert_pass_input = QLineEdit()
        self.cert_pass_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("Contraseña del certificado:", self.cert_pass_input)
        layout.addLayout(form_layout)

        btn_sign = QPushButton("Firmar masivamente")
        btn_sign.clicked.connect(self.start_signing)
        layout.addWidget(btn_sign)

        # Botón para abrir carpeta de firmados
        btn_open_folder = QPushButton("Ver firmados")
        btn_open_folder.clicked.connect(self.open_firmados_folder)
        layout.addWidget(btn_open_folder)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)
        layout.addWidget(self.status_box)

        self.setLayout(layout)
    
    def show_instructions(self):
        instrucciones = (
            "📘 Instrucciones de uso:\n\n"
            "1. Seleccione el rol (Director o Coordinador).\n"
            "2️. Haga clic en 'Seleccionar archivos' y cargue documentos PDF o Word.\n"
            "   - Los .doc y .docx se convierten automáticamente a PDF.\n"
            "3️. Seleccione su certificado (.pfx) y escriba la contraseña.\n"
            "4️. Pulse 'Firmar masivamente'.\n"
            "   - El sistema normaliza cada PDF.\n"
            "   - Inserta el logo correspondiente al rol.\n"
            "   - Firma digitalmente con el certificado.\n"
            "5️. Revise la carpeta 'firmados' para encontrar los documentos.\n"
            "6️. Use el botón 'Ver firmados' para abrir directamente la carpeta.\n\n"
            "⚠ Nota: Si ya existen archivos con el mismo nombre, se le preguntará "
            "si desea reemplazarlos."
        )
        QMessageBox.information(self, "Guía de uso", instrucciones)

    def log_message(self, message, error=False):
        if error:
            self.status_box.append(f"❌ ERROR: {message}")
        else:
            self.status_box.append(f"✅ {message}")

    def load_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Seleccionar archivos", "", "Documentos (*.pdf *.docx *.doc)"
        )
        if files:
            self.files_to_sign = files
            self.files_display.clear()
            self.files_display.append("\n".join(files))
            self.log_message(f"{len(files)} archivos cargados correctamente.")

    def load_cert(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar certificado", "", "Certificados (*.pfx)")
        if file:
            self.cert_path = file
            self.cert_display.setText(file)
            self.log_message("Certificado cargado correctamente.")

    def start_signing(self):
        if not self.files_to_sign:
            self.log_message("Debe seleccionar al menos un archivo.", True)
            return

        if not self.cert_path:
            self.log_message("Debe seleccionar un certificado.", True)
            return

        self.cert_password = self.cert_pass_input.text().encode("utf-8")

        try:
            with open(self.cert_path, "rb") as f:
                private_key, cert, extra_certs = pkcs12.load_key_and_certificates(
                    f.read(), self.cert_password
                )

            cert_asn1 = x509.Certificate.load(cert.public_bytes(serialization.Encoding.DER))
            extra_asn1 = [x509.Certificate.load(c.public_bytes(serialization.Encoding.DER)) for c in (extra_certs or [])]

            private_bytes = private_key.private_bytes(
                encoding=serialization.Encoding.DER,
                format=serialization.PrivateFormat.PKCS8,
                encryption_algorithm=serialization.NoEncryption()
            )
            key_asn1 = PrivateKeyInfo.load(private_bytes)

            cert_store = SimpleCertificateStore()
            for c in extra_asn1:
                cert_store.register(c)

            signer = signers.SimpleSigner(
                signing_cert=cert_asn1,
                signing_key=key_asn1,
                cert_registry=cert_store
            )

        except Exception as e:
            self.log_message(f"No se pudo preparar el certificado: {e}", True)
            traceback.print_exc()
            return

        root_dir = os.path.join(os.getcwd(), "firmados")
        os.makedirs(root_dir, exist_ok=True)

        options = ["Guardar en la carpeta raíz", "Guardar en subcarpeta existente", "Crear nueva subcarpeta"]
        choice, ok = QInputDialog.getItem(self, "Guardar archivos firmados", "Seleccione opción de guardado:", options, 0, False)

        if not ok:
            self.log_message("Proceso cancelado por el usuario.", True)
            return

        if choice == options[0]:
            self.output_dir = root_dir
        elif choice == options[1]:
            sub_dir = QFileDialog.getExistingDirectory(self, "Seleccione subcarpeta existente dentro de 'firmados'", root_dir)
            if not sub_dir:
                self.log_message("No se seleccionó carpeta. Proceso cancelado.", True)
                return
            self.output_dir = sub_dir
        else:
            folder_name, ok = QInputDialog.getText(self, "Crear nueva subcarpeta", "Nombre de la nueva carpeta dentro de 'firmados':")
            if not ok or not folder_name.strip():
                self.log_message("No se proporcionó nombre de carpeta. Proceso cancelado.", True)
                return
            self.output_dir = os.path.join(root_dir, folder_name.strip())
            os.makedirs(self.output_dir, exist_ok=True)

        # --------------------------
        # Comprobación de archivos existentes
        # --------------------------
        # Generar la lista de rutas de salida esperadas
        expected_outputs = []
        for f in self.files_to_sign:
            name = os.path.basename(f)
            outname = name.replace(".docx", ".pdf").replace(".doc", ".pdf")
            expected_outputs.append(os.path.join(self.output_dir, outname))

        existing = [p for p in expected_outputs if os.path.exists(p)]

        if existing:
            # Mostrar cuadro de advertencia en hilo principal
            reply = QMessageBox.question(
                self,
                "Archivos existentes",
                f"Se encontraron {len(existing)} archivo(s) con el mismo nombre en la carpeta de salida.\n"
                "¿Desea reemplazarlos? (Sí = reemplazar, No = cancelar operación)",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )

            if reply == QMessageBox.No:
                self.log_message("Operación cancelada por el usuario (no se reemplazarán archivos existentes).", True)
                return
            else:
                # usuario seleccionó reemplazar -> se sobrescribirán
                self.log_message(f"{len(existing)} archivo(s) existentes serán reemplazados.", False)

        # Iniciar hilo
        self.sign_thread = SignThread(self.files_to_sign, self.cert_path, self.cert_password, self.logo_path, self.output_dir, signer)
        self.sign_thread.progress.connect(self.progress_bar.setValue)
        self.sign_thread.message.connect(self.log_message)
        self.sign_thread.finished_process.connect(self.final_message)
        self.sign_thread.start()

    def final_message(self):
        QMessageBox.information(self, "Proceso finalizado!", "Revise la ventana de la interfaz para ver logs")

    def open_firmados_folder(self):
        if self.output_dir and os.path.exists(self.output_dir):
            if sys.platform == "win32":
                os.startfile(self.output_dir)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self.output_dir])
            else:
                subprocess.Popen(["xdg-open", self.output_dir])
        else:
            QMessageBox.warning(self, "Carpeta no disponible", "No se ha definido la carpeta de documentos firmados.")


# =====================================
# MAIN
# =====================================
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # pedir rol antes de crear la ventana principal
    seleccionar_rol()

    ventana = FirmaDigitalApp()
    ventana.show()
    sys.exit(app.exec_())
