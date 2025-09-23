import sys
import os
import fitz
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

# TARGET_TEXT = "WOLFANG ALBERTO LATORRE MARTINEZ" # - COORDINADOR
TARGET_TEXT = "GERARDO ARTURO MEDINA ROSAS" # - DIRECTOR

class SignThread(QThread):
    progress = pyqtSignal(int)
    message = pyqtSignal(str, bool)
    finished_process = pyqtSignal() # Señal para indicar que finalizo el proceso

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
                #conversion de Word -> PDF
                if file.endswith((".docx", ".doc")):
                    self.message.emit(f"Convirtiendo {filename} a PDF...", False)
                    temp_pdf = os.path.join(tempfile.gettempdir(), filename + ".pdf")
                    docx2pdf.convert(file, temp_pdf)
                    file = temp_pdf
                    self.message.emit(f"Archivo convertido: {file}", False)

                # Normalizar pdf
                self.message.emit(f"Normalizando PDF {filename}...", False)
                doc = fitz.open(file)
                temp_normalized = os.path.join(tempfile.gettempdir(), filename + "_normalized.pdf")
                doc.save(temp_normalized)
                doc.close()
                file = temp_normalized
                self.message.emit(f"PDF normalizado {file}", False)

                # Insertar logo dinamico en la ultima pagina
                if self.logo_path:
                    self.message.emit(f"Inserción de logo en {filename}...", False)
                    self.insert_logo_dynamic(file)
                    self.message.emit(f"Logo insertado (o no) segun deteccion.", False)

                #Firmar PDF
                self.message.emit(f"Firmando {filename}...", False)
                with open(file, "rb") as inf:
                    w = IncrementalPdfFileWriter(inf)
                    append_signature_field(w, SigFieldSpec(sig_field_name="FirmaDigital"))

                    # Metadata que ayuda a que acrobat muestre info al hacer click en la firma - Director
                    metadata = signers.PdfSignatureMetadata(
                        field_name="FirmaDigital",
                        name="Gerardo Arturo Medina Rosas",
                        reason="Documento firmado digitalmente",
                        location="Bogotá, Colombia",
                        contact_info="Walatorrem@sena.edu.co",
                        certify=False,  # TRUE (CERTIFICADO) || FALSE (VALIDADO)  (ROLES DIFERENTES)
                        docmdp_permissions=MDPPerm.NO_CHANGES
                    )

                    '''
                    # Metadata que ayuda a que Acrobat muestre info al hacer clic en la firma - Coordinador
                    metadata = signers.PdfSignatureMetadata(
                        field_name="FirmaDigital",
                        name="Wolfang Alberto Latorre Martinez",
                        reason="Documento firmado digitalmente",
                        location="Bogotá, Colombia",
                        contact_info="walatorrem@sena.edu.co",
                        certify=False,  # firma de Aprobación (False) || Certificación (True)  (Roles diferentes)
                        docmdp_permissions=MDPPerm.NO_CHANGES 
                    )
                    ''' 
                    pdf_signed = signers.sign_pdf(w, metadata, signer=self.signer, new_field_spec=None)

                    with open(output_file, "wb") as outf:
                        outf.write(pdf_signed.read())

                    self.message.emit(f"Documento firmado correctamente: {output_file}", False)
            
            except Exception as e:
                self.message.emit(f"Error firmando {filename}: {e}", True)
                traceback.print_exc()
            
            self.progress.emit(int((idx + 1) / total * 100))
        
        # Emitir señal de finalizacion
        self.finished_process.emit()
    
    # Metodo que coloca el logo dinamicamente, buscando variantes del nombre ignoranod tildes y mayusculas
    def insert_logo_dynamic(self, pdf_path):
        """
        Abre pdf_path, busca en la ultima pagina alguno de los nombres de la lista
        (ignorando tildes y mayusculas) y coloca el logo a la derecha y por encima
        del nombre (posocion vertical ajustada dinamicamente).
        Sino encuentra emite un mensaje, no hay fallback
        """
        doc = fitz.open(pdf_path)
        page = doc[-1] # Ultima pagina

        # Lista de posibles nombres
        posibles_nombres = [
            "GERARDO ARTURO MEDINA ROSAS",
            "Gerardo Arturo Medina Rosas",
            "gerardo arturo medina rosas",
            "G. A. Medina Rosas",
            "Gerardo A. Medina",
            "Medina Rosas Gerardo Arturo"
        ]
        '''

        # Lista de posibles nombres - Coordinador
        posibles_nombres = [
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
        '''
        # Normalizar funcion: quitar tildes y pasa a minusculas
        def normalize_text(t: str) -> str:
            if t is None:
                return ""
            # Desconponer y eliminar marcas diacriticas
            t_norm = unicodedata.normalize("NFD", t)
            t_no_accents = "".join(ch for ch in t_norm if unicodedata.category(ch) != "Mn")
            return t_no_accents.lower().strip()
        
        # normalizar la lsita de nombres una sola vez
        nombres_norm = [normalize_text(n) for n in posibles_nombres]
        logo_inserted = False

        # Recorrer los spans (mas preciso que bloques) para encontrar la mejor posicion
        # get_text("dict") devuelve estructura con paginas, bloques, lines, spans con bbox
        try:
            page_dict = page.get_text("dict")
        except Exception:
            #fallback mas simple si get_text("dict") falla
            page_dict = {"blocks": page.get_text("blocks")}

        # Iterar bloques -> lineas -> spans
        for block in page_dict.get("blocks", []):
            if logo_inserted:
                break
            # algunos bloques (imagenes) no tienen "lines"
            lines = block.get("lines", [])
            for line in lines:
                if logo_inserted:
                    break
                for span in line.get("spans", []):
                    raw_text = span.get("text", "")
                    if not raw_text:
                        continue

                    text_norm = normalize_text(raw_text)

                    # Si cualquier nombre normalizado esta contenido en el span normalizado -> coincidencia
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
                        logo_width = 165
                        logo_height = 55

                        # Separacion horizontal y vertical respecto al texto detectado
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

class FirmaDigitalApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Firma Digital Masiva - Certificación de documentos  By: Erick.Z")
        self.resize(600, 600)

        self.logo_path = os.path.join(os.getcwd(), "firma", "logo.png")
        if not os.path.exists(self.logo_path):
            self.logo_path = None
        
        self.cert_path = None
        self.cert_password = None
        self.files_to_sign = []
        self.output_dir = None # Guardar carpeta de salida

        layout = QVBoxLayout()

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

        # Boton para abrir carpeta firmados
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
    
    def log_message(self, message, error=False):
        if error:
            self.status_box.append(f"❌ ERROR: {message}")
        else:
            self.status_box.append(f"✅ {message}")
    
    def load_files(self):
        files, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivos", "", "Documentos (*.pdf *.docx *.doc)"
        )
        if files:
            self.files_to_sign = files
            self.files_display.clear()
            self.files_display.append("\n".join(files))
            self.log_message(f"{len(files)} archivos cargados correctamente.")
    
    def load_cert(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar certificado", "", "Certificados (.pfx)")
        if file:
            self.cert_path = file
            self.cert_display.setText(file)
            self.log_message("Certificado cargado correctamente.")
    
    def start_signing(self):
        if not self.files_to_sign:
            self.log_message("Debe seleccionar un certificado.", True)
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
            extra_asn1 = [x509.Certificate.load(c.public_bytes(serialization.Encoding.DER))for c in (extra_certs or [])]

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
            self.log_message(f"No se pudo preparar el certififcado || Contraseña incorrecta: {e}", True)
            traceback.print_exc()
            return
        
        root_dir = os.path.join(os.getcwd(), "firmados")
        os.makedirs(root_dir, exist_ok=True)

        options = ["Guardar en la arpeta raiz", "Guardar en subcarpeta existente", "Crear nueva subcarpeta"]
        choice, ok = QInputDialog.getItem(self, "Guardar archivos firmados", "Seleccione opción de guardado:", options, 0, False)

        if not ok:
            self.log_message("Proceso cancelado por usuario.", True)
            return
        
        if choice == options[0]:
            self.output_dir = root_dir
        elif choice == options [1]:
            sub_dir = QFileDialog.getExistingDirectory(self, "Seleccione una subcarpeta dentro de firmados", root_dir)
            if not sub_dir:
                self.log_message("No se selecciono carpeta. Proceso cancelado.", True)
                return
            self.output_dir = sub_dir
        else:
            folder_name, ok = QInputDialog.getText(self, "Crear nueva subcarpeta", "Nombre de la nueva carpeta dentro de firmados:")
            if not ok or not folder_name.strip():
                self.log_message("No se proporciono nombre de carpeta. Proceso cancelado.", True)
                return
            self.output_dir = os.path.join(root_dir, folder_name.strip())
            os.makedirs(self.output_dir, exist_ok=True)

        #-----------------------
        # Comprobacion de archivos existentes
        #-----------------------
        #Generar la lista de rutas de salida esperada
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
                f"Se econtraron {len(existing)} archivo (s) con el mismo nombre en la carpeta de salida. \n"
                "¿Desea reemplazarlos? (Si = reemplazar, No = Cancelar operacion)",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )

            if reply == QMessageBox.No:
                self.log_message("Operacion cancelada por el usuario (No se reemplazaran archivos existentes)", True)
                return
            else:
                # Si selecciona reemplazar
                self.log_message(f"{len(existing)} archivo (s) existentes seran reemplazados." , False)
        
        # Iniciar hilo
        self.sign_thread = SignThread(self.files_to_sign, self.cert_path, self.cert_password, self.logo_path, self.output_dir, signer)
        self.sign_thread.progress.connect(self.progress_bar.setValue)
        self.sign_thread.message.connect(self.log_message)
        self.sign_thread.finished_process.connect(self.final_message)
        self.sign_thread.start()

    def final_message(self):
        QMessageBox.information(self, "Proceso finalzado!", "Revise la ventana para ver logs")

    def open_firmados_folder(self):
        if self.output_dir and os.path.exists(self.output_dir):
            if sys.plataform == "win32":
                os.startfile(self.output_dir)
            elif sys.plataform == "darwin":
                subprocess.Popen([open], self.output_dir)
            else:
                subprocess.Pooen(["xdg-open"], self.output_dir)
        else:
            QMessageBox.warning(self, "Carpeta no disponible", "No se ha definido la carpeta documntos firmados.")
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    ventana = FirmaDigitalApp()
    ventana.show()
    sys.exit(app.exec_())