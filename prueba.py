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
    message = pyqtSignal (str, bool)
    finished_progress = pyqtSignal() # Señal para indicar que finalizo el proceso

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
                if file.endswith((".docx", ".pdf")):
                    self.message.emit(f"Convirtiendo {filename} a PDF...", False)
                    temp_pdf = os.path.join(file, temp_pdf)
                    docx2pdf.convert(file, temp_pdf)
                    file = temp_pdf
                    self.message.emit(f"Archivo convertido: {file}", False)

                # Normalizar pdf
                self.message.emit(f"Normalizando PDF {filename}...", False)
                doc = 