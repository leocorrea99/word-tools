"""Diálogo de ativação de licença."""
from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from app.license_check import save_key, validate_key


class LicenseDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Word Tools — Ativação")
        self.setFixedWidth(460)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 28, 32, 28)
        layout.setSpacing(16)

        title = QLabel("Insira sua chave de licença")
        title.setFont(QFont("", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        subtitle = QLabel(
            "Você pode obter sua chave gratuita em nosso site.\n"
            "A chave foi enviada para o seu e-mail no momento do download."
        )
        subtitle.setWordWrap(True)
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color: #666;")
        layout.addWidget(subtitle)

        self._key_input = QLineEdit()
        self._key_input.setPlaceholderText("Ex.: a1b2c3d4e5f6...")
        self._key_input.setMinimumHeight(36)
        self._key_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self._key_input)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)

        self._activate_btn = QPushButton("Ativar")
        self._activate_btn.setMinimumHeight(36)
        self._activate_btn.setStyleSheet(
            "QPushButton { background:#2563EB; color:white; border-radius:6px; font-weight:bold; }"
            "QPushButton:hover { background:#1D4ED8; }"
        )
        self._activate_btn.clicked.connect(self._on_activate)

        btn_row.addWidget(self._activate_btn)
        layout.addLayout(btn_row)

        self._status_label = QLabel("")
        self._status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._status_label.setWordWrap(True)
        layout.addWidget(self._status_label)

    def _on_activate(self) -> None:
        key = self._key_input.text().strip()
        if not key:
            self._status_label.setText("Por favor, insira a chave de licença.")
            self._status_label.setStyleSheet("color: #DC2626;")
            return

        self._activate_btn.setEnabled(False)
        self._status_label.setText("Validando...")
        self._status_label.setStyleSheet("color: #2563EB;")
        from PyQt6.QtWidgets import QApplication
        QApplication.processEvents()

        ok, msg = validate_key(key)

        self._activate_btn.setEnabled(True)

        if ok:
            save_key(key)
            self._status_label.setText("Licença ativada com sucesso!")
            self._status_label.setStyleSheet("color: #16A34A;")
            self.accept()
        else:
            self._status_label.setText(msg)
            self._status_label.setStyleSheet("color: #DC2626;")
