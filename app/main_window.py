from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QDoubleSpinBox, QCheckBox, QFrame, QStatusBar,
    QMessageBox, QComboBox, QStackedWidget, QGridLayout,
    QButtonGroup, QRadioButton, QLineEdit,
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QFont
from typing import Optional
from app import word_bridge


class _Worker(QThread):
    finished = pyqtSignal(bool, str)

    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self._func = func
        self._args = args
        self._kwargs = kwargs

    def run(self):
        self.finished.emit(*self._func(*self._args, **self._kwargs))


class _DocsWorker(QThread):
    finished = pyqtSignal(bool, list)

    def run(self):
        ok, names = word_bridge.get_open_documents()
        self.finished.emit(ok, names)


# ---------------------------------------------------------------------------
# Componentes reutilizáveis
# ---------------------------------------------------------------------------

class _ToolCard(QFrame):
    """Card clicável na tela inicial."""
    clicked = pyqtSignal()

    def __init__(self, icon: str, title: str, desc: str, parent=None):
        super().__init__(parent)
        self.setFixedHeight(140)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._normal = "QFrame { background: white; border: 1px solid #dee2e6; border-radius: 12px; }"
        self._hover  = "QFrame { background: #e8f4fd; border: 2px solid #0078d4; border-radius: 12px; }"
        self.setStyleSheet(self._normal)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(5)
        layout.setContentsMargins(10, 12, 10, 12)

        icon_lbl = QLabel(icon)
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_lbl.setStyleSheet("font-size: 30px; background: transparent; border: none;")

        title_lbl = QLabel(title)
        title_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        f = QFont()
        f.setBold(True)
        f.setPointSize(11)
        title_lbl.setFont(f)
        title_lbl.setStyleSheet("background: transparent; border: none;")

        desc_lbl = QLabel(desc)
        desc_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc_lbl.setWordWrap(True)
        desc_lbl.setStyleSheet(
            "color: #6c757d; font-size: 10px; background: transparent; border: none;"
        )

        layout.addWidget(icon_lbl)
        layout.addWidget(title_lbl)
        layout.addWidget(desc_lbl)

    def enterEvent(self, event):
        self.setStyleSheet(self._hover)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.setStyleSheet(self._normal)
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)


def _primary_btn(text: str) -> QPushButton:
    btn = QPushButton(text)
    btn.setMinimumHeight(38)
    btn.setStyleSheet("""
        QPushButton { background-color: #0078d4; color: white;
            border-radius: 6px; font-weight: bold; font-size: 13px; }
        QPushButton:hover   { background-color: #106ebe; }
        QPushButton:pressed { background-color: #005a9e; }
        QPushButton:disabled{ background-color: #b0b0b0; color: #e0e0e0; }
    """)
    return btn


def _back_btn() -> QPushButton:
    btn = QPushButton("← Voltar")
    btn.setFixedHeight(28)
    btn.setStyleSheet(
        "QPushButton { font-size: 12px; color: #0078d4; border: none;"
        " background: transparent; font-weight: bold; }"
        "QPushButton:hover { color: #005a9e; }"
    )
    return btn


def _sep() -> QFrame:
    f = QFrame()
    f.setFrameShape(QFrame.Shape.HLine)
    f.setStyleSheet("color: #dee2e6;")
    return f


# ---------------------------------------------------------------------------
# Janela principal
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word Tools")
        self.setMinimumSize(440, 540)
        self.resize(460, 580)
        self._build_ui()
        self._run_connection_test(silent=True)

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(16, 16, 16, 8)
        root.setSpacing(8)

        # Título
        title = QLabel("Word Tools")
        f = QFont()
        f.setPointSize(17)
        f.setBold(True)
        title.setFont(f)
        root.addWidget(title)

        # Status Word
        status_row = QHBoxLayout()
        self._word_status_lbl = QLabel("Verificando Word…")
        self._word_status_lbl.setStyleSheet("color: #6c757d; font-size: 12px;")
        status_row.addWidget(self._word_status_lbl)
        status_row.addStretch()
        test_btn = self._small_btn("Testar Conexão")
        test_btn.clicked.connect(lambda: self._run_connection_test(silent=False))
        status_row.addWidget(test_btn)
        root.addLayout(status_row)

        # Seletor de documento
        doc_row = QHBoxLayout()
        doc_row.addWidget(QLabel("Documento:"))
        self._doc_combo = QComboBox()
        self._doc_combo.setPlaceholderText("Nenhum documento aberto")
        self._doc_combo.setStyleSheet(
            "QComboBox { padding: 3px 6px; border: 1px solid #adb5bd;"
            " border-radius: 4px; background: white; min-height: 24px; }"
        )
        doc_row.addWidget(self._doc_combo, stretch=1)
        refresh = QPushButton("↻")
        refresh.setFixedSize(28, 28)
        refresh.setToolTip("Atualizar lista de documentos")
        refresh.setStyleSheet(
            "QPushButton { font-size: 15px; border: 1px solid #adb5bd;"
            " border-radius: 4px; background: white; }"
            "QPushButton:hover { background: #e9ecef; }"
        )
        refresh.clicked.connect(self._refresh_documents)
        doc_row.addWidget(refresh)
        root.addLayout(doc_row)

        root.addWidget(_sep())

        # Páginas
        self._stack = QStackedWidget()
        self._stack.addWidget(self._page_launcher())        # 0
        self._stack.addWidget(self._page_resize())          # 1
        self._stack.addWidget(self._page_caption())         # 2
        self._stack.addWidget(self._page_paragraph_config()) # 3
        self._stack.addWidget(self._page_indent())          # 4
        root.addWidget(self._stack)

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Pronto")

    # ------------------------------------------------------------------
    # Página 0 — Launcher
    # ------------------------------------------------------------------
    def _page_launcher(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 4, 0, 0)
        layout.setSpacing(12)

        sub = QLabel("Selecione uma ferramenta")
        sub.setStyleSheet("color: #6c757d; font-size: 12px;")
        layout.addWidget(sub)

        grid = QGridLayout()
        grid.setSpacing(12)

        c1 = _ToolCard("⤢", "Redimensionar Imagens", "Ajusta largura e altura de imagens")
        c1.clicked.connect(lambda: self._stack.setCurrentIndex(1))

        c2 = _ToolCard("🏷", "Adicionar Legendas", "Insere legendas numeradas automáticas")
        c2.clicked.connect(lambda: self._stack.setCurrentIndex(2))

        c3 = _ToolCard("¶", "Configurar Parágrafo", "Alinhamento, recuos e espaçamento das imagens")
        c3.clicked.connect(lambda: self._stack.setCurrentIndex(3))

        c4 = _ToolCard("↵", "Recuo de Imagens", "Define recuo de primeira linha das imagens")
        c4.clicked.connect(lambda: self._stack.setCurrentIndex(4))

        grid.addWidget(c1, 0, 0)
        grid.addWidget(c2, 0, 1)
        grid.addWidget(c3, 1, 0)
        grid.addWidget(c4, 1, 1)
        layout.addLayout(grid)
        layout.addStretch()
        return page

    # ------------------------------------------------------------------
    # Página 1 — Redimensionar imagens
    # ------------------------------------------------------------------
    def _page_resize(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        hdr = QHBoxLayout()
        b = _back_btn()
        b.clicked.connect(lambda: self._stack.setCurrentIndex(0))
        hdr.addWidget(b)
        hdr.addStretch()
        lbl = QLabel("⤢  Redimensionar Imagens")
        f = QFont(); f.setBold(True); f.setPointSize(12)
        lbl.setFont(f)
        hdr.addWidget(lbl)
        hdr.addStretch()
        layout.addLayout(hdr)
        layout.addWidget(_sep())

        row_w = QHBoxLayout()
        row_w.addWidget(QLabel("Largura (cm):"))
        self._width_spin = QDoubleSpinBox()
        self._width_spin.setRange(0.5, 50.0)
        self._width_spin.setValue(10.0)
        self._width_spin.setSingleStep(0.5)
        self._width_spin.setDecimals(1)
        row_w.addWidget(self._width_spin)
        layout.addLayout(row_w)

        self._ratio_check = QCheckBox("Manter proporção  (altura calculada automaticamente)")
        self._ratio_check.setChecked(True)
        self._ratio_check.toggled.connect(self._toggle_height_row)
        layout.addWidget(self._ratio_check)

        self._height_row = QWidget()
        row_h = QHBoxLayout(self._height_row)
        row_h.setContentsMargins(0, 0, 0, 0)
        row_h.addWidget(QLabel("Altura (cm):"))
        self._height_spin = QDoubleSpinBox()
        self._height_spin.setRange(0.5, 50.0)
        self._height_spin.setValue(7.5)
        self._height_spin.setSingleStep(0.5)
        self._height_spin.setDecimals(1)
        row_h.addWidget(self._height_spin)
        self._height_row.setVisible(False)
        layout.addWidget(self._height_row)

        # Alinhamento
        layout.addWidget(_sep())
        layout.addWidget(QLabel("Alinhamento do parágrafo:"))
        aln_row = QHBoxLayout()
        self._resize_aln_none   = QRadioButton("Não alterar")
        self._resize_aln_none.setChecked(True)
        self._resize_aln_left   = QRadioButton("Esquerda")
        self._resize_aln_center = QRadioButton("Centralizado")
        self._resize_aln_right  = QRadioButton("Direita")
        self._resize_aln_grp = QButtonGroup()
        for btn in (self._resize_aln_none, self._resize_aln_left,
                    self._resize_aln_center, self._resize_aln_right):
            self._resize_aln_grp.addButton(btn)
            aln_row.addWidget(btn)
        aln_row.addStretch()
        layout.addLayout(aln_row)
        self._resize_aln_note = QLabel("O documento será salvo e reaberto para aplicar o alinhamento.")
        self._resize_aln_note.setStyleSheet("color: #6c757d; font-size: 11px;")
        self._resize_aln_note.setWordWrap(True)
        self._resize_aln_note.setVisible(False)
        layout.addWidget(self._resize_aln_note)
        self._resize_aln_grp.buttonToggled.connect(
            lambda _btn, _chk: self._resize_aln_note.setVisible(not self._resize_aln_none.isChecked())
        )

        self._resize_btn = _primary_btn("Aplicar a Todas as Imagens")
        self._resize_btn.clicked.connect(self._run_resize)
        layout.addWidget(self._resize_btn)
        layout.addStretch()
        return page

    # ------------------------------------------------------------------
    # Página 2 — Adicionar legendas
    # ------------------------------------------------------------------
    def _page_caption(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        hdr = QHBoxLayout()
        b = _back_btn()
        b.clicked.connect(lambda: self._stack.setCurrentIndex(0))
        hdr.addWidget(b)
        hdr.addStretch()
        lbl = QLabel("🏷  Adicionar Legendas")
        f = QFont(); f.setBold(True); f.setPointSize(12)
        lbl.setFont(f)
        hdr.addWidget(lbl)
        hdr.addStretch()
        layout.addLayout(hdr)
        layout.addWidget(_sep())

        # Prefixo
        prefix_row = QHBoxLayout()
        prefix_row.addWidget(QLabel("Prefixo:"))
        self._btn_figura = QRadioButton("Figura")
        self._btn_figura.setChecked(True)
        self._btn_foto = QRadioButton("Foto")
        self._caption_group = QButtonGroup()
        self._caption_group.addButton(self._btn_figura)
        self._caption_group.addButton(self._btn_foto)
        prefix_row.addWidget(self._btn_figura)
        prefix_row.addWidget(self._btn_foto)
        prefix_row.addStretch()
        layout.addLayout(prefix_row)

        # Alinhamento
        align_lbl = QLabel("Alinhamento:")
        layout.addWidget(align_lbl)
        align_row = QHBoxLayout()
        self._btn_left      = QRadioButton("Esquerda")
        self._btn_center    = QRadioButton("Centralizado")
        self._btn_center.setChecked(True)
        self._btn_right     = QRadioButton("Direita")
        self._btn_justified = QRadioButton("Justificado")
        self._align_group = QButtonGroup()
        for btn in (self._btn_left, self._btn_center, self._btn_right, self._btn_justified):
            self._align_group.addButton(btn)
            align_row.addWidget(btn)
        align_row.addStretch()
        layout.addLayout(align_row)

        # Texto personalizado opcional
        text_row = QHBoxLayout()
        self._caption_text_chk = QCheckBox("Adicionar texto à legenda:")
        self._caption_text_edit = QLineEdit()
        self._caption_text_edit.setPlaceholderText("ex: — Vista frontal do equipamento")
        self._caption_text_edit.setEnabled(False)
        self._caption_text_chk.toggled.connect(self._caption_text_edit.setEnabled)
        text_row.addWidget(self._caption_text_chk)
        layout.addLayout(text_row)
        layout.addWidget(self._caption_text_edit)
        hint_text = QLabel("O texto é inserido após o número  →  Figura 1 — Vista frontal do equipamento")
        hint_text.setStyleSheet("color: #6c757d; font-size: 10px;")
        layout.addWidget(hint_text)

        note = QLabel("O documento será salvo e reaberto automaticamente.")
        note.setStyleSheet("color: #6c757d; font-size: 11px;")
        note.setWordWrap(True)
        layout.addWidget(note)

        self._btn_add_captions = _primary_btn("Adicionar Legendas a Todas as Imagens")
        self._btn_add_captions.clicked.connect(self._run_add_captions)
        layout.addWidget(self._btn_add_captions)
        layout.addStretch()
        return page

    # ------------------------------------------------------------------
    # Página 3 — Configurar parágrafo das imagens
    # ------------------------------------------------------------------
    def _page_paragraph_config(self) -> QWidget:
        from PyQt6.QtWidgets import QScrollArea
        outer = QWidget()
        outer_layout = QVBoxLayout(outer)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 4, 0)
        layout.setSpacing(10)

        hdr = QHBoxLayout()
        b = _back_btn()
        b.clicked.connect(lambda: self._stack.setCurrentIndex(0))
        hdr.addWidget(b)
        hdr.addStretch()
        lbl = QLabel("¶  Configurar Parágrafo")
        f = QFont(); f.setBold(True); f.setPointSize(12)
        lbl.setFont(f)
        hdr.addWidget(lbl)
        hdr.addStretch()
        layout.addLayout(hdr)
        layout.addWidget(_sep())

        # Alinhamento
        layout.addWidget(QLabel("Alinhamento:"))
        aln_row1 = QHBoxLayout()
        self._para_aln_none    = QRadioButton("Não alterar")
        self._para_aln_none.setChecked(True)
        self._para_aln_left    = QRadioButton("Esquerda")
        self._para_aln_center  = QRadioButton("Centralizado")
        aln_row2 = QHBoxLayout()
        self._para_aln_right   = QRadioButton("Direita")
        self._para_aln_just    = QRadioButton("Justificado")
        self._para_aln_grp = QButtonGroup()
        for btn in (self._para_aln_none, self._para_aln_left, self._para_aln_center,
                    self._para_aln_right, self._para_aln_just):
            self._para_aln_grp.addButton(btn)
        for btn in (self._para_aln_none, self._para_aln_left, self._para_aln_center):
            aln_row1.addWidget(btn)
        aln_row1.addStretch()
        for btn in (self._para_aln_right, self._para_aln_just):
            aln_row2.addWidget(btn)
        aln_row2.addStretch()
        layout.addLayout(aln_row1)
        layout.addLayout(aln_row2)

        # Recuos e espaçamentos
        def _opt_spin_row(label: str):
            row = QHBoxLayout()
            chk = QCheckBox(label)
            spin = QDoubleSpinBox()
            spin.setRange(0.0, 20.0)
            spin.setSingleStep(0.25)
            spin.setDecimals(2)
            spin.setValue(0.0)
            spin.setFixedWidth(70)
            spin.setEnabled(False)
            unit = QLabel("cm")
            chk.toggled.connect(spin.setEnabled)
            row.addWidget(chk)
            row.addWidget(spin)
            row.addWidget(unit)
            row.addStretch()
            return row, chk, spin

        row, self._para_li_chk, self._para_li_spin = _opt_spin_row("Recuo esquerdo:")
        layout.addLayout(row)
        row, self._para_ri_chk, self._para_ri_spin = _opt_spin_row("Recuo direito:")
        layout.addLayout(row)
        row, self._para_sb_chk, self._para_sb_spin = _opt_spin_row("Espaço antes:")
        layout.addLayout(row)
        row, self._para_sa_chk, self._para_sa_spin = _opt_spin_row("Espaço depois:")
        layout.addLayout(row)

        # Espaço entre linhas
        layout.addWidget(QLabel("Espaço entre linhas:"))
        ls_row = QHBoxLayout()
        self._para_ls_none   = QRadioButton("Não alterar")
        self._para_ls_none.setChecked(True)
        self._para_ls_single = QRadioButton("Simples")
        self._para_ls_15     = QRadioButton("1,5")
        self._para_ls_grp = QButtonGroup()
        for btn in (self._para_ls_none, self._para_ls_single, self._para_ls_15):
            self._para_ls_grp.addButton(btn)
            ls_row.addWidget(btn)
        ls_row.addStretch()
        layout.addLayout(ls_row)

        note_p = QLabel("O documento será salvo e reaberto automaticamente.")
        note_p.setStyleSheet("color: #6c757d; font-size: 11px;")
        layout.addWidget(note_p)

        btn = _primary_btn("Aplicar Configurações")
        btn.clicked.connect(self._run_paragraph_config)
        layout.addWidget(btn)
        layout.addStretch()

        scroll.setWidget(page)
        outer_layout.addWidget(scroll)
        return outer

    # ------------------------------------------------------------------
    # Página 4 — Recuo de imagens
    # ------------------------------------------------------------------
    def _page_indent(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        hdr = QHBoxLayout()
        b = _back_btn()
        b.clicked.connect(lambda: self._stack.setCurrentIndex(0))
        hdr.addWidget(b)
        hdr.addStretch()
        lbl = QLabel("↵  Recuo de Imagens")
        f = QFont(); f.setBold(True); f.setPointSize(12)
        lbl.setFont(f)
        hdr.addWidget(lbl)
        hdr.addStretch()
        layout.addLayout(hdr)
        layout.addWidget(_sep())

        layout.addWidget(QLabel("Recuo da primeira linha:"))

        self._indent_none = QRadioButton("Sem recuo  (remove recuo existente)")
        self._indent_first = QRadioButton("Primeira linha:")
        self._indent_first.setChecked(True)
        self._indent_grp = QButtonGroup()
        self._indent_grp.addButton(self._indent_none)
        self._indent_grp.addButton(self._indent_first)
        layout.addWidget(self._indent_none)

        first_row = QHBoxLayout()
        first_row.addWidget(self._indent_first)
        self._indent_spin = QDoubleSpinBox()
        self._indent_spin.setRange(0.1, 10.0)
        self._indent_spin.setValue(1.25)
        self._indent_spin.setSingleStep(0.25)
        self._indent_spin.setDecimals(2)
        self._indent_spin.setFixedWidth(70)
        first_row.addWidget(self._indent_spin)
        first_row.addWidget(QLabel("cm"))
        first_row.addStretch()
        layout.addLayout(first_row)
        self._indent_none.toggled.connect(lambda checked: self._indent_spin.setEnabled(not checked))

        note_i = QLabel("O documento será salvo e reaberto automaticamente.")
        note_i.setStyleSheet("color: #6c757d; font-size: 11px;")
        layout.addWidget(note_i)

        btn = _primary_btn("Aplicar Recuo")
        btn.clicked.connect(self._run_indent)
        layout.addWidget(btn)
        layout.addStretch()
        return page

    # ------------------------------------------------------------------
    # Conexão com Word
    # ------------------------------------------------------------------
    def _run_connection_test(self, silent: bool = False):
        self._word_status_lbl.setText("Verificando Word…")
        self._conn_worker = _Worker(word_bridge.test_connection)
        self._conn_worker.finished.connect(
            lambda ok, msg: self._on_connection_result(ok, msg, silent)
        )
        self._conn_worker.start()

    def _on_connection_result(self, ok: bool, message: str, silent: bool):
        if ok:
            self._word_status_lbl.setText(f"● {message}")
            self._word_status_lbl.setStyleSheet("color: #198754; font-size: 12px;")
            self._refresh_documents()
        else:
            self._word_status_lbl.setText("● Word inacessível")
            self._word_status_lbl.setStyleSheet("color: #dc3545; font-size: 12px;")
            self._doc_combo.clear()
            if not silent:
                self._show_error("Erro de conexão com Word", message)

    # ------------------------------------------------------------------
    # Seletor de documento
    # ------------------------------------------------------------------
    def _refresh_documents(self):
        self._docs_worker = _DocsWorker()
        self._docs_worker.finished.connect(self._on_documents_loaded)
        self._docs_worker.start()

    def _on_documents_loaded(self, ok: bool, names: list):
        current = self._doc_combo.currentText()
        self._doc_combo.clear()
        if not ok or not names:
            return
        self._doc_combo.addItems(names)
        if current in names:
            self._doc_combo.setCurrentText(current)

    def _selected_doc(self) -> Optional[str]:
        text = self._doc_combo.currentText()
        return text if text else None

    # ------------------------------------------------------------------
    # Ações — Redimensionar
    # ------------------------------------------------------------------
    def _toggle_height_row(self, checked: bool):
        self._height_row.setVisible(not checked)

    def _run_resize(self):
        doc = self._selected_doc()
        if doc is None:
            self._show_error("Nenhum documento selecionado", "Selecione um documento no seletor acima.")
            return
        if self._resize_aln_left.isChecked():
            alignment = "left"
        elif self._resize_aln_center.isChecked():
            alignment = "center"
        elif self._resize_aln_right.isChecked():
            alignment = "right"
        else:
            alignment = None
        self.statusBar().showMessage("Processando…")
        self._resize_worker = _Worker(
            word_bridge.resize_all_images,
            self._width_spin.value(),
            None if self._ratio_check.isChecked() else self._height_spin.value(),
            self._ratio_check.isChecked(),
            doc,
            "all",
            alignment,
        )
        self._resize_worker.finished.connect(self._on_resize_done)
        self._resize_worker.start()

    def _on_resize_done(self, ok: bool, message: str):
        if ok:
            self.statusBar().showMessage(f"✓  {message}", 6000)
        else:
            self.statusBar().showMessage("✗  Falhou — veja o detalhe", 6000)
            self._show_error("Erro ao redimensionar imagens", message)

    # ------------------------------------------------------------------
    # Ações — Legendas
    # ------------------------------------------------------------------
    def _selected_alignment(self) -> str:
        if self._btn_left.isChecked():
            return "left"
        if self._btn_right.isChecked():
            return "right"
        if self._btn_justified.isChecked():
            return "justified"
        return "center"

    def _run_add_captions(self):
        doc = self._selected_doc()
        if doc is None:
            self._show_error("Nenhum documento selecionado", "Selecione um documento no seletor acima.")
            return
        label = "Figura" if self._btn_figura.isChecked() else "Foto"
        alignment = self._selected_alignment()
        caption_text = self._caption_text_edit.text().strip() if self._caption_text_chk.isChecked() else ""
        self.statusBar().showMessage("Adicionando legendas…")
        self._caption_worker = _Worker(word_bridge.add_captions_to_images, label, alignment, doc, caption_text)
        self._caption_worker.finished.connect(self._on_captions_done)
        self._caption_worker.start()

    def _on_captions_done(self, ok: bool, message: str):
        if ok:
            self.statusBar().showMessage("✓  Legendas adicionadas", 6000)
            self._run_connection_test(silent=True)
        else:
            self.statusBar().showMessage("✗  Falhou — veja o detalhe", 6000)
            self._show_error("Erro ao adicionar legendas", message)

    # ------------------------------------------------------------------
    # Ações — Configurar parágrafo
    # ------------------------------------------------------------------
    def _run_paragraph_config(self):
        doc = self._selected_doc()
        if doc is None:
            self._show_error("Nenhum documento selecionado", "Selecione um documento no seletor acima.")
            return

        aln = None
        if self._para_aln_left.isChecked():    aln = "left"
        elif self._para_aln_center.isChecked(): aln = "center"
        elif self._para_aln_right.isChecked():  aln = "right"
        elif self._para_aln_just.isChecked():   aln = "justified"

        ls = None
        if self._para_ls_single.isChecked(): ls = "single"
        elif self._para_ls_15.isChecked():   ls = "1.5"

        self.statusBar().showMessage("Aplicando configurações…")
        self._para_worker = _Worker(
            word_bridge.configure_image_paragraphs,
            aln,
            self._para_li_spin.value() if self._para_li_chk.isChecked() else None,
            self._para_ri_spin.value() if self._para_ri_chk.isChecked() else None,
            self._para_sb_spin.value() if self._para_sb_chk.isChecked() else None,
            self._para_sa_spin.value() if self._para_sa_chk.isChecked() else None,
            ls,
            "all",
            doc,
        )
        self._para_worker.finished.connect(self._on_paragraph_config_done)
        self._para_worker.start()

    def _on_paragraph_config_done(self, ok: bool, message: str):
        if ok:
            self.statusBar().showMessage(f"✓  {message}", 6000)
        else:
            self.statusBar().showMessage("✗  Falhou — veja o detalhe", 6000)
            self._show_error("Erro ao configurar parágrafo", message)

    # ------------------------------------------------------------------
    # Ações — Recuo
    # ------------------------------------------------------------------
    def _run_indent(self):
        doc = self._selected_doc()
        if doc is None:
            self._show_error("Nenhum documento selecionado", "Selecione um documento no seletor acima.")
            return
        indent_cm = 0.0 if self._indent_none.isChecked() else self._indent_spin.value()
        self.statusBar().showMessage("Aplicando recuo…")
        self._indent_worker = _Worker(
            word_bridge.set_image_first_line_indent,
            indent_cm,
            "all",
            doc,
        )
        self._indent_worker.finished.connect(self._on_indent_done)
        self._indent_worker.start()

    def _on_indent_done(self, ok: bool, message: str):
        if ok:
            self.statusBar().showMessage(f"✓  {message}", 6000)
        else:
            self.statusBar().showMessage("✗  Falhou — veja o detalhe", 6000)
            self._show_error("Erro ao aplicar recuo", message)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _show_error(self, title: str, message: str):
        dlg = QMessageBox(self)
        dlg.setWindowTitle(title)
        dlg.setIcon(QMessageBox.Icon.Warning)
        dlg.setText(message)
        dlg.exec()

    @staticmethod
    def _small_btn(text: str) -> QPushButton:
        btn = QPushButton(text)
        btn.setFixedHeight(26)
        btn.setStyleSheet(
            "QPushButton { font-size: 11px; padding: 0 8px; border: 1px solid #adb5bd;"
            " border-radius: 4px; background: white; }"
            "QPushButton:hover { background: #e9ecef; }"
        )
        return btn
