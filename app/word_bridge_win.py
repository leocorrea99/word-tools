"""Controle do Microsoft Word via COM automation (Windows)."""
from __future__ import annotations

import os
import time
from typing import Optional


# ---------------------------------------------------------------------------
# Helpers internos
# ---------------------------------------------------------------------------

def _get_word():
    """
    Retorna o objeto COM do Word ativo.
    Chama CoInitialize para compatibilidade com threads (QThread).
    """
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("Word.Application")
    except Exception:
        raise RuntimeError("word_not_running")


def _get_doc(word, doc_name: Optional[str] = None):
    if word.Documents.Count == 0:
        raise RuntimeError("no_document")
    if doc_name:
        for i in range(1, word.Documents.Count + 1):
            if word.Documents(i).Name == doc_name:
                return word.Documents(i)
        raise RuntimeError(f"Documento '{doc_name}' não encontrado.")
    return word.ActiveDocument


def _friendly_error(raw: str) -> str:
    if "word_not_running" in raw:
        return "Microsoft Word não está em execução. Abra o Word e tente novamente."
    if "no_document" in raw:
        return "Nenhum documento Word aberto."
    if "Permission" in raw or "Access" in raw:
        return "Permissão negada. Verifique se o Word está autorizado."
    return raw


def _close_and_reopen(file_path: str) -> tuple[bool, str]:
    try:
        word = _get_word()
        for i in range(1, word.Documents.Count + 1):
            try:
                if os.path.normcase(word.Documents(i).FullName) == os.path.normcase(file_path):
                    word.Documents(i).Close(SaveChanges=0)  # wdDoNotSaveChanges
                    break
            except Exception:
                pass
        time.sleep(0.4)
        word.Documents.Open(file_path)
        word.Activate()
        return True, "ok"
    except Exception as e:
        return False, _friendly_error(str(e))


# ---------------------------------------------------------------------------
# API pública — espelho de word_bridge_mac.py
# ---------------------------------------------------------------------------

def test_connection() -> tuple[bool, str]:
    try:
        word = _get_word()
        if word.Documents.Count == 0:
            return True, "Word acessível — nenhum documento aberto"
        return True, f"Word acessível — documento: {word.ActiveDocument.Name}"
    except Exception as e:
        return False, _friendly_error(str(e))


def get_open_documents() -> tuple[bool, list[str]]:
    try:
        word = _get_word()
        names = [word.Documents(i).Name for i in range(1, word.Documents.Count + 1)]
        return True, names
    except Exception:
        return True, []


def save_and_get_path(doc_name: Optional[str] = None) -> tuple[bool, str]:
    try:
        word = _get_word()
        doc = _get_doc(word, doc_name)
        full = doc.FullName
        if not full or full == doc.Name:
            return False, (
                "O documento ainda não foi salvo.\n\n"
                "Salve o arquivo com Ctrl+S e tente novamente."
            )
        doc.Save()
        return True, full
    except Exception as e:
        return False, _friendly_error(str(e))


def resize_all_images(
    width_cm: float,
    height_cm: Optional[float] = None,
    maintain_ratio: bool = True,
    doc_name: Optional[str] = None,
    scope: str = "all",
    alignment: Optional[str] = None,
) -> tuple[bool, str]:
    # msoTrue = -1, msoFalse = 0
    MSO_TRUE = -1
    MSO_FALSE = 0

    try:
        word = _get_word()
        doc = _get_doc(word, doc_name)

        width_pt = width_cm * 28.3465
        total = doc.InlineShapes.Count
        if total == 0:
            return False, "Nenhuma imagem (inline) encontrada no documento."

        for i in range(1, total + 1):
            shape = doc.InlineShapes(i)
            if maintain_ratio:
                shape.LockAspectRatio = MSO_TRUE
                shape.Width = width_pt
            else:
                height_pt = (height_cm or width_cm) * 28.3465
                shape.LockAspectRatio = MSO_FALSE
                shape.Width = width_pt
                shape.Height = height_pt

        count_msg = f"{total} imagem(ns) redimensionada(s)"

        if not alignment:
            return True, f"{count_msg} com sucesso!"

        # Alinhamento via python-docx (requer salvar/fechar/reabrir)
        from app.docx_utils import configure_image_paragraphs_in_file
        ok_path, path = save_and_get_path(doc_name)
        if not ok_path:
            return True, f"{count_msg} com sucesso! (Alinhamento não aplicado: {path})"
        if not path.lower().endswith(".docx"):
            return True, f"{count_msg} com sucesso! (Alinhamento ignorado — requer .docx)"
        try:
            configure_image_paragraphs_in_file(path, alignment=alignment)
        except Exception as e:
            return True, f"{count_msg} com sucesso! (Erro no alinhamento: {e})"
        ok2, msg2 = _close_and_reopen(path)
        if not ok2:
            return True, f"{count_msg} com alinhamento aplicado, mas falha ao recarregar:\n{msg2}"
        return True, f"{count_msg} com alinhamento aplicado!"

    except Exception as e:
        return False, _friendly_error(str(e))


def add_captions_to_images(
    label_name: str = "Figura",
    alignment: str = "center",
    doc_name: Optional[str] = None,
    caption_text: str = "",
) -> tuple[bool, str]:
    from app.docx_utils import add_captions_to_file

    ok, path = save_and_get_path(doc_name)
    if not ok:
        return False, path
    if not path.lower().endswith(".docx"):
        return False, "Esta operação requer formato .docx.\n\nSalve como .docx e tente novamente."

    try:
        added, skipped = add_captions_to_file(path, label_name, alignment, caption_text)
    except Exception as e:
        return False, f"Erro ao processar o arquivo: {e}"

    ok2, msg2 = _close_and_reopen(path)
    if not ok2:
        return False, (
            f"Legendas gravadas, mas falha ao recarregar no Word:\n{msg2}\n\n"
            "Feche e reabra o documento manualmente."
        )

    parts = []
    if added:
        parts.append(f"{added} legenda(s) adicionada(s)")
    if skipped:
        parts.append(f"{skipped} imagem(ns) já tinham legenda (ignoradas)")
    msg = "; ".join(parts) if parts else "Nenhuma legenda nova adicionada"
    msg += "\n\nPressione Ctrl+A → F9 no Word para atualizar os números."
    return True, msg


def configure_image_paragraphs(
    alignment: Optional[str] = None,
    left_indent_cm: Optional[float] = None,
    right_indent_cm: Optional[float] = None,
    space_before_cm: Optional[float] = None,
    space_after_cm: Optional[float] = None,
    line_spacing: Optional[str] = None,
    scope: str = "all",
    doc_name: Optional[str] = None,
) -> tuple[bool, str]:
    from app.docx_utils import configure_image_paragraphs_in_file

    if not any([alignment, left_indent_cm is not None, right_indent_cm is not None,
                space_before_cm is not None, space_after_cm is not None, line_spacing]):
        return False, "Nenhuma configuração selecionada para aplicar."

    ok, path = save_and_get_path(doc_name)
    if not ok:
        return False, path
    if not path.lower().endswith(".docx"):
        return False, "Esta operação requer formato .docx.\n\nSalve como .docx e tente novamente."

    try:
        count = configure_image_paragraphs_in_file(
            path, alignment, left_indent_cm, right_indent_cm,
            space_before_cm, space_after_cm, line_spacing, None
        )
    except Exception as e:
        return False, f"Erro ao processar o arquivo: {e}"

    ok2, msg2 = _close_and_reopen(path)
    if not ok2:
        return False, (
            f"Configurações aplicadas, mas falha ao recarregar:\n{msg2}\n\n"
            "Feche e reabra o documento manualmente."
        )
    return True, f"Configurações aplicadas em {count} parágrafo(s) com imagem!"


def set_image_first_line_indent(
    indent_cm: float = 0.0,
    scope: str = "all",
    doc_name: Optional[str] = None,
) -> tuple[bool, str]:
    from app.docx_utils import set_image_first_line_indent_in_file

    ok, path = save_and_get_path(doc_name)
    if not ok:
        return False, path
    if not path.lower().endswith(".docx"):
        return False, "Esta operação requer formato .docx.\n\nSalve como .docx e tente novamente."

    try:
        count = set_image_first_line_indent_in_file(path, indent_cm, None)
    except Exception as e:
        return False, f"Erro ao processar o arquivo: {e}"

    ok2, msg2 = _close_and_reopen(path)
    if not ok2:
        return False, (
            f"Recuo aplicado, mas falha ao recarregar:\n{msg2}\n\n"
            "Feche e reabra o documento manualmente."
        )

    indent_label = f"{indent_cm:.2f} cm" if indent_cm > 0 else "removido"
    return True, f"Recuo ({indent_label}) aplicado em {count} parágrafo(s) com imagem!"
