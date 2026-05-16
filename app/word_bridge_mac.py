import os
import subprocess
import tempfile
from typing import Optional


def _run_applescript(script: str) -> tuple[bool, str]:
    # Usa arquivo temporário (mais confiável que -e para scripts multilinhas)
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".applescript", delete=False, encoding="utf-8"
    ) as f:
        f.write(script)
        tmp = f.name
    try:
        result = subprocess.run(
            ["osascript", tmp],
            capture_output=True,
            text=True,
        )
        if result.returncode == 0:
            return True, result.stdout.strip()
        err = (result.stderr or result.stdout or "Erro desconhecido").strip()
        return False, err
    finally:
        os.unlink(tmp)


def _friendly_error(raw: str) -> str:
    """Converte erros técnicos do AppleScript em mensagens legíveis."""
    if "-1743" in raw:
        return (
            "Permissão negada.\n\n"
            "Vá em: Preferências do Sistema → Privacidade e Segurança "
            "→ Automação → habilite o Terminal (ou o app que rodou esta ferramenta) "
            "para controlar o Microsoft Word."
        )
    if "-600" in raw or "is not running" in raw.lower():
        return "Microsoft Word não está em execução. Abra o Word e tente novamente."
    if "no_document" in raw:
        return "Nenhum documento Word aberto."
    if "no_images" in raw:
        return "Nenhuma imagem (inline) encontrada no documento."
    return raw


def test_connection() -> tuple[bool, str]:
    """Verifica se consegue se comunicar com o Word."""
    script = """
tell application "Microsoft Word"
    if (count of documents) = 0 then
        return "ok:no_doc"
    end if
    return "ok:" & (name of active document)
end tell
"""
    ok, out = _run_applescript(script)
    if not ok:
        return False, _friendly_error(out)
    if out.startswith("ok:no_doc"):
        return True, "Word acessível — nenhum documento aberto"
    doc = out.removeprefix("ok:")
    return True, f"Word acessível — documento: {doc}"


def get_open_documents() -> tuple[bool, list[str]]:
    """Retorna lista com os nomes de todos os documentos abertos no Word."""
    script = """
tell application "Microsoft Word"
    if (count of documents) = 0 then
        return ""
    end if
    set docList to ""
    repeat with i from 1 to count of documents
        if i > 1 then
            set docList to docList & "|||"
        end if
        set docList to docList & (name of document i)
    end repeat
    return docList
end tell
"""
    ok, out = _run_applescript(script)
    if not ok:
        return False, []
    if not out:
        return True, []
    return True, out.split("|||")


def _doc_target(doc_name: Optional[str]) -> str:
    """Retorna a referência AppleScript ao documento correto."""
    if doc_name:
        safe = doc_name.replace('"', '\\"')
        return f'document "{safe}"'
    return "active document"


def save_and_get_path(doc_name: Optional[str] = None) -> tuple[bool, str]:
    """
    Salva o documento e retorna seu caminho POSIX.
    Retorna (True, '/caminho/arquivo.docx') ou (False, mensagem_erro).
    """
    target = _doc_target(doc_name)
    script = f"""
tell application "Microsoft Word"
    if (count of documents) = 0 then
        return "NO_DOC"
    end if
    set targetDoc to {target}
    set docFullName to full name of targetDoc
    if docFullName = "" then
        return "NOT_SAVED"
    end if
    save targetDoc
    return POSIX path of (docFullName as alias)
end tell
"""
    ok, out = _run_applescript(script)
    if not ok:
        return False, _friendly_error(out)
    if out == "NO_DOC":
        return False, "Nenhum documento Word aberto."
    if out == "NOT_SAVED":
        return False, (
            "O documento ainda não foi salvo.\n\n"
            "Salve o arquivo com Cmd+S e tente novamente."
        )
    return True, out.strip()


def _close_and_reopen(file_path: str) -> tuple[bool, str]:
    """Fecha o documento pelo caminho e o reabre."""
    safe = file_path.replace('"', '\\"')
    script = f"""
tell application "Microsoft Word"
    set docToClose to missing value
    set n to count of documents
    repeat with i from 1 to n
        set d to document i
        set dp to ""
        try
            set dp to POSIX path of (full name of d as alias)
        end try
        if dp = "{safe}" then
            set docToClose to d
            exit repeat
        end if
    end repeat
    if docToClose is not missing value then
        close docToClose saving no
    end if
    open POSIX file "{safe}"
    activate
    return "ok"
end tell
"""
    ok, out = _run_applescript(script)
    if not ok:
        return False, _friendly_error(out)
    return True, "ok"


def get_selected_shape_indices(doc_name: Optional[str] = None) -> tuple[bool, list[int]]:
    """
    Retorna os índices 1-based das inline shapes selecionadas no Word.
    Retorna (True, []) se nada selecionado; (False, []) em caso de erro.
    """
    target = _doc_target(doc_name)
    script = f"""
tell application "Microsoft Word"
    if (count of documents) = 0 then return "NO_DOC"
    tell {target}
        set selShapes to inline shapes of selection
        set selCount to count of selShapes
        if selCount = 0 then return "NONE"
        set selNames to {{}}
        repeat with s in selShapes
            set end of selNames to (name of s as string)
        end repeat
        set result to ""
        repeat with i from 1 to count of inline shapes
            set shapeName to name of inline shape i as string
            repeat with sn in selNames
                if shapeName = (sn as string) then
                    if result is not "" then set result to result & ","
                    set result to result & (i as string)
                    exit repeat
                end if
            end repeat
        end repeat
        if result is "" then return "NONE"
        return result
    end tell
end tell
"""
    ok, out = _run_applescript(script)
    if not ok:
        return False, []
    if out in ("NONE", "NO_DOC", ""):
        return True, []
    try:
        return True, [int(x) for x in out.split(",") if x.strip()]
    except ValueError:
        return False, []


def add_captions_to_selected_images(
    label_name: str = "Figura",
    alignment: str = "center",
    doc_name: Optional[str] = None,
    caption_text: str = "",
) -> tuple[bool, str]:
    """
    Adiciona legenda apenas às imagens selecionadas no Word.
    O usuário deve selecionar as imagens (clique / Cmd+clique) antes de executar.
    """
    from app.docx_utils import add_captions_to_selected_in_file

    ok, indices = get_selected_shape_indices(doc_name)
    if not ok:
        return False, "Erro ao detectar seleção no Word."
    if not indices:
        return False, (
            "Nenhuma imagem selecionada no Word.\n\n"
            "Clique em uma imagem para selecioná-la "
            "(ou Cmd+clique para selecionar várias) e tente novamente."
        )

    ok, path = save_and_get_path(doc_name)
    if not ok:
        return False, path

    if not path.lower().endswith(".docx"):
        return False, (
            "Esta operação requer formato .docx.\n\n"
            "Salve o documento como .docx e tente novamente."
        )

    try:
        added, skipped = add_captions_to_selected_in_file(path, label_name, alignment, set(indices), caption_text)
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


def add_captions_to_images(
    label_name: str = "Figura",
    alignment: str = "center",
    doc_name: Optional[str] = None,
    caption_text: str = "",
) -> tuple[bool, str]:
    """
    Adiciona legenda nativa do Word (campo SEQ auto-numerado) em imagens
    que ainda não têm legenda. Requer documento salvo como .docx.
    """
    from app.docx_utils import add_captions_to_file

    ok, path = save_and_get_path(doc_name)
    if not ok:
        return False, path

    if not path.lower().endswith(".docx"):
        return False, (
            "Esta operação requer formato .docx.\n\n"
            "Salve o documento como .docx e tente novamente."
        )

    try:
        added, skipped = add_captions_to_file(path, label_name, alignment, caption_text)
    except Exception as e:
        return False, f"Erro ao processar o arquivo: {e}"

    ok2, msg2 = _close_and_reopen(path)
    if not ok2:
        return False, (
            f"Legendas gravadas no arquivo, mas falha ao recarregar no Word:\n{msg2}\n\n"
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
    """
    Configura o formato do parágrafo das imagens via python-docx (OOXML).
    Salva, modifica e reabre o documento.
    """
    from app.docx_utils import configure_image_paragraphs_in_file

    if not any([alignment, left_indent_cm is not None, right_indent_cm is not None,
                space_before_cm is not None, space_after_cm is not None, line_spacing]):
        return False, "Nenhuma configuração selecionada para aplicar."

    selected = None
    if scope == "selected":
        ok_sel, indices = get_selected_shape_indices(doc_name)
        if not ok_sel:
            return False, "Erro ao detectar seleção no Word."
        if not indices:
            return False, "Nenhuma imagem selecionada no Word.\n\nClique em uma imagem e tente novamente."
        selected = set(indices)

    ok, path = save_and_get_path(doc_name)
    if not ok:
        return False, path
    if not path.lower().endswith(".docx"):
        return False, "Esta operação requer formato .docx.\n\nSalve como .docx e tente novamente."

    try:
        count = configure_image_paragraphs_in_file(
            path, alignment, left_indent_cm, right_indent_cm,
            space_before_cm, space_after_cm, line_spacing, selected
        )
    except Exception as e:
        return False, f"Erro ao processar o arquivo: {e}"

    ok2, msg2 = _close_and_reopen(path)
    if not ok2:
        return False, f"Configurações aplicadas, mas falha ao recarregar:\n{msg2}\n\nFeche e reabra o documento manualmente."

    label = "selecionado(s)" if selected else "com imagem"
    return True, f"Configurações aplicadas em {count} parágrafo(s) {label}!"


def set_image_first_line_indent(
    indent_cm: float = 0.0,
    scope: str = "all",
    doc_name: Optional[str] = None,
) -> tuple[bool, str]:
    """
    Define o recuo de primeira linha dos parágrafos de imagens via python-docx.
    indent_cm=0 remove o recuo. scope: "all" ou "selected".
    """
    from app.docx_utils import set_image_first_line_indent_in_file

    selected = None
    if scope == "selected":
        ok_sel, indices = get_selected_shape_indices(doc_name)
        if not ok_sel:
            return False, "Erro ao detectar seleção no Word."
        if not indices:
            return False, "Nenhuma imagem selecionada no Word.\n\nClique em uma imagem e tente novamente."
        selected = set(indices)

    ok, path = save_and_get_path(doc_name)
    if not ok:
        return False, path
    if not path.lower().endswith(".docx"):
        return False, "Esta operação requer formato .docx.\n\nSalve como .docx e tente novamente."

    try:
        count = set_image_first_line_indent_in_file(path, indent_cm, selected)
    except Exception as e:
        return False, f"Erro ao processar o arquivo: {e}"

    ok2, msg2 = _close_and_reopen(path)
    if not ok2:
        return False, f"Recuo aplicado, mas falha ao recarregar:\n{msg2}\n\nFeche e reabra o documento manualmente."

    indent_label = f"{indent_cm:.2f} cm" if indent_cm > 0 else "removido"
    scope_label = "selecionado(s)" if selected else "com imagem"
    return True, f"Recuo ({indent_label}) aplicado em {count} parágrafo(s) {scope_label}!"


def resize_all_images(
    width_cm: float,
    height_cm: Optional[float] = None,
    maintain_ratio: bool = True,
    doc_name: Optional[str] = None,
    scope: str = "all",
    alignment: Optional[str] = None,
) -> tuple[bool, str]:
    """
    Redimensiona imagens inline. scope: "all" ou "selected".
    Unidade entrada: cm. Internamente converte para pontos (1 cm = 28.3465 pt).
    """
    width_pt = round(width_cm * 28.3465, 2)
    target = _doc_target(doc_name)

    def _shape_list(sel: bool) -> str:
        if sel:
            return (
                "        set selShapes to inline shapes of selection\n"
                "        set shapeCount to count of selShapes\n"
                "        if shapeCount = 0 then\n"
                "            return \"no_images\"\n"
                "        end if"
            )
        return (
            "        set shapeCount to count of inline shapes\n"
            "        if shapeCount = 0 then\n"
            "            return \"no_images\"\n"
            "        end if"
        )

    def _shape_ref(sel: bool) -> str:
        return "item i of selShapes" if sel else "inline shape i"

    sel = (scope == "selected")
    shape_list = _shape_list(sel)
    shape_ref = _shape_ref(sel)

    if maintain_ratio:
        script = f"""
tell application "Microsoft Word"
    if (count of documents) = 0 then
        return "no_document"
    end if
    tell {target}
{shape_list}
        repeat with i from 1 to shapeCount
            set aShape to {shape_ref}
            set lock aspect ratio of aShape to true
            set width of aShape to {width_pt}
        end repeat
        return shapeCount as string
    end tell
end tell
"""
    else:
        height_pt = round((height_cm or width_cm) * 28.3465, 2)
        script = f"""
tell application "Microsoft Word"
    if (count of documents) = 0 then
        return "no_document"
    end if
    tell {target}
{shape_list}
        repeat with i from 1 to shapeCount
            set aShape to {shape_ref}
            set lock aspect ratio of aShape to false
            set width of aShape to {width_pt}
            set height of aShape to {height_pt}
        end repeat
        return shapeCount as string
    end tell
end tell
"""

    ok, out = _run_applescript(script)

    if not ok:
        return False, _friendly_error(out)
    if out == "no_document":
        return False, "Nenhum documento Word aberto."
    if out == "no_images":
        return False, "Nenhuma imagem (inline) encontrada no documento."

    try:
        count = int(out)
        count_msg = f"{count} imagem(ns) redimensionada(s)"
    except ValueError:
        count_msg = "Imagens redimensionadas"

    if not alignment:
        return True, f"{count_msg} com sucesso!"

    # Aplica alinhamento via python-docx (requer salvar/fechar/reabrir)
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
