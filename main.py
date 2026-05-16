import sys

from PyQt6.QtWidgets import QApplication, QMessageBox

from app.license_check import get_stored_key, validate_key
from app.license_dialog import LicenseDialog
from app.main_window import MainWindow


def _check_license(app: QApplication) -> bool:
    key = get_stored_key()
    if key:
        ok, msg = validate_key(key)
        if ok:
            return True
        # Chave inválida ou desativada — pede nova chave
        if msg != "Sem conexão com a internet e chave não armazenada.":
            QMessageBox.warning(
                None,
                "Licença inválida",
                f"{msg}\n\nPor favor, insira uma chave válida.",
            )

    dialog = LicenseDialog()
    result = dialog.exec()
    return result == LicenseDialog.DialogCode.Accepted


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Word Tools")

    if not _check_license(app):
        sys.exit(0)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
