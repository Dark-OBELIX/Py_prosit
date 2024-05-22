import sys
from PyQt5.QtWidgets import QApplication, QFileDialog, QWidget
from PyQt5.QtCore import QDir

app = QApplication(sys.argv)
dialog = QFileDialog()  # Définissez la boîte de dialogue comme une variable globale

def get_docx_file():
    dialog.setFileMode(QFileDialog.AnyFile)
    dialog.setFilter(QDir.Files)

    if dialog.exec_():
        file_name = dialog.selectedFiles()
        if file_name[0].endswith('.docx'):
            print(file_name[0])
            return True, str(file_name[0])
        else:
            print("Erreur : Le fichier sélectionné n'est pas un docx")
            return False
    else:
        print("Erreur : Pas de fichier sélectionné")
        return False

importOk, doc_name = get_docx_file()
print(doc_name)

# Fermez explicitement la boîte de dialogue en indiquant que l'opération a été acceptée.
dialog.setResult(QFileDialog.Accepted)

