import os
import sys
from win32com.client import Dispatch

# Définir manuellement les constantes Word
wdExportFormatPDF = 17
wdExportOptimizeForPrint = 0
wdExportCreateNoBookmarks = 0

def convert_word_to_pdf(source_folder, dest_folder):
    # Vérifie si les dossiers existent
    if not os.path.exists(source_folder):
        print(f"Le dossier source '{source_folder}' n'existe pas.")
        return

    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)
        print(f"Le dossier destination '{dest_folder}' a été créé.")

    # Initialisation de Word
    word = Dispatch("Word.Application")
    word.Visible = False  # Rendre Word invisible pendant l'exécution

    # Parcourt tous les fichiers dans le dossier source
    for file_name in os.listdir(source_folder):
        # Ignore les fichiers temporaires (~$)
        if file_name.startswith("~$"):
            continue

        # Vérifie les extensions .doc et .docx
        if file_name.lower().endswith((".doc", ".docx")):
            doc_path = os.path.join(source_folder, file_name)
            pdf_name = os.path.splitext(file_name)[0][:200] + ".pdf"  # Tronque les noms trop longs
            pdf_path = os.path.join(dest_folder, pdf_name)

            # Vérifie si le fichier source existe et est accessible
            if not os.path.exists(doc_path):
                print(f"Le fichier n'existe pas ou est inaccessible : {doc_path}")
                continue

            try:
                # Conversion en PDF
                print(f"Conversion en cours : {doc_path} -> {pdf_path}")
                doc = word.Documents.Open(os.path.abspath(doc_path), False, True, False)
                doc.ExportAsFixedFormat(
                    OutputFileName=os.path.abspath(pdf_path),
                    ExportFormat=wdExportFormatPDF,
                    OpenAfterExport=False,
                    OptimizeFor=wdExportOptimizeForPrint,
                    CreateBookmarks=wdExportCreateNoBookmarks,
                )
            except Exception as e:
                print(f"Erreur lors de la conversion de {doc_path} : {e}")
            finally:
                if 'doc' in locals() and doc is not None:
                    doc.Close(False)

    # Quitte Word
    word.Quit()
    print("Conversion terminée.")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Utilisation : python script.py <source_folder> <dest_folder>")
    else:
        source_folder = sys.argv[1]
        dest_folder = sys.argv[2]
        convert_word_to_pdf(source_folder, dest_folder)
