import os
import camelot

# 🔹 Ajouter Ghostscript au PATH
GS_PATH = r"C:\Program Files\gs\gs10.06.0\bin"
os.environ['PATH'] += f";{GS_PATH}"

def pdf_to_excel(pdf_path, excel_path):
    try:
        print(f"📂 Lecture du fichier PDF : {pdf_path} ...")

        # Utilisation de flavor='stream' pour détecter les colonnes multiples
        tables = camelot.read_pdf(pdf_path, pages="all", flavor="stream")

        if tables.n == 0:
            print("❌ Aucun tableau trouvé dans le PDF.")
            return

        # Exporter tous les tableaux dans un fichier Excel
        tables.export(excel_path, f="excel")
        print(f"✅ Conversion terminée ! Résultat enregistré dans : {excel_path}")

    except Exception as e:
        print(f"⚠️ Erreur lors de la conversion : {e}")

if __name__ == "__main__":
    pdf_file = "JUIN_2025.pdf"
    excel_file = "resultat.xlsx"
    pdf_to_excel(pdf_file, excel_file)
