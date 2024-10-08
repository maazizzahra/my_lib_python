import openpyxl
import logging

class ExcelProcessor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.workbook = None
        self.sheet = None

    def __enter__(self):
        self.workbook = openpyxl.load_workbook(self.file_path)
        self.sheet = self.workbook.active
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.workbook:
            self.workbook.close()

    def calculate_totals(self, data_column, result_column):
        """
        Calcule le total des valeurs dans la colonne 'data_column' et écrit le résultat dans la colonne 'result_column'.
        """
        total = 0
        for row in range(2, self.sheet.max_row + 1):
            value = self.sheet[f"{data_column}{row}"].value
            if isinstance(value, (int, float)):
                total += value
        self.sheet[f"{result_column}{1}"] = "Total"
        self.sheet[f"{result_column}{2}"] = total
        logging.info(f"Total des colonnes {data_column} calculé: {total}")

    def save(self, output_file):
        """
        Sauvegarde le fichier Excel modifié.
        """
        self.workbook.save(output_file)
        logging.info(f"Fichier Excel enregistré sous {output_file}")
