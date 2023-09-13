import os
import xlrd
from openpyxl import Workbook
import shutil

class FileConverter:
    def __init__(self, input_folder: str, output_folder: str) -> None:
        """
        Инициализация конвертера файлов.

        :param input_folder: Папка с исходными файлами.
        :param output_folder: Папка для сохранения результата.
        """
        self.input_folder = input_folder
        self.output_folder = output_folder

    def convert_and_move_files(self) -> None:
        """
        Преобразует файлы формата .xls в .xlsx и перемещает .xlsx файлы в целевую папку.
        """
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)

        files = os.listdir(self.input_folder)

        for file in files:
            if file.endswith('.xls'):
                self._convert_xls_to_xlsx(file)
            elif file.endswith('.xlsx'):
                self._move_xlsx(file)

        print('Преобразование и перенос завершены')

    def _convert_xls_to_xlsx(self, file: str) -> None:
        """
        Преобразует файл формата .xls в .xlsx.

        :param file: Имя файла для преобразования.
        """
        xls_file_path = os.path.join(self.input_folder, file)
        new_xlsx_file_name = f"{file.replace('.xls', '.xlsx')}"
        new_xlsx_file_path = os.path.join(self.output_folder, new_xlsx_file_name)

        xls_workbook = xlrd.open_workbook(xls_file_path)
        xlsx_workbook = Workbook()
        xlsx_sheet = xlsx_workbook.active

        for sheet in xls_workbook.sheets():
            for row_idx in range(sheet.nrows):
                row_data = sheet.row_values(row_idx)
                xlsx_sheet.append(row_data)

        xlsx_workbook.save(new_xlsx_file_path)

        print(f'Файл {file} успешно преобразован и сохранен как {new_xlsx_file_name} в папке {self.output_folder}')

    def _move_xlsx(self, file: str) -> None:
        """
        Перемещает файл формата .xlsx в целевую папку.

        :param file: Имя файла для перемещения.
        """
        xlsx_file_path = os.path.join(self.input_folder, file)
        new_xlsx_file_path = os.path.join(self.output_folder, file)

        shutil.move(xlsx_file_path, new_xlsx_file_path)

        print(f'Файл {file} успешно перенесен в папку {self.output_folder}')

if __name__ == "__main__":
    converter = FileConverter('Prices', 'Prices/xlsx')
    converter.convert_and_move_files()
