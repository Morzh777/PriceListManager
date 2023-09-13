import pandas as pd
from datetime import datetime
from data.data import column_mappings
import os

class PriceFileProcessor:
    def __init__(self) -> None:
        """
        Инициализация обработчика файлов с ценами.
        """
        self.columns = ["Штрихкод", "Наименование", "Производитель", "Цена", "Остаток", "Срок годности", "Поставщик", "ЖВ", "Цена реестра", "НДС"]

    def process_price_files(self) -> None:
        """
        Обрабатывает файлы с ценами, объединяя их в один сводный файл.
        """
        current_date = datetime.now().strftime("%Y-%m-%d")
        your_file_name = f"Сводный на {current_date}.xlsx"
        your_df = pd.DataFrame(columns=self.columns)
        suppliers = []

        price_files = sorted([file for file in os.listdir("Prices/xlsx") if file.endswith(".xlsx")])

        for file in price_files:
            file_path = os.path.join("Prices/xlsx", file)
            try:
                price_df = self._read_and_process_file(file_path)
            except Exception as e:
                print(f"Ошибка при чтении файла {file_path}: {e}")
                continue

            new_rows = self._map_columns(price_df, file)
            your_df = pd.concat([your_df, pd.DataFrame(new_rows)], ignore_index=True)
            suppliers.extend([file] * len(new_rows))

        your_df["Поставщик"] = suppliers
        your_df.to_excel(your_file_name, index=False)
        print(f"Файл {your_file_name} успешно создан с данными из всех файлов в папке Prices.")

    def _read_and_process_file(self, file_path: str) -> pd.DataFrame:
        """
        Читает и обрабатывает файл с ценами.

        :param file_path: Путь к файлу для чтения.
        :return: Обработанный DataFrame.
        """
        price_df = pd.read_excel(file_path)

        # Преобразование строковых значений в числовой формат
        numeric_columns = ["Штрихкод", "Остаток", "Цена реестра", "Цена"]
        for column in numeric_columns:
            if column in price_df.columns:
                price_df[column] = pd.to_numeric(price_df[column], errors='coerce')

        # Преобразование столбца "Срок годности" в формат день месяц год
        if "Срок годности" in price_df.columns:
            price_df["Срок годности"] = pd.to_datetime(price_df["Срок годности"], errors='coerce', format='%d.%m.%Y')

        return price_df

    def _map_columns(self, price_df: pd.DataFrame, file: str) -> list:
        """
        Преобразует столбцы в соответствии с маппингом.

        :param price_df: DataFrame с исходными данными.
        :param file: Имя файла.
        :return: Список новых строк для добавления в сводный файл.
        """
        new_rows = []
        for index, row in price_df.iterrows():
            new_row = {}
            for target_column, source_columns in column_mappings.items():
                if target_column != "Поставщик":
                    for source_column in source_columns:
                        if source_column in price_df.columns:
                            new_row[target_column] = row[source_column]
                            break
                else:
                    new_row[target_column] = file
            new_rows.append(new_row)
        return new_rows

if __name__ == "__main__":
    processor = PriceFileProcessor()
    processor.process_price_files()
