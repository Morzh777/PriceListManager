from modules.FileConverter import FileConverter
from modules.PriceFileProcessor import PriceFileProcessor

def main():
    # Первый этап: конвертация и перемещение файлов
    input_folder = 'Prices'
    output_folder = 'Prices/xlsx'
    converter = FileConverter(input_folder, output_folder)
    converter.convert_and_move_files()
    
    # Второй этап: обработка файлов и формирование данных
    processor = PriceFileProcessor()
    processor.process_price_files()

if __name__ == "__main__":
    main()
