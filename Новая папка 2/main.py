import csv
import xlsxwriter
import argparse


def convert_csv_to_xlsx(input_file, output_file, delimiter=';'):
    # Создание книги Excel
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    # Чтение данных из CSV файла
    with open(input_file, 'r') as csv_file:
        reader = csv.reader(csv_file, delimiter=delimiter)
        row = 0
        for line in reader:
            col = 0
            for item in line:
                worksheet.write(row, col, item)
                col += 1
            row += 1

    # Закрытие книги Excel
    workbook.close()
    print(f"Файл {output_file} успешно создан.")


if __name__ == "__main__":
    # Парсинг аргументов командной строки
    parser = argparse.ArgumentParser(description='Конвертер CSV в XLSX')
    parser.add_argument('input_file', help='Имя входного CSV файла')
    parser.add_argument('output_file', help='Имя выходного XLSX файла')
    parser.add_argument('-d', '--delimiter', default=';', help='Разделитель в CSV файле (по умолчанию - ;)')
    args = parser.parse_args()

    # Вызов функции конвертации
    convert_csv_to_xlsx(args.input_file, args.output_file, args.delimiter)

    #python main.py test_table.csv test_table.xls -d ';'

