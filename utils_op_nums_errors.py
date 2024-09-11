import pandas as pd
from rapidfuzz import fuzz
import os
import openpyxl


BAN_LIST = []  # чтения списка игнорируемых листов файла
with open('op_ban_list.txt', 'r', encoding='utf-8') as ban_file:
    for line in ban_file.readlines():
        BAN_LIST.append(line.strip())


def compare_headers_across_sheets(file_path: str) -> bool:
    """
    Проверяет одинаковость заголовков внутри файла Excel. За эталон взят первый лист.
    Возвращает False если хотя бы в одном листе есть несоответствие.
    """
    xlsx = pd.ExcelFile(file_path)
    reference_headers = None
    all_sheets_ok = []  # bool список безошибочных листов
    # Проход по каждому листу
    for sheet_name in xlsx.sheet_names:
        if sheet_name in BAN_LIST:
            continue
        # Чтение заголовков листа с указанием, что заголовки начинаются со строки №4
        df = pd.read_excel(xlsx, sheet_name=sheet_name, header=3)

        # Убираем столбцы с именами Unnamed
        current_headers = [col for col in df.columns if not col.startswith("Unnamed")]

        # Если эталонный заголовок еще не задан, назначаем его
        if reference_headers is None:
            reference_headers = current_headers
        else:
            # Сравнение текущих заголовков с эталоном
            if current_headers != reference_headers:
                all_sheets_ok.append(False)
                print(f"Различие в заголовках на листе '{sheet_name}':")
                for ref_col, cur_col in zip(reference_headers, current_headers):
                    if ref_col != cur_col:
                        print(f"Ожидаемый заголовок: '{ref_col}', а полученный: '{cur_col}'")
                # Проверяем, есть ли лишние или отсутствующие столбцы
                if len(reference_headers) != len(current_headers):
                    if len(reference_headers) > len(current_headers):
                        print("Отсутствуют заголовки:", reference_headers[len(current_headers):])
                    else:
                        print("Лишние заголовки:", current_headers[len(reference_headers):])
                print()
            else:
                all_sheets_ok.append(True)
    if all(all_sheets_ok):
        print(f'Все заголовки в порядке для {file_path} - продолжение расчёта.')
    return all(all_sheets_ok)


def compare_rows_across_sheets(file_path: str, exclude_columns: list = None, threshold: int = 80) -> list:
    """
    Сравнивает значения строк с номером п/п во всех листах Excel, исключая заданные колонки.
    Строки сравниваются с учетом нечеткости с порогом схожести threshold.
    """
    if exclude_columns is None:
        exclude_columns = []
    # Загрузить Excel-файл
    xls = pd.ExcelFile(file_path)
    # Инициализация словаря для хранения строк по номерам п/п
    pp_dict = {}
    # Проход по каждому листу
    for sheet_name in xls.sheet_names:
        if sheet_name in BAN_LIST:
            continue
        df = pd.read_excel(xls, sheet_name=sheet_name, header=3)
        # обрезка df до строки с ИТОГО
        end_index = df.index[df.iloc[:, 1].str.contains('ИТОГО', na=False)].tolist()[0]
        df = df.head(end_index)
        # Исключение Unnamed колонок
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        # Проход по строкам текущего листа
        for idx, row in df.iterrows():
            pp_num = row['п/п']
            if pd.isna(pp_num):  # исключаем pp_num, равные NaN
                continue  # новая строка
            # Если номер п/п уже есть в словаре, добавить строку к этому номеру
            if pp_num in pp_dict:
                pp_dict[pp_num].append((sheet_name, row))
            else:
                # Иначе создать новый список для этого номера п/п
                pp_dict[pp_num] = [(sheet_name, row)]

    # Сохранение результатов в файл
    base_sheet_name = xls.sheet_names[0]
    report_lines = []
    result_wb = openpyxl.Workbook()
    result_sh = result_wb.active
    # Сравнение строк с одинаковыми номерами п/п
    for pp_num, rows in pp_dict.items():
        if len(rows) > 1:
            # Получение эталонной строки (первой) для сравнения
            base_row = rows[0][1]
            for sheet_name, row in rows[1:]:
                # Сравнение строки с эталоном
                differences = []
                for col in df.columns:
                    if col not in exclude_columns:  # Исключаем указанные колонки
                        base_value = str(base_row[col])
                        current_value = str(row[col])
                        # Сравнение строк с использованием rapidfuzz
                        similarity = fuzz.ratio(base_value, current_value)
                        if similarity < threshold:
                            differences.append(
                                f"Столбец '{col}': '{base_value}' != '{current_value}' (схожесть: {similarity}%)")

                # Если найдены различия, вывести их
                if differences:
                    comparison_result = (f"Различия найдены для п/п {pp_num} "
                                         f"в исходном листе '{base_sheet_name}' с листом '{sheet_name}':")
                    print(comparison_result)
                    result_sh.append([comparison_result])  # сохраняем в excel
                    report_lines.append(comparison_result)

                    for diff in differences:
                        print(diff)
                        result_sh.append([diff])
                        report_lines.append(diff)
                    result_sh.append([''])  # пустая строка
                    report_lines.append('')

    # Сохранение отчета в xlsx файл
    report_filename = os.path.join(
        os.path.dirname(file_path),
        f"Отличия листов {os.path.basename(file_path)}.xlsx"
    )
    result_wb.save(report_filename)
    return report_lines


def validate_file_path(file_path: str) -> bool:
    """
    Валидация пути к файлу file_path: файл должен быть xlsx.
    """
    try:
        # Проверка, что файл существует и является .xlsx файлом
        if file_path.endswith('.xlsx'):
            pd.ExcelFile(file_path)  # Попытка открыть файл
            return True
        else:
            return False
    except Exception as e:
        return False


def validate_sheet_names(file_path: str, source_sheet_name: str) -> bool:
    """
    Валидация имени листа источника копирования
    """
    try:
        # Открываем файл и получаем список листов
        xls = pd.ExcelFile(file_path)
        available_sheets = xls.sheet_names
        # Проверяем, что лист-источник существует
        if source_sheet_name not in available_sheets:
            return False
        return True
    except Exception as e:
        return False


if __name__ == '__main__':
    pass
    # tst_file_path = r'D:\АСУП\Python\Projects\OmzitTerminal\misc\Трудоемкость серия М (в работе).xlsx'
    # tst_exclude_columns = ('ГОСТ и тип сварочного шва',
    #                        'Стоимость часа, руб',
    #                        'Трудоёмкость на 1 котел, чел/час',
    #                        '№ рабочего центра',
    #                        'Загрузка оборудования на 1 котел, часов',
    #                        'Кол-во ед./заготовок на 1 котел',
    #                        'Ссылка на чертежи',
    #                        'Расценка за объем работ, руб.',
    #                        'Рабочий центр',
    #                        'Объём работ (максимальный в смену)',
    #                        'Трудоёмкость на 1 ед./заготовку, чел/час',
    #                        'ед.изм.',
    #                        'Численность, чел.',
    #                        'Трудоёмкость в смену, час'
    #                        )
    # tst_threshold = 80
    # if compare_headers_across_sheets(tst_file_path):
    #     compare_rows_across_sheets(tst_file_path, tst_exclude_columns, threshold=tst_threshold)
