import openpyxl


def copy_columns_to_sheets(file_path: str, source_sheet_name: str, exclude_sheets: tuple = None):
    """
    Копирование колонок из одного листа source_sheet_name excel файла file_path за исключением списка
    exclude_sheets листов
    """
    # TODO вынести имена колонок для копирования во внешние параметры
    # Открываем файл с помощью openpyxl
    book = openpyxl.load_workbook(file_path)
    source_sheet = book[source_sheet_name]
    result_str = []
    # Определение листов, которые необходимо исключить
    if exclude_sheets is None:
        exclude_sheets = tuple()

    # Считывание данных из исходного листа
    source_data = {}
    for row in source_sheet.iter_rows(min_row=4, values_only=True):
        p_num = row[0]
        # TODO подставить новые координаты значений для копирования
        tech_queue = row[2]
        next_sz = row[3]
        if p_num is not None:
            source_data[p_num] = (tech_queue, next_sz)

    # Проход по каждому листу
    for sheet_name in book.sheetnames:
        if sheet_name != source_sheet_name and sheet_name not in exclude_sheets:
            sheet = book[sheet_name]
            # Запись данных в соответствующие ячейки
            # TODO откорректировать интервал
            for row_idx, row in enumerate(sheet.iter_rows(min_row=4, max_col=1, values_only=True), start=4):
                p_num = row[0]
                if p_num in source_data:
                    tech_queue, next_sz = source_data[p_num]
                    # Пропускаем объединенные ячейки
                    for cell in sheet[row_idx]:
                        if cell.coordinate in sheet.merged_cells:
                            continue
                        # TODO подставить реальные номера колонок
                        if cell.column == 3:  # Очерёдность
                            cell.value = tech_queue
                        elif cell.column == 4:  # Следующее СЗ
                            cell.value = next_sz
                    if tech_queue:
                        success_copy_str = (f"Скопировано в {sheet_name}: {p_num} - Очередность = {tech_queue},"
                                            f" Следующее СЗ = {next_sz}")  # строка успехов для вывода в streamlit
                        result_str.append(success_copy_str)
    book.save(file_path)
    book.close()
    return result_str


if __name__ == '__main__':
    # tst_file_path = r'C:\Users\user-18\Desktop\1\Трудоемкость серия М (в работе).xlsx'
    # tst_source_sheet_name = '12000М+'
    # tst_exclude_sheets = ('Интерполяция М', 'Интерполяция R', 'Лист1', 'Sheet1', 'Интерполяция SV', 'Интерполяция P')
    # copy_columns_to_sheets(tst_file_path, tst_source_sheet_name, tst_exclude_sheets)
    pass
