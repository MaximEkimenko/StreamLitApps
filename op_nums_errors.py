import streamlit as st
from pathlib import Path

from copy_columns_to_sheets import copy_columns_to_sheets
from utils_op_nums_errors import compare_headers_across_sheets, compare_rows_across_sheets, validate_file_path
from utils_op_nums_errors import validate_sheet_names
from streamlit_styles import streamlit_app_style


def main():
    """
    Реализация интерфейса для работы с набором утилит по редактированию технологических процессов
    """
    st.set_page_config(page_icon=None, layout='wide', initial_sidebar_state='auto')
    st.title("Утилиты для работы с таблицами xlsx технологических процессов.")
    # стилизация
    st.markdown(streamlit_app_style, unsafe_allow_html=True)
    # утилита обработки ошибок нумерации операций
    st.markdown("### Утилита для проверки схожести номеров операций.")
    main_word = 'Трудоемкость'  # обязательное слово имени файла для расчёта
    ban_word = 'Отличия листов'  # слово в имени файла для исключения из файла расчёта
    all_files = []
    # st.markdown("##### Укажите директорию для проверки Excel файлов на равенство номеров операций:")
    # input для ввода директории
    directory = st.text_input("Укажите директорию для проверки Excel файлов на равенство номеров операций:",
                              key="directory_input", placeholder="Введите путь к директории")
    if st.button("Проверить все файлы Трудоёмкости xlsx в директории"):
        if directory:
            directory_path = Path(directory)
            if directory_path.is_dir():
                # формирование списка файлов
                for file in list(directory_path.glob('*.xlsx')):
                    if main_word in str(file) and ban_word not in str(file):
                        all_files.append(file)
                if not all_files:
                    st.error("В указанной директории нет файлов Excel с Трудоемкостью в именем .")
                else:
                    # список колонок для исключения из расчёта
                    tst_exclude_columns = ('ГОСТ и тип сварочного шва',
                                           'Стоимость часа, руб',
                                           'Трудоёмкость на 1 котел, чел/час',
                                           '№ рабочего центра',
                                           'Загрузка оборудования на 1 котел, часов',
                                           'Кол-во ед./заготовок на 1 котел',
                                           'Ссылка на чертежи',
                                           'Расценка за объем работ, руб.',
                                           'Рабочий центр',
                                           'Объём работ (максимальный в смену)',
                                           'Трудоёмкость на 1 ед./заготовку, чел/час',
                                           'ед.изм.',
                                           'Численность, чел.',
                                           'Трудоёмкость в смену, час'
                                           )
                    for file_path in all_files:
                        tst_threshold = 80  # минимально допустимый уровень сходства
                        if compare_headers_across_sheets(file_path):
                            st.success(f"Начало расчёта для {file_path}")
                            to_streamlit = compare_rows_across_sheets(file_path, tst_exclude_columns,
                                                                      threshold=tst_threshold)
                        else:
                            to_streamlit = f'Обнаружены ошибки в заголовках файла {file_path}. Смотри консоль.'

                        st.write(to_streamlit)
                st.success('Расчёт окончен.')
            else:
                st.error("Указанная директория не существует.")
        else:
            st.error("Вы не выбрали директорию.")
    # утилита копирования данных внутри таблицы excel
    st.markdown("### Утилита для копирования колонок в файле excel.")
    # Поле для ввода пути к файлу
    file_path = st.text_input("Введите путь к файлу .xlsx", placeholder="Введите путь к файлу .xlsx")
    # Поле для ввода имени листа источника
    source_sheet_name = st.text_input("Введите имя листа источника", placeholder="Введите имя листа источника")
    # Поле для ввода списка листов для исключения (через запятую)
    default_exclude_sheets = ('Интерполяция М', 'Интерполяция R', 'Лист1', 'Sheet1',
                              'Интерполяция SV', 'Интерполяция P')
    exclude_sheets_input = st.text_input(f"Введите имена листов для исключения (через запятую). "
                                         f"Список по умолчанию: {default_exclude_sheets}",
                                         placeholder=default_exclude_sheets)
    st.warning('Перед запуском копирования закройте файл excel в котором происходит копирование!')
    if st.button("Запустить копирование"):
        # Обрабатываем список листов для исключения
        if exclude_sheets_input:
            exclude_sheets = tuple([sheet.strip() for sheet in exclude_sheets_input.split(",")]
                                   if exclude_sheets_input else [])
        else:
            exclude_sheets = default_exclude_sheets

        # Проверяем корректность всех данных
        if validate_file_path(file_path) and validate_sheet_names(file_path, source_sheet_name):
            st.success("Все данные введены корректно. Запуск копирования.")
            # Вызываем функцию копирования с параметрами
            try:
                to_streamlit = copy_columns_to_sheets(file_path, source_sheet_name, exclude_sheets)
            except PermissionError:
                st.error("Закройте файл Excel!")
                return
            st.success("Копирование завершено.")
            st.write(to_streamlit)
        else:
            st.error("Ошибка в вводе данных. Проверьте введённые значения.")


if __name__ == "__main__":
    main()
