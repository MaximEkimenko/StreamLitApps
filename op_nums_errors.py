import streamlit as st
import pandas as pd
from pathlib import Path
from utils_op_nums_errors import compare_headers_across_sheets, compare_rows_across_sheets


# Streamlit
def main():
    st.set_page_config(page_icon=None, layout='wide', initial_sidebar_state='auto')
    st.title("Анализ ошибок нумерации операций технологического процесса.")

    st.markdown(
        """
        <style>
        .styled-text-input {
            border: 2px solid #4CAF50;
            padding: 10px;
            border-radius: 5px;
            box-shadow: 2px 2px 5px rgba(0, 0, 0, 0.1);
            width: 100%;
        }
        /* Скрываем меню Streamlit */
        #MainMenu {visibility: hidden;}
        /* Скрываем иконку Streamlit в браузере */
        header {visibility: hidden;}
        </style>
        """,
        unsafe_allow_html=True
    )
    st.markdown("### Укажите директорию для чтения всех Excel файлов:")
    directory = st.text_input("", key="directory_input", placeholder="Введите путь к директории")

    st.markdown(
        """
        <style>
        div.stTextInput > div > input {
            border: 2px solid #4CAF50;
            padding: 10px;
            border-radius: 5px;
            box-shadow: 2px 2px 5px rgba(0, 0, 0, 0.1);
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    main_word = 'Трудоемкость'
    ban_word = 'Отличия листов'
    all_files = []
    if st.button("Обработать все файлы Трудоёмкости xlsx в директории"):
        if directory:
            directory_path = Path(directory)
            if directory_path.is_dir():
                for file in list(directory_path.glob('*.xlsx')):
                    if main_word in str(file) and ban_word not in str(file):
                        all_files.append(file)
                if not all_files:
                    st.write("В указанной директории нет файлов Excel с Трудоемкостью в именем .")
                else:
                    for file_path in all_files:
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
                        tst_threshold = 80
                        if compare_headers_across_sheets(file_path):
                            st.write(f"Начало расчёта для {file_path}")
                            to_streamlit = compare_rows_across_sheets(file_path, tst_exclude_columns,
                                                                      threshold=tst_threshold)
                        else:
                            to_streamlit = f'Обнаружены ошибки в заголовках файла {file_path}. Смотри консоль.'

                        st.write(to_streamlit)
                st.write('Расчёт окончен.')
            else:
                st.write("Указанная директория не существует.")
        else:
            st.write("Вы не выбрали директорию.")


if __name__ == "__main__":
    main()
