import streamlit as st
import pandas as pd
from pathlib import Path


def split_excel_sheets(file_path):
    """
    # Функция для разбиения листов Excel на отдельные файлы
    :param file_path:
    :return:
    """
    xls = pd.ExcelFile(file_path)

    output_directory = file_path.parent / f'Листы {file_path.stem}'
    output_directory.mkdir(exist_ok=True)
    created_files = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        output_file = output_directory / f"{sheet_name}.xlsx"

        if output_file.exists():
            output_file.unlink()

        output_file = output_directory / f"{sheet_name}.xlsx"
        if df.columns.isnull().all() or (df.columns.str.contains('^Unnamed', regex=True).all()):
            df.to_excel(output_file, index=False, header=False, sheet_name=sheet_name)
        else:
            df.to_excel(output_file, index=False, header=True, sheet_name=sheet_name)

        created_files.append(output_file)

    return created_files


# Streamlit
st.set_page_config(page_icon=None, layout='wide', initial_sidebar_state='auto')
st.title("Создание файлов excel из листов всех файлов xlsx в директории.")

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

if st.button("Обработать все файлы в директории"):
    if directory:
        directory_path = Path(directory)
        if directory_path.is_dir():
            all_files = list(directory_path.glob('*.xlsx'))
            if not all_files:
                st.write("В указанной директории нет файлов Excel.")
            else:
                for file_path in all_files:
                    created_files = split_excel_sheets(file_path)
                    st.write(f"Из файла `{file_path.name}` созданы файлы в папке `{file_path.stem}`:")
                    for file in created_files:
                        st.write(f"- {file.resolve()}")
        else:
            st.write("Указанная директория не существует.")
    else:
        st.write("Вы не выбрали директорию.")
