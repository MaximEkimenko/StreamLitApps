import streamlit as st
import pandas as pd
from pathlib import Path
from streamlit_styles import streamlit_app_style


def split_excel_sheets(file_path: Path) -> list:
    """
    Функция для разбиения листов файла file_path Excel на отдельные файлы
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


def main():
    """
    Интерфейс streamlit для использования функции split_excel_sheets
    :return:
    """
    # стилизация
    st.markdown(streamlit_app_style, unsafe_allow_html=True)
    # основное приложение
    st.title("Создание файлов excel из листов всех файлов xlsx в директории.")
    st.markdown("### Укажите директорию для чтения всех Excel файлов:")
    directory = st.text_input("", key="directory_input", placeholder="Введите путь к директории")
    if st.button("Обработать все файлы в директории"):
        if directory:
            directory_path = Path(directory)
            if directory_path.is_dir():
                all_files = list(directory_path.glob('*.xlsx'))
                if not all_files:
                    st.error("В указанной директории нет файлов Excel.")
                else:
                    for file_path in all_files:
                        created_files = split_excel_sheets(file_path)
                        st.success(f"Из файла `{file_path.name}` созданы файлы в папке `{file_path.stem}`:")
                        for file in created_files:
                            st.success(f"- {file.resolve()}")
            else:
                st.error("Указанная директория не существует.")
        else:
            st.error("Вы не выбрали директорию.")


if __name__ == "__main__":
    main()
