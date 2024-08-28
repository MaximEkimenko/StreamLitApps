chcp 65001
@echo off

:: Путь к виртуальной среде
set VENV_PATH="D:\АСУП\Python\Projects\StreamlitApps\venv\Scripts\activate"

:: Активировать виртуальную среду
call %VENV_PATH%

:: Запуск Streamlit приложения
streamlit run "D:\АСУП\Python\Projects\StreamlitApps\split_excel_sheets.py"

:: Завершение работы
deactivate




