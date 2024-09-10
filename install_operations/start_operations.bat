chcp 65001
@echo off

:: Путь к виртуальной среде
set VENV_PATH="D:\АСУП\Python\Projects\StreamlitApps\venv\Scripts\activate"

:: Активировать виртуальную среду
call %VENV_PATH%

:: Запуск Streamlit приложения
streamlit run "D:\АСУП\Python\Projects\StreamlitApps\op_nums_errors.py" --server.port 8501
pause
:: Завершение работы
deactivate




