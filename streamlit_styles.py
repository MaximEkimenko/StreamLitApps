# здесь хранится переменная для дополнительной ручной стилизации css приложения streamlit

streamlit_app_style = """
    <style>
    /* Скрываем меню Streamlit */
    #MainMenu {visibility: hidden;}
    /* Скрываем иконку Streamlit в браузере */
    header {visibility: hidden;}
    /* other css */
    div.stTextInput div input 
        {
            border: 1px solid lightgrey;
            border-radius: 5px; 
        }
     .stButton button
        {
        color: green;
        border: 1px solid black; 
        }
    
    </style>
"""
# for test
