import streamlit as st
from streamlit_option_menu import option_menu
import Form308, Form813 ,Form815,Form003004,Form309

st.set_page_config(page_title="Batch Upload")

class MultiApp:
    def __init__(self):
        self.apps = []

    def add_app(self, title, func):
        self.apps.append({
            "title": title,
            "function": func
        })

    def run(self):
        app = None
        with st.sidebar:
            st.markdown("<h1 style='color: orange; font-weight: bold;'>Batch Uploading</h1>", unsafe_allow_html=True)
            app = option_menu(
                menu_title="",
                options=["Form308","Form309", "Form813","Form815","Form003004"],
                icons=["house-fill", "house-fill","house-fill", "house-fill","house-fill"],
                menu_icon="chat-text-fill",
                default_index=1,
                styles={
                    "container": {"padding": "5!important", "background-color": 'black'},
                    "icon": {"color": "white", "font-size": "23px"},
                    "nav-link": {"color": "white", "font-size": "15px", "text-align": "left", "margin": "0px", "--hover-color": "blue"},
                    "nav-link-selected": {"background-color": "#02ab21"},
                }
            )
        for app_dict in self.apps:
            if app == app_dict["title"]:
                app_dict["function"]()

# Create an instance of MultiApp
multi_app = MultiApp()

# Add apps to the MultiApp instance
multi_app.add_app("Form308", Form308.app)
multi_app.add_app("Form309", Form309.app)
multi_app.add_app("Form813", Form813.app)
multi_app.add_app("Form815", Form815.app)
multi_app.add_app("Form003004", Form003004.app)


# Run the MultiApp instance
multi_app.run()
