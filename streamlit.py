import streamlit as st
import pandas as pd
import datetime
import tempfile
import os

# Your scheduling engine here (import it or copy logic)



from scheduler_logic import generate_schedule  # Placeholder — we'll adjust!

# --- Streamlit App Starts Here ---
st.set_page_config(page_title="Grafik weterynaryjny", layout="centered")

st.title("Pomoc w układaniu grafiku")
st.subheader("Prześlij plik excel z dostępnościami i jednym kliknięciem stwórz grafik!")
st.write("Aplikacja pomoże Ci w szybki sposób stworzyć grafik na dany miesiąc. Wystarczy, że prześlesz plik excel z dostępnościami, ograniczeniami i preferencjami, a my zajmiemy się resztą!")

# --- Upload Excel file ---
uploaded_file = st.file_uploader("Prześlij odpowiedni plik excel", type=["xlsx"])

# --- Select Year and Month ---
current_year = datetime.datetime.now().year
year = st.selectbox("Wybierz rok", [current_year, current_year + 1])
month = st.selectbox("Wybierz miesiąc", list(range(1, 13)))

# --- Generate Schedule Button ---
if uploaded_file:
    if st.button("Utwórz grafik"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            temp_input_path = tmp.name
            temp_output_path = temp_input_path.replace(".xlsx", "_output.xlsx")
            # Save uploaded file temporarily
            tmp.write(uploaded_file.read())
        
        try:
            # Run your existing scheduling function
            generate_schedule(temp_input_path, temp_output_path, year, month)

            # Allow download of the generated schedule
            with open(temp_output_path, "rb") as file:
                st.success("Grafik został wygenerowany pomyślnie!")
                st.download_button(
                    label="Pobierz grafik",
                    data=file,
                    file_name="grafik.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Napotkano błąd podczas tworzenia grafiku: {str(e)}")
        finally:
            os.remove(temp_input_path)
            if os.path.exists(temp_output_path):
                os.remove(temp_output_path)
else:
    st.info("Prześlij plik excel, aby rozpocząć!")

