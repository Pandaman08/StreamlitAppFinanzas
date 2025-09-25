import streamlit as st

st.title("Bienvenido a mi primera App en Streamlit 🚀")
st.write("Esta es la página de inicio.")

menu = ["Inicio", "Otra Página"]
choice = st.sidebar.selectbox("Navegación", menu)

if choice == "Inicio":
    st.subheader("Página Principal")
    st.write("Aquí puedes mostrar gráficos, tablas, etc.")
elif choice == "Otra Página":
    st.subheader("Aquí va otra página")
    st.write("Puedes dividir tu app en múltiples páginas usando la carpeta `pages`.")
