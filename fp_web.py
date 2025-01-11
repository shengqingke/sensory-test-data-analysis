# the version is developed based streamlit

import streamlit as st

# st.write("Hello world")

st.header("st.button")

if st.button("Say Hello"):
    st.write("why hello")
else:
    st.write("Goodbye")