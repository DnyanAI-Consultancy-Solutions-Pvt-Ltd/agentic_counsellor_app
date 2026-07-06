import os
import requests
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

load_dotenv()
st.set_page_config(page_title="MHT-CET & JEE AI Counselor Client", page_icon="🎓", layout="wide")

FASTAPI_URL = "http://127.0.0.1:8000"

st.title("🎓 MHT-CET & JEE Engineering CAP Round AI Counselor")
st.markdown("Interact naturally with the AI Agent or use the manual sidebar controllers to extract high-end option preferences.")

if "chat_history" not in st.session_state: st.session_state["chat_history"] = []
if "last_generated_file" not in st.session_state: st.session_state["last_generated_file"] = None
if "display_preview" not in st.session_state: st.session_state["display_preview"] = None

# Sidebar Configuration Control Panel
with st.sidebar:
    st.header("📋 Candidate Profile Controls")
    ui_percentile = st.number_input("MHT-CET Percentile Score", min_value=0.0, max_value=100.0, value=94.25, step=0.01)
    ui_jee_percentile = st.number_input("JEE Percentile Score", min_value=0.0, max_value=100.0, value=92.10, step=0.01)
    ui_category = st.selectbox("Caste Reservation Pool", ["OPEN", "OBC", "SC", "ST", "EWS", "TFWS"])
    ui_stream = st.text_input("Target Branches (Keywords)", value="CS or IT")
    ui_pwd = st.toggle("Registered PWD Status Profile", value=False)
    ui_university = st.text_input("Preferred University", value="Any")
    ui_city = st.text_input("Preferred City Name Location", value="Any")
    ui_gender = st.radio("Gender Category Pool Allocation", ["General Pool", "Female"])
    
    st.markdown("---")
    if st.button("🚀 Direct Generate Report", use_container_width=True):
        with st.spinner("Analyzing choices against historical cutoffs..."):
            chat_payload = {
                "prompt": f"Generate a completely optimized selection table context containing exactly 30 preferences based on category distribution parameters.",
                "ui_percentile": ui_percentile, 
                "ui_jee_percentile": ui_jee_percentile, 
                "ui_category": ui_category,
                "ui_stream": ui_stream, 
                "ui_pwd": ui_pwd, 
                "ui_city": ui_city, 
                "ui_gender": ui_gender, 
                "ui_university": ui_university
            }
            try:
                res_api = requests.post(f"{FASTAPI_URL}/api/v1/chat", json=chat_payload)
                if res_api.status_code == 200:
                    res = res_api.json()
                    if res.get("filepath"):
                        st.session_state["last_generated_file"] = res["filepath"]
                        st.session_state["display_preview"] = pd.DataFrame(res["preview_data"])
                        st.success("Preference list generated flawlessly!")
                    else:
                        st.error("Engine compiled but no data generated. Check if the PDF has data matching your reserves selection.")
                else:
                    st.error("Error processing request. Ensure the backend engine is running successfully.")
            except requests.exceptions.ConnectionError:
                st.error("FastAPI backend server is not running on port 8000.")

# --- Preview Dashboard Section ---
st.subheader("📊 Output Preview Dashboard")
if st.session_state["last_generated_file"] and os.path.exists(st.session_state["last_generated_file"]):
    with open(st.session_state["last_generated_file"], "rb") as file:
        st.download_button(label="📥 Download Preference Sheet (Excel)", data=file, file_name=os.path.basename(st.session_state["last_generated_file"]), type="primary", use_container_width=True)
    if st.session_state["display_preview"] is not None:
        st.dataframe(st.session_state["display_preview"], use_container_width=True, hide_index=True, height=400)
else:
    st.info("No active preferences portfolio loaded yet. Run a prompt below or click generate above.")

# --- AI Chatbot Interface Section ---
st.markdown("---")
st.subheader("💬 AI Admission Chatbot")
chat_container = st.container(height=300)

with chat_container:
    for msg in st.session_state["chat_history"]:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])
            
if prompt := st.chat_input("Ask about cutoffs or assignments (e.g., 'Re-optimize my sheet with more dream options'):"):
    st.session_state["chat_history"].append({"role": "user", "content": prompt})
    with chat_container:
        with st.chat_message("user"): st.markdown(prompt)
            
    with chat_container:
        with st.chat_message("assistant"):
            with st.spinner("Analyzing data..."):
                chat_payload = {
                    "prompt": prompt, "ui_percentile": ui_percentile, "ui_jee_percentile": ui_jee_percentile, 
                    "ui_category": ui_category, "ui_stream": ui_stream, "ui_pwd": ui_pwd, 
                    "ui_city": ui_city, "ui_gender": ui_gender, "ui_university": ui_university
                }
                try:
                    res_api = requests.post(f"{FASTAPI_URL}/api/v1/chat", json=chat_payload)
                    if res_api.status_code == 200:
                        res = res_api.json()
                        if res.get("filepath"):
                            st.session_state["last_generated_file"] = res["filepath"]
                            st.session_state["display_preview"] = pd.DataFrame(res["preview_data"])
                            
                        st.markdown(res["reply"])
                        st.session_state["chat_history"].append({"role": "assistant", "content": res["reply"]})
                        st.rerun()
                except Exception as e:
                    st.error(f"Error handling request session: {str(e)}")