import streamlit as st
import pandas as pd
import gspread
import google.generativeai as genai
from google.oauth2.service_account import Credentials
from docx import Document
from datetime import datetime, timedelta
import json
import io
import zipfile 

# 1. SETUP & PAGE CONFIG
st.set_page_config(page_title="Coaching Generator", page_icon="🏆", layout="wide")

# --- 🧹 JANITOR FUNCTION ---
def clean_text(text):
    """Handles lists, removes brackets, quotes, and markdown asterisks."""
    if isinstance(text, list):
        text = "\n".join([str(item) for item in text])
    text = str(text)
    for junk in ["**", "['", "']", '["', '"]']:
        text = text.replace(junk, "")
    return text.strip()

# --- 🔒 PASSWORD PROTECTION ---
if "APP_PASSWORD" in st.secrets:
    password = st.sidebar.text_input("Enter Password", type="password")
    if password != st.secrets["APP_PASSWORD"]:
        st.warning("Please enter the correct password to access the tools.")
        st.stop()
else:
    st.error("⚠️ Password not set in Secrets!")
    st.stop()

st.title("🏆 Coaching Plan Generator")

# --- 📋 TEAM DATA ---
# Try to load from secrets, otherwise default to an empty list
if "team" in st.secrets and "agents" in st.secrets["team"]:
    MY_TEAM = sorted(st.secrets["team"]["agents"])
else:
    st.error("⚠️ Team list not found in Secrets!")
    MY_TEAM = []

# Initialize Session States
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = []
if 'batch_complete' not in st.session_state:
    st.session_state.batch_complete = False

# --- 🔄 SELECT ALL CALLBACK ---
def toggle_all():
    """Forces all team checkboxes to match the 'Select All' state."""
    st.session_state.batch_complete = False
    for agent in MY_TEAM:
        st.session_state[f"check_{agent}"] = st.session_state.select_all_team

# 2. AUTHENTICATION
try:
    creds_info = st.secrets["GAA_JSON"]
    creds = Credentials.from_service_account_info(creds_info, scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ])
    gc = gspread.authorize(creds)
    genai.configure(api_key=st.secrets["GEMINI_KEY"])
except Exception as e:
    st.error(f"Authentication Error: {e}")
    st.stop()

# 3. 🧠 DYNAMIC MODEL FINDER
@st.cache_resource
def get_valid_gemini_model():
    try:
        all_models = [m for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        gemini_models = [m.name for m in all_models if 'gemini' in m.name]
        if not gemini_models: return 'models/gemini-1.5-flash'
        flash_model = next((m for m in gemini_models if 'flash' in m), None)
        return flash_model if flash_model else gemini_models[0]
    except:
        return 'models/gemini-1.5-flash'

active_model_name = get_valid_gemini_model()
model = genai.GenerativeModel(active_model_name)
st.sidebar.success(f"🤖 Connected")
#st.sidebar.success(f"🤖 Connected to: {active_model_name}")

# 4. CONFIGURATION SIDEBAR
with st.sidebar:
    st.header("Settings")
    sheet_url = st.text_input("Google Sheet URL", value="", placeholder="Paste Link Here...")
    template_path = "template.docx" 
    sidebar_date = st.date_input("Global Coaching Date", datetime.today())
    limit = st.number_input("Lookback Rows", value=5, min_value=1)

# 5. CORE LOGIC
if sheet_url:
    try:
        sh = gc.open_by_url(sheet_url)
        worksheet = sh.get_worksheet(0)
        data = worksheet.get_all_values()
        headers = data.pop(0)
        df = pd.DataFrame(data, columns=headers)
        
        if "ES Last Name, First Name" in df.columns:
            st.subheader("👥 Select Agents & Dates")
            view_option = st.radio("Selection View:", ["My Team Only", "Search All Agents from Sheet"], horizontal=True)
            
            selected_configs = [] 

            if view_option == "My Team Only":
                # Select All button
                st.checkbox("Select All My Team", key="select_all_team", on_change=toggle_all)
                st.divider()
                
                for agent in MY_TEAM:
                    if agent in df["ES Last Name, First Name"].values:
                        col_check, col_name, col_date = st.columns([0.5, 3, 2])
                        with col_check:
                            is_selected = st.checkbox("", key=f"check_{agent}")
                        with col_name:
                            st.markdown(f"**{agent}**")
                        with col_date:
                            agent_date = st.date_input(f"Date for {agent}", value=sidebar_date, key=f"date_{agent}", label_visibility="collapsed")
                        
                        if is_selected:
                            selected_configs.append((agent, agent_date))
            else:
                all_names = sorted(df["ES Last Name, First Name"].unique())
                search_selection = st.multiselect("Search for agents:", all_names)
                if search_selection:
                    st.divider()
                    for agent in search_selection:
                        col_check, col_name, col_date = st.columns([0.5, 3, 2])
                        with col_check:
                            st.checkbox("", value=True, key=f"check_search_{agent}", disabled=True)
                        with col_name:
                            st.markdown(f"**{agent}**")
                        with col_date:
                            agent_date = st.date_input(f"Date for {agent}", value=sidebar_date, key=f"date_search_{agent}", label_visibility="collapsed")
                        selected_configs.append((agent, agent_date))

            # THE RUN BUTTON
            if selected_configs:
                st.divider()
                if st.button("⚡ Generate Plans", type="primary", use_container_width=True):
                    st.session_state.generated_files = []
                    st.session_state.batch_complete = False
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    for i, (agent, final_date) in enumerate(selected_configs):
                        status_text.text(f"⏳ Processing {agent}...")
                        
                        # A. GET DATA
                        agent_rows = df[df["ES Last Name, First Name"] == agent].tail(limit)
                        missed = [str(val) for col in df.columns if "Skill Performance Area Missed" in col for val in agent_rows[col] if val]
                        strengths = [str(val) for col in df.columns if "Strength" in col for val in agent_rows[col] if val]
                        
                        # B. AI GENERATION (RESTORED PROMPTS EXACTLY)
                        prompt = f"""
                        You are Khori, an expert QA Coach.
                        INPUT DATA: 
                        "MISSED OPPORTUNITIES": { ' || '.join(missed) }
                        "STRENGTHS": { ' || '.join(strengths) }
                        
                        TASK: Output JSON following these strict rules:
                        1. **Issues:** Identify 3 DISTINCT and DIFFERENT critical issues from MISSED OPPORTUNITIES.
                        2. **Trend Identification:** Analyze the data to find a pattern. Do not just list random errors; identify the specific "Stimulus" (trigger) that causes the agent's performance to break down. Find the common thread for each. Do not repeat the same behavior or trend for multiple issues.
                        3. **Behavior Syntax (CRITICAL):** For the "issue" fields, you MUST strictly follow this format: "Whenever <STIMULI>, <SYMPTOM> by <ACTION>". 
                            - <STIMULI>: The situation/trigger (e.g., "the customer is in a hurry").
                            - <SYMPTOM>: The high-level failure (e.g., "lack of urgency").
                            - <ACTION>: The observable behavior (e.g., "ignoring cues and reading the full script slowly").
                        4. **Sort:** Sort by severity: AF > Total Resolution > Professionalism > Sincerity.
                        5. **Primary Focus:** MUST be the most critical issue (Issue 1) based on the identified trend, create a high-level summary title (Category) that describes the main area of improvement (e.g., "Resolution Accuracy," "Engagement & Tone," or "Process Efficiency"). Do NOT just copy Issue 1.
                        6. **Tone:** SIMPLE, DIRECT, CONVERSATIONAL. No big words.
                        7. **Quick Fixes:** Provide a short, simple corrective sentence for each issue.
                        8. **Habits:** Map "Missed" to Reference List (Essential Habit). Map "Strength" to Reference List (Essential Habit Performed).
                        9. **Constraints:** - 'action_plan' must be under 245 characters.
                            - 'impact_question' must be under 245 characters.
                            - Do NOT use markdown asterisks (**) or bracket in any of the output text. Keep it clean.
                        
                        OUTPUT JSON KEYS:
                        {{
                          "primary_focus": "Same as Issue 1 name. A high-level category title summarizing the main trend.",
                          "why_matters": "Importance of fixing this trend.",
                          "action_plan": "SMART plan (Max 245 chars).",
                          "impact_question": "A question to help the agent self-reflect, followed by your insight on how improving this behavior will have a positive impact on their KPIs and the customer experience. (Max 245 chars).",
                          "essential_habit": "From Reference List (Matches Issue 1).",
                          "essential_habit_performed": "From Reference List (Matches Strength).",
                          "likely_root_cause": "The underlying skill or will gap.",
                          "root_cause_questions": "3 questions to ask the agent.",
                          "final_thoughts": "Closing encouragement from coach to agent.",
                          "issue_1": "Whenever <STIMULI>, <SYMPTOM> by <ACTION>", 
                          "comment_1": "A professional coach's insight analyzing the behavior. Do NOT repeat the problem; provide unique insight into why this behavior is detrimental to the customer experience.", 
                          "fix_1": "Simple fix.",
                          "issue_2": "Whenever <STIMULI>, <SYMPTOM> by <ACTION>", 
                          "comment_2": "A professional coach's insight analyzing the behavior. Do NOT repeat the problem; provide unique insight into why this behavior is detrimental to the customer experience.", 
                          "fix_2": "Simple fix.",
                          "issue_3": "Whenever <STIMULI>, <SYMPTOM> by <ACTION>", 
                          "comment_3": "A professional coach's insight analyzing the behavior. Do NOT repeat the problem; provide unique insight into why this behavior is detrimental to the customer experience.", 
                          "fix_3": "Simple fix."
                        }}

                        REFERENCE LIST (Use EXACTLY):
                        - Establish Credibility - Listen to the needs
                        - Establish Credibility - Demonstrate common courtesy
                        - Establish Credibility - Choose Language to optimize compression
                        - Ask Insightful Questions - Ask Insightful Questions
                        - Ask Insightful Questions - Informative and Persuasive Language
                        - Ask Insightful Questions - Verbal Matching
                        - Make Things Easy - Vocal Delivery
                        - Make Things Easy - Manage discussions
                        - Make Things Easy - Minimize Future effort
                        - Be Present - Responding Immediately
                        - Be Present - Demonstrate Understanding
                        - Be Present - Provide Personalized responses
                        - Communicate Optimism - Taking responsibility
                        - Communicate Optimism - Framing optimistically
                        - Communicate Optimism - Focusing on what can be done
                        - Build Rapport - Respond to disclosures
                        - Build Rapport - Engage in small talk
                        - Build Rapport - Protect and promote self-image
                        """
                        
                        try:
                            response = model.generate_content(prompt)
                            clean_json = response.text.replace('```json', '').replace('```', '').strip()
                            ai_data = json.loads(clean_json)
                        except:
                            ai_data = {}

                        # C. CREATE DOC
                        end_date = final_date + timedelta(days=21)
                        f_up = final_date + timedelta(days=7)
                        doc = Document(template_path)
                        
                        replacements = {
                            "{{Agent Name}}": agent,
                            "{{Date}}": final_date.strftime("%m/%d/%Y"),
                            "{{End Date}}": end_date.strftime("%m/%d/%Y"),
                            "{{Follow Up Date}}": f_up.strftime("%m/%d/%Y"),
                            "{{Primary Focus}}": ai_data.get('primary_focus', ''),
                            "{{Why Matters}}": ai_data.get('why_matters', ''),
                            "{{Action Plan}}": ai_data.get('action_plan', ''),
                            "{{Impact}}": ai_data.get('impact_question', ''),
                            "{{Essential Habit}}": ai_data.get('essential_habit', ''),
                            "{{Essential Habit Performed}}": ai_data.get('essential_habit_performed', ''),
                            "{{Issue 1}}": ai_data.get('issue_1', ''), "{{Comment 1}}": ai_data.get('comment_1', ''), "{{Fix 1}}": ai_data.get('fix_1', ''),
                            "{{Issue 2}}": ai_data.get('issue_2', ''), "{{Comment 2}}": ai_data.get('comment_2', ''), "{{Fix 2}}": ai_data.get('fix_2', ''),
                            "{{Issue 3}}": ai_data.get('issue_3', ''), "{{Comment 3}}": ai_data.get('comment_3', ''), "{{Fix 3}}": ai_data.get('fix_3', ''),
                            "{{Root Cause}}": ai_data.get('likely_root_cause', ''), "{{Root Questions}}": ai_data.get('root_cause_questions', ''), "{{Final Thoughts}}": ai_data.get('final_thoughts', '')
                        }

                        for p in doc.paragraphs:
                            for tag, val in replacements.items():
                                if tag in p.text:
                                    p.text = p.text.replace(tag, clean_text(val))
                        
                        # D. SAVE TO MEMORY
                        bio = io.BytesIO()
                        doc.save(bio)
                        bio.seek(0)
                        
                        file_name = f"Coaching Plan - {agent}.docx"
                        st.session_state.generated_files.append({"name": file_name, "data": bio})
                        progress_bar.progress((i + 1) / len(selected_configs))
                    
                    st.session_state.batch_complete = True
                    status_text.text("✅ All Done!")

            # ---------------------------------------------------------
            # ⬇️ DOWNLOAD SECTION
            # ---------------------------------------------------------
            if st.session_state.batch_complete and st.session_state.generated_files:
                st.success("Analysis Complete! Download your files below.")
                st.markdown("---")
                
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for f in st.session_state.generated_files:
                        zf.writestr(f["name"], f["data"].getvalue())
                
                st.download_button(
                    label="📦 Download ALL (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Coaching_Batch_{sidebar_date.strftime('%Y-%m-%d')}.zip",
                    mime="application/zip",
                    type="primary"
                )
                
                st.markdown("---")
                cols = st.columns(2) 
                for idx, f in enumerate(st.session_state.generated_files):
                    with cols[idx % 2]:
                        st.download_button(
                            label=f"📄 {f['name']}",
                            data=f["data"],
                            file_name=f['name'],
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
        
    except Exception as e:
        st.error(f"Error: {e}")
