import streamlit as st
import pandas as pd
import openpyxl
from pathlib import Path
from io import BytesIO

st.set_page_config(layout="wide", page_title="AI Jaadu")

# ---------------- STYLE ----------------
st.markdown("""
<style>

/* Animation */
@keyframes zoomPulse {
    0% { transform: scale(0.8); opacity: 0; }
    60% { transform: scale(1.08); opacity: 1; }
    100% { transform: scale(1); }
}

.hero {
 text-align:center;
 padding:40px;
 animation: zoomPulse 1s ease;
}

.hero h1 {
 font-size:60px;
 background:linear-gradient(90deg,#6366f1,#ec4899);
 -webkit-background-clip:text;
 -webkit-text-fill-color:transparent;
}

.card {
 background:#0f172a;
 padding:20px;
 border-radius:14px;
 margin-bottom:12px;
 transition:0.3s;
}
.card:hover { transform:scale(1.05); }

</style>
""", unsafe_allow_html=True)

# ---------------- CLEAN ----------------
def clean(v):
    return "" if pd.isna(v) else str(v)

# ---------------- LOAD ----------------
FILE = Path("AI Jaadu.xlsx")

@st.cache_data
def load():
    wb = openpyxl.load_workbook(FILE)
    data = {}
    for ws in wb.worksheets:
        rows = list(ws.values)
        df = pd.DataFrame(rows[1:], columns=rows[0])
        df = df.dropna(how="all").fillna("")
        data[ws.title] = df
    return data

if "data" not in st.session_state:
    st.session_state.data = load()

data = st.session_state.data

# ---------------- NAV ----------------
if "page" not in st.session_state:
    st.session_state.page = "Home"

nav = st.columns(5)

if nav[0].button("🏠 Home"): st.session_state.page = "Home"
if nav[1].button("📋 Agenda"): st.session_state.page = "Agenda"
if nav[2].button("🧠 Problems"): st.session_state.page = "Problems"
if nav[3].button("🎓 Trainings"): st.session_state.page = "Trainings"
if nav[4].button("📦 Licenses"): st.session_state.page = "Licenses"

page = st.session_state.page

# ================= HOME =================
if page == "Home":

    df = data["Problem Statement"]

    st.markdown('<div class="hero"><h1>AI Jaadu</h1></div>', unsafe_allow_html=True)

    c1,c2,c3 = st.columns(3)

    c1.markdown(f'<div class="card"><h2>{len(df)}</h2><p>Problems</p></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="card"><h2>{len(data["Trainings & Videos"])}</h2><p>Trainings</p></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="card"><h2>{len(data["Agenda"])}</h2><p>Agenda</p></div>', unsafe_allow_html=True)

    st.subheader("Summary")

    if "Problem Category / Theme" in df.columns:
        top = df["Problem Category / Theme"].value_counts().head(3)
        for k,v in top.items():
            st.markdown(f'<div class="card">{k} → {v} Problems</div>', unsafe_allow_html=True)

# ================= AGENDA =================
elif page == "Agenda":

    st.title("Agenda")
    df = data["Agenda"].reset_index(drop=True)

    delete_index = None
    new_rows = []

    for i in range(len(df)):
        col1, col2 = st.columns([9,1])

        val = col1.text_input("", clean(df.loc[i,"Agenda Points"]), key=f"A{i}")

        if col2.button("❌", key=f"delA{i}"):
            delete_index = i

        new_rows.append({
            "Sr. No.": i+1,
            "Agenda Points": val
        })

    # DELETE instantly
    if delete_index is not None:
        df = df.drop(delete_index).reset_index(drop=True)
        st.session_state.data["Agenda"] = df
        st.rerun()

    # UPDATE
    if st.button("✅ Update Agenda"):
        st.session_state.data["Agenda"] = pd.DataFrame(new_rows)

    # ADD
    if st.button("➕ Add Agenda"):
        df.loc[len(df)] = [len(df)+1, "New Item"]
        st.session_state.data["Agenda"] = df
        st.rerun()

# ================= PROBLEMS =================
elif page == "Problems":

    st.title("Problem Statements")
    df = data["Problem Statement"].reset_index(drop=True)

    delete_index = None
    new_rows = []

    for i in range(len(df)):

        col_left, col_right = st.columns([10,1])

        with col_left:
            with st.expander(f"{i+1}. {clean(df.loc[i,'Problem Statement'])[:60]}"):

                row = {}

                row["Sr. No."] = i+1

                row["Problem Category / Theme"] = st.text_input(
                    "Category", clean(df.loc[i,"Problem Category / Theme"]), key=f"cat{i}"
                )

                row["Problem Statement"] = st.text_area(
                    "Problem", clean(df.loc[i,"Problem Statement"]), key=f"prob{i}"
                )

                row["Proposed Solution statement"] = st.text_area(
                    "Solution", clean(df.loc[i,"Proposed Solution statement"]), key=f"sol{i}"
                )

                row["Solution category"] = st.text_input(
                    "Solution Category", clean(df.loc[i,"Solution category"]), key=f"sc{i}"
                )

                row["Reporting Person"] = st.text_input(
                    "Person", clean(df.loc[i,"Reporting Person"]), key=f"p{i}"
                )

                new_rows.append(row)

        # ✅ DELETE RIGHT SIDE
        if col_right.button("❌", key=f"del{i}"):
            delete_index = i

    if delete_index is not None:
        df = df.drop(delete_index).reset_index(drop=True)
        st.session_state.data["Problem Statement"] = df
        st.rerun()

    if st.button("✅ Update Problems"):
        st.session_state.data["Problem Statement"] = pd.DataFrame(new_rows)

    if st.button("➕ Add Problem"):
        df.loc[len(df)] = ["","","","","",""]
        st.session_state.data["Problem Statement"] = df
        st.rerun()

# ================= TRAININGS =================
elif page == "Trainings":

    st.title("Trainings")
    df = data["Trainings & Videos"].reset_index(drop=True)

    delete_index = None
    new_rows = []

    for i in range(len(df)):

        t = st.text_input("Training", clean(df.loc[i,"Training"]), key=f"T{i}")
        l = st.text_input("Link", clean(df.loc[i,"Weblink"]), key=f"L{i}")

        if l.startswith("http"):
            st.markdown(f"{l}")

        if st.button("❌", key=f"delT{i}"):
            delete_index = i

        new_rows.append({"Training":t,"Weblink":l})

    if delete_index is not None:
        df = df.drop(delete_index).reset_index(drop=True)
        st.session_state.data["Trainings & Videos"] = df
        st.rerun()

    if st.button("✅ Update Trainings"):
        st.session_state.data["Trainings & Videos"] = pd.DataFrame(new_rows)

    if st.button("➕ Add Training"):
        df.loc[len(df)] = ["","","","","","",""]
        st.session_state.data["Trainings & Videos"] = df
        st.rerun()

# ================= LICENSE =================
elif page == "Licenses":

    st.title("Licenses")
    df = data["AI Licence & offers"].reset_index(drop=True)

    delete_index = None
    new_rows = []

    for i in range(len(df)):

        s = st.text_input("Subscription", clean(df.loc[i,"AI Subscriptions"]), key=f"S{i}")
        c = st.text_input("Core", clean(df.loc[i,"Core area"]), key=f"C{i}")

        if st.button("❌", key=f"delL{i}"):
            delete_index = i

        new_rows.append({"AI Subscriptions":s,"Core area":c})

    if delete_index is not None:
        df = df.drop(delete_index).reset_index(drop=True)
        st.session_state.data["AI Licence & offers"] = df
        st.rerun()

    if st.button("✅ Update Licenses"):
        st.session_state.data["AI Licence & offers"] = pd.DataFrame(new_rows)

    if st.button("➕ Add License"):
        df.loc[len(df)] = ["","",""]
        st.session_state.data["AI Licence & offers"] = df
        st.rerun()

# ================= SAVE =================
st.write("---")

if st.button("💾 Save Changes"):

    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")

    for name, df in st.session_state.data.items():
        df.to_excel(writer, sheet_name=name, index=False)

    writer.close()

    st.download_button("Download Excel", output.getvalue(), "Updated.xlsx")