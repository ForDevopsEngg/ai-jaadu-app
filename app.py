import streamlit as st
import pandas as pd
import openpyxl
from pathlib import Path
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="AI Jaadu")

# ---------------- STYLE ----------------
st.markdown("""
<style>

@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700;800&family=Playfair+Display:wght@700&display=swap');

:root {
 --bg-a:#f7f4ef;
 --bg-b:#e7eef8;
 --accent-a:#0ea5e9;
 --accent-b:#f97316;
 --card:#0b1328;
 --card-border:#1f2a44;
 --card-text:#f8fafc;
}

[data-testid="stAppViewContainer"] {
 background:
  radial-gradient(circle at 12% 18%, rgba(249,115,22,0.18), transparent 34%),
  radial-gradient(circle at 82% 22%, rgba(14,165,233,0.20), transparent 36%),
  linear-gradient(140deg, var(--bg-a), var(--bg-b));
}

/* Animation */
@keyframes zoomPulse {
    0% { transform: scale(0.8); opacity: 0; }
    60% { transform: scale(1.08); opacity: 1; }
    100% { transform: scale(1); }
}

@keyframes glowDrift {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}

@keyframes floatSoft {
    0% { transform: translateY(0px); }
    50% { transform: translateY(-4px); }
    100% { transform: translateY(0px); }
}

@keyframes navSheen {
    0% { background-position: 0% 50%; }
    100% { background-position: 100% 50%; }
}

.hero {
 text-align:center;
 padding:46px 20px 30px;
 animation: zoomPulse 1s ease;
}

.hero h1 {
 font-family:'Playfair Display', serif;
 letter-spacing:0.5px;
 font-size:64px;
 margin:0;
 background:linear-gradient(110deg,var(--accent-a),#22d3ee,var(--accent-b));
 background-size:220% 220%;
 animation: glowDrift 8s ease infinite;
 -webkit-background-clip:text;
 -webkit-text-fill-color:transparent;
}

.card {
 font-family:'Manrope', sans-serif;
 background:linear-gradient(150deg,var(--card),#131d36);
 border:1px solid var(--card-border);
 box-shadow:0 14px 30px rgba(15,23,42,0.24);
 padding:20px;
 border-radius:16px;
 margin-bottom:14px;
 color:var(--card-text);
 transition:transform 0.25s ease, box-shadow 0.25s ease;
}
.card h2 {
 color:#ffffff;
 font-size:2rem;
 margin:0 0 6px;
}

.card p {
 color:#cbd5e1;
 margin:0;
 font-weight:600;
}

.card:hover {
 transform:translateY(-3px) scale(1.02);
 box-shadow:0 20px 38px rgba(15,23,42,0.30);
}

.insight-card {
 font-family:'Manrope', sans-serif;
 background:rgba(255,255,255,0.84);
 border:1px solid rgba(30,41,59,0.12);
 border-radius:16px;
 padding:16px;
 box-shadow:0 8px 24px rgba(15,23,42,0.10);
 animation:floatSoft 6s ease-in-out infinite;
 margin-bottom:14px;
}

.insight-title {
 margin:0 0 8px;
 color:#0f172a;
 font-size:1.02rem;
 font-weight:800;
}

.insight-value {
 margin:0;
 color:#1e293b;
 font-size:1.15rem;
 font-weight:700;
}

.hero-tagline {
 text-align:center;
 margin-top:-8px;
 color:#334155;
 font-weight:600;
}

.page-banner {
 font-family:'Manrope', sans-serif;
 border:1px solid rgba(30,41,59,0.12);
 background:linear-gradient(120deg, rgba(255,255,255,0.92), rgba(255,255,255,0.72));
 border-radius:14px;
 padding:14px 16px;
 margin:4px 0 16px;
 box-shadow:0 8px 22px rgba(15,23,42,0.08);
}

.page-title {
 font-size:1.25rem;
 color:#0f172a;
 font-weight:800;
 margin:0;
}

.page-subtitle {
 margin-top:4px;
 color:#475569;
 font-weight:600;
 font-size:0.95rem;
}

/* Top navigation: live-touch hover animation */
div[data-testid="stHorizontalBlock"]:first-of-type div[data-testid="stButton"] > button {
 border-radius:14px;
 border:1px solid rgba(51,65,85,0.22);
 background:linear-gradient(120deg, rgba(255,255,255,0.96) 0%, rgba(248,250,252,0.92) 42%, rgba(219,234,254,0.96) 50%, rgba(248,250,252,0.92) 58%, rgba(255,255,255,0.96) 100%);
 background-size:220% 220%;
 box-shadow:0 8px 18px rgba(15,23,42,0.10);
 font-weight:700;
 transition:transform 0.18s ease, box-shadow 0.22s ease, border-color 0.22s ease, background 0.22s ease;
}

div[data-testid="stHorizontalBlock"]:first-of-type div[data-testid="stButton"] > button:hover {
 transform:translateY(-2px) scale(1.02);
 border-color:rgba(14,165,233,0.65);
 background:linear-gradient(120deg, rgba(240,249,255,0.98) 0%, rgba(224,242,254,0.95) 42%, rgba(186,230,253,0.98) 50%, rgba(224,242,254,0.95) 58%, rgba(240,249,255,0.98) 100%);
 background-size:230% 230%;
 animation:navSheen 0.9s linear 1;
 box-shadow:0 14px 28px rgba(14,165,233,0.20);
}

div[data-testid="stHorizontalBlock"]:first-of-type div[data-testid="stButton"] > button:active {
 transform:translateY(0px) scale(0.98);
 box-shadow:0 6px 14px rgba(15,23,42,0.15);
}

@media (max-width: 768px) {
 .hero h1 {
  font-size:44px;
 }
 .hero {
  padding:34px 12px 22px;
 }
}

</style>
""", unsafe_allow_html=True)

# ---------------- CLEAN ----------------
def clean(v):
    return "" if pd.isna(v) else str(v)


def push_notice(msg, level="success"):
    st.session_state["_notice"] = {"msg": msg, "level": level}


def show_notice():
    notice = st.session_state.pop("_notice", None)
    if not notice:
        return

    icon = "✅" if notice["level"] == "success" else "⚠️"
    if hasattr(st, "toast"):
        st.toast(notice["msg"], icon=icon)
    elif notice["level"] == "success":
        st.success(notice["msg"])
    else:
        st.warning(notice["msg"])


def page_banner(icon, title, subtitle):
    st.markdown(
        f'<div class="page-banner"><div class="page-title">{icon} {title}</div><div class="page-subtitle">{subtitle}</div></div>',
        unsafe_allow_html=True,
    )

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

if nav[0].button("✨ Home"): st.session_state.page = "Home"
if nav[1].button("🗓 Agenda"): st.session_state.page = "Agenda"
if nav[2].button("🧩 Problems"): st.session_state.page = "Problems"
if nav[3].button("🎬 Trainings"): st.session_state.page = "Trainings"
if nav[4].button("🪪 Licenses"): st.session_state.page = "Licenses"

page = st.session_state.page
show_notice()

# ================= HOME =================
if page == "Home":

    df = data["Problem Statement"].copy()
    training_df = data["Trainings & Videos"].copy()
    agenda_df = data["Agenda"].copy()
    license_df = data["AI Licence & offers"].copy()

    total_problems = len(df)
    solutions_filled = 0
    summary_completion = 0
    top_theme = "No themes available"
    top_reporter = "No reporter data"

    if "Proposed Solution statement" in df.columns and total_problems > 0:
        solutions_filled = df["Proposed Solution statement"].fillna("").astype(str).str.strip().ne("").sum()
        summary_completion = int(round((solutions_filled / total_problems) * 100, 0))

    if "Problem Category / Theme" in df.columns and total_problems > 0:
        top_themes = df["Problem Category / Theme"].fillna("").replace("", "Uncategorized").value_counts().head(3)
        if len(top_themes) > 0:
            top_theme = f"{top_themes.index[0]} ({top_themes.iloc[0]})"

    if "Reporting Person" in df.columns and total_problems > 0:
        top_people = df["Reporting Person"].fillna("").replace("", "Unknown").value_counts().head(3)
        if len(top_people) > 0:
            top_reporter = f"{top_people.index[0]} ({top_people.iloc[0]})"

    valid_links = 0
    if "Weblink" in training_df.columns and len(training_df) > 0:
        valid_links = training_df["Weblink"].fillna("").astype(str).str.startswith(("http://", "https://")).sum()

    st.markdown('<div class="hero"><h1>AI Jaadu</h1></div>', unsafe_allow_html=True)
    st.markdown('<p class="hero-tagline">Your AI operations cockpit for agenda, challenges, learning, and licenses.</p>', unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)

    c1.markdown(f'<div class="card"><h2>{total_problems}</h2><p>Problems Logged</p></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="card"><h2>{solutions_filled}</h2><p>Solutions Drafted</p></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="card"><h2>{len(training_df)}</h2><p>Training Resources</p></div>', unsafe_allow_html=True)
    c4.markdown(f'<div class="card"><h2>{len(agenda_df)}</h2><p>Agenda Items</p></div>', unsafe_allow_html=True)

    st.subheader("Summary Snapshot")

    s1, s2, s3, s4 = st.columns(4)

    s1.markdown(
        f'<div class="insight-card"><p class="insight-title">Solution Coverage</p><p class="insight-value">{summary_completion}%</p></div>',
        unsafe_allow_html=True
    )
    s2.markdown(
        f'<div class="insight-card"><p class="insight-title">Top Theme</p><p class="insight-value">{top_theme}</p></div>',
        unsafe_allow_html=True
    )
    s3.markdown(
        f'<div class="insight-card"><p class="insight-title">Most Active Reporter</p><p class="insight-value">{top_reporter}</p></div>',
        unsafe_allow_html=True
    )
    s4.markdown(
        f'<div class="insight-card"><p class="insight-title">Training Links Live</p><p class="insight-value">{valid_links}/{len(training_df)}</p></div>',
        unsafe_allow_html=True
    )

    st.markdown("### Quick Insights")
    st.markdown(f"- {len(license_df)} AI subscriptions tracked in Licenses")
    st.markdown(f"- {total_problems - solutions_filled} problems still need proposed solutions")
    st.markdown("- Use Problems page to complete solution fields and improve coverage")

# ================= AGENDA =================
elif page == "Agenda":

    page_banner("🗓", "Agenda", "Track and refine meeting points before execution.")
    df = data["Agenda"].reset_index(drop=True)

    delete_index = None
    pending_key = "pending_delete_agenda"
    new_rows = []

    for i in range(len(df)):
        col1, col2 = st.columns([9,1])

        val = col1.text_input("", clean(df.loc[i,"Agenda Points"]), key=f"A{i}")

        if col2.button("🗑 Delete", key=f"delA{i}"):
            st.session_state[pending_key] = i

        new_rows.append({
            "Sr. No.": i+1,
            "Agenda Points": val
        })

    pending_idx = st.session_state.get(pending_key)
    if pending_idx is not None:
        st.warning(f"Delete Agenda item #{pending_idx + 1}? This cannot be undone.")
        conf_col, cancel_col = st.columns(2)
        if conf_col.button("Confirm Delete", key="confirm_delA"):
            delete_index = pending_idx
            st.session_state[pending_key] = None
        if cancel_col.button("Cancel", key="cancel_delA"):
            st.session_state[pending_key] = None
            st.rerun()

    # DELETE instantly
    if delete_index is not None:
        df = df.drop(delete_index).reset_index(drop=True)
        st.session_state.data["Agenda"] = df
        push_notice("Agenda item deleted")
        st.rerun()

    # UPDATE
    if st.button("✅ Update Agenda"):
        st.session_state.data["Agenda"] = pd.DataFrame(new_rows)
        push_notice("Agenda updated")
        st.rerun()

    # ADD
    if st.button("➕ Add Agenda"):
        df.loc[len(df)] = [len(df)+1, "New Item"]
        st.session_state.data["Agenda"] = df
        push_notice("New agenda item added")
        st.rerun()

# ================= PROBLEMS =================
elif page == "Problems":

    page_banner("🧩", "Problem Statements", "Capture challenges, ownership, and solution direction.")
    df = data["Problem Statement"].reset_index(drop=True)

    delete_index = None
    pending_key = "pending_delete_problem"
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
        if col_right.button("🗑 Delete", key=f"del{i}"):
            st.session_state[pending_key] = i

    pending_idx = st.session_state.get(pending_key)
    if pending_idx is not None:
        st.warning(f"Delete Problem #{pending_idx + 1}? This cannot be undone.")
        conf_col, cancel_col = st.columns(2)
        if conf_col.button("Confirm Delete", key="confirm_delP"):
            delete_index = pending_idx
            st.session_state[pending_key] = None
        if cancel_col.button("Cancel", key="cancel_delP"):
            st.session_state[pending_key] = None
            st.rerun()

    if delete_index is not None:
        df = df.drop(delete_index).reset_index(drop=True)
        st.session_state.data["Problem Statement"] = df
        push_notice("Problem statement deleted")
        st.rerun()

    if st.button("✅ Update Problems"):
        st.session_state.data["Problem Statement"] = pd.DataFrame(new_rows)
        push_notice("Problem statements updated")
        st.rerun()

    if st.button("➕ Add Problem"):
        df.loc[len(df)] = ["","","","","",""]
        st.session_state.data["Problem Statement"] = df
        push_notice("New problem row added")
        st.rerun()

# ================= TRAININGS =================
elif page == "Trainings":

    page_banner("🎬", "Trainings", "Curate learning resources and verify active links.")
    df = data["Trainings & Videos"].reset_index(drop=True)

    delete_index = None
    pending_key = "pending_delete_training"
    new_rows = []

    for i in range(len(df)):
        st.write(f"**Training #{i+1}**")
        col_t, col_l, col_open, col_del = st.columns([5,5,2,2])

        t = col_t.text_input("Training", clean(df.loc[i,"Training"]), key=f"T{i}")
        l = col_l.text_input("Link", clean(df.loc[i,"Weblink"]), key=f"L{i}")

        if l.startswith("http"):
            col_open.link_button("Open Link", l, use_container_width=True)
        else:
            col_open.caption("No valid URL")

        if col_del.button("🗑 Delete", key=f"delT{i}", use_container_width=True):
            st.session_state[pending_key] = i

        new_rows.append({"Training":t,"Weblink":l})
        st.write("---")

    pending_idx = st.session_state.get(pending_key)
    if pending_idx is not None:
        label = clean(df.loc[pending_idx, "Training"]) if pending_idx < len(df) else "this training"
        st.warning(f"Delete training '{label}'? This cannot be undone.")
        conf_col, cancel_col = st.columns(2)
        if conf_col.button("Confirm Delete", key="confirm_delT"):
            delete_index = pending_idx
            st.session_state[pending_key] = None
        if cancel_col.button("Cancel", key="cancel_delT"):
            st.session_state[pending_key] = None
            st.rerun()

    if delete_index is not None:
        df = df.drop(delete_index).reset_index(drop=True)
        st.session_state.data["Trainings & Videos"] = df
        push_notice("Training deleted")
        st.rerun()

    if st.button("✅ Update Trainings"):
        st.session_state.data["Trainings & Videos"] = pd.DataFrame(new_rows)
        push_notice("Trainings updated")
        st.rerun()

    if st.button("➕ Add Training"):
        df.loc[len(df)] = ["","","","","","",""]
        st.session_state.data["Trainings & Videos"] = df
        push_notice("New training row added")
        st.rerun()

# ================= LICENSE =================
elif page == "Licenses":

    page_banner("🪪", "Licenses", "Manage subscriptions by core area and usage.")
    df = data["AI Licence & offers"].reset_index(drop=True)

    delete_index = None
    pending_key = "pending_delete_license"
    new_rows = []

    for i in range(len(df)):

        s = st.text_input("Subscription", clean(df.loc[i,"AI Subscriptions"]), key=f"S{i}")
        c = st.text_input("Core", clean(df.loc[i,"Core area"]), key=f"C{i}")

        if st.button("🗑 Delete", key=f"delL{i}"):
            st.session_state[pending_key] = i

        new_rows.append({"AI Subscriptions":s,"Core area":c})

    pending_idx = st.session_state.get(pending_key)
    if pending_idx is not None:
        st.warning(f"Delete License #{pending_idx + 1}? This cannot be undone.")
        conf_col, cancel_col = st.columns(2)
        if conf_col.button("Confirm Delete", key="confirm_delL"):
            delete_index = pending_idx
            st.session_state[pending_key] = None
        if cancel_col.button("Cancel", key="cancel_delL"):
            st.session_state[pending_key] = None
            st.rerun()

    if delete_index is not None:
        df = df.drop(delete_index).reset_index(drop=True)
        st.session_state.data["AI Licence & offers"] = df
        push_notice("License row deleted")
        st.rerun()

    if st.button("✅ Update Licenses"):
        st.session_state.data["AI Licence & offers"] = pd.DataFrame(new_rows)
        push_notice("Licenses updated")
        st.rerun()

    if st.button("➕ Add License"):
        df.loc[len(df)] = ["","",""]
        st.session_state.data["AI Licence & offers"] = df
        push_notice("New license row added")
        st.rerun()

# ================= SAVE =================
st.write("---")

if st.button("💾 Save Changes"):

    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")

    for name, df in st.session_state.data.items():
        df.to_excel(writer, sheet_name=name, index=False)

    writer.close()
    filename = f"Updated_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    push_notice("Workbook is ready to download")
    st.download_button("Download Excel", output.getvalue(), filename)