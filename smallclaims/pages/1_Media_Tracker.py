"""
BHU Media Tracker — public coverage page.

Shows media coverage of homelessness & street vendor issues in Berkeley/Oakland,
tracked automatically by the BHU media scanner
(https://github.com/robbiepowelson-code/bhu-media-tracker — scans news feeds daily).

Deliberately public-safe: reporter names, outlets and activity only.
No email addresses or phone numbers are published on this page.
"""

import json
import urllib.request
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

st.set_page_config(page_title="BHU Media Tracker", page_icon="📰", layout="wide")

RAW = "https://raw.githubusercontent.com/robbiepowelson-code/bhu-media-tracker/main/data/"


@st.cache_data(ttl=1800, show_spinner="Loading coverage data…")
def load(name):
    req = urllib.request.Request(RAW + name, headers={"User-Agent": "BHU-Streamlit/1.0"})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode("utf-8"))


try:
    coverage = load("coverage.json")
    press = load("press_list.json")
except Exception:
    st.title("📰 Media Coverage Tracker")
    st.warning(
        "Coverage data is temporarily unavailable. "
        "The tracker updates every morning — check back soon."
    )
    st.stop()

articles = coverage.get("articles", [])
contacts = press.get("contacts", [])
removed = press.get("removed", [])

st.title("📰 Media Coverage Tracker")
st.caption(
    "Homelessness & street vendor coverage in Berkeley and Oakland, "
    "scanned automatically every morning · Berkeley Homeless Union"
)

# ---------------------------------------------------------------- stat tiles
today = date.today()


def days_ago(s):
    try:
        return (today - datetime.strptime(s, "%Y-%m-%d").date()).days
    except Exception:
        return 9999


news = [a for a in articles if a.get("kind") == "news"]
last30 = [a for a in articles if days_ago(a.get("date", "")) <= 30]
active = sum(1 for c in contacts if c.get("status") == "active")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Stories · last 30 days", len(last30))
c2.metric("Stories tracked", len(articles))
c3.metric("Reporters & outlets tracked", len(contacts))
c4.metric("Verified active reporters", active)

# ---------------------------------------------------------------- weekly chart
weeks = []
for i in range(11, -1, -1):
    end = today - timedelta(days=i * 7)
    start = end - timedelta(days=7)
    n = sum(
        1 for a in articles
        if start < datetime.strptime(a["date"], "%Y-%m-%d").date() <= end
    ) if articles else 0
    weeks.append({"week of": end.strftime("%b %d"), "stories": n})
st.subheader("Stories found per week (last 12 weeks)")
chart_df = pd.DataFrame(weeks).set_index("week of")
st.bar_chart(chart_df, color="#2a78d6", height=200)

# ---------------------------------------------------------------- coverage feed
st.subheader("Recent coverage")
fc1, fc2, fc3 = st.columns([2, 1, 1])
q = fc1.text_input("Search stories", "", placeholder="Search title, outlet, reporter…")
kind = fc2.selectbox("Type", ["All", "News only", "Social only", "Notable only"])
rng = fc3.selectbox("Range", ["Last 30 days", "Last 7 days", "Last 90 days", "All time"], index=0)
rdays = {"Last 7 days": 7, "Last 30 days": 30, "Last 90 days": 90, "All time": 99999}[rng]

shown = 0
for a in articles:
    if days_ago(a.get("date", "")) > rdays:
        continue
    if kind == "News only" and a.get("kind") != "news":
        continue
    if kind == "Social only" and a.get("kind") != "social":
        continue
    if kind == "Notable only" and not a.get("notable"):
        continue
    hay = " ".join([a.get("title", ""), a.get("outlet") or "", " ".join(a.get("authors", []))]).lower()
    if q and q.lower() not in hay:
        continue
    tags = (" · ⭐ notable" if a.get("notable") else "") + (" · 💬 social" if a.get("kind") == "social" else "")
    by = f" · by {', '.join(a['authors'])}" if a.get("authors") else ""
    st.markdown(
        f"**[{a['title']}]({a['url']})**  \n"
        f"{a.get('outlet') or '—'}{by} · {a.get('date', '')}{tags}"
    )
    shown += 1
if not shown:
    st.info("No stories match those filters yet.")

# ---------------------------------------------------------------- reporter activity
st.subheader("Reporters on this beat")
st.caption(
    "Tracked automatically from bylines. Contact details are kept in BHU's "
    "internal press list — this public page shows activity only."
)
rows = [
    {
        "Reporter": c.get("name") or "(desk)",
        "Outlet": c.get("outlet") or "—",
        "Beat": c.get("beat") or "—",
        "Status": c.get("status") or "—",
        "Last byline seen": c.get("last_byline") or "—",
        "Bylines tracked": c.get("byline_count", 0),
    }
    for c in contacts
    if (c.get("category") or "").startswith("Reporter") or c.get("byline_count", 0) > 0
]
rows.sort(key=lambda r: (r["Last byline seen"] == "—", r["Last byline seen"]), reverse=False)
df = pd.DataFrame([r for r in rows if True])
df = df.sort_values(["Last byline seen"], ascending=False)
st.dataframe(df, use_container_width=True, hide_index=True)

with st.expander(f"No longer on the list ({len(removed)})"):
    for r in removed:
        st.markdown(f"- **{r.get('name') or '—'}** ({r.get('outlet') or '—'}) — {r.get('reason', '')}")

st.divider()
st.caption(
    "Reporters are added automatically when their byline appears on a tracked story, "
    "and rotated off after 6 months without one. "
    "Data refreshes every morning from the BHU media scanner."
)
