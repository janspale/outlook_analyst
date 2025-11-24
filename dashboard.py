import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

raw_path = os.path.join(BASE_DIR, "mail_raw.csv")
df_raw = pd.read_csv(raw_path)
df_raw["datetime"] = pd.to_datetime(df_raw["datetime"], errors="coerce")
df_raw = df_raw.dropna(subset=["datetime"])
st.set_page_config(page_title="Outlook Mail Analytics", layout="wide")

# ---------- LOAD DATA ----------
per_day = pd.read_csv("mail_counts_per_day.csv")
per_week = pd.read_csv("mail_counts_per_week.csv")

per_day["date"] = pd.to_datetime(per_day["date"])
per_week["week_start"] = (
    pd.to_datetime(per_week["year"].astype(str) + "-W" + per_week["week"].astype(str) + "-1",
                   format="%G-W%V-%u")
)

def calendar_heatmap_from_raw(df_raw, year, directions=("incoming","outgoing"),
                              hours=None, weekdays_only=False,
                              title="Calendar heatmap"):
    """
    df_raw columns: ['direction','datetime']
    hours: tuple (start_hour, end_hour) inclusive start, exclusive end
           e.g. (8,18) keeps 08:00-17:59
    weekdays_only: if True, keeps Mon-Fri
    """
    df = df_raw.copy()
    df["datetime"] = pd.to_datetime(df["datetime"])
    df = df[df["datetime"].dt.year == year]
    df = df[df["direction"].isin(directions)]

    if hours is not None:
        h0, h1 = hours
        df = df[(df["datetime"].dt.hour >= h0) & (df["datetime"].dt.hour < h1)]

    if weekdays_only:
        df = df[df["datetime"].dt.weekday < 5]

    if df.empty:
        st.warning(f"No raw data for year {year} / directions {directions}.")
        return

    # count per day
    per_day_raw = (
        df.assign(date=df["datetime"].dt.floor("D"))
          .groupby("date", as_index=False)
          .size()
          .rename(columns={"size":"count"})
    )

    iso = per_day_raw["date"].dt.isocalendar()
    per_day_raw["iso_week"] = iso.week.astype(int)
    per_day_raw["iso_year"] = iso.year.astype(int)
    per_day_raw["weekday"] = per_day_raw["date"].dt.weekday  # Mon=0

    # pad ISO weeks spilling into neighbor year
    per_day_raw.loc[per_day_raw["iso_year"] < year, "iso_week"] = 0
    max_week = int(pd.Timestamp(year=year, month=12, day=28).isocalendar().week)
    per_day_raw.loc[per_day_raw["iso_year"] > year, "iso_week"] = max_week + 1

    weeks = list(range(0, max_week + 2))
    weekdays = list(range(7))

    grid = pd.DataFrame(index=weekdays, columns=weeks, data=0)

    for _, r in per_day_raw.iterrows():
        grid.at[int(r["weekday"]), int(r["iso_week"])] = r["count"]

    # hover text (date per cell)
    date_map = per_day_raw.set_index(["weekday","iso_week"])["date"].to_dict()
    hover = []
    for wd in weekdays:
        row = []
        for wk in weeks:
            d = date_map.get((wd, wk))
            row.append("" if d is None else d.strftime("%Y-%m-%d"))
        hover.append(row)

    fig = go.Figure(
        data=go.Heatmap(
            z=grid.values,
            x=weeks,
            y=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"],
            customdata=hover,
            hovertemplate="Date: %{customdata}<br>Mails: %{z}<extra></extra>"
        )
    )

    fig.update_layout(
        title=title,
        xaxis_title="ISO week",
        yaxis_title="Weekday",
        height=280,
        margin=dict(l=40, r=10, t=50, b=40),
    )
    fig.update_xaxes(showgrid=False, tickmode="array",
                     tickvals=[w for w in weeks if w % 4 == 0])
    fig.update_yaxes(showgrid=False)

    st.plotly_chart(fig, use_container_width=True)


# if you have raw data, the heatmap will be more accurate
raw_available = False
try:
    df = pd.read_csv("mail_raw.csv")
    df["datetime"] = pd.to_datetime(df["datetime"])
    raw_available = True
except FileNotFoundError:
    df = None

# ---------- SIDEBAR FILTERS ----------
st.sidebar.header("Filters")
directions = st.sidebar.multiselect(
    "Direction", options=per_day["direction"].unique().tolist(),
    default=per_day["direction"].unique().tolist()
)
start = st.sidebar.date_input("From", per_day["date"].min().date())
end = st.sidebar.date_input("To", per_day["date"].max().date())

mask = (
    per_day["direction"].isin(directions) &
    (per_day["date"].dt.date >= start) &
    (per_day["date"].dt.date <= end)
)
pd_f = per_day[mask].copy()

# ---------- KPI ROW ----------
total_in = per_day[per_day.direction=="incoming"]["count"].sum()
total_out = per_day[per_day.direction=="outgoing"]["count"].sum()
total = total_in + total_out

col1, col2, col3, col4 = st.columns(4)
col1.metric("Total mails", f"{total:,}".replace(",", " "))
col2.metric("Incoming", f"{total_in:,}".replace(",", " "))
col3.metric("Outgoing", f"{total_out:,}".replace(",", " "))
col4.metric("In/Out ratio", f"{total_in/total_out:.2f}" if total_out else "∞")

st.divider()

# ---------- 1) DAILY TREND ----------
st.subheader("Daily trend")
pd_plot = pd_f.pivot_table(index="date", columns="direction", values="count", fill_value=0)

# MA30
for c in pd_plot.columns:
    pd_plot[c + "_ma30"] = pd_plot[c].rolling(30, min_periods=1).mean()

fig = px.line(pd_plot, x=pd_plot.index, y=pd_plot.columns,
              labels={"value":"mail count", "date":"date"},
              title="Daily counts + MA30")
st.plotly_chart(fig, use_container_width=True)

# ---------- 2) WEEKLY COUNTS ----------
st.subheader("Weekly counts")
pw_f = per_week[per_week["direction"].isin(directions)].copy()
fig2 = px.bar(
    pw_f, x="week_start", y="count", color="direction", barmode="group",
    labels={"week_start":"week", "count":"count"},
    title="Mail counts per week"
)
st.plotly_chart(fig2, use_container_width=True)

# ---------- 3) HEATMAP DAY x HOUR ----------
st.subheader("Weekly rhythm (day × hour)")
if raw_available:
    df_f = df[df["direction"].isin(directions)].copy()
    df_f["hour"] = df_f["datetime"].dt.hour
    df_f["weekday"] = df_f["datetime"].dt.day_name()

    for direction in directions:
        sub = df_f[df_f.direction==direction]
        mat = sub.groupby(["weekday","hour"]).size().reset_index(name="count")
        pivot = mat.pivot(index="weekday", columns="hour", values="count").fillna(0)
        # order days
        order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
        pivot = pivot.reindex(order)

        fig3 = px.imshow(
            pivot, aspect="auto",
            labels=dict(x="hour", y="weekday", color="count"),
            title=f"Heatmap: {direction}"
        )
        st.plotly_chart(fig3, use_container_width=True)
else:
    st.info("Heatmap requires mail_raw.csv (with timestamps of individual emails).")

# ---------- 4) CALENDAR HEATMAP ----------
st.subheader("Calendar heatmap (from raw data)")

available_years = sorted(df_raw["datetime"].dt.year.unique().tolist())
year = st.selectbox("Year", available_years, index=len(available_years)-1)

directions_sel = st.multiselect(
    "Direction",
    options=["incoming","outgoing"],
    default=["incoming","outgoing"]
)

colA, colB, colC = st.columns([1,1,2])
with colA:
    weekdays_only = st.checkbox("Weekdays only (Mon–Fri)", value=False)
with colB:
    work_hours = st.checkbox("Working hours only", value=False)
with colC:
    if work_hours:
        h0, h1 = st.slider("Hours", 0, 24, (8, 18))
        hours = (h0, h1)
    else:
        hours = None

calendar_heatmap_from_raw(
    df_raw=df_raw,
    year=year,
    directions=tuple(directions_sel),
    hours=hours,
    weekdays_only=weekdays_only,
    title=f"Email activity in {year} ({', '.join(directions_sel)})"
)


st.caption("Data are loaded from CSVs generated by the COM script.")
