import win32com.client
import pandas as pd
import matplotlib.pyplot as plt

# ========= SETTINGS =========
START_DATE = "01/01/2025"   # Outlook filter expects mm/dd/yyyy
END_DATE   = "12/31/2025"

# ========= Outlook session =========
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

def fetch_folder(folder_id, incoming=True):
    folder = outlook.GetDefaultFolder(folder_id)
    items = folder.Items

    # Outlook filtering
    if incoming:
        filt = f"[ReceivedTime] >= '{START_DATE}' AND [ReceivedTime] <= '{END_DATE}'"
        items.Sort("[ReceivedTime]", True)
    else:
        filt = f"[SentOn] >= '{START_DATE}' AND [SentOn] <= '{END_DATE}'"
        items.Sort("[SentOn]", True)

    items = items.Restrict(filt)

    rows = []
    for item in items:
        try:
            # only real mail items
            if item.Class != 43:
                continue

            dt = item.ReceivedTime if incoming else item.SentOn
            if dt is None:
                continue

            rows.append({
                "direction": "incoming" if incoming else "outgoing",
                "datetime": pd.Timestamp(dt)
            })
        except Exception:
            # sometimes COM raises an error on a special item
            continue

    return pd.DataFrame(rows)

# ========= Load Inbox / Sent =========
df_in  = fetch_folder(6, incoming=True)    # Inbox
df_out = fetch_folder(5, incoming=False)  # Sent Items

df = pd.concat([df_in, df_out], ignore_index=True)

# ... after df = pd.concat([...])

if df.empty:
    print("DataFrame is empty – Outlook returned no items in the specified interval.")
    exit()

from datetime import datetime
import pandas as pd

def safe_naive_dt(x):
    if x is None:
        return pd.NaT
    try:
        ts = pd.Timestamp(x)  # handles pywintypes.TimeType as well
        # if tz-aware, drop timezone
        if ts.tzinfo is not None:
            ts = ts.tz_convert(None)
        # return plain python datetime
        return ts.to_pydatetime()
    except Exception:
        return pd.NaT

# --- after df = pd.concat([...]) and empty-check ---
df["datetime"] = df["datetime"].astype("object").apply(safe_naive_dt)
df["datetime"] = pd.to_datetime(df["datetime"], errors="coerce")
df = df.dropna(subset=["datetime"])


# ========= Aggregation =========
df["date"] = df["datetime"].dt.date
df["year"] = df["datetime"].dt.year
df["week"] = df["datetime"].dt.isocalendar().week
df["weekday"] = df["datetime"].dt.day_name()

per_day = df.groupby(["direction", "date"]).size().reset_index(name="count")
per_week = df.groupby(["direction", "year", "week"]).size().reset_index(name="count")
per_weekday = df.groupby(["direction", "weekday"]).size().reset_index(name="count")

# ========= Export =========
per_day.to_csv("mail_counts_per_day.csv", index=False)
per_week.to_csv("mail_counts_per_week.csv", index=False)
per_weekday.to_csv("mail_counts_per_weekday.csv", index=False)

print("Per-day head:\n", per_day.head())
print("Per-week head:\n", per_week.head())

# ========= Plots =========
for direction in ["incoming", "outgoing"]:
    sub = per_day[per_day["direction"] == direction].copy()
    sub["date"] = pd.to_datetime(sub["date"])
    sub = sub.sort_values("date")

    plt.figure()
    plt.plot(sub["date"], sub["count"])
    plt.title(f"Number of emails per day ({direction})")
    plt.xlabel("Date")
    plt.ylabel("Number of emails")
    plt.tight_layout()
    plt.show()

# per_day: columns [direction, date, count]
per_day["date"] = pd.to_datetime(per_day["date"])
per_day = per_day.sort_values("date")

fig, axes = plt.subplots(2, 1, sharex=True, figsize=(12, 7))

for ax, direction in zip(axes, ["incoming", "outgoing"]):
    sub = per_day[per_day.direction == direction].copy()
    sub["ma30"] = sub["count"].rolling(30, min_periods=1).mean()

    ax.plot(sub["date"], sub["count"], alpha=0.35, linewidth=1)
    ax.plot(sub["date"], sub["ma30"], linewidth=2)
    ax.set_title(f"{direction}: emails per day + MA30")
    ax.set_ylabel("count")
    ax.grid(True, alpha=0.2)

axes[-1].set_xlabel("date")
plt.tight_layout()
plt.savefig("daily_trends_ma30.png")
plt.show()

# per_week: [direction, year, week, count]
pw = per_week.copy()
pw["week_label"] = pw["year"].astype(str) + "-W" + pw["week"].astype(str)

pivot = pw.pivot_table(index=["year","week"], columns="direction", values="count", fill_value=0).reset_index()

plt.figure(figsize=(12,5))
plt.bar(pivot.index-0.2, pivot["incoming"], width=0.4, label="incoming")
plt.bar(pivot.index+0.2, pivot["outgoing"], width=0.4, label="outgoing")
plt.title("Number of emails per week")
plt.xlabel("week of year")
plt.ylabel("count")
plt.legend()
plt.tight_layout()
plt.savefig("weekly_mail_counts.png")
plt.show()

df["hour"] = df["datetime"].dt.hour
df["weekday"] = df["datetime"].dt.weekday  # 0=Mon
weekday_names = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]

def heat(direction):
    sub = df[df.direction == direction]
    mat = sub.groupby(["weekday","hour"]).size().unstack(fill_value=0)
    mat = mat.reindex(range(7)).reindex(columns=range(24), fill_value=0)
    return mat

mat_in = heat("incoming")
mat_out = heat("outgoing")

fig, axes = plt.subplots(1,2, figsize=(14,4), sharey=True)

for ax, mat, title in zip(axes, [mat_in, mat_out], ["Incoming", "Outgoing"]):
    im = ax.imshow(mat.values, aspect="auto")
    ax.set_title(title)
    ax.set_xticks(range(24))
    ax.set_xlabel("hour")
    ax.set_yticks(range(7))
    ax.set_yticklabels(weekday_names)

fig.colorbar(im, ax=axes.ravel().tolist(), shrink=0.8, label="number of emails")
plt.tight_layout()
plt.savefig("heatmap_day_hour.png")
plt.show()

pivot_d = per_day.pivot_table(index="date", columns="direction", values="count", fill_value=0)
pivot_d.index = pd.to_datetime(pivot_d.index)
pivot_d = pivot_d.sort_index()

pivot_d_ma = pivot_d.rolling(14, min_periods=1).mean()

plt.figure(figsize=(12,4))
plt.stackplot(pivot_d_ma.index, pivot_d_ma["incoming"], pivot_d_ma["outgoing"], labels=["incoming","outgoing"], alpha=0.8)
plt.title("Incoming vs Outgoing (MA14)")
plt.ylabel("number of emails")
plt.legend(loc="upper right")
plt.tight_layout()
plt.savefig("incoming_outgoing_ma14.png")
plt.show()

plt.figure(figsize=(10,4))
for direction in ["incoming","outgoing"]:
    sub = per_day[per_day.direction==direction]["count"]
    plt.hist(sub, bins=30, alpha=0.5, label=direction)

plt.title("Distribution of daily email counts")
plt.xlabel("emails per day")
plt.ylabel("number of days")
plt.legend()
plt.tight_layout()
plt.savefig("daily_mail_count_distribution.png")
plt.show()

df.to_csv("mail_raw.csv", index=False)
