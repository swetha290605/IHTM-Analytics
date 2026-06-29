import json
from pathlib import Path
import numpy as np
import pandas as pd
from datetime import datetime
import sys
sys.stdout.reconfigure(encoding='utf-8')

# ─────────────────────────────────────────────────────────────────────
# PATHS
# ─────────────────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUTPUT_JSON = ROOT / "analytics_data.json"

DATA_DIR.mkdir(exist_ok=True)

# ─────────────────────────────────────────────────────────────────────
# LOAD ALL EXCELS
# ─────────────────────────────────────────────────────────────────────

def load_excel(filename):
    path = DATA_DIR / filename
    if path.exists():
        df = pd.read_excel(path)
        print(f"✅ Loaded: {filename}")
        return df
    else:
        print(f"⚠️  Missing: {filename}")
        return None

df_query   = load_excel("Query_Analytics.xlsx")
df_safety  = load_excel("Safety_Status.xlsx")
df_revenue = load_excel("IHTM_Revenue.xlsx")
df_issues  = load_excel("Field_Issues.xlsx")
df_cost    = load_excel("Cost_Competency.xlsx")
df_punch   = load_excel("Project_Punch_Points.xlsx")
df_tools   = load_excel("Tools_Under_Progress.xlsx")
df_keshkomi = load_excel("KeshKomi.xlsx")
df_ecr      = load_excel("Enquiry_Conversion_Ratio.xlsx")

# ─────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────

def safe_records(df_sub):
    return json.loads(df_sub.to_json(orient="records"))

def safe_list(data):
    return json.loads(json.dumps(data, default=str))

def to_num(series):
    """Coerce a string series to numeric, silently turning non-numeric to NaN."""
    return pd.to_numeric(series, errors='coerce')

# ─────────────────────────────────────────────────────────────────────
# 1. QUERY ANALYTICS
# ─────────────────────────────────────────────────────────────────────

def build_query_analytics():
    if df_query is None:
        return {}

    d = df_query.copy()

    stage_dist = (
        d.groupby("Stage").size()
        .reset_index(name="count")
        .assign(pct=lambda x: (x["count"] / x["count"].sum() * 100).round(1))
        .sort_values("count", ascending=False)
    )

    plant_dist = (
        d.groupby("Plant").size()
        .reset_index(name="count")
        .assign(pct=lambda x: (x["count"] / x["count"].sum() * 100).round(1))
        .sort_values("count", ascending=False)
    )

    return {
        "stage_distribution": safe_records(stage_dist),
        "plant_distribution": safe_records(plant_dist),
        "total_queries": int(len(d)),
    }

# ─────────────────────────────────────────────────────────────────────
# 2. SAFETY STATUS
# ─────────────────────────────────────────────────────────────────────

def build_safety():
    if df_safety is None:
        return {}

    d = df_safety.copy()
    d["No of Accidents"] = to_num(d["No of Accidents"])
    d = d.dropna(subset=["No of Accidents"])

    # Add Year column so JS year-filter works
    MONTH_TO_YEAR = {
        'January': '2026', 'February': '2026', 'March': '2026', 'April': '2026',
        'May': '2026', 'June': '2026', 'July': '2026', 'August': '2026',
        'September': '2026', 'October': '2026', 'November': '2026', 'December': '2026',
    }
    if "Year" not in d.columns:
        d["Year"] = d["Month"].map(MONTH_TO_YEAR).fillna('2026')

    month_accidents = safe_records(d[["Month", "Year", "No of Accidents"]])

    safety_calendar = []
    for _, row in d.iterrows():
        accidents = int(row["No of Accidents"])
        if accidents == 0:
            status, label = "green", "No Accident"
        elif accidents == 1:
            status, label = "blue", "NLWD"
        else:
            status, label = "red", f"LWD ({accidents})"
        safety_calendar.append({
            "month": row["Month"],
            "accidents": accidents,
            "status": status,
            "label": label,
        })

    return {
        "month_accidents": month_accidents,
        "safety_calendar": safety_calendar,
    }

# ─────────────────────────────────────────────────────────────────────
# 3. IHTM REVENUE
# ─────────────────────────────────────────────────────────────────────

def build_revenue():
    if df_revenue is None:
        return {}

    d = df_revenue.copy()
    d["Revenue (Million Rs)"] = to_num(d["Revenue (Million Rs)"])
    d["Target (Million Rs)"]  = to_num(d["Target (Million Rs)"])
    total_revenue = d["Revenue (Million Rs)"].dropna().sum()

    target_values = d["Target (Million Rs)"].dropna()
    total_target = float(target_values.iloc[0]) if not target_values.empty else 0

    target_pct = round(total_revenue / total_target * 100, 1) if total_target > 0 else 0

    months_left = 12 - len(d)
    target_remaining = total_target - total_revenue

    needed_per_3months = (
    round(target_remaining / (months_left / 3), 1)
    if months_left > 0
    else 0
    )

    # Build per-month array for the chart
    rev_by_month = safe_records(d[["Month", "Revenue (Million Rs)", "Target (Million Rs)"]].dropna(subset=["Month"]))

    return {
        "total_revenue": round(float(total_revenue), 1),
        "yearly_target": float(total_target),
        "target_achievement_pct": target_pct,
        "needed_per_3months": needed_per_3months,
        "revenue_by_month": rev_by_month,
    }

# ─────────────────────────────────────────────────────────────────────
# 4. FIELD ISSUES
# ─────────────────────────────────────────────────────────────────────

def build_field_issues():
    if df_issues is None:
        return {}

    d = df_issues.copy()
    for col in ["Identified", "Solved", "Cum Identified", "Cum Solved"]:
        if col in d.columns:
            d[col] = to_num(d[col])

    # Add Year column so JS year-filter works
    if "Year" not in d.columns:
        MONTH_YEAR = {
            'January': '2026', 'February': '2026', 'March': '2026', 'April': '2026',
            'May': '2026', 'June': '2026', 'July': '2026', 'August': '2026',
            'September': '2026', 'October': '2026', 'November': '2026', 'December': '2026',
        }
        d["Year"] = d["Month"].map(MONTH_YEAR).fillna('2026')

    cols = ["Month", "Year", "Identified", "Solved", "Cum Identified", "Cum Solved"]
    cols = [c for c in cols if c in d.columns]
    return {
        "issues_by_month": safe_records(d[cols]),
    }

# ─────────────────────────────────────────────────────────────────────
# 5. COST COMPETENCY
# ─────────────────────────────────────────────────────────────────────

def build_cost_competency():
    if df_cost is None:
        return {}

    d = df_cost.copy()
    for col in ["Local Value", "Target Value", "IHTM Value", "Cost in Manufacturing"]:
        if col in d.columns:
            d[col] = to_num(d[col])
    return {
        "cost_by_month": safe_records(
            d[["Month", "Local Value", "Target Value", "IHTM Value", "Cost in Manufacturing"]]
        ),
    }

# ─────────────────────────────────────────────────────────────────────
# 6. PROJECT PUNCH POINTS
# ─────────────────────────────────────────────────────────────────────

def build_punch_points():
    if df_punch is None:
        return {}

    d = df_punch.copy()
    item_col   = d.columns[0]
    month_cols = [c for c in d.columns[1:] if str(c).strip() != ""]

    d_long = d.melt(
        id_vars=[item_col],
        value_vars=month_cols,
        var_name="Month",
        value_name="Item Count",
    ).rename(columns={item_col: "Item"})

    d_long = d_long.dropna(subset=["Item Count"])
    d_long["Item Count"] = to_num(d_long["Item Count"])
    d_long = d_long.dropna(subset=["Item Count"])
    d_long = d_long[d_long["Item Count"] != 0]
    d_long["Punch Points"] = d_long["Item Count"]

    return {
        "punch_by_month": safe_records(
            d_long[["Month", "Item", "Item Count", "Punch Points"]]
        ),
    }

# ─────────────────────────────────────────────────────────────────────
# 7. TOOLS UNDER PROGRESS
# ─────────────────────────────────────────────────────────────────────

def build_tools():
    if df_tools is None:
        return {}

    df = df_tools.copy()
    df["In Mill"]     = to_num(df["In Mill"])
    df["No of Tools"] = to_num(df["No of Tools"])
    total_cost  = round(float(df["In Mill"].sum()), 1)
    df = df.dropna(subset=["No of Tools"])
    df = df.iloc[:-1]  
    total_tools = int(df["No of Tools"].sum())

    plant_grp = df.groupby("Plant", as_index=False).agg(
        cost=("In Mill", "sum"),
        tools=("No of Tools", "sum"),
    )
    plant_grp["cost"] = plant_grp["cost"].round(1)
    plant_grp["pct"]  = (plant_grp["cost"] / total_cost * 100).round(1)

    shop_grp = df.groupby(["Plant", "Shop"], as_index=False).agg(
        cost=("In Mill", "sum"),
        tools=("No of Tools", "sum"),
    )
    shop_grp["cost"] = shop_grp["cost"].round(1)
    shop_grp["pct"]  = (shop_grp["cost"] / total_cost * 100).round(1)
    shop_grp["name"] = shop_grp["Plant"] + " - " + shop_grp["Shop"]

    rev_grp = df.groupby("Budget Type", as_index=False).agg(
        cost=("In Mill", "sum"),
        tools=("No of Tools", "sum"),
    )
    rev_grp["cost"] = rev_grp["cost"].round(1)
    rev_grp["pct"]  = (rev_grp["cost"] / total_cost * 100).round(1)

    # Build raw entries for table preview
    raw_entries = safe_records(df)

    return {
        "total_cost_million": total_cost,
        "total_tools": total_tools,
        "plant_distribution": plant_grp.to_dict(orient="records"),
        "plant_shop_distribution": shop_grp.to_dict(orient="records"),
        "budget_distribution" : rev_grp.to_dict(orient = "records"),
        "raw_entries": raw_entries,
    }

# ─────────────────────────────────────────────────────────────────────
# 8. ENQUIRY CONVERSION RATIO (FIXED)
# ─────────────────────────────────────────────────────────────────────

def build_ecr():
    if df_ecr is None or df_ecr.empty:
        return {}

    d = df_ecr.copy()

    # Get actual column names
    cols = list(d.columns)
    if len(cols) < 3:
        return {}

    # Map by position: col 0 = Month, col 1 = Total Queries, col 2 = Queries Accepted
    col_map = {
        'Month': cols[0],
        'Total Queries': cols[1],
        'Queries Accepted': cols[2],
    }
    if len(cols) >= 4:
        col_map['Conversion Ratio'] = cols[3]
    if len(cols) >= 5:
        col_map['Average Conversion Ratio'] = cols[4]

    # Rename columns to standard names
    rename_map = {v: k for k, v in col_map.items()}
    d = d.rename(columns=rename_map)

    # Drop rows with missing Month
    d = d.dropna(subset=['Month'])

    # Convert numeric columns
    d['Total Queries'] = pd.to_numeric(d['Total Queries'], errors='coerce').fillna(0).astype(int)
    d['Queries Accepted'] = pd.to_numeric(d['Queries Accepted'], errors='coerce').fillna(0).astype(int)

    if 'Conversion Ratio' in d.columns:
        d['Conversion Ratio'] = pd.to_numeric(d['Conversion Ratio'], errors='coerce').fillna(0).round(1)
    if 'Average Conversion Ratio' in d.columns:
        d['Average Conversion Ratio'] = pd.to_numeric(d['Average Conversion Ratio'], errors='coerce').fillna(0).round(1)

    # Build output columns
    output_cols = ['Month', 'Total Queries', 'Queries Accepted']
    if 'Conversion Ratio' in d.columns:
        output_cols.append('Conversion Ratio')
    if 'Average Conversion Ratio' in d.columns:
        output_cols.append('Average Conversion Ratio')

    # Get average conversion from first row if available
    avg_conversion = 0.0
    if 'Average Conversion Ratio' in d.columns and len(d) > 0:
        val = d['Average Conversion Ratio'].iloc[0]
        if pd.notna(val):
            avg_conversion = round(float(val), 1)

    # Convert to plain Python objects for JSON - FIX: format dates as strings
    by_month = []
    for _, row in d[output_cols].iterrows():
        record = {}
        for col in output_cols:
            val = row[col]
            if col == 'Month':
                # Format month as string - handle both datetime and string
                if pd.isna(val):
                    record[col] = ''
                elif hasattr(val, 'strftime'):
                    # It's a datetime - format it
                    record[col] = val.strftime('%b-%y')
                else:
                    record[col] = str(val).strip()
            elif pd.isna(val):
                record[col] = 0
            elif isinstance(val, (np.integer, np.int64)):
                record[col] = int(val)
            elif isinstance(val, (np.floating, np.float64)):
                record[col] = float(val)
            else:
                record[col] = val
        by_month.append(record)

    return {
        "by_month": by_month,
        "avg_conversion": avg_conversion,
        "total_queries": int(d['Total Queries'].sum()),
        "total_accepted": int(d['Queries Accepted'].sum()),
    }

# ─────────────────────────────────────────────────────────────────────
# 9. KESHKOMI
# ─────────────────────────────────────────────────────────────────────

def build_kesh_komi():
    if df_keshkomi is None:
        return {}

    d = df_keshkomi.copy()

    # Normalise column names to snake_case so keys match JS + server API
    col_rename = {}
    for col in d.columns:
        key = str(col).strip().lower().replace(' ', '_')
        col_rename[col] = key
    d = d.rename(columns=col_rename)

    # Map known aliases
    alias = {
        'kadai_point': ['kadai_point', 'kadai point', 'point', 'issue'],
        'action_plan': ['action_plan', 'action plan', 'action'],
        'responsibility': ['responsibility', 'responsible', 'owner'],
        'target_date': ['target_date', 'target date', 'date'],
        'sl_no.': 'sl_no',
    }
    for old, new in [('sl_no.', 'sl_no')]:
        if old in d.columns:
            d = d.rename(columns={old: new})

    # Ensure status column
    if 'status' in d.columns:
        d['status'] = d['status'].astype('object')
        legacy_hp = d['status'] == 'High Priority'
        d.loc[legacy_hp, 'status'] = 'Incomplete'
        if 'high_priority' not in d.columns:
            d['high_priority'] = False
        d.loc[legacy_hp, 'high_priority'] = True
    else:
        d['status'] = 'Incomplete'
        d['high_priority'] = False

    # Ensure high_priority as bool
    if 'high_priority' in d.columns:
        d['high_priority'] = d['high_priority'].fillna(False).astype(bool)
    else:
        d['high_priority'] = False

    # Parse target_date for Month grouping
    if 'target_date' in d.columns:
        d['target_date'] = pd.to_datetime(d['target_date'], errors='coerce')
        d['month'] = d['target_date'].dt.strftime('%b %Y')
        d['target_date'] = d['target_date'].dt.strftime('%Y-%m-%d').where(d['target_date'].notna(), None)
    else:
        d['month'] = None
        d['target_date'] = None

    # Status distribution
    status_dist = (
        d.groupby("status").size()
        .reset_index(name="count")
        .assign(pct=lambda x: (x["count"] / x["count"].sum() * 100).round(1))
        .sort_values("count", ascending=False)
    )

    hp_count = int(d['high_priority'].sum())

    # Monthly breakdown
    monthly = pd.DataFrame(columns=['month', 'status', 'count'])
    monthly_hp = pd.DataFrame(columns=['month', 'hp_count'])
    if 'month' in d.columns:
        monthly = (
            d.dropna(subset=['month'])
            .groupby(['month', 'status']).size()
            .reset_index(name='count')
        )
        monthly_hp = (
            d[d['high_priority']].dropna(subset=['month'])
            .groupby('month').size()
            .reset_index(name='hp_count')
        )

    return {
        "entries": json.loads(d.to_json(orient="records", default_handler=str)),
        "status_distribution": safe_records(status_dist),
        "monthly_breakdown": safe_records(monthly),
        "monthly_hp_breakdown": safe_records(monthly_hp),
        "total_entries": int(len(d)),
        "high_priority": hp_count,
        "completed":  int(len(d[d['status'] == 'Completed'])),
        "partial":    int(len(d[d['status'] == 'Partially Completed'])),
        "incomplete": int(len(d[d['status'] == 'Incomplete'])),
    }

# ─────────────────────────────────────────────────────────────────────
# BUILD GLOBAL KPIs
# ─────────────────────────────────────────────────────────────────────

safety_accidents = 0
try:
    if df_safety is not None and "No of Accidents" in df_safety.columns:
        safety_accidents = int(pd.to_numeric(df_safety["No of Accidents"], errors='coerce').sum())
except:
    pass

global_kpis = {
    "timestamp": datetime.now().isoformat(),
    "total_queries": build_query_analytics().get("total_queries", 0),
    "total_tools":   build_tools().get("total_tools", 0),
    "safety_accidents": safety_accidents,
}

# ─────────────────────────────────────────────────────────────────────
# BUILD RESULT
# ─────────────────────────────────────────────────────────────────────

result = {
    "meta": {
        "generated_at": datetime.now().isoformat(),
        "source": "Manufacturing Analytics",
    },
    "global_kpis":       global_kpis,
    "query_analytics":   build_query_analytics(),
    "safety":            build_safety(),
    "revenue":           build_revenue(),
    "field_issues":      build_field_issues(),
    "cost_competency":   build_cost_competency(),
    "punch_points":      build_punch_points(),
    "tools":             build_tools(),
    "kesh_komi":         build_kesh_komi(),
    "ecr":               build_ecr(),
}

OUTPUT_JSON.write_text(json.dumps(result, indent=2, default=str), encoding="utf-8")
print(f"\n✅ Analytics data written to: {OUTPUT_JSON}")
print(f"   Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")