import os
import numpy as np
import pandas as pd
import random
import importlib.util

# ------------------------
# BI dataset generator with debug logs
# ------------------------

# Reproducible results
np.random.seed(42)
random.seed(42)

print("Starting dataset generation...")

# --- Dimension tables ---
dates = pd.date_range(start="2024-01-01", end="2024-12-31", freq="D")
print(f"Generated {len(dates)} dates for dim_date")

dim_date = pd.DataFrame({
    "date": dates,
})
dim_date["year"] = dim_date["date"].dt.year
dim_date["quarter"] = dim_date["date"].dt.quarter
dim_date["month"] = dim_date["date"].dt.month
dim_date["month_name"] = dim_date["date"].dt.month_name()
dim_date["day"] = dim_date["date"].dt.day
dim_date["weeknum"] = dim_date["date"].dt.isocalendar().week.astype(int)
dim_date["is_weekend"] = dim_date["date"].dt.dayofweek >= 5

# Product dimension
products = [
    ("P001", "Laptop 14\"", "Electronics", "Computers"),
    ("P002", "Laptop 16\"", "Electronics", "Computers"),
    ("P003", "Smartphone X", "Electronics", "Mobiles"),
    ("P004", "Smartphone Y", "Electronics", "Mobiles"),
    ("P005", "Office Chair", "Furniture", "Seating"),
    ("P006", "Standing Desk", "Furniture", "Desks"),
    ("P007", "Monitor 27\"", "Electronics", "Displays"),
    ("P008", "Headset Pro", "Accessories", "Audio"),
    ("P009", "Keyboard MX", "Accessories", "Input"),
    ("P010", "Mouse Ergo", "Accessories", "Input"),
    ("P011", "Docking Station", "Accessories", "Connectivity"),
    ("P012", "Webcam 4K", "Accessories", "Video"),
]
dim_product = pd.DataFrame(
    products, columns=["product_id", "product_name", "category", "subcategory"])
print(f"Loaded {len(dim_product)} products")

# Region dimension
regions = [
    ("R01", "APAC", "India", "Karnataka", "Bengaluru"),
    ("R02", "APAC", "India", "Maharashtra", "Mumbai"),
    ("R03", "APAC", "Singapore", "-", "Singapore"),
    ("R04", "EMEA", "UAE", "Dubai", "Dubai"),
    ("R05", "EMEA", "UK", "England", "London"),
    ("R06", "Americas", "USA", "California", "San Francisco"),
    ("R07", "Americas", "USA", "New York", "New York"),
    ("R08", "Americas", "Canada", "Ontario", "Toronto"),
]
dim_region = pd.DataFrame(
    regions, columns=["region_id", "region_group", "country", "state", "city"])
print(f"Loaded {len(dim_region)} regions")

channels = ["Online", "Retail", "Wholesale", "Partners"]

# --- Fact: Sales ---
n_sales = 1200
print("Generating sales fact table...")
sales_dates = np.random.choice(dates, size=n_sales)
sales_products = np.random.choice(dim_product["product_id"], size=n_sales)
sales_regions = np.random.choice(dim_region["region_id"], size=n_sales)
sales_channels = np.random.choice(
    channels, size=n_sales, p=[0.5, 0.25, 0.15, 0.10])

base_price = {
    "P001": 800, "P002": 1200, "P003": 700, "P004": 600, "P005": 150, "P006": 400,
    "P007": 300, "P008": 120, "P009": 90, "P010": 60, "P011": 180, "P012": 130
}

units = np.random.poisson(lam=5, size=n_sales) + 1
discount_rate = np.clip(np.random.normal(
    loc=0.07, scale=0.05, size=n_sales), 0, 0.25)
cost_margin = np.clip(np.random.normal(
    loc=0.65, scale=0.08, size=n_sales), 0.45, 0.85)

prices = np.array([base_price[pid] for pid in sales_products])
gross = units * prices * (1 - discount_rate)
cost = units * prices * cost_margin
fact_sales = pd.DataFrame({
    "sale_id": range(1, n_sales+1),
    "date": sales_dates,
    "product_id": sales_products,
    "region_id": sales_regions,
    "channel": sales_channels,
    "units": units,
    "discount_pct": np.round(discount_rate, 3),
    "revenue": np.round(gross, 2),
    "cost": np.round(cost, 2)
})
fact_sales["profit"] = np.round(fact_sales["revenue"] - fact_sales["cost"], 2)
print(f"Generated {len(fact_sales)} sales records")

# --- Marketing Spend ---
print("Generating marketing spend data...")
campaigns = ["Brand Push", "Summer Sale",
             "Diwali Promo", "Back to School", "Clearance"]
n_mkt = 400
mkt_dates = np.random.choice(dates, size=n_mkt)
mkt_regions = np.random.choice(dim_region["region_id"], size=n_mkt)
mkt_channels = np.random.choice(
    ["Search", "Social", "Email", "Affiliate", "Events"], size=n_mkt)
mkt_campaign = np.random.choice(campaigns, size=n_mkt)
spend = np.round(np.random.gamma(shape=3, scale=200, size=n_mkt), 2)
leads = np.random.poisson(lam=spend/80).astype(int)
marketing_spend = pd.DataFrame({
    "mkt_id": range(1, n_mkt+1),
    "date": mkt_dates,
    "region_id": mkt_regions,
    "channel": mkt_channels,
    "campaign": mkt_campaign,
    "spend": spend,
    "leads": leads
})
print(f"Generated {len(marketing_spend)} marketing spend records")

# --- Finance Actuals ---
print("Generating finance actuals data...")
departments = ["Sales", "Marketing", "Operations", "HR", "Finance", "IT"]
months = pd.date_range("2024-01-01", "2024-12-01", freq="MS")
rows = []
for d in departments:
    for m in months:
        revenue = np.random.gamma(9, 10000) if d == "Sales" else 0
        opex = np.random.gamma(7, 2000)
        capex = np.random.gamma(2, 5000) if d in [
            "IT", "Operations"] else np.random.gamma(1, 2000)
        rows.append([d, m, round(revenue, 2), round(opex, 2), round(capex, 2)])
finance_actuals = pd.DataFrame(
    rows, columns=["department", "month", "revenue", "opex", "capex"])
print(f"Generated finance actuals with {len(finance_actuals)} records")

# --- HR Roster ---
print("Generating HR roster data...")
n_emp = 80
first_names = ["Aarav", "Isha", "Kabir", "Diya", "Rohan", "Ananya", "Vihaan", "Meera",
               "Arjun", "Zoya", "Rahul", "Sneha", "Karan", "Priya", "Ishan", "Nisha", "Aman", "Aisha"]
last_names = ["Singh", "Sharma", "Patel", "Gupta",
              "Khan", "Iyer", "Reddy", "Kapoor", "Mehta", "Das"]
genders = ["Male", "Female", "Other"]
emp_ids = [f"E{1000+i}" for i in range(n_emp)]
departments_emp = np.random.choice(departments, size=n_emp, p=[
                                   0.25, 0.15, 0.25, 0.1, 0.1, 0.15])
hire_dates = pd.to_datetime(np.random.choice(
    pd.date_range("2021-01-01", "2024-12-31"), size=n_emp))

term_dates = []
termination_probability = 0.15
cutoff_date = pd.Timestamp("2024-12-31")
for hd in hire_dates:
    if np.random.rand() < termination_probability:
        earliest_td = hd + pd.Timedelta(days=30)
        random_offset = int(np.random.randint(100, 1200))
        candidate_td = hd + pd.Timedelta(days=random_offset)
        td = min(candidate_td, cutoff_date)
        if td > earliest_td:
            term_dates.append(td)
        else:
            term_dates.append(pd.NaT)
    else:
        term_dates.append(pd.NaT)

salaries = np.round(np.random.normal(1200000, 350000, size=n_emp), -3)
hr_roster = pd.DataFrame({
    "employee_id": emp_ids,
    "full_name": [f"{random.choice(first_names)} {random.choice(last_names)}" for _ in range(n_emp)],
    "gender": np.random.choice(genders, size=n_emp, p=[0.55, 0.43, 0.02]),
    "department": departments_emp,
    "hire_date": hire_dates,
    "termination_date": term_dates,
    "salary_inr": salaries
})
print(f"Generated HR roster with {len(hr_roster)} employees")

# --- Operations KPIs ---
print("Generating operations KPIs data...")
plants = ["BLR-Plant", "MUM-Plant", "SFO-Plant", "LDN-Plant"]
ops_rows = []
for r in dim_region["region_id"]:
    for pl in np.random.choice(plants, size=2, replace=False):
        for m in months:
            inv = int(np.random.normal(5000, 800))
            ontime = round(np.clip(np.random.normal(0.93, 0.05), 0.6, 1.0), 3)
            defects = int(np.clip(np.random.normal(220, 60), 50, 600))
            ops_rows.append([m, r, pl, inv, ontime, defects])
operations_kpis = pd.DataFrame(ops_rows, columns=[
                               "month", "region_id", "plant", "inventory_units", "on_time_rate", "defects_ppm"])
print(f"Generated {len(operations_kpis)} operations KPI records")

# ------------------------
# Excel export handling
# ------------------------


def detect_excel_engine():
    if importlib.util.find_spec("openpyxl") is not None:
        return "openpyxl"
    if importlib.util.find_spec("xlsxwriter") is not None:
        return "xlsxwriter"
    if importlib.util.find_spec("xlwt") is not None:
        return "xlwt"
    return None


# Prepare sheets dictionary
sheets = {
    "dim_date": dim_date,
    "dim_product": dim_product,
    "dim_region": dim_region,
    "fact_sales": fact_sales,
    "marketing_spend": marketing_spend,
    "finance_actuals": finance_actuals,
    "hr_roster": hr_roster,
    "operations_kpis": operations_kpis,
}

base_path = "/mnt/data"
if not os.path.isdir(base_path):
    base_path = "."

engine = detect_excel_engine()
print(f"Detected Excel engine: {engine}")
if engine is not None:
    ext = ".xls" if engine == "xlwt" else ".xlsx"
    out_path = os.path.join(base_path, f"data/pbixlexcelfile{ext}")
    try:
        with pd.ExcelWriter(out_path, engine=engine) as writer:
            for sheet_name, df in sheets.items():
                print(f"Writing sheet: {sheet_name} ({len(df)} rows)")
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Wrote Excel using engine '{engine}' to: {out_path}")
    except (OSError, ValueError, ImportError) as e:
        print(f"Error writing with engine {engine}: {e}")
else:
    print("No Excel writing engine available")

print("Dataset generation complete.")
