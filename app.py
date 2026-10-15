from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# ==========================================================
# PAGE CONFIG
# ==========================================================
st.set_page_config(
    page_title="Freight Cost Analytics",
    page_icon="🚚",
    layout="wide",
)

st.title("🚚 Freight Cost Analytics Dashboard")
st.markdown(
    """
Upload:
- Manual Accruals file
- SAP ERP file

Master files are automatically loaded from the repository's `master_data` folder.
"""
)


# ==========================================================
# MASTER-DATA CONFIGURATION
# ==========================================================
APP_DIR = Path(__file__).resolve().parent
MASTER_FOLDER = APP_DIR / "master_data"

CARRIER_FILE = MASTER_FOLDER / "ERP Carrier Name.xlsx"
CUSTOMERS_FILE = MASTER_FOLDER / "ERP Customers.xlsx"
PLANTS_FILE = MASTER_FOLDER / "ERP Plants.xlsx"
SHIPPING_FILE = MASTER_FOLDER / "ERP Shipping Point.xlsx"
CURRENCY_FILE = MASTER_FOLDER / "Currency conversion table.xlsx"

MASTER_FILES = {
    "carrier": CARRIER_FILE,
    "customers": CUSTOMERS_FILE,
    "plants": PLANTS_FILE,
    "shipping points": SHIPPING_FILE,
    "currency conversion": CURRENCY_FILE,
}


# ==========================================================
# HELPERS
# ==========================================================
@st.cache_data
def load_excel(file_path: Path) -> pd.DataFrame:
    return pd.read_excel(file_path, engine="openpyxl")


@st.cache_data
def load_master_data():
    if not MASTER_FOLDER.is_dir():
        raise FileNotFoundError(
            f"Master-data folder not found: {MASTER_FOLDER}"
        )

    missing_files = [
        path.name for path in MASTER_FILES.values() if not path.is_file()
    ]

    if missing_files:
        available_files = sorted(
            path.name for path in MASTER_FOLDER.iterdir() if path.is_file()
        )
        raise FileNotFoundError(
            "One or more master-data files are missing. "
            f"Expected folder: {MASTER_FOLDER}. "
            f"Missing: {missing_files}. "
            f"Available: {available_files}."
        )

    return (
        load_excel(CARRIER_FILE),
        load_excel(CUSTOMERS_FILE),
        load_excel(PLANTS_FILE),
        load_excel(SHIPPING_FILE),
        load_excel(CURRENCY_FILE),
    )


def clean_carrier(value):
    if pd.isna(value):
        return value
    value = str(value)
    if "/" in value:
        return value.split("/")[0].strip()
    return value.strip()


def normalize_currency(value):
    if pd.isna(value):
        return None
    return str(value).strip().upper()


def classify_goods(value):
    if pd.isna(value):
        return "NFG"
    value = str(value).strip()
    if value.startswith("1") or value.startswith("4"):
        return "FG"
    return "NFG"


# ==========================================================
# LOAD STATIC FILES
# ==========================================================
try:
    (
        carrier_master,
        customer_master,
        plant_master,
        shipping_master,
        currency_master,
    ) = load_master_data()
except Exception as exc:
    st.error(f"Unable to load master files: {exc}")
    st.stop()


# ==========================================================
# FILE UPLOADS
# ==========================================================
manual_file = st.file_uploader(
    "Upload Manual Accruals", type=["xlsx", "xls"]
)
sap_file = st.file_uploader(
    "Upload SAP ERP", type=["xlsx", "xls"]
)

if not manual_file or not sap_file:
    st.stop()


# ==========================================================
# LOAD DATA
# ==========================================================
try:
    manual_raw = pd.read_excel(manual_file)
    sap_raw = pd.read_excel(sap_file)
except Exception as exc:
    st.error(f"Unable to read an uploaded Excel file: {exc}")
    st.stop()


# ==========================================================
# MANUAL ACCRUALS TRANSFORMATION
# ==========================================================
manual = pd.DataFrame()
manual["Carrier"] = (
    manual_raw["Carrier Description"].astype(str).apply(clean_carrier)
)
manual["Type of Goods"] = manual_raw["PBU"]
manual["Date"] = pd.to_datetime(
    manual_raw["Planned Arrival Date-Last Stop"], errors="coerce"
)
manual["Origin"] = manual_raw["Source Location Description"]
manual["Destination"] = manual_raw["Destination Location Descripti"]
manual["Value"] = pd.to_numeric(
    manual_raw["Net Amt in Doc Crcy"], errors="coerce"
)
manual["Currency"] = (
    manual_raw["Currency"].astype(str).apply(normalize_currency)
)
manual["Source"] = "Manual Accruals"


# ==========================================================
# SAP ERP TRANSFORMATION
# ==========================================================
sap = sap_raw.copy()

# Carrier
carrier_master["CarrierKey"] = (
    carrier_master["ServcAgent"].astype(str).str.strip()
)
sap["CarrierKey"] = sap["ServcAgent"].astype(str).str.strip()
sap = sap.merge(
    carrier_master[["CarrierKey", "Name 1"]].rename(
        columns={"Name 1": "CarrierName"}
    ),
    on="CarrierKey",
    how="left",
)
sap["Carrier"] = sap["CarrierName"]

# Type of goods
sap["Type of Goods"] = sap["Product Hierarchy"].apply(classify_goods)

# Date
sap["Date"] = pd.to_datetime(sap["Deliv.Date"], errors="coerce")

# Origin
shipping_master["ShPtKey"] = (
    shipping_master["ShPt"].astype(str).str.strip()
)
sap["ShPtKey"] = sap["ShPt"].astype(str).str.strip()
sap = sap.merge(
    shipping_master[["ShPtKey", "Description"]].rename(
        columns={"Description": "ShippingPointDescription"}
    ),
    on="ShPtKey",
    how="left",
)
sap["Origin"] = np.where(
    sap["ShPt"].isna(), "Import", sap["ShippingPointDescription"]
)

# Destination
customer_master["ShipToKey"] = (
    customer_master["Ship-To"].astype(str).str.strip()
)
sap["ShipToKey"] = sap["Ship-To"].astype(str).str.strip()
sap = sap.merge(
    customer_master[["ShipToKey", "Name 1"]].rename(
        columns={"Name 1": "CustomerName"}
    ),
    on="ShipToKey",
    how="left",
)

plant_master["PlantKey"] = plant_master["Plnt"].astype(str).str.strip()
sap["PlantKey"] = sap["Plnt"].astype(str).str.strip()
sap = sap.merge(
    plant_master[["PlantKey", "Name 1"]].rename(
        columns={"Name 1": "PlantName"}
    ),
    on="PlantKey",
    how="left",
)
sap["Destination"] = np.where(
    sap["Ship-To"].isna(), sap["PlantName"], sap["CustomerName"]
)

# Value and currency
sap["Value"] = pd.to_numeric(sap["Loc.curr.amount"], errors="coerce")
sap["Currency"] = (
    sap["Local Curr."].astype(str).apply(normalize_currency)
)
sap["Source"] = "SAP ERP"

sap_final = sap[
    [
        "Carrier",
        "Type of Goods",
        "Date",
        "Origin",
        "Destination",
        "Value",
        "Currency",
        "Source",
    ]
]


# ==========================================================
# EUR CONVERSION
# ==========================================================
currency_master.columns = [
    str(column).strip() for column in currency_master.columns
]
currency_master["Currency"] = (
    currency_master["Currency"].astype(str).str.strip().str.upper()
)

combined = pd.concat([manual, sap_final], ignore_index=True)
combined = combined.merge(currency_master, on="Currency", how="left")

rate_column = next(
    (column for column in currency_master.columns if "rate" in column.lower()),
    None,
)

if rate_column is None:
    st.error(
        "Currency conversion file must contain a column with 'Rate' in the name."
    )
    st.stop()

combined["EUR Rate"] = pd.to_numeric(
    combined[rate_column], errors="coerce"
)
combined["Value EUR"] = np.where(
    combined["Currency"] == "EUR",
    combined["Value"],
    combined["Value"] * combined["EUR Rate"],
)
combined["Year"] = combined["Date"].dt.year
combined["Month"] = combined["Date"].dt.strftime("%Y-%m")


# ==========================================================
# SIDEBAR FILTERS
# ==========================================================
st.sidebar.header("Filters")
carrier_filter = st.sidebar.multiselect(
    "Carrier", sorted(combined["Carrier"].dropna().unique())
)
goods_filter = st.sidebar.multiselect(
    "Type of Goods", sorted(combined["Type of Goods"].dropna().unique())
)
origin_filter = st.sidebar.multiselect(
    "Origin", sorted(combined["Origin"].dropna().unique())
)
destination_filter = st.sidebar.multiselect(
    "Destination", sorted(combined["Destination"].dropna().unique())
)
source_filter = st.sidebar.multiselect(
    "Source", sorted(combined["Source"].dropna().unique())
)

filtered = combined.copy()
if carrier_filter:
    filtered = filtered[filtered["Carrier"].isin(carrier_filter)]
if goods_filter:
    filtered = filtered[filtered["Type of Goods"].isin(goods_filter)]
if origin_filter:
    filtered = filtered[filtered["Origin"].isin(origin_filter)]
if destination_filter:
    filtered = filtered[filtered["Destination"].isin(destination_filter)]
if source_filter:
    filtered = filtered[filtered["Source"].isin(source_filter)]


# ==========================================================
# KPIs
# ==========================================================
k1, k2, k3, k4 = st.columns(4)
k1.metric("Spend EUR", f"€ {filtered['Value EUR'].sum():,.0f}")
k2.metric("Shipments", f"{len(filtered):,}")
k3.metric("Carriers", filtered["Carrier"].nunique())
k4.metric("Origins", filtered["Origin"].nunique())


# ==========================================================
# CHARTS
# ==========================================================
st.subheader("Monthly Spend")
monthly = filtered.groupby("Month", as_index=False)["Value EUR"].sum()
fig = px.line(monthly, x="Month", y="Value EUR", markers=True)
st.plotly_chart(fig, use_container_width=True)

col1, col2 = st.columns(2)

with col1:
    spend_by_carrier = (
        filtered.groupby("Carrier", as_index=False)["Value EUR"]
        .sum()
        .sort_values("Value EUR", ascending=False)
        .head(15)
    )
    fig = px.bar(
        spend_by_carrier,
        x="Carrier",
        y="Value EUR",
        title="Top Carriers",
    )
    st.plotly_chart(fig, use_container_width=True)

with col2:
    spend_by_goods = filtered.groupby(
        "Type of Goods", as_index=False
    )["Value EUR"].sum()
    fig = px.pie(
        spend_by_goods,
        names="Type of Goods",
        values="Value EUR",
        title="FG vs NFG",
    )
    st.plotly_chart(fig, use_container_width=True)

route_spend = (
    filtered.groupby(["Origin", "Destination"], as_index=False)["Value EUR"]
    .sum()
    .sort_values("Value EUR", ascending=False)
    .head(20)
)
fig = px.bar(
    route_spend,
    x="Origin",
    y="Value EUR",
    color="Destination",
    title="Top Freight Lanes",
)
st.plotly_chart(fig, use_container_width=True)


# ==========================================================
# DETAIL TABLE
# ==========================================================
st.subheader("Detailed Freight Cost Data")
st.dataframe(
    filtered.sort_values("Date", ascending=False),
    use_container_width=True,
    height=500,
)


# ==========================================================
# EXPORT
# ==========================================================
csv = filtered.to_csv(index=False).encode("utf-8")
st.download_button(
    label="Download filtered data as CSV",
    data=csv,
    file_name="freight_cost_data.csv",
    mime="text/csv",
)
