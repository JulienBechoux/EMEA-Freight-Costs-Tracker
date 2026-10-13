import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from pathlib import Path

st.set_page_config(
    page_title="Freight Cost Dashboard",
    page_icon="🚛",
    layout="wide"
)

# =====================================================================================
# CONFIG
# =====================================================================================

MASTER_DATA = Path("master_data")

CARRIER_FILE = MASTER_DATA / "ERP Carrier Name.xlsx"
CUSTOMERS_FILE = MASTER_DATA / "ERP Customers.xlsx"
PLANTS_FILE = MASTER_DATA / "ERP Plants.xlsx"
SHIPPING_FILE = MASTER_DATA / "ERP Shipping Point.xlsx"
CURRENCY_FILE = MASTER_DATA / "Currency conversion table.xlsx"

# =====================================================================================
# HELPERS
# =====================================================================================

@st.cache_data
def load_master_data():

    carrier = pd.read_excel(CARRIER_FILE)
    customers = pd.read_excel(CUSTOMERS_FILE)
    plants = pd.read_excel(PLANTS_FILE)
    shipping = pd.read_excel(SHIPPING_FILE)
    currency = pd.read_excel(CURRENCY_FILE)

    return carrier, customers, plants, shipping, currency


def clean_carrier_name(value):

    if pd.isna(value):
        return ""

    value = str(value)

    if "/" in value:
        value = value.split("/")[0]

    return value.strip()


def get_eur_rate(currency_df):

    cols = currency_df.columns.tolist()

    currency_col = cols[0]
    rate_col = cols[1]

    rates = (
        currency_df[[currency_col, rate_col]]
        .dropna()
        .rename(
            columns={
                currency_col: "Currency",
                rate_col: "RateToEUR"
            }
        )
    )

    rates["Currency"] = rates["Currency"].astype(str).str.upper()

    return rates


def convert_to_eur(df, rates):

    df["Currency"] = df["Currency"].astype(str).str.upper()

    merged = df.merge(
        rates,
        on="Currency",
        how="left"
    )

    merged["RateToEUR"] = merged["RateToEUR"].fillna(1)

    merged["Value_EUR"] = (
        merged["Value"] * merged["RateToEUR"]
    )

    return merged


# =====================================================================================
# MANUAL ACCRUALS
# =====================================================================================

def process_manual_accruals(file, currency_rates):

    df = pd.read_excel(file)

    result = pd.DataFrame()

    result["Carrier"] = df["Carrier Description"].apply(clean_carrier_name)

    result["Type of Goods"] = df["PBU"]

    result["Date"] = pd.to_datetime(
        df["Planned Arrival Date-Last Stop"],
        errors="coerce"
    )

    result["Origin"] = df["Source Location Description"]

    result["Destination"] = df["Destination Location Description"]

    result["Value"] = pd.to_numeric(
        df["Net Amt in Doc Crcy"],
        errors="coerce"
    )

    result["Currency"] = df["Currency"]

    result["Source"] = "Manual Accruals"

    result = convert_to_eur(result, currency_rates)

    return result


# =====================================================================================
# SAP ERP
# =====================================================================================

def process_sap_erp(
    file,
    carrier_map,
    customer_map,
    plant_map,
    shipping_map,
    currency_rates
):

    df = pd.read_excel(file)

    result = pd.DataFrame()

    # Carrier

    carrier_lookup = dict(
        zip(
            carrier_map.iloc[:, 0].astype(str),
            carrier_map["Name 1"]
        )
    )

    result["Carrier"] = (
        df["ServcAgent"]
        .astype(str)
        .map(carrier_lookup)
        .fillna(df["ServcAgent"].astype(str))
    )

    # Goods Type

    product = df["Product Hierarchy"].astype(str)

    result["Type of Goods"] = np.where(
        product.str.startswith(("1", "4")),
        "FG",
        "NFG"
    )

    # Date

    result["Date"] = pd.to_datetime(
        df["Deliv.Date"],
        errors="coerce"
    )

    # Origin

    ship_lookup = dict(
        zip(
            shipping_map.iloc[:, 0].astype(str),
            shipping_map["Description"]
        )
    )

    origin = (
        df["ShPt"]
        .astype(str)
        .map(ship_lookup)
    )

    result["Origin"] = origin.fillna("Import")

    # Destination

    customer_lookup = dict(
        zip(
            customer_map.iloc[:, 0].astype(str),
            customer_map["Name 1"]
        )
    )

    plant_lookup = dict(
        zip(
            plant_map.iloc[:, 0].astype(str),
            plant_map["Name 1"]
        )
    )

    destination = (
        df["Ship-To"]
        .astype(str)
        .map(customer_lookup)
    )

    plant_destination = (
        df["Plnt"]
        .astype(str)
        .map(plant_lookup)
    )

    result["Destination"] = np.where(
        df["Ship-To"].isna(),
        plant_destination,
        destination
    )

    # Cost

    result["Value"] = pd.to_numeric(
        df["Loc.curr.amount"],
        errors="coerce"
    )

    result["Currency"] = df["Local Curr."]

    result["Source"] = "SAP ERP"

    result = convert_to_eur(result, currency_rates)

    return result


# =====================================================================================
# UI
# =====================================================================================

st.title("🚛 Freight Cost Analytics Dashboard")

carrier_master,\
customer_master,\
plant_master,\
shipping_master,\
currency_master = load_master_data()

currency_rates = get_eur_rate(currency_master)

manual_file = st.sidebar.file_uploader(
    "Upload Manual Accruals",
    type=["xlsx"]
)

sap_file = st.sidebar.file_uploader(
    "Upload SAP ERP",
    type=["xlsx"]
)

if manual_file and sap_file:

    manual_df = process_manual_accruals(
        manual_file,
        currency_rates
    )

    sap_df = process_sap_erp(
        sap_file,
        carrier_master,
        customer_master,
        plant_master,
        shipping_master,
        currency_rates
    )

    data = pd.concat(
        [manual_df, sap_df],
        ignore_index=True
    )

    # FILTERS

    st.sidebar.header("Filters")

    carriers = st.sidebar.multiselect(
        "Carrier",
        sorted(data["Carrier"].dropna().unique())
    )

    goods = st.sidebar.multiselect(
        "Type of Goods",
        sorted(data["Type of Goods"].dropna().unique())
    )

    origins = st.sidebar.multiselect(
        "Origin",
        sorted(data["Origin"].dropna().unique())
    )

    destinations = st.sidebar.multiselect(
        "Destination",
        sorted(data["Destination"].dropna().unique())
    )

    source = st.sidebar.multiselect(
        "Source",
        sorted(data["Source"].dropna().unique())
    )

    filtered = data.copy()

    if carriers:
        filtered = filtered[
            filtered["Carrier"].isin(carriers)
        ]

    if goods:
        filtered = filtered[
            filtered["Type of Goods"].isin(goods)
        ]

    if origins:
        filtered = filtered[
            filtered["Origin"].isin(origins)
        ]

    if destinations:
        filtered = filtered[
            filtered["Destination"].isin(destinations)
        ]

    if source:
        filtered = filtered[
            filtered["Source"].isin(source)
        ]

    # KPIs

    k1, k2, k3, k4 = st.columns(4)

    k1.metric(
        "Total Freight Cost EUR",
        f"€ {filtered['Value_EUR'].sum():,.0f}"
    )

    k2.metric(
        "Shipments",
        f"{len(filtered):,}"
    )

    k3.metric(
        "Carriers",
        filtered["Carrier"].nunique()
    )

    k4.metric(
        "Origins",
        filtered["Origin"].nunique()
    )

    st.divider()

    # Charts

    col1, col2 = st.columns(2)

    with col1:

        carrier_cost = (
            filtered.groupby("Carrier")["Value_EUR"]
            .sum()
            .reset_index()
            .sort_values("Value_EUR", ascending=False)
            .head(20)
        )

        fig = px.bar(
            carrier_cost,
            x="Carrier",
            y="Value_EUR",
            title="Top Carriers"
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:

        goods_cost = (
            filtered.groupby("Type of Goods")
            ["Value_EUR"]
            .sum()
            .reset_index()
        )

        fig = px.pie(
            goods_cost,
            names="Type of Goods",
            values="Value_EUR",
            title="FG vs NFG"
        )

        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    monthly = (
        filtered
        .dropna(subset=["Date"])
        .assign(Month=lambda x: x["Date"].dt.to_period("M").astype(str))
        .groupby("Month")["Value_EUR"]
        .sum()
        .reset_index()
    )

    fig = px.line(
        monthly,
        x="Month",
        y="Value_EUR",
        markers=True,
        title="Monthly Freight Spend"
    )

    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    lane = (
        filtered.groupby(
            ["Origin", "Destination"]
        )["Value_EUR"]
        .sum()
        .reset_index()
        .sort_values(
            "Value_EUR",
            ascending=False
        )
        .head(50)
    )

    st.subheader("Top Freight Lanes")

    st.dataframe(
        lane,
        use_container_width=True
    )

    st.subheader("Detailed Transactions")

    st.dataframe(
        filtered,
        use_container_width=True
    )

    csv = filtered.to_csv(index=False)

    st.download_button(
        "Download Data",
        csv,
        "freight_costs.csv",
        "text/csv"
    )

else:

    st.info("Please upload both files to start the analysis.")
