from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# =============================================================================
# PAGE CONFIGURATION
# =============================================================================

st.set_page_config(
    page_title="Freight Cost Analytics",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =============================================================================
# REPOSITORY FILE LOCATIONS
# =============================================================================

# Resolve all static files relative to app.py. This avoids dependency on the
# current working directory and works locally and on repository deployments.
APP_DIR = Path(__file__).resolve().parent
MASTER_DATA_DIR = APP_DIR / "master_data"

MASTER_FILES = {
    "ERP Carrier Name": MASTER_DATA_DIR / "ERP Carrier Name.xlsx",
    "ERP Customers": MASTER_DATA_DIR / "ERP Customers.xlsx",
    "ERP Plants": MASTER_DATA_DIR / "ERP Plants.xlsx",
    "ERP Shipping Point": MASTER_DATA_DIR / "ERP Shipping Point.xlsx",
    "Currency conversion table": MASTER_DATA_DIR
    / "Currency conversion table.xlsx",
}

OUTPUT_COLUMNS = [
    "Carrier",
    "Type of Goods",
    "Date",
    "Origin",
    "Destination",
    "Value",
    "Currency",
    "EUR Rate",
    "Value EUR",
    "Source",
]


# =============================================================================
# EXPECTED SOURCE COLUMNS
# =============================================================================

MANUAL_REQUIRED_COLUMNS = [
    "Carrier Description",
    "PBU",
    "Planned Arrival Date-Last Stop",
    "Source Location Description",
    "Destination Location Descripti",
    "Net Amt in Doc Crcy",
    "Currency",
]

SAP_REQUIRED_COLUMNS = [
    "ServcAgent",
    "Product Hierarchy",
    "Deliv.Date",
    "ShPt",
    "Ship-To",
    "Plnt",
    "Loc.curr.amount",
    "Local Curr.",
]

MASTER_REQUIRED_COLUMNS = {
    "ERP Carrier Name": ["ServcAgent", "Name 1"],
    "ERP Customers": ["Ship-To", "Name 1"],
    "ERP Plants": ["Plnt", "Name 1"],
    "ERP Shipping Point": ["ShPt", "Description"],
    "Currency conversion table": ["Currency"],
}


# =============================================================================
# USER INTERFACE HEADER
# =============================================================================

st.title("🚚 Freight Cost Analytics Dashboard")
st.caption(
    "Consolidate Manual Accruals and SAP ERP freight costs, enrich the data "
    "with repository master files, and convert all spend to EUR."
)

with st.expander("Files used by this application", expanded=False):
    st.markdown(
        """
**Files uploaded by the user**

- Manual Accruals
- SAP ERP

**Static files loaded automatically from the repository**

- `master_data/Currency conversion table.xlsx`
- `master_data/ERP Carrier Name.xlsx`
- `master_data/ERP Customers.xlsx`
- `master_data/ERP Plants.xlsx`
- `master_data/ERP Shipping Point.xlsx`
"""
    )


# =============================================================================
# GENERIC HELPERS
# =============================================================================


def clean_column_names(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with leading and trailing spaces removed from headers."""
    cleaned = dataframe.copy()
    cleaned.columns = [str(column).strip() for column in cleaned.columns]
    return cleaned


@st.cache_data(show_spinner=False)
def read_repository_excel(file_path: str) -> pd.DataFrame:
    """Read a static Excel file stored in the application repository."""
    return clean_column_names(pd.read_excel(file_path, engine="openpyxl"))


@st.cache_data(show_spinner=False)
def read_uploaded_excel(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    """Read an uploaded XLSX or XLS file from its bytes."""
    suffix = Path(file_name).suffix.lower()
    engine = "xlrd" if suffix == ".xls" else "openpyxl"
    return clean_column_names(pd.read_excel(BytesIO(file_bytes), engine=engine))


def validate_columns(
    dataframe: pd.DataFrame,
    required_columns: Iterable[str],
    file_label: str,
) -> None:
    """Stop the app and show a clear message when columns are missing."""
    missing_columns = [
        column for column in required_columns if column not in dataframe.columns
    ]

    if missing_columns:
        st.error(f"{file_label} is missing required columns.")
        st.markdown("**Missing columns:**")
        st.code("\n".join(missing_columns), language="text")
        st.markdown("**Columns found in the file:**")
        st.code("\n".join(map(str, dataframe.columns)), language="text")
        st.stop()


def normalize_key(series: pd.Series) -> pd.Series:
    """
    Normalize IDs used in Excel lookups.

    This removes whitespace and Excel's common trailing .0 from numeric IDs
    while preserving missing values as empty strings.
    """
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


def missing_or_blank(series: pd.Series) -> pd.Series:
    """Return True for null, empty, or textual nan values."""
    normalized = series.fillna("").astype(str).str.strip()
    return normalized.eq("") | normalized.str.lower().isin(["nan", "none", "nat"])


def clean_text(series: pd.Series, fallback: str = "Unknown") -> pd.Series:
    """Trim text and replace missing or blank values with a fallback."""
    cleaned = series.fillna("").astype(str).str.strip()
    return cleaned.mask(missing_or_blank(cleaned), fallback)


def clean_manual_carrier(series: pd.Series) -> pd.Series:
    """Keep text before the first slash in Manual Accruals carriers."""
    cleaned = series.fillna("").astype(str).str.split("/", n=1).str[0].str.strip()
    return cleaned.mask(missing_or_blank(cleaned), "Unknown Carrier")


def classify_goods(series: pd.Series) -> pd.Series:
    """Classify Product Hierarchy beginning with 1 or 4 as FG, else NFG."""
    hierarchy = normalize_key(series)
    return pd.Series(
        np.where(hierarchy.str.startswith(("1", "4")), "FG", "NFG"),
        index=series.index,
    )


def normalize_currency(series: pd.Series) -> pd.Series:
    """Normalize ISO-style currency codes."""
    normalized = series.fillna("").astype(str).str.strip().str.upper()
    return normalized.mask(missing_or_blank(normalized), pd.NA)


def parse_amount(series: pd.Series) -> pd.Series:
    """
    Parse numeric amounts while supporting common Excel text formats.

    Native numeric cells are preserved. Text values containing spaces or
    thousands separators such as 1,234.56 are also supported.
    """
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    cleaned = (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
    )

    # Parentheses are treated as negative amounts.
    negative_mask = cleaned.str.match(r"^\(.*\)$")
    cleaned = cleaned.str.replace("(", "", regex=False).str.replace(
        ")", "", regex=False
    )

    # Standard format expected by pandas: commas as grouping and dot decimal.
    cleaned = cleaned.str.replace(",", "", regex=False)
    result = pd.to_numeric(cleaned, errors="coerce")
    result.loc[negative_mask & result.notna()] *= -1
    return result


def make_lookup(
    dataframe: pd.DataFrame,
    key_column: str,
    value_column: str,
    output_value_column: str,
) -> pd.DataFrame:
    """Create a clean, unique two-column lookup table."""
    lookup = dataframe[[key_column, value_column]].copy()
    lookup["_LookupKey"] = normalize_key(lookup[key_column])
    lookup[output_value_column] = clean_text(lookup[value_column], fallback="")
    lookup = lookup.loc[lookup["_LookupKey"].ne("")]
    lookup = lookup[["_LookupKey", output_value_column]].drop_duplicates(
        subset=["_LookupKey"], keep="first"
    )
    return lookup


def find_rate_column(currency_master: pd.DataFrame) -> str:
    """Find the EUR conversion-rate column in the currency master."""
    columns = list(currency_master.columns)
    preferred_names = [
        "EUR Rate",
        "Conversion Rate",
        "Conversion rate",
        "Rate",
        "Exchange Rate",
        "Exchange rate",
    ]

    for preferred in preferred_names:
        if preferred in columns:
            return preferred

    candidate_columns = [
        column
        for column in columns
        if column != "Currency"
        and any(
            token in str(column).lower()
            for token in ("rate", "conversion", "exchange", "eur")
        )
    ]

    if len(candidate_columns) == 1:
        return candidate_columns[0]

    st.error(
        "The Currency conversion table must contain one conversion-rate "
        "column, preferably named 'Rate' or 'EUR Rate'."
    )
    st.markdown("**Columns found:**")
    st.code("\n".join(map(str, columns)), language="text")
    st.stop()


def format_eur(value: float) -> str:
    """Format an amount as EUR for KPI cards."""
    return f"€{value:,.2f}"


def safe_sorted_values(series: pd.Series) -> list:
    """Produce stable sorted filter options."""
    return sorted(series.dropna().unique().tolist(), key=lambda item: str(item))


# =============================================================================
# LOAD AND VALIDATE STATIC REPOSITORY FILES
# =============================================================================


def load_master_files() -> dict[str, pd.DataFrame]:
    """Validate presence and content of all repository master files."""
    missing_files = [path for path in MASTER_FILES.values() if not path.is_file()]

    if missing_files:
        st.error(
            "One or more required static master files are missing from the "
            "repository."
        )
        st.markdown("**Expected missing paths, relative to `app.py`:**")
        st.code(
            "\n".join(
                str(path.relative_to(APP_DIR))
                if path.is_relative_to(APP_DIR)
                else str(path)
                for path in missing_files
            ),
            language="text",
        )
        st.info(
            "Add these files to the master_data folder, commit and push the "
            "folder to the repository, and redeploy the application."
        )
        st.stop()

    loaded: dict[str, pd.DataFrame] = {}

    for label, path in MASTER_FILES.items():
        try:
            dataframe = read_repository_excel(str(path))
        except Exception as exc:
            st.error(f"Unable to read repository master file: {label}")
            st.code(str(path), language="text")
            st.exception(exc)
            st.stop()

        validate_columns(dataframe, MASTER_REQUIRED_COLUMNS[label], label)
        loaded[label] = dataframe

    return loaded


with st.spinner("Loading repository master data..."):
    masters = load_master_files()

carrier_master = masters["ERP Carrier Name"]
customer_master = masters["ERP Customers"]
plant_master = masters["ERP Plants"]
shipping_master = masters["ERP Shipping Point"]
currency_master = masters["Currency conversion table"]


# =============================================================================
# FILE UPLOADS
# =============================================================================

st.subheader("1. Upload period data")
upload_left, upload_right = st.columns(2)

with upload_left:
    manual_file = st.file_uploader(
        "Manual Accruals",
        type=["xlsx", "xls"],
        key="manual_accruals",
        help="Upload the Manual Accruals file for the period to analyze.",
    )

with upload_right:
    sap_file = st.file_uploader(
        "SAP ERP",
        type=["xlsx", "xls"],
        key="sap_erp",
        help="Upload the SAP ERP freight cost file for the period to analyze.",
    )

if manual_file is None or sap_file is None:
    st.info(
        "Upload both the Manual Accruals file and the SAP ERP file to build "
        "the dashboard. The five master files are loaded automatically from "
        "the repository."
    )
    st.stop()


# =============================================================================
# READ AND VALIDATE USER FILES
# =============================================================================

try:
    manual_raw = read_uploaded_excel(manual_file.getvalue(), manual_file.name)
except Exception as exc:
    st.error("Unable to read the Manual Accruals file.")
    st.exception(exc)
    st.stop()

try:
    sap_raw = read_uploaded_excel(sap_file.getvalue(), sap_file.name)
except Exception as exc:
    st.error("Unable to read the SAP ERP file.")
    st.exception(exc)
    st.stop()

validate_columns(manual_raw, MANUAL_REQUIRED_COLUMNS, "Manual Accruals")
validate_columns(sap_raw, SAP_REQUIRED_COLUMNS, "SAP ERP")


# =============================================================================
# MANUAL ACCRUALS TRANSFORMATION
# =============================================================================

manual = pd.DataFrame(index=manual_raw.index)
manual["Carrier"] = clean_manual_carrier(manual_raw["Carrier Description"])
manual["Type of Goods"] = clean_text(manual_raw["PBU"])
manual["Date"] = pd.to_datetime(
    manual_raw["Planned Arrival Date-Last Stop"], errors="coerce"
)
manual["Origin"] = clean_text(
    manual_raw["Source Location Description"], fallback="Unknown Origin"
)
manual["Destination"] = clean_text(
    manual_raw["Destination Location Descripti"],
    fallback="Unknown Destination",
)
manual["Value"] = parse_amount(manual_raw["Net Amt in Doc Crcy"])
manual["Currency"] = normalize_currency(manual_raw["Currency"])
manual["Source"] = "Manual Accruals"


# =============================================================================
# SAP ERP TRANSFORMATION AND MASTER-DATA ENRICHMENT
# =============================================================================

sap = sap_raw.copy()

# Carrier: SAP ServcAgent -> ERP Carrier Name Name 1
carrier_lookup = make_lookup(
    carrier_master,
    key_column="ServcAgent",
    value_column="Name 1",
    output_value_column="_CarrierName",
)
sap["_CarrierKey"] = normalize_key(sap["ServcAgent"])
sap = sap.merge(
    carrier_lookup,
    left_on="_CarrierKey",
    right_on="_LookupKey",
    how="left",
)
sap["Carrier"] = clean_text(sap["_CarrierName"], fallback="Unknown Carrier")

# Type of goods: Product Hierarchy beginning with 1 or 4 -> FG, otherwise NFG
sap["Type of Goods"] = classify_goods(sap["Product Hierarchy"])

# Date
sap["Date"] = pd.to_datetime(sap["Deliv.Date"], errors="coerce")

# Origin: SAP ShPt -> ERP Shipping Point Description; blank ShPt -> Import
shipping_lookup = make_lookup(
    shipping_master,
    key_column="ShPt",
    value_column="Description",
    output_value_column="_ShippingPointDescription",
)
sap["_ShPtKey"] = normalize_key(sap["ShPt"])
sap = sap.merge(
    shipping_lookup,
    left_on="_ShPtKey",
    right_on="_LookupKey",
    how="left",
    suffixes=("", "_shipping"),
)
blank_shipping_point = missing_or_blank(sap["ShPt"])
sap["Origin"] = clean_text(
    sap["_ShippingPointDescription"], fallback="Unmapped Shipping Point"
)
sap.loc[blank_shipping_point, "Origin"] = "Import"

# Destination primary rule: SAP Ship-To -> ERP Customers Name 1
customer_lookup = make_lookup(
    customer_master,
    key_column="Ship-To",
    value_column="Name 1",
    output_value_column="_CustomerName",
)
sap["_ShipToKey"] = normalize_key(sap["Ship-To"])
sap = sap.merge(
    customer_lookup,
    left_on="_ShipToKey",
    right_on="_LookupKey",
    how="left",
    suffixes=("", "_customer"),
)

# Destination fallback rule: if Ship-To is blank, SAP Plnt -> ERP Plants Name 1
plant_lookup = make_lookup(
    plant_master,
    key_column="Plnt",
    value_column="Name 1",
    output_value_column="_PlantName",
)
sap["_PlantKey"] = normalize_key(sap["Plnt"])
sap = sap.merge(
    plant_lookup,
    left_on="_PlantKey",
    right_on="_LookupKey",
    how="left",
    suffixes=("", "_plant"),
)
blank_ship_to = missing_or_blank(sap["Ship-To"])
customer_destination = clean_text(
    sap["_CustomerName"], fallback="Unmapped Customer"
)
plant_destination = clean_text(sap["_PlantName"], fallback="Unmapped Plant")
sap["Destination"] = np.where(
    blank_ship_to,
    plant_destination,
    customer_destination,
)

# Value and currency
sap["Value"] = parse_amount(sap["Loc.curr.amount"])
sap["Currency"] = normalize_currency(sap["Local Curr."])
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
].copy()


# =============================================================================
# CONSOLIDATE BOTH SOURCES
# =============================================================================

combined = pd.concat([manual, sap_final], ignore_index=True)
combined["Date"] = pd.to_datetime(combined["Date"], errors="coerce")
combined["Value"] = pd.to_numeric(combined["Value"], errors="coerce")
combined["Currency"] = normalize_currency(combined["Currency"])


# =============================================================================
# CURRENCY CONVERSION TO EUR
# =============================================================================

currency_master = currency_master.copy()
rate_column = find_rate_column(currency_master)
currency_master["_CurrencyKey"] = normalize_currency(currency_master["Currency"])
currency_master["_EURRate"] = parse_amount(currency_master[rate_column])

currency_lookup = (
    currency_master[["_CurrencyKey", "_EURRate"]]
    .dropna(subset=["_CurrencyKey"])
    .drop_duplicates(subset=["_CurrencyKey"], keep="first")
)

combined = combined.merge(
    currency_lookup,
    left_on="Currency",
    right_on="_CurrencyKey",
    how="left",
)

# EUR is always 1 even if EUR is not included in the conversion file.
combined["EUR Rate"] = combined["_EURRate"]
combined.loc[combined["Currency"].eq("EUR"), "EUR Rate"] = 1.0
combined["Value EUR"] = combined["Value"] * combined["EUR Rate"]

combined["Year"] = combined["Date"].dt.year.astype("Int64")
combined["Month"] = combined["Date"].dt.to_period("M").astype("string")
combined["Route"] = combined["Origin"] + " → " + combined["Destination"]


# =============================================================================
# DATA QUALITY SUMMARY
# =============================================================================

invalid_amount_count = int(combined["Value"].isna().sum())
invalid_date_count = int(combined["Date"].isna().sum())
missing_currency_count = int(combined["Currency"].isna().sum())
unmapped_rate_count = int(
    (combined["Currency"].notna() & combined["EUR Rate"].isna()).sum()
)

unmapped_currencies = safe_sorted_values(
    combined.loc[
        combined["Currency"].notna() & combined["EUR Rate"].isna(),
        "Currency",
    ]
)

with st.expander("Data quality checks", expanded=unmapped_rate_count > 0):
    quality_col1, quality_col2, quality_col3, quality_col4 = st.columns(4)
    quality_col1.metric("Invalid or blank values", f"{invalid_amount_count:,}")
    quality_col2.metric("Invalid or blank dates", f"{invalid_date_count:,}")
    quality_col3.metric("Missing currencies", f"{missing_currency_count:,}")
    quality_col4.metric("Missing EUR rates", f"{unmapped_rate_count:,}")

    if unmapped_currencies:
        st.warning(
            "No EUR conversion rate was found for: "
            + ", ".join(map(str, unmapped_currencies))
            + ". These rows remain visible but are excluded from EUR spend totals."
        )

    st.caption(
        "Unmapped SAP lookups are labeled as Unknown Carrier, Unmapped "
        "Shipping Point, Unmapped Customer, or Unmapped Plant."
    )


# =============================================================================
# FILTERS
# =============================================================================

st.sidebar.header("Filters")

valid_dates = combined["Date"].dropna()
if not valid_dates.empty:
    minimum_date = valid_dates.min().date()
    maximum_date = valid_dates.max().date()
    selected_dates = st.sidebar.date_input(
        "Date range",
        value=(minimum_date, maximum_date),
        min_value=minimum_date,
        max_value=maximum_date,
    )
else:
    selected_dates = ()
    st.sidebar.warning("No valid dates were found in the uploaded files.")

selected_carriers = st.sidebar.multiselect(
    "Carrier",
    options=safe_sorted_values(combined["Carrier"]),
)
selected_goods = st.sidebar.multiselect(
    "Type of Goods",
    options=safe_sorted_values(combined["Type of Goods"]),
)
selected_origins = st.sidebar.multiselect(
    "Origin",
    options=safe_sorted_values(combined["Origin"]),
)
selected_destinations = st.sidebar.multiselect(
    "Destination",
    options=safe_sorted_values(combined["Destination"]),
)
selected_currencies = st.sidebar.multiselect(
    "Original Currency",
    options=safe_sorted_values(combined["Currency"]),
)
selected_sources = st.sidebar.multiselect(
    "Source",
    options=safe_sorted_values(combined["Source"]),
)

filtered = combined.copy()

if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
    start_date = pd.Timestamp(selected_dates[0])
    end_date = pd.Timestamp(selected_dates[1])
    filtered = filtered.loc[filtered["Date"].between(start_date, end_date)]

if selected_carriers:
    filtered = filtered.loc[filtered["Carrier"].isin(selected_carriers)]
if selected_goods:
    filtered = filtered.loc[filtered["Type of Goods"].isin(selected_goods)]
if selected_origins:
    filtered = filtered.loc[filtered["Origin"].isin(selected_origins)]
if selected_destinations:
    filtered = filtered.loc[
        filtered["Destination"].isin(selected_destinations)
    ]
if selected_currencies:
    filtered = filtered.loc[filtered["Currency"].isin(selected_currencies)]
if selected_sources:
    filtered = filtered.loc[filtered["Source"].isin(selected_sources)]


# =============================================================================
# DASHBOARD
# =============================================================================

st.subheader("2. Freight spend overview")

if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

spend_eur = filtered["Value EUR"].sum(min_count=1)
spend_eur = 0.0 if pd.isna(spend_eur) else float(spend_eur)

kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
kpi1.metric("Freight Spend", format_eur(spend_eur))
kpi2.metric("Cost Records", f"{len(filtered):,}")
kpi3.metric("Carriers", f"{filtered['Carrier'].nunique():,}")
kpi4.metric("Routes", f"{filtered['Route'].nunique():,}")
kpi5.metric("Destinations", f"{filtered['Destination'].nunique():,}")

chart_left, chart_right = st.columns((2, 1))

with chart_left:
    monthly_spend = (
        filtered.dropna(subset=["Date", "Value EUR"])
        .assign(MonthStart=lambda frame: frame["Date"].dt.to_period("M").dt.to_timestamp())
        .groupby("MonthStart", as_index=False)["Value EUR"]
        .sum()
        .sort_values("MonthStart")
    )

    if monthly_spend.empty:
        st.info("No valid dated EUR values are available for the monthly trend.")
    else:
        figure_monthly = px.line(
            monthly_spend,
            x="MonthStart",
            y="Value EUR",
            markers=True,
            title="Monthly Freight Spend",
            labels={"MonthStart": "Month", "Value EUR": "Spend in EUR"},
        )
        figure_monthly.update_layout(hovermode="x unified")
        st.plotly_chart(figure_monthly, use_container_width=True)

with chart_right:
    source_spend = (
        filtered.dropna(subset=["Value EUR"])
        .groupby("Source", as_index=False)["Value EUR"]
        .sum()
    )

    if source_spend.empty:
        st.info("No converted values are available for source analysis.")
    else:
        figure_source = px.pie(
            source_spend,
            names="Source",
            values="Value EUR",
            hole=0.45,
            title="Spend by Source",
        )
        st.plotly_chart(figure_source, use_container_width=True)

analysis_left, analysis_right = st.columns(2)

with analysis_left:
    carrier_spend = (
        filtered.dropna(subset=["Value EUR"])
        .groupby("Carrier", as_index=False)["Value EUR"]
        .sum()
        .sort_values("Value EUR", ascending=False)
        .head(15)
        .sort_values("Value EUR")
    )

    if carrier_spend.empty:
        st.info("No converted values are available for carrier analysis.")
    else:
        figure_carrier = px.bar(
            carrier_spend,
            x="Value EUR",
            y="Carrier",
            orientation="h",
            title="Top 15 Carriers",
            labels={"Value EUR": "Spend in EUR"},
        )
        st.plotly_chart(figure_carrier, use_container_width=True)

with analysis_right:
    goods_spend = (
        filtered.dropna(subset=["Value EUR"])
        .groupby("Type of Goods", as_index=False)["Value EUR"]
        .sum()
        .sort_values("Value EUR", ascending=False)
    )

    if goods_spend.empty:
        st.info("No converted values are available for goods analysis.")
    else:
        figure_goods = px.bar(
            goods_spend,
            x="Type of Goods",
            y="Value EUR",
            color="Type of Goods",
            title="Spend by Type of Goods",
            labels={"Value EUR": "Spend in EUR"},
        )
        figure_goods.update_layout(showlegend=False)
        st.plotly_chart(figure_goods, use_container_width=True)

lane_spend = (
    filtered.dropna(subset=["Value EUR"])
    .groupby(["Origin", "Destination", "Route"], as_index=False)["Value EUR"]
    .sum()
    .sort_values("Value EUR", ascending=False)
    .head(20)
    .sort_values("Value EUR")
)

if lane_spend.empty:
    st.info("No converted values are available for lane analysis.")
else:
    figure_lanes = px.bar(
        lane_spend,
        x="Value EUR",
        y="Route",
        orientation="h",
        color="Origin",
        title="Top 20 Freight Lanes",
        labels={"Value EUR": "Spend in EUR"},
        hover_data=["Origin", "Destination"],
    )
    st.plotly_chart(figure_lanes, use_container_width=True)


# =============================================================================
# DETAILED DATA AND EXPORT
# =============================================================================

st.subheader("3. Detailed freight cost data")

export_data = filtered[OUTPUT_COLUMNS].copy()
export_data = export_data.sort_values("Date", ascending=False, na_position="last")

st.dataframe(
    export_data,
    use_container_width=True,
    height=550,
    hide_index=True,
    column_config={
        "Date": st.column_config.DateColumn("Date", format="YYYY-MM-DD"),
        "Value": st.column_config.NumberColumn("Original Value", format="%.2f"),
        "EUR Rate": st.column_config.NumberColumn("EUR Rate", format="%.6f"),
        "Value EUR": st.column_config.NumberColumn("Value EUR", format="€ %.2f"),
    },
)

csv_data = export_data.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    label="⬇️ Download filtered data as CSV",
    data=csv_data,
    file_name="freight_costs_filtered.csv",
    mime="text/csv",
    use_container_width=False,
)

st.caption(
    "Freight spend totals include only records with a valid amount, currency, "
    "and EUR conversion rate. Rows with data-quality issues remain available "
    "in the detailed table for review."
)
