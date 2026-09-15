from __future__ import annotations

import io
import re
from dataclasses import dataclass
from datetime import date
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Actual Freight Cost Dashboard", page_icon="🚚", layout="wide")

STANDARD_COLUMNS = [
    "Carrier", "Type of goods", "Date", "Origin", "Destination",
    "Value", "Currency", "Value EUR", "Source", "Freight Document"
]

ALIASES = {
    "freight_document": ["Freight Document", "Freight Document Number", "Freight Doc.", "Freight Doc", "Freight Order", "Freight Booking"],
    "conversion_currency": ["Currency", "From Currency", "Source Currency", "Curr.", "Local Curr."],
    "conversion_rate": ["Conversion Rate", "Rate", "Rate to EUR", "EUR Rate", "Exchange Rate"],
}

@dataclass(frozen=True)
class UploadSpec:
    label: str
    key: str
    help: str

UPLOADS = [
    UploadSpec("Manual accruals", "manual", "Manual accrual cost extract"),
    UploadSpec("SAP ERP", "erp", "SAP ERP freight cost extract"),
    UploadSpec("SAP TM", "tm", "SAP TM freight cost extract"),
    UploadSpec("ERP Carrier Name", "carrier", "ServcAgent to Name 1 mapping"),
    UploadSpec("ERP Shipping Point", "shipping", "ShPt to Description mapping"),
    UploadSpec("ERP Customers", "customer", "Ship-To to Name 1 mapping"),
    UploadSpec("ERP Plants", "plant", "Plnt to Name 1 mapping"),
    UploadSpec("TM FO", "tm_fo", "TM details for freight documents starting with 68"),
    UploadSpec("TM FB", "tm_fb", "TM details for freight documents starting with 69"),
    UploadSpec("Currency conversion table", "fx", "Currency-to-EUR conversion rates"),
]


def norm_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def clean_key(series: pd.Series) -> pd.Series:
    def one(value: object) -> object:
        if pd.isna(value):
            return pd.NA
        text = str(value).strip()
        text = re.sub(r"\.0$", "", text)
        return text if text else pd.NA
    return series.map(one).astype("string")


def find_col(df: pd.DataFrame, names: Sequence[str], required: bool = True) -> Optional[str]:
    by_norm = {norm_header(c): c for c in df.columns}
    for name in names:
        if norm_header(name) in by_norm:
            return by_norm[norm_header(name)]
    if required:
        raise KeyError(f"Missing column. Expected one of: {', '.join(names)}")
    return None


def read_uploaded(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    raw = uploaded.getvalue()
    if name.endswith((".xlsx", ".xlsm")):
        return pd.read_excel(io.BytesIO(raw), engine="openpyxl", dtype=object)
    if name.endswith(".xls"):
        return pd.read_excel(io.BytesIO(raw), engine="xlrd", dtype=object)
    if name.endswith(".csv"):
        for encoding in ("utf-8-sig", "utf-8", "cp1252"):
            try:
                return pd.read_csv(io.BytesIO(raw), sep=None, engine="python", encoding=encoding, dtype=object)
            except UnicodeDecodeError:
                continue
        raise ValueError("CSV encoding could not be detected.")
    raise ValueError("Supported formats are .xlsx, .xlsm, .xls and .csv")


def parse_amount(series: pd.Series) -> pd.Series:
    def one(value: object) -> float:
        if pd.isna(value) or str(value).strip() == "":
            return float("nan")
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return float(value)
        text = str(value).strip().replace("\u00a0", "").replace(" ", "")
        negative = text.startswith("(") and text.endswith(")")
        text = text.strip("()")
        text = re.sub(r"[^0-9,\.\-]", "", text)
        if "," in text and "." in text:
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", "")
        elif "," in text:
            lhs, rhs = text.rsplit(",", 1)
            text = lhs.replace(",", "") + ("." + rhs if len(rhs) <= 2 else rhs)
        number = float(text)
        return -number if negative else number
    return series.map(one)


def parse_dates(series: pd.Series) -> pd.Series:
    # Handles native Excel dates, common text dates and Excel serial dates.
    numeric = pd.to_numeric(series, errors="coerce")
    parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)
    serial_mask = parsed.isna() & numeric.between(1, 100000)
    parsed.loc[serial_mask] = pd.to_datetime(numeric.loc[serial_mask], unit="D", origin="1899-12-30", errors="coerce")
    return parsed


def clean_dimension(series: pd.Series, fallback: str = "Unknown") -> pd.Series:
    return series.astype("string").str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA}).fillna(fallback)


def build_lookup(df: pd.DataFrame, key_names: Sequence[str], value_names: Sequence[str]) -> Dict[str, str]:
    key_col = find_col(df, key_names)
    value_col = find_col(df, value_names)
    temp = pd.DataFrame({"key": clean_key(df[key_col]), "value": clean_dimension(df[value_col], "")})
    temp = temp.dropna(subset=["key"]).drop_duplicates("key", keep="last")
    return dict(zip(temp["key"], temp["value"]))


def base_output(df: pd.DataFrame, source: str) -> pd.DataFrame:
    out = df.copy()
    out["Source"] = source
    for col in STANDARD_COLUMNS:
        if col not in out:
            out[col] = pd.NA
    return out[STANDARD_COLUMNS]


def transform_manual(df: pd.DataFrame) -> pd.DataFrame:
    carrier = find_col(df, ["Carrier Description"])
    goods = find_col(df, ["PBU"])
    dt = find_col(df, ["Planned Arrival Date-Last Stop", "Planned Arrival Date Last Stop"])
    origin = find_col(df, ["Source Location Description"])
    destination = find_col(df, ["Destination Location Descripti", "Destination Location Description"])
    value = find_col(df, ["Net Amt in Doc Crcy"])
    currency = find_col(df, ["Currency"])
    out = pd.DataFrame({
        "Carrier": df[carrier].astype("string").str.split("/", n=1).str[0].str.strip(),
        "Type of goods": df[goods],
        "Date": parse_dates(df[dt]),
        "Origin": df[origin],
        "Destination": df[destination],
        "Value": parse_amount(df[value]),
        "Currency": df[currency],
    })
    return base_output(out, "Manual accruals")


def transform_erp(df: pd.DataFrame, carrier_df: pd.DataFrame, shipping_df: pd.DataFrame,
                  customer_df: pd.DataFrame, plant_df: pd.DataFrame) -> pd.DataFrame:
    service = find_col(df, ["ServcAgent", "Service Agent"])
    goods = find_col(df, ["Product Hierarchy"])
    dt = find_col(df, ["Deliv.Date", "Delivery Date"])
    shpt = find_col(df, ["ShPt", "Shipping Point"])
    ship_to = find_col(df, ["Ship-To", "Ship To"])
    plant = find_col(df, ["Plnt", "Plant"])
    value = find_col(df, ["Loc.curr.amount", "Local Currency Amount"])
    currency = find_col(df, ["Local Curr.", "Local Currency"])

    carrier_map = build_lookup(carrier_df, ["ServcAgent", "Service Agent", "Number", "Vendor"], ["Name 1"])
    shipping_map = build_lookup(shipping_df, ["ShPt", "Shipping Point", "Code"], ["Description"])
    customer_map = build_lookup(customer_df, ["Ship-To", "Ship To", "Customer", "Number"], ["Name 1"])
    plant_map = build_lookup(plant_df, ["Plnt", "Plant", "Code"], ["Name 1"])

    service_key = clean_key(df[service])
    shpt_key = clean_key(df[shpt])
    ship_to_key = clean_key(df[ship_to])
    plant_key = clean_key(df[plant])
    hierarchy = clean_key(df[goods]).fillna("")

    mapped_carrier = service_key.map(carrier_map).fillna(service_key)
    mapped_origin = shpt_key.map(shipping_map)
    mapped_origin = mapped_origin.where(shpt_key.notna(), "Import").fillna(shpt_key)
    mapped_destination = ship_to_key.map(customer_map)
    plant_destination = plant_key.map(plant_map).fillna(plant_key)
    mapped_destination = mapped_destination.where(ship_to_key.notna(), plant_destination).fillna(ship_to_key)

    out = pd.DataFrame({
        "Carrier": mapped_carrier,
        "Type of goods": hierarchy.str.startswith(("1", "4")).map({True: "FG", False: "NFG"}),
        "Date": parse_dates(df[dt]),
        "Origin": mapped_origin,
        "Destination": mapped_destination,
        "Value": parse_amount(df[value]),
        "Currency": df[currency],
    })
    return base_output(out, "SAP ERP")


def tm_detail_lookup(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    doc_col = find_col(df, ALIASES["freight_document"])
    if kind == "FO":
        actual = find_col(df, ["Actual Delivered Date"], required=False)
        planned = find_col(df, ["Planned Arrival Date-Last Stop", "Planned Arrival Date Last Stop"])
        actual_dates = parse_dates(df[actual]) if actual else pd.Series(pd.NaT, index=df.index)
        chosen_date = actual_dates.fillna(parse_dates(df[planned]))
    else:
        expected = find_col(df, ["Expected Arrival Date"])
        chosen_date = parse_dates(df[expected])
    origin = find_col(df, ["Source Location Description"])
    destination = find_col(df, ["Destination Location Descripti", "Destination Location Description"])
    detail = pd.DataFrame({
        "Freight Document": clean_key(df[doc_col]),
        "TM Date": chosen_date,
        "TM Origin": df[origin],
        "TM Destination": df[destination],
    })
    return detail.dropna(subset=["Freight Document"]).drop_duplicates("Freight Document", keep="last")


def transform_tm(df: pd.DataFrame, fo_df: pd.DataFrame, fb_df: pd.DataFrame) -> pd.DataFrame:
    doc = find_col(df, ALIASES["freight_document"])
    carrier = find_col(df, ["Invoicing Party"])
    value = find_col(df, ["Net Amt in Doc Crcy"])
    currency = find_col(df, ["Currency"])

    cost = pd.DataFrame({
        "Freight Document": clean_key(df[doc]),
        "Carrier": df[carrier],
        "Value": parse_amount(df[value]),
        "Currency": df[currency],
    })
    details = pd.concat([tm_detail_lookup(fo_df, "FO"), tm_detail_lookup(fb_df, "FB")], ignore_index=True)
    details = details.drop_duplicates("Freight Document", keep="last")
    merged = cost.merge(details, on="Freight Document", how="left", validate="m:1")
    prefixes = merged["Freight Document"].fillna("").str[:2]
    merged.loc[~prefixes.isin(["68", "69"]), ["TM Date", "TM Origin", "TM Destination"]] = pd.NA
    out = pd.DataFrame({
        "Carrier": merged["Carrier"],
        "Type of goods": "Unknown",
        "Date": merged["TM Date"],
        "Origin": merged["TM Origin"],
        "Destination": merged["TM Destination"],
        "Value": merged["Value"],
        "Currency": merged["Currency"],
        "Freight Document": merged["Freight Document"],
    })
    return base_output(out, "SAP TM")


def build_fx(df: pd.DataFrame) -> Tuple[Dict[str, float], str]:
    currency_col = find_col(df, ALIASES["conversion_currency"])
    rate_col = find_col(df, ALIASES["conversion_rate"])
    currency = df[currency_col].astype("string").str.upper().str.strip()
    rate = parse_amount(df[rate_col])
    temp = pd.DataFrame({"Currency": currency, "Rate": rate}).dropna()
    temp = temp.drop_duplicates("Currency", keep="last")
    rates = dict(zip(temp["Currency"], temp["Rate"]))
    rates["EUR"] = 1.0
    return rates, rate_col


def add_eur(data: pd.DataFrame, fx_df: pd.DataFrame, convention: str) -> Tuple[pd.DataFrame, List[str]]:
    rates, _ = build_fx(fx_df)
    out = data.copy()
    out["Currency"] = out["Currency"].astype("string").str.upper().str.strip()
    out["FX Rate"] = out["Currency"].map(rates)
    if convention == "1 unit of currency = rate EUR":
        out["Value EUR"] = out["Value"] * out["FX Rate"]
    else:
        out["Value EUR"] = out["Value"] / out["FX Rate"]
    missing = sorted(out.loc[out["FX Rate"].isna() & out["Currency"].notna(), "Currency"].dropna().unique().tolist())
    return out, missing


def style_clean(data: pd.DataFrame) -> pd.DataFrame:
    out = data.copy()
    for col in ["Carrier", "Type of goods", "Origin", "Destination", "Currency", "Source"]:
        out[col] = clean_dimension(out[col])
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    return out


def multiselect_filter(label: str, values: pd.Series, key: str) -> List[str]:
    options = sorted(values.dropna().astype(str).unique().tolist())
    return st.sidebar.multiselect(label, options, default=[], key=key, placeholder="All")


def apply_filters(data: pd.DataFrame) -> pd.DataFrame:
    filtered = data.copy()
    st.sidebar.header("Filters")
    valid_dates = filtered["Date"].dropna()
    if not valid_dates.empty:
        min_date, max_date = valid_dates.min().date(), valid_dates.max().date()
        chosen = st.sidebar.date_input("Period", value=(min_date, max_date), min_value=min_date, max_value=max_date)
        if isinstance(chosen, (tuple, list)) and len(chosen) == 2:
            filtered = filtered[filtered["Date"].dt.date.between(chosen[0], chosen[1])]
    for label, col, key in [
        ("Source", "Source", "f_source"), ("Carrier", "Carrier", "f_carrier"),
        ("Type of goods", "Type of goods", "f_goods"), ("Origin", "Origin", "f_origin"),
        ("Destination", "Destination", "f_destination"), ("Original currency", "Currency", "f_currency")
    ]:
        selected = multiselect_filter(label, filtered[col], key)
        if selected:
            filtered = filtered[filtered[col].astype(str).isin(selected)]
    return filtered


st.title("🚚 Actual Freight Cost Dashboard")
st.caption("Consolidates Manual accruals, SAP ERP and SAP TM; maps master data; converts values to EUR; and provides interactive analysis.")

with st.sidebar:
    st.header("1. Upload files")
    files = {}
    for spec in UPLOADS:
        files[spec.key] = st.file_uploader(spec.label, type=["xlsx", "xlsm", "xls", "csv"], key=spec.key, help=spec.help)
    st.header("2. Currency rule")
    fx_convention = st.radio(
        "Interpret conversion rate as",
        ["1 unit of currency = rate EUR", "1 EUR = rate units of currency"],
        help="Choose the convention used by the Currency conversion table."
    )
    build = st.button("Build dashboard", type="primary", use_container_width=True)

if not build:
    st.info("Upload all ten files in the sidebar, select the conversion-rate convention, then click **Build dashboard**.")
    st.stop()

missing_files = [spec.label for spec in UPLOADS if files[spec.key] is None]
if missing_files:
    st.error("Missing required files: " + ", ".join(missing_files))
    st.stop()

try:
    frames = {key: read_uploaded(upload) for key, upload in files.items()}
    manual = transform_manual(frames["manual"])
    erp = transform_erp(frames["erp"], frames["carrier"], frames["shipping"], frames["customer"], frames["plant"])
    tm = transform_tm(frames["tm"], frames["tm_fo"], frames["tm_fb"])
    combined = style_clean(pd.concat([manual, erp, tm], ignore_index=True))
    combined, missing_fx = add_eur(combined, frames["fx"], fx_convention)
except Exception as exc:
    st.exception(exc)
    st.stop()

if missing_fx:
    st.warning("No EUR conversion rate found for: " + ", ".join(missing_fx) + ". Those rows remain visible but have no Value EUR.")

quality = {
    "Rows loaded": len(combined),
    "Missing date": int(combined["Date"].isna().sum()),
    "Missing FX": int(combined["FX Rate"].isna().sum()),
    "Missing EUR value": int(combined["Value EUR"].isna().sum()),
}

filtered = apply_filters(combined)

st.subheader("Overview")
k1, k2, k3, k4 = st.columns(4)
k1.metric("Freight cost (EUR)", f"€{filtered['Value EUR'].sum():,.2f}")
k2.metric("Transactions", f"{len(filtered):,}")
k3.metric("Carriers", f"{filtered['Carrier'].nunique():,}")
k4.metric("Average cost / transaction", f"€{filtered['Value EUR'].mean():,.2f}" if filtered["Value EUR"].notna().any() else "n/a")

if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

monthly = (filtered.dropna(subset=["Date", "Value EUR"])
           .assign(Month=lambda x: x["Date"].dt.to_period("M").dt.to_timestamp())
           .groupby(["Month", "Source"], as_index=False)["Value EUR"].sum())
carrier_cost = (filtered.groupby("Carrier", as_index=False)["Value EUR"].sum()
                .sort_values("Value EUR", ascending=False).head(15))
route_cost = (filtered.assign(Route=filtered["Origin"] + " → " + filtered["Destination"])
              .groupby("Route", as_index=False)["Value EUR"].sum()
              .sort_values("Value EUR", ascending=False).head(15))
source_cost = filtered.groupby("Source", as_index=False)["Value EUR"].sum()

c1, c2 = st.columns(2)
with c1:
    st.plotly_chart(px.line(monthly, x="Month", y="Value EUR", color="Source", markers=True,
                            title="Monthly freight costs", labels={"Value EUR": "Cost (EUR)"}), use_container_width=True)
with c2:
    st.plotly_chart(px.bar(carrier_cost.sort_values("Value EUR"), x="Value EUR", y="Carrier", orientation="h",
                           title="Top 15 carriers", labels={"Value EUR": "Cost (EUR)"}), use_container_width=True)

c3, c4 = st.columns(2)
with c3:
    st.plotly_chart(px.bar(route_cost.sort_values("Value EUR"), x="Value EUR", y="Route", orientation="h",
                           title="Top 15 routes", labels={"Value EUR": "Cost (EUR)"}), use_container_width=True)
with c4:
    st.plotly_chart(px.pie(source_cost, names="Source", values="Value EUR", hole=0.45,
                           title="Cost split by source"), use_container_width=True)

st.subheader("Detailed transactions")
display_cols = ["Source", "Freight Document", "Carrier", "Type of goods", "Date", "Origin", "Destination", "Value", "Currency", "FX Rate", "Value EUR"]
st.dataframe(filtered[display_cols].sort_values("Date", ascending=False), use_container_width=True, hide_index=True,
             column_config={
                 "Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
                 "Value": st.column_config.NumberColumn(format="%.2f"),
                 "FX Rate": st.column_config.NumberColumn(format="%.6f"),
                 "Value EUR": st.column_config.NumberColumn(format="€%.2f"),
             })

csv = filtered[display_cols].to_csv(index=False).encode("utf-8-sig")
st.download_button("Download filtered data (CSV)", csv, "actual_freight_costs_filtered.csv", "text/csv")

with st.expander("Data quality summary"):
    qcols = st.columns(len(quality))
    for col, (name, value) in zip(qcols, quality.items()):
        col.metric(name, f"{value:,}")
    st.write("Unmapped master-data keys fall back to their original code. Empty dimensions are shown as Unknown. SAP ERP empty ShPt is shown as Import.")
