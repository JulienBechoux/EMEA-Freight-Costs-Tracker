# Actual Freight Cost Dashboard

A Streamlit application that consolidates actual freight costs from **Manual accruals**, **SAP ERP**, and **SAP TM**, enriches ERP and TM records from reference files, converts all amounts to EUR, and provides an interactive dashboard with filters and exports.

## 1. What the application does

The application:

1. Accepts ten Excel or CSV uploads.
2. Standardizes all cost records into one model:
   - Carrier
   - Type of goods
   - Date
   - Origin
   - Destination
   - Original value
   - Original currency
   - EUR value
   - Source system
   - Freight Document, where available
3. Applies the requested cleansing and lookup rules.
4. Converts each transaction to EUR using the uploaded currency conversion table.
5. Provides filters for period, source, carrier, goods type, origin, destination, and currency.
6. Displays KPIs and charts for monthly spend, carriers, routes, and source-system split.
7. Allows the filtered detail to be downloaded as CSV.
8. Shows basic data-quality indicators for missing dates and exchange rates.

## 2. Project structure

```text
freight-cost-dashboard/
├── app.py
├── requirements.txt
└── README.md
```

## 3. Required input files

All ten files are required by the current implementation. Supported formats are `.xlsx`, `.xlsm`, `.xls`, and `.csv`.

### Cost files

- Manual accruals
- SAP ERP
- SAP TM

### Reference files

- ERP Carrier Name
- ERP Shipping Point
- ERP Customers
- ERP Plants
- TM FO
- TM FB
- Currency conversion table

Column matching is case-insensitive and ignores spaces and punctuation, but the business column names below should still be used whenever possible.

## 4. Transformation rules

### 4.1 Manual accruals

- **Carrier**: `Carrier Description`. The first slash and all text after it are removed. For example, `Carrier ABC / 123` becomes `Carrier ABC`.
- **Type of goods**: `PBU`.
- **Date**: `Planned Arrival Date-Last Stop`.
- **Origin**: `Source Location Description`.
- **Destination**: `Destination Location Descripti` or `Destination Location Description`.
- **Value**: `Net Amt in Doc Crcy`.
- **Currency**: `Currency`.

### 4.2 SAP ERP

- **Carrier**: `ServcAgent` is matched to the carrier reference file and returns `Name 1`. If no match is found, the original service-agent code is retained.
- **Type of goods**: `Product Hierarchy` starting with `1` or `4` becomes `FG`; all other values become `NFG`.
- **Date**: `Deliv.Date`.
- **Origin**: `ShPt` is matched to `Description` in the shipping-point reference file. If `ShPt` is empty, origin is `Import`. If a non-empty code is not mapped, the original code is retained.
- **Destination**: `Ship-To` is matched to `Name 1` in the customer reference file. If `Ship-To` is empty, `Plnt` is matched to `Name 1` in the plant reference file. Unmapped non-empty codes are retained.
- **Value**: `Loc.curr.amount`.
- **Currency**: `Local Curr.`.

The lookup file needs a key column. The app accepts common alternatives:

- Carrier key: `ServcAgent`, `Service Agent`, `Number`, or `Vendor`
- Shipping-point key: `ShPt`, `Shipping Point`, or `Code`
- Customer key: `Ship-To`, `Ship To`, `Customer`, or `Number`
- Plant key: `Plnt`, `Plant`, or `Code`

### 4.3 SAP TM

The SAP TM cost file supplies:

- **Carrier**: `Invoicing Party`
- **Value**: `Net Amt in Doc Crcy`
- **Currency**: `Currency`
- **Type of goods**: `Unknown`, because the source does not provide it

A freight-document column is required. Accepted names include `Freight Document`, `Freight Document Number`, `Freight Doc.`, `Freight Doc`, `Freight Order`, and `Freight Booking`.

For freight-document numbers starting with `68`, TM FO provides:

- **Date**: `Actual Delivered Date`; if empty, `Planned Arrival Date-Last Stop`
- **Origin**: `Source Location Description`
- **Destination**: `Destination Location Descripti` or its full spelling

For freight-document numbers starting with `69`, TM FB provides:

- **Date**: `Expected Arrival Date`
- **Origin**: `Source Location Description`
- **Destination**: `Destination Location Descripti` or its full spelling

Documents not starting with `68` or `69` remain in the output, but TM date, origin, and destination are left empty and displayed as `Unknown` where appropriate.

## 5. Currency conversion

The currency table must contain:

- A currency column named `Currency`, `From Currency`, `Source Currency`, `Curr.`, or `Local Curr.`
- A rate column named `Conversion Rate`, `Rate`, `Rate to EUR`, `EUR Rate`, or `Exchange Rate`

EUR is automatically assigned a rate of `1.0`.

The sidebar offers two rate conventions:

1. **1 unit of currency = rate EUR**
   - Formula: `Value EUR = Value × FX Rate`
2. **1 EUR = rate units of currency**
   - Formula: `Value EUR = Value ÷ FX Rate`

Select the convention that matches the uploaded table. A missing rate leaves `Value EUR` blank and triggers a warning.

> Important: The current application uses one rate per currency. If the conversion table contains several dated rates for the same currency, the last row for that currency is used. For historical daily or monthly rates, extend the lookup to include an effective date and merge on both currency and date.

## 6. Installation

### Prerequisites

- Python 3.10 or later
- `pip`

### Create and activate a virtual environment

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Windows Command Prompt:

```bat
python -m venv .venv
.venv\Scripts\activate.bat
```

macOS or Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### Install dependencies

```bash
pip install -r requirements.txt
```

## 7. Run the application

From the project folder:

```bash
streamlit run app.py
```

Streamlit normally opens the app at `http://localhost:8501`.

## 8. Using the dashboard

1. Open the application.
2. Upload each of the ten files in the sidebar.
3. Select the currency-rate convention.
4. Click **Build dashboard**.
5. Review any missing-rate warning or column error.
6. Use sidebar filters to slice the consolidated data.
7. Review KPIs and charts.
8. Open the detailed transaction table.
9. Click **Download filtered data (CSV)** to export the current selection.

An empty multi-select means **All**. Filters are applied sequentially, so later filter choices show values available after earlier filters have been applied.

## 9. Dashboard content

### KPIs

- Freight cost in EUR
- Number of transactions
- Number of distinct carriers
- Average EUR cost per transaction

### Visuals

- Monthly freight cost by source
- Top 15 carriers by EUR cost
- Top 15 origin-to-destination routes by EUR cost
- Cost split by source system

### Detail and controls

- Interactive transaction table
- Period, source, carrier, goods type, origin, destination, and currency filters
- CSV export of filtered rows
- Data-quality summary

## 10. Data quality and reconciliation

Before publishing the dashboard, complete these checks:

1. Reconcile transaction counts and original-currency totals against each input file.
2. Confirm whether cost values can be negative and whether credits should offset spend.
3. Confirm the exchange-rate convention and rate effective date.
4. Review unmapped ERP carrier, shipping point, customer, and plant codes.
5. Review TM documents without FO or FB detail matches.
6. Review missing or invalid dates.
7. Confirm whether duplicate freight documents are valid. The master-data and TM-detail lookups keep the last record for duplicate keys.
8. Validate totals after EUR conversion with Finance.
9. Confirm whether manual accruals and posted ERP/TM costs can overlap. The app consolidates records but does not automatically eliminate cross-source duplicates.

## 11. Error handling

The app stops and displays a detailed exception when:

- A required file is missing
- A required column cannot be found
- A spreadsheet cannot be read
- A numeric amount cannot be parsed
- A lookup cannot be built

For missing FX rates, the application continues, warns the user, and leaves the EUR value blank.

## 12. Parsing behavior

- Headers are matched after converting to lowercase and removing punctuation and spaces.
- Lookup keys are converted to text, trimmed, and trailing `.0` is removed to reduce Excel numeric-format mismatches.
- Amount parsing supports numbers, spaces, currency symbols, parentheses for negatives, and common comma/decimal conventions.
- Dates support native Excel dates, common day-first text dates, and Excel serial dates.
- Blank descriptive dimensions become `Unknown` in the dashboard.

## 13. Security and privacy

- Uploaded files are processed in the running Streamlit session.
- The supplied code does not write uploads to disk or call an external API.
- Deployment logging, authentication, backups, retention, and infrastructure security depend on the hosting environment.
- For production use, deploy behind corporate authentication and restrict access to authorized freight and finance users.
- Do not expose commercially sensitive carrier rates or personally identifiable information unnecessarily.

## 14. Production deployment options

The app may be hosted on an approved internal platform, a container service, or a virtual machine.

Example local Dockerfile, if your environment permits containers:

```dockerfile
FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py .
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0", "--server.port=8501"]
```

For production, also configure:

- Single sign-on or reverse-proxy authentication
- TLS
- Maximum upload size
- Resource limits
- Monitoring and health checks
- Dependency scanning
- Controlled release and rollback procedures

## 15. Known assumptions

- Each uploaded workbook uses its first worksheet.
- The currency table contains one usable rate per currency.
- The requested three cost sources are additive.
- `Freight Document` values are preserved as text, including meaningful leading zeros when the source itself preserves them.
- TM details are joined many-to-one by freight-document number.
- When duplicate lookup keys exist, the last occurrence is used.
- The app does not derive Type of goods for SAP TM and uses `Unknown`.
- The app does not deduplicate possible overlap between manual accruals, SAP ERP, and SAP TM.

## 16. Customization points

Common extensions include:

- Add dated FX rates and date-effective conversion
- Add user-selectable workbook sheets
- Add automatic cross-source duplicate detection
- Add budget, forecast, or prior-year comparisons
- Add lane, country, region, or business-unit mappings
- Store curated data in a database rather than uploading files each session
- Schedule ingestion and publish a governed semantic model
- Add role-based access and audit logging
- Add tests for transformations using representative masked data

## 17. Troubleshooting

### Missing column error

Compare the error's expected column names with the actual file header. Remove merged cells and title rows above the header. Ensure the correct worksheet is first.

### Lookup values do not map

Check leading zeros, hidden spaces, and whether the reference file uses the expected key. The app normalizes trimming and `.0`, but it cannot reconstruct leading zeros already removed by Excel.

### EUR totals are too high or too low

Switch the rate convention and verify the definition of the uploaded rate. Also check whether the table is direct-to-EUR or requires a cross-rate.

### Dates look incorrect

Provide real Excel dates or an unambiguous text date. Text dates are interpreted day-first by default.

### CSV does not open correctly

The export uses UTF-8 with a byte-order mark so European characters normally display correctly in Excel.

## 18. Suggested acceptance criteria

- All ten supported files load successfully.
- Manual carrier descriptions are truncated at the first slash.
- ERP carrier, origin, and destination names map correctly.
- Empty ERP shipping points become `Import`.
- ERP goods type is classified as FG or NFG correctly.
- TM 68 and 69 documents use their respective detail source and date logic.
- Every currency either has a valid EUR rate or is clearly flagged.
- Dashboard totals reconcile with an independently calculated sample.
- Every filter updates KPIs, visuals, detail, and export consistently.
- Downloaded data matches the visible filtered selection.
