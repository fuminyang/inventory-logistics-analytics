"""
Project 2: Optimizing Inventory and International Logistics
for Manufacturing Operations

Author: Fu-Min Yang

This single-file Python project:
1. Loads all worksheets from the Excel workbook
2. Cleans the data
3. Creates summary tables
4. Analyzes orders, delivery performance, carriers, plants,
   warehouse cost/capacity, and freight rates
5. Demonstrates ABC analysis, EOQ, Safety Stock, and Reorder Point using clearly stated assumptions
6. Calculates operational KPIs including warehouse utilization and on-time fulfillment proxy
7. Saves charts and CSV outputs

Required folder structure:

Project-2/
├── project2_analysis.py
├── data/
│   └── Supply chain logistics problem.xlsx
├── images/
└── outputs/
"""

from pathlib import Path
import math
import re

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


# =========================================================
# 1. PROJECT PATHS
# =========================================================

try:
    PROJECT_DIR = Path(__file__).resolve().parent
except NameError:
    PROJECT_DIR = Path.cwd()

DATA_DIR = PROJECT_DIR / "data"
IMAGES_DIR = PROJECT_DIR / "images"
OUTPUTS_DIR = PROJECT_DIR / "outputs"

IMAGES_DIR.mkdir(parents=True, exist_ok=True)
OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

DATA_FILE_CANDIDATES = [
    DATA_DIR / "Supply chain logistics problem.xlsx",
    DATA_DIR / "Supply chain logisitcs problem.xlsx",
    DATA_DIR / "Supply chain logisitcs problem(1).xlsx",
]

DATA_FILE = next((path for path in DATA_FILE_CANDIDATES if path.exists()), DATA_FILE_CANDIDATES[0])


# =========================================================
# 2. SETTINGS FOR DEMONSTRATION CALCULATIONS
# =========================================================

# The dataset does not provide ordering cost or annual holding
# cost. These values are assumptions used only to demonstrate EOQ.
ORDERING_COST = 100
ANNUAL_HOLDING_COST_PER_UNIT = 2

# Safety Stock uses a 95% service-level assumption. The workbook does
# not contain supplier lead time, so observed transportation time (TPT)
# is used as a lead-time proxy for this portfolio demonstration.
SERVICE_LEVEL_Z = 1.645


# =========================================================
# 3. HELPER FUNCTIONS
# =========================================================

def clean_column_name(column):
    """Convert column names to lowercase snake_case."""
    column = str(column).strip()
    column = re.sub(r"[^A-Za-z0-9]+", "_", column)
    return column.strip("_").lower()


def clean_dataframe(df):
    """Apply simple and conservative cleaning."""
    df = df.copy()

    df.columns = [clean_column_name(column) for column in df.columns]
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df = df.drop_duplicates().reset_index(drop=True)

    text_columns = df.select_dtypes(include=["object", "string"]).columns
    for column in text_columns:
        df[column] = df[column].apply(
            lambda value: value.strip() if isinstance(value, str) else value
        )

    return df


def save_chart(filename):
    """Save the current chart and close it."""
    plt.tight_layout()
    plt.savefig(IMAGES_DIR / filename, dpi=300, bbox_inches="tight")
    plt.close()


def save_table(df, filename):
    """Save a DataFrame as a CSV file."""
    df.to_csv(OUTPUTS_DIR / filename, index=False)


def print_section(title):
    """Print a clear section heading."""
    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)


# =========================================================
# 4. LOAD DATA
# =========================================================

if not DATA_FILE.exists():
    raise FileNotFoundError(
        f"Excel file not found:\n{DATA_FILE}\n\n"
        "Place the workbook inside the data folder. Recommended filename: "
        "'Supply chain logistics problem.xlsx'."
    )

print_section("Loading Excel Workbook")

excel_data = pd.read_excel(DATA_FILE, sheet_name=None)

data = {
    sheet_name: clean_dataframe(dataframe)
    for sheet_name, dataframe in excel_data.items()
}

orders = data["OrderList"]
freight = data["FreightRates"]
warehouse_costs = data["WhCosts"]
warehouse_capacities = data["WhCapacities"]
products_per_plant = data["ProductsPerPlant"]
vmi_customers = data["VmiCustomers"]
plant_ports = data["PlantPorts"]

orders["order_date"] = pd.to_datetime(orders["order_date"], errors="coerce")

numeric_order_columns = [
    "tpt",
    "ship_ahead_day_count",
    "ship_late_day_count",
    "unit_quantity",
    "weight",
]

for column in numeric_order_columns:
    if column in orders.columns:
        orders[column] = pd.to_numeric(orders[column], errors="coerce")

numeric_freight_columns = [
    "minm_wgh_qty",
    "max_wgh_qty",
    "minimum_cost",
    "rate",
    "tpt_day_cnt",
]

for column in numeric_freight_columns:
    if column in freight.columns:
        freight[column] = pd.to_numeric(freight[column], errors="coerce")

warehouse_costs["cost_unit"] = pd.to_numeric(
    warehouse_costs["cost_unit"], errors="coerce"
)

warehouse_capacities["daily_capacity"] = pd.to_numeric(
    warehouse_capacities["daily_capacity"], errors="coerce"
)

print("Workbook loaded successfully.")
for sheet_name, dataframe in data.items():
    print(f"{sheet_name}: {dataframe.shape[0]:,} rows × {dataframe.shape[1]} columns")


# =========================================================
# 5. DATA QUALITY SUMMARY
# =========================================================

print_section("Data Quality Summary")

quality_rows = []

for sheet_name, dataframe in data.items():
    quality_rows.append(
        {
            "sheet": sheet_name,
            "rows": len(dataframe),
            "columns": len(dataframe.columns),
            "missing_values": int(dataframe.isna().sum().sum()),
            "duplicate_rows_after_cleaning": int(dataframe.duplicated().sum()),
        }
    )

data_quality = pd.DataFrame(quality_rows)
save_table(data_quality, "data_quality_summary.csv")
print(data_quality.to_string(index=False))


# =========================================================
# 6. ORDER OVERVIEW
# =========================================================

print_section("Order Overview")

order_summary = pd.DataFrame(
    {
        "metric": [
            "Number of Orders",
            "Total Units",
            "Total Weight",
            "Unique Products",
            "Unique Customers",
            "Unique Plants",
            "Unique Carriers",
            "Average Transit Time",
            "Average Late Days",
        ],
        "value": [
            orders["order_id"].nunique(),
            orders["unit_quantity"].sum(),
            orders["weight"].sum(),
            orders["product_id"].nunique(),
            orders["customer"].nunique(),
            orders["plant_code"].nunique(),
            orders["carrier"].nunique(),
            orders["tpt"].mean(),
            orders["ship_late_day_count"].mean(),
        ],
    }
)

order_summary["value"] = order_summary["value"].round(2)
save_table(order_summary, "order_summary.csv")
print(order_summary.to_string(index=False))


# =========================================================
# 7. MONTHLY ORDER VOLUME
# =========================================================

monthly_orders = (
    orders.dropna(subset=["order_date"])
    .assign(month=lambda df: df["order_date"].dt.to_period("M").astype(str))
    .groupby("month", as_index=False)
    .agg(
        total_units=("unit_quantity", "sum"),
        shipment_records=("order_id", "count"),
    )
)

save_table(monthly_orders, "monthly_order_volume.csv")

plt.figure(figsize=(12, 6))
plt.plot(monthly_orders["month"], monthly_orders["total_units"], marker="o")
plt.title("Monthly Order Volume")
plt.xlabel("Month")
plt.ylabel("Total Unit Quantity")
plt.xticks(rotation=45)
save_chart("monthly_order_volume.png")


# =========================================================
# 8. PRODUCT ANALYSIS
# =========================================================

product_analysis = (
    orders.groupby("product_id", as_index=False)
    .agg(
        total_units=("unit_quantity", "sum"),
        total_weight=("weight", "sum"),
        shipment_records=("order_id", "count"),
    )
    .sort_values("total_units", ascending=False)
)

save_table(product_analysis, "product_analysis.csv")

top_products = product_analysis.head(10)

plt.figure(figsize=(11, 6))
plt.bar(top_products["product_id"].astype(str), top_products["total_units"])
plt.title("Top 10 Products by Unit Quantity")
plt.xlabel("Product ID")
plt.ylabel("Total Unit Quantity")
plt.xticks(rotation=45)
save_chart("top_products_by_units.png")


# =========================================================
# 9. ABC INVENTORY CLASSIFICATION
# =========================================================

# Unit quantity is used as the available measure of product activity.
# This is an activity-based ABC demonstration because the workbook
# does not contain product revenue or unit value.

abc_analysis = product_analysis[
    ["product_id", "total_units", "shipment_records"]
].copy()

abc_analysis = abc_analysis.sort_values(
    "total_units", ascending=False
).reset_index(drop=True)

total_units_all_products = abc_analysis["total_units"].sum()

if total_units_all_products > 0:
    abc_analysis["unit_share"] = (
        abc_analysis["total_units"] / total_units_all_products
    )
else:
    abc_analysis["unit_share"] = 0

abc_analysis["cumulative_share"] = abc_analysis["unit_share"].cumsum()

abc_analysis["abc_class"] = np.select(
    [
        abc_analysis["cumulative_share"] <= 0.80,
        abc_analysis["cumulative_share"] <= 0.95,
    ],
    ["A", "B"],
    default="C",
)

save_table(abc_analysis, "abc_inventory_classification.csv")

abc_summary = (
    abc_analysis.groupby("abc_class", as_index=False)
    .agg(
        number_of_products=("product_id", "count"),
        total_units=("total_units", "sum"),
    )
    .sort_values("abc_class")
)

save_table(abc_summary, "abc_summary.csv")

plt.figure(figsize=(8, 5))
plt.bar(abc_summary["abc_class"], abc_summary["number_of_products"])
plt.title("ABC Product Classification")
plt.xlabel("ABC Class")
plt.ylabel("Number of Products")
save_chart("abc_product_classification.png")


# =========================================================
# 10. EOQ DEMONSTRATION
# =========================================================

# Annual demand is estimated from observed unit quantity.
# EOQ = sqrt((2 × annual demand × ordering cost)
#            / annual holding cost per unit)

eoq_analysis = product_analysis[
    ["product_id", "total_units", "shipment_records"]
].copy()

eoq_analysis = eoq_analysis.rename(
    columns={"total_units": "estimated_annual_demand"}
)

eoq_analysis["assumed_ordering_cost"] = ORDERING_COST
eoq_analysis["assumed_holding_cost_per_unit"] = (
    ANNUAL_HOLDING_COST_PER_UNIT
)

eoq_analysis["demonstration_eoq"] = np.sqrt(
    (
        2
        * eoq_analysis["estimated_annual_demand"]
        * ORDERING_COST
    )
    / ANNUAL_HOLDING_COST_PER_UNIT
).round(0)

eoq_analysis = eoq_analysis.sort_values(
    "estimated_annual_demand", ascending=False
)

save_table(eoq_analysis, "eoq_demonstration.csv")


# =========================================================
# 11. SAFETY STOCK AND REORDER POINT DEMONSTRATION
# =========================================================

# The workbook does not contain supplier lead time or inventory-on-hand.
# Therefore, Safety Stock and ROP are demonstrated using observed daily
# product demand and transportation time (TPT) as a lead-time proxy.
#
# Safety Stock = Z × standard deviation of daily demand × sqrt(lead time)
# ROP = average daily demand × lead time + Safety Stock

order_dates = orders["order_date"].dropna()
analysis_days = max((order_dates.max() - order_dates.min()).days + 1, 1)

daily_product_demand = (
    orders.dropna(subset=["order_date", "product_id"])
    .assign(order_day=lambda df: df["order_date"].dt.normalize())
    .groupby(["product_id", "order_day"], as_index=False)
    .agg(daily_units=("unit_quantity", "sum"))
)

# Reindex each product across the complete observed date range so days with
# no recorded demand are represented as zero demand.
demand_stats_rows = []
full_date_range = pd.date_range(order_dates.min().normalize(), order_dates.max().normalize(), freq="D")

for product_id, group in daily_product_demand.groupby("product_id"):
    series = (
        group.set_index("order_day")["daily_units"]
        .reindex(full_date_range, fill_value=0)
    )
    demand_stats_rows.append(
        {
            "product_id": product_id,
            "average_daily_demand": series.mean(),
            "daily_demand_std": series.std(ddof=0),
        }
    )

demand_stats = pd.DataFrame(demand_stats_rows)

product_lead_time = (
    orders.groupby("product_id", as_index=False)
    .agg(average_lead_time_days=("tpt", "mean"))
)

safety_stock_rop = demand_stats.merge(product_lead_time, on="product_id", how="left")
safety_stock_rop["average_lead_time_days"] = (
    safety_stock_rop["average_lead_time_days"].fillna(1).clip(lower=1)
)
safety_stock_rop["service_level_z"] = SERVICE_LEVEL_Z
safety_stock_rop["safety_stock_units"] = (
    SERVICE_LEVEL_Z
    * safety_stock_rop["daily_demand_std"]
    * np.sqrt(safety_stock_rop["average_lead_time_days"])
).round(0)
safety_stock_rop["reorder_point_units"] = (
    safety_stock_rop["average_daily_demand"]
    * safety_stock_rop["average_lead_time_days"]
    + safety_stock_rop["safety_stock_units"]
).round(0)

safety_stock_rop = safety_stock_rop.merge(
    abc_analysis[["product_id", "abc_class"]],
    on="product_id",
    how="left",
).sort_values("reorder_point_units", ascending=False)

for column in ["average_daily_demand", "daily_demand_std", "average_lead_time_days"]:
    safety_stock_rop[column] = safety_stock_rop[column].round(2)

save_table(safety_stock_rop, "safety_stock_reorder_point.csv")


# =========================================================
# 12. DELIVERY PERFORMANCE
# =========================================================

orders["delivery_status"] = np.where(
    orders["ship_late_day_count"].fillna(0) > 0,
    "Late",
    "On Time",
)

delivery_status = (
    orders["delivery_status"]
    .value_counts()
    .rename_axis("delivery_status")
    .reset_index(name="shipment_records")
)

delivery_status["percentage"] = (
    delivery_status["shipment_records"]
    / delivery_status["shipment_records"].sum()
    * 100
).round(2)

save_table(delivery_status, "delivery_status_summary.csv")

plt.figure(figsize=(8, 5))
plt.bar(
    delivery_status["delivery_status"],
    delivery_status["shipment_records"],
)
plt.title("Delivery Status")
plt.xlabel("Delivery Status")
plt.ylabel("Shipment Records")
save_chart("delivery_status.png")


# =========================================================
# 12. SERVICE LEVEL ANALYSIS
# =========================================================

service_level_analysis = (
    orders.groupby("service_level", as_index=False)
    .agg(
        shipment_records=("order_id", "count"),
        total_units=("unit_quantity", "sum"),
        average_transit_time=("tpt", "mean"),
        average_late_days=("ship_late_day_count", "mean"),
    )
    .sort_values("shipment_records", ascending=False)
)

service_level_analysis[
    ["average_transit_time", "average_late_days"]
] = service_level_analysis[
    ["average_transit_time", "average_late_days"]
].round(2)

save_table(service_level_analysis, "service_level_analysis.csv")

plt.figure(figsize=(9, 5))
plt.bar(
    service_level_analysis["service_level"].astype(str),
    service_level_analysis["shipment_records"],
)
plt.title("Shipping Records by Service Level")
plt.xlabel("Service Level")
plt.ylabel("Shipment Records")
save_chart("shipping_records_by_service_level.png")


# =========================================================
# 13. CARRIER PERFORMANCE
# =========================================================

carrier_performance = (
    orders.groupby("carrier", as_index=False)
    .agg(
        shipment_records=("order_id", "count"),
        total_units=("unit_quantity", "sum"),
        total_weight=("weight", "sum"),
        average_transit_time=("tpt", "mean"),
        average_late_days=("ship_late_day_count", "mean"),
    )
    .sort_values("shipment_records", ascending=False)
)

carrier_performance[
    ["total_weight", "average_transit_time", "average_late_days"]
] = carrier_performance[
    ["total_weight", "average_transit_time", "average_late_days"]
].round(2)

save_table(carrier_performance, "carrier_performance.csv")

top_carriers = carrier_performance.head(10)

plt.figure(figsize=(11, 6))
plt.bar(top_carriers["carrier"], top_carriers["shipment_records"])
plt.title("Top Carriers by Shipment Records")
plt.xlabel("Carrier")
plt.ylabel("Shipment Records")
plt.xticks(rotation=45)
save_chart("top_carriers.png")


# =========================================================
# 14. PLANT PERFORMANCE
# =========================================================

plant_performance = (
    orders.groupby("plant_code", as_index=False)
    .agg(
        shipment_records=("order_id", "count"),
        total_units=("unit_quantity", "sum"),
        total_weight=("weight", "sum"),
        unique_products=("product_id", "nunique"),
        average_late_days=("ship_late_day_count", "mean"),
    )
)

plant_performance = plant_performance.merge(
    warehouse_costs.rename(columns={"wh": "plant_code"}),
    on="plant_code",
    how="left",
)

plant_performance = plant_performance.merge(
    warehouse_capacities.rename(columns={"plant_id": "plant_code"}),
    on="plant_code",
    how="left",
)

plant_performance["estimated_warehouse_cost"] = (
    plant_performance["total_units"]
    * plant_performance["cost_unit"]
)

plant_performance["units_per_daily_capacity_point"] = (
    plant_performance["total_units"]
    / plant_performance["daily_capacity"].replace(0, np.nan)
)

plant_performance = plant_performance.sort_values(
    "total_units", ascending=False
)

numeric_plant_columns = [
    "total_weight",
    "average_late_days",
    "cost_unit",
    "estimated_warehouse_cost",
    "units_per_daily_capacity_point",
]

for column in numeric_plant_columns:
    plant_performance[column] = plant_performance[column].round(2)

save_table(plant_performance, "plant_performance.csv")

plt.figure(figsize=(12, 6))
plt.bar(
    plant_performance["plant_code"],
    plant_performance["total_units"],
)
plt.title("Order Volume by Plant")
plt.xlabel("Plant")
plt.ylabel("Total Unit Quantity")
plt.xticks(rotation=45)
save_chart("order_volume_by_plant.png")


# =========================================================
# 15. WAREHOUSE COST AND CAPACITY
# =========================================================

warehouse_analysis = warehouse_costs.rename(
    columns={"wh": "plant_code"}
).merge(
    warehouse_capacities.rename(columns={"plant_id": "plant_code"}),
    on="plant_code",
    how="outer",
)

warehouse_analysis["cost_capacity_ratio"] = (
    warehouse_analysis["cost_unit"]
    / warehouse_analysis["daily_capacity"].replace(0, np.nan)
)

warehouse_analysis = warehouse_analysis.sort_values(
    "cost_unit", ascending=False
)

warehouse_analysis[
    ["cost_unit", "cost_capacity_ratio"]
] = warehouse_analysis[
    ["cost_unit", "cost_capacity_ratio"]
].round(4)

save_table(warehouse_analysis, "warehouse_cost_capacity_analysis.csv")

plt.figure(figsize=(11, 6))
plt.bar(
    warehouse_analysis["plant_code"],
    warehouse_analysis["cost_unit"],
)
plt.title("Warehouse Cost per Unit by Plant")
plt.xlabel("Plant")
plt.ylabel("Cost per Unit")
plt.xticks(rotation=45)
save_chart("warehouse_cost_per_unit.png")

capacity_sorted = warehouse_analysis.sort_values(
    "daily_capacity", ascending=False
)

plt.figure(figsize=(11, 6))
plt.bar(
    capacity_sorted["plant_code"],
    capacity_sorted["daily_capacity"],
)
plt.title("Daily Warehouse Capacity by Plant")
plt.xlabel("Plant")
plt.ylabel("Daily Capacity")
plt.xticks(rotation=45)
save_chart("warehouse_daily_capacity.png")


# =========================================================
# 16. FREIGHT RATE ANALYSIS
# =========================================================

freight_mode_analysis = (
    freight.groupby("mode_dsc", as_index=False)
    .agg(
        rate_records=("rate", "count"),
        average_rate=("rate", "mean"),
        average_minimum_cost=("minimum_cost", "mean"),
        average_transit_days=("tpt_day_cnt", "mean"),
    )
    .sort_values("average_rate")
)

freight_mode_analysis[
    ["average_rate", "average_minimum_cost", "average_transit_days"]
] = freight_mode_analysis[
    ["average_rate", "average_minimum_cost", "average_transit_days"]
].round(2)

save_table(freight_mode_analysis, "freight_mode_analysis.csv")

plt.figure(figsize=(9, 5))
plt.bar(
    freight_mode_analysis["mode_dsc"].astype(str),
    freight_mode_analysis["average_rate"],
)
plt.title("Average Freight Rate by Transportation Mode")
plt.xlabel("Transportation Mode")
plt.ylabel("Average Rate")
save_chart("average_freight_rate_by_mode.png")


# =========================================================
# 17. FREIGHT ROUTE ANALYSIS
# =========================================================

freight_route_analysis = (
    freight.groupby(
        ["orig_port_cd", "dest_port_cd"],
        as_index=False,
    )
    .agg(
        rate_records=("rate", "count"),
        average_rate=("rate", "mean"),
        average_minimum_cost=("minimum_cost", "mean"),
        average_transit_days=("tpt_day_cnt", "mean"),
    )
)

freight_route_analysis["route"] = (
    freight_route_analysis["orig_port_cd"].astype(str)
    + " → "
    + freight_route_analysis["dest_port_cd"].astype(str)
)

freight_route_analysis = freight_route_analysis.sort_values(
    "average_rate", ascending=False
)

freight_route_analysis[
    ["average_rate", "average_minimum_cost", "average_transit_days"]
] = freight_route_analysis[
    ["average_rate", "average_minimum_cost", "average_transit_days"]
].round(2)

save_table(freight_route_analysis, "freight_route_analysis.csv")

top_expensive_routes = freight_route_analysis.head(10)

plt.figure(figsize=(12, 6))
plt.bar(
    top_expensive_routes["route"],
    top_expensive_routes["average_rate"],
)
plt.title("Top 10 Routes by Average Freight Rate")
plt.xlabel("Route")
plt.ylabel("Average Rate")
plt.xticks(rotation=45, ha="right")
save_chart("top_routes_by_average_rate.png")


# =========================================================
# 18. PLANT, PRODUCT, CUSTOMER, AND PORT COVERAGE
# =========================================================

plant_coverage = (
    products_per_plant.groupby("plant_code", as_index=False)
    .agg(number_of_products=("product_id", "nunique"))
)

vmi_coverage = (
    vmi_customers.groupby("plant_code", as_index=False)
    .agg(number_of_vmi_customers=("customers", "nunique"))
)

port_coverage = (
    plant_ports.groupby("plant_code", as_index=False)
    .agg(number_of_ports=("port", "nunique"))
)

plant_network = (
    plant_coverage
    .merge(vmi_coverage, on="plant_code", how="outer")
    .merge(port_coverage, on="plant_code", how="outer")
    .fillna(0)
    .sort_values("number_of_products", ascending=False)
)

save_table(plant_network, "plant_network_coverage.csv")


# =========================================================
# 20. OPERATIONAL KPI SUMMARY
# =========================================================

# Warehouse utilization is estimated as average daily shipped units divided
# by listed daily capacity. It is an operational proxy, not physical storage
# occupancy, because inventory-on-hand data are not available.
plant_daily_units = (
    orders.dropna(subset=["order_date"])
    .assign(order_day=lambda df: df["order_date"].dt.normalize())
    .groupby(["plant_code", "order_day"], as_index=False)
    .agg(daily_units=("unit_quantity", "sum"))
)

average_daily_units_by_plant = (
    plant_daily_units.groupby("plant_code", as_index=False)
    .agg(average_daily_units=("daily_units", "mean"))
)

warehouse_utilization = warehouse_capacities.rename(
    columns={"plant_id": "plant_code"}
).merge(average_daily_units_by_plant, on="plant_code", how="left")
warehouse_utilization["average_daily_units"] = warehouse_utilization["average_daily_units"].fillna(0)
warehouse_utilization["utilization_proxy_pct"] = (
    warehouse_utilization["average_daily_units"]
    / warehouse_utilization["daily_capacity"].replace(0, np.nan)
    * 100
).round(2)

save_table(warehouse_utilization, "warehouse_utilization_proxy.csv")

# The dataset has shipment timing but no explicit fulfilled/cancelled flag.
# On-time shipment rate is therefore reported as a fulfillment-service proxy.
on_time_records = int((orders["delivery_status"] == "On Time").sum())
total_shipment_records = int(len(orders))
on_time_fulfillment_proxy = (
    on_time_records / total_shipment_records * 100
    if total_shipment_records else np.nan
)

kpi_summary = pd.DataFrame(
    [
        {
            "kpi": "On-time fulfillment proxy",
            "value": round(on_time_fulfillment_proxy, 2),
            "unit": "%",
            "note": "On-time shipment rate; used because explicit fulfillment status is unavailable.",
        },
        {
            "kpi": "Products with Safety Stock / ROP estimates",
            "value": len(safety_stock_rop),
            "unit": "products",
            "note": "95% service-level assumption; TPT used as lead-time proxy.",
        },
        {
            "kpi": "Products with EOQ estimates",
            "value": len(eoq_analysis),
            "unit": "products",
            "note": "Uses stated ordering-cost and holding-cost assumptions.",
        },
    ]
)

save_table(kpi_summary, "operational_kpi_summary.csv")


# =========================================================
# 21. KEY FINDINGS
# =========================================================

print_section("Key Findings")

late_row = delivery_status[
    delivery_status["delivery_status"] == "Late"
]

late_percentage = (
    float(late_row["percentage"].iloc[0])
    if not late_row.empty
    else 0
)

top_product_id = product_analysis.iloc[0]["product_id"]
top_product_units = product_analysis.iloc[0]["total_units"]

top_plant = plant_performance.iloc[0]["plant_code"]
top_plant_units = plant_performance.iloc[0]["total_units"]

lowest_rate_mode = freight_mode_analysis.iloc[0]["mode_dsc"]
lowest_average_rate = freight_mode_analysis.iloc[0]["average_rate"]

highest_capacity_row = warehouse_analysis.sort_values(
    "daily_capacity", ascending=False
).iloc[0]

findings = [
    {
        "finding": "Highest-volume product",
        "result": (
            f"Product {top_product_id} with "
            f"{top_product_units:,.0f} units"
        ),
    },
    {
        "finding": "Highest-volume plant",
        "result": (
            f"{top_plant} with "
            f"{top_plant_units:,.0f} units"
        ),
    },
    {
        "finding": "Late shipment percentage",
        "result": f"{late_percentage:.2f}%",
    },
    {
        "finding": "Transportation mode with lowest average rate",
        "result": (
            f"{lowest_rate_mode} with an average rate of "
            f"{lowest_average_rate:,.2f}"
        ),
    },
    {
        "finding": "Plant with highest listed daily capacity",
        "result": (
            f"{highest_capacity_row['plant_code']} with "
            f"{highest_capacity_row['daily_capacity']:,.0f}"
        ),
    },
]

key_findings = pd.DataFrame(findings)
save_table(key_findings, "key_findings.csv")
print(key_findings.to_string(index=False))


# =========================================================
# 22. FINAL MESSAGE
# =========================================================

print_section("Analysis Completed")

print(f"Charts saved to: {IMAGES_DIR}")
print(f"CSV outputs saved to: {OUTPUTS_DIR}")
print(
    "\nNote: ABC classification uses unit quantity because product "
    "value/revenue is unavailable. EOQ uses stated ordering-cost and "
    "holding-cost assumptions. Safety Stock and ROP use a 95% service "
    "level and observed TPT as a lead-time proxy because supplier lead "
    "time is unavailable."
)
