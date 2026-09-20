# Optimizing Inventory and International Logistics for Manufacturing Operations

**Business Analytics & Supply Chain Analytics Portfolio Project**

## Project Overview

Efficient inventory management and international logistics are essential
to manufacturing operations. Poor inventory planning can increase
operating costs and stockout risk, while inefficient logistics can lead
to shipment delays, higher transportation costs, and reduced customer
satisfaction.

This project analyzes a **public Supply Chain Logistics dataset** using
**Python and Microsoft Excel**. It evaluates order activity, product
demand, delivery performance, carrier activity, plant operations,
warehouse cost and capacity, freight rates, transportation routes, and
plant network coverage.

The project was inspired by a supply chain topic discussed during a
manufacturing internship. To protect company confidentiality, the
analysis uses only publicly available data.

## Business Problem

Manufacturing companies may face operational challenges such as:

-   Excess inventory and inventory shortages
-   Long or variable lead times
-   Late shipments
-   High transportation costs
-   Inefficient warehouse capacity utilization
-   Differences in carrier, plant, and route performance

Business Analytics can help identify these issues and support more
consistent, data-driven supply chain decisions.

## Project Objectives

-   Analyze order volume and product demand
-   Evaluate delivery, carrier, and service-level performance
-   Compare plant activity and warehouse cost/capacity
-   Analyze freight rates and transportation routes
-   Demonstrate ABC Inventory Classification
-   Demonstrate Economic Order Quantity (EOQ)
-   Estimate Safety Stock and Reorder Point (ROP)
-   Calculate operational KPI proxies for warehouse utilization and
    on-time fulfillment
-   Generate practical business recommendations

## Dataset

The project uses the Excel workbook:

``` text
Supply chain logistics problem.xlsx
```

Worksheets used in the analysis:

-   `OrderList`
-   `FreightRates`
-   `WhCosts`
-   `WhCapacities`
-   `ProductsPerPlant`
-   `PlantPorts`
-   `VmiCustomers`

The workbook contains order, warehouse, freight, plant, customer, and
logistics information used throughout the project.

## Tools & Technologies

-   Python
-   Pandas
-   NumPy
-   Matplotlib
-   Microsoft Excel
-   GitHub

## Project Structure

``` text
Project-2/
├── data/
│   └── Supply chain logistics problem.xlsx
├── images/
├── outputs/
├── report/
├── project2_analysis.py
├── README.md
├── requirements.txt
└── LICENSE
```

## Python Analysis

The Python script performs:

-   Data loading from all required Excel worksheets
-   Data cleaning and data-quality summary
-   Order overview
-   Monthly order-volume analysis
-   Product-demand analysis
-   ABC Inventory Classification
-   EOQ demonstration
-   Safety Stock and Reorder Point (ROP) demonstration
-   Delivery-performance analysis
-   Service-level analysis
-   Carrier-performance analysis
-   Plant-performance analysis
-   Warehouse cost and capacity analysis
-   Freight-rate analysis
-   Freight-route analysis
-   Plant, product, customer, and port network coverage
-   Warehouse-utilization proxy
-   On-time fulfillment proxy
-   Key-findings summary
-   Automatic export of charts and CSV tables

## Inventory Optimization Assumptions

Some variables required for a complete inventory model are not available
in the public dataset. The project therefore uses clearly stated
assumptions for demonstration purposes.

### ABC Classification

ABC classification is based on **unit quantity and product activity**
because product revenue and unit value are not available in the dataset.

### Economic Order Quantity (EOQ)

EOQ uses observed unit quantity as estimated annual demand and the
following demonstration assumptions:

-   Ordering cost: **100**
-   Annual holding cost per unit: **2**

These assumptions are used to demonstrate the EOQ method and are not
actual company costs.

### Safety Stock and Reorder Point

The dataset does not contain supplier lead-time or inventory-on-hand
data. Therefore:

-   Observed transportation time (`TPT`) is used as a **lead-time
    proxy**
-   A **95% service level** is assumed (`Z = 1.645`)
-   Safety Stock is estimated from daily demand variability and the
    lead-time proxy
-   Reorder Point combines average daily demand, the lead-time proxy,
    and Safety Stock

These calculations are portfolio demonstrations rather than production
inventory policies.

## Operational KPI Notes

The project also calculates operational KPI proxies where the public
dataset does not contain all information required for a traditional KPI.

-   **Warehouse Utilization Proxy:** average daily shipped units divided
    by listed daily capacity
-   **On-Time Fulfillment Proxy:** percentage of shipment records
    classified as on time

The warehouse-utilization measure is an operational throughput proxy
rather than physical storage occupancy, and the on-time fulfillment
measure is used because the dataset does not include an explicit
fulfilled/cancelled status.

## Key Outputs

### Charts

-   Monthly Order Volume
-   Top 10 Products by Unit Quantity
-   ABC Product Classification
-   Delivery Status
-   Shipping Records by Service Level
-   Top Carriers by Shipment Records
-   Order Volume by Plant
-   Warehouse Cost per Unit by Plant
-   Daily Warehouse Capacity by Plant
-   Average Freight Rate by Transportation Mode
-   Top 10 Routes by Average Freight Rate

### CSV Tables

-   Data Quality Summary
-   Order Summary
-   Monthly Order Volume
-   Product Analysis
-   ABC Inventory Classification
-   ABC Summary
-   EOQ Demonstration
-   Safety Stock and Reorder Point
-   Delivery Status Summary
-   Service Level Analysis
-   Carrier Performance
-   Plant Performance
-   Warehouse Cost and Capacity Analysis
-   Freight Mode Analysis
-   Freight Route Analysis
-   Plant Network Coverage
-   Warehouse Utilization Proxy
-   Operational KPI Summary
-   Key Findings

## Key Findings

The analysis is designed to identify:

-   High-volume products that may require closer inventory monitoring
-   Product priorities through ABC classification
-   Safety Stock and Reorder Point estimates under stated assumptions
-   Shipment reliability and late-delivery patterns
-   Differences in carrier and service-level activity
-   High-volume plants and differences in warehouse cost/capacity
-   Transportation modes and routes with different freight-rate
    characteristics
-   Plant-level product, customer, and port coverage
-   Operational warehouse-utilization and on-time fulfillment proxies

The Python script automatically exports a `key_findings.csv` file
containing selected quantitative findings from the analysis.

## Business Recommendations

Based on the analytical framework used in this project:

-   Adopt inventory policies based on data rather than intuition
-   Review EOQ assumptions as more accurate ordering and holding cost
    data become available
-   Use ABC Classification to prioritize high-volume products for closer
    monitoring
-   Use Safety Stock and ROP estimates as a starting point for inventory
    planning when better lead-time data become available
-   Monitor carrier and delivery performance using shipment volume and
    late-delivery indicators
-   Evaluate transportation modes and routes according to cost and
    service requirements
-   Improve warehouse visibility through KPI monitoring and dashboard
    reporting
-   Standardize inventory review processes across warehouses
-   Continue using Python for automated reporting and consider SQL for
    future data management
-   Develop interactive Tableau dashboards for management reporting
-   Expand future analysis with demand forecasting and ERP integration

## Future Improvements

-   Add SQL for data management and analysis
-   Build interactive Tableau dashboards
-   Integrate real-time ERP data
-   Incorporate supplier-performance metrics
-   Add demand forecasting
-   Explore machine-learning forecasting models
-   Expand analysis to procurement optimization and production
    scheduling
-   Explore AI-based inventory optimization

## Author

**Fu-Min Yang**\
Business Analytics\
San José State University

## License

This project is released under the MIT License.
