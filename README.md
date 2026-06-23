# Automated Price Analysis & Purchase Order Engine

A high-performance Excel application powered by **Power Query** and **VBA**. It is designed to automate complex price matrix evaluations, utilize local memory caching for speed, and dynamically generate optimized supplier purchase orders (RFQ) for the automotive parts market.

*Note: The tool was originally developed for the Polish automotive market, so product descriptions are kept in Polish within the dataset, while the technical architecture and codebase are fully documented in English.*

## 🚀 Key Features

* **High-Performance Memory Caching:** Utilizes VBA `Scripting.Dictionary` objects to load and map thousands of price records instantly in-memory, bypassing slow standard Excel sheet lookups (`VLOOKUP`/`XLOOKUP`).
* **Smart ETL Integration:** Power Query automatically consolidates multiple decentralized price source files (simulated SharePoint environment) into a unified analytical structure.
* **Fail-Safe Global Reset:** Includes interactive dashboard controls linked to a macro that loops through the session memory and instantly restores all adjusted item rows back to their optimal system-calculated minimums.
* **Dynamic Visual Formatting:** Custom static gradient rendering using native VBA color mapping, optimized for large datasets (5000+ rows) without the performance overhead of standard Excel Conditional Formatting.
* **One-Click System Export:** Instantly converts approved order rows into a clean, ERP-ready CSV file via `FileSystemObject` handling.

## 🛠️ Tech Stack

* **Excel Front-End:** Event-driven user interface with automated data validation inputs.
* **VBA (Visual Basic for Applications):** Advanced memory handling, dynamic UI rendering, and event automation (`Worksheet_Change`).
* **Power Query (M Language):** ETL process for extracting, transforming, and loading decentralized price tables.

## 🧩 Core Modules

### 1. The Price Engine (`Analiza`)
The primary module evaluates multiple pricing tiers and acquisition channels to dynamically determine the optimal buying strategy. The data architecture includes:
* **Stock Price:** Baseline fixed price for standard delivery terms.
* **Urgent Price:** Premium fixed price for expedited/urgent delivery.
* **Retail Price:** Benchmark end-user price used for margin calculations.
* **Campaign Price:** Promotional pricing applied to a specific pool of items, updated on a quarterly basis.
* **Negotiated Price & Date:** Historical prices secured through direct supplier negotiations.
* **Bulk Price & Date:** Heavily discounted pricing triggered exclusively by large order volumes.
* **Sub-Wholesale:** Segmented pricing framework available for specific B2B customer networks.

### 2. Interactive RFQ & Negotiation Engine (`Zapytanie Ofertowe`)
The second module acts as an interactive Procurement Interface (Request for Quote builder) that allows procurement managers to dynamically compile and fine-tune supplier inquiries:
* **Automated Best-Price Selection:** By default, the system evaluates all hidden price matrices and automatically extracts the absolute minimum available price for the target order.
* **Smart Dropdown Overrides:** Powered by an event-driven macro, users can manually override the automated selection via a dynamic data-validated dropdown.
* **Manual Price Logic:** If a completely arbitrary price is typed into the final price column, the VBA engine captures the input, automatically overrides the source selection to `MANUAL`, and flushes out stale validation flags.
* **Visual Optimizations:** Optimal target prices are rendered with a high-contrast pastel green formatting (`ERP Success Green`), ensuring quick visibility during fast-paced negotiation processes.


📦 How to Test the Project
1. Clone this repository or download the Price_Engine.xlsm file [Download file]([https://githubusercontent.com](https://github.com/gnovac/auto-price-analyser/raw/refs/heads/master/Price_Analyser.xlsm).
2. Open the file in Microsoft Excel and Enable Macros.
3. Optional: The core data sheets are hidden for UI clarity. To view the raw schema, unhide the data sheets or inspect the data/ folder in this repository.
4. Go to the Analiza sheet, enter any part number from the provided [Download file]( https://github.com/gnovac/auto-price-analyser/raw/refs/heads/master/data/cennik_fake_data.xlsx) samples into Column B, and see the automated price mapping and styling engine in action.
5. Navigate to the Zapytanie Ofertowe sheet to interact with the RFQ builder, test the manual price overrides, and use the global reset button.



