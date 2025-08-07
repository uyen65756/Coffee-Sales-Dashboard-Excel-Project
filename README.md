
This project demonstrates the complete workflow of creating a coffee sales dashboard:

1. **Data Gathering**  
   - Import raw orders, customers, and products data.
   - Use `XLOOKUP` to populate fields like Customer Name, Email, Country, Loyalty Card.
   - Use `INDEX` + `MATCH` to dynamically retrieve Product details (coffee type, roast, size, price, etc.).

2. **Data Wrangling & Formatting**  
   - Calculate `Sales = Unit Price × Quantity`.
   - Convert shorthand codes to full names (e.g., “ROB” → “Robusta”, roast codes to light/medium/dark).
   - Apply user-friendly formatting:
     - Date as `DD‑MMM‑YYYY`
     - Size in kilos (e.g., `0.5 kilo`)
     - Prices and Sales in USD with currency formatting
   - Remove duplicates and convert the data range into an Excel Table (`Ctrl+T`) for dynamic referencing.

3. **Dashboard Creation**  
   - Build pivot tables/charts:
     - **Line chart** for total sales over time, segmented by coffee type.
     - **Bar chart** for sales by country.
     - **Top‑5 customers** bar chart.
   - Insert interactive elements:
     - **Timeline** for date filtering.
     - **Slicers** for roast type, size, and loyalty card.
   - Apply cohesive formatting (purple-themed styles, fonts, borders).
   - Layout everything on a single “Dashboard” worksheet, with usable UI:
     - Hide gridlines, formula bar, sheet tabs, and (optionally) scrollbars for a clean presentation.
     - Ensure all filters (timeline + slicers) control every chart via “Report Connections.”
