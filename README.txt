TAQUION CASHFLOW SYSTEM v3
==========================

Application Type: Single-file HTML5 app with embedded JSON data
Size: 802 KB
Framework: Vanilla JavaScript + Chart.js CDN
Theme: Dark mode (GitHub-style)

FEATURES IMPLEMENTED:
=====================

1. DASHBOARD
   - KPI cards (Ingresos, Gastos, Flujo Neto del último trimestre)
   - Reference banner (Real: Mar 2024 - Ago 2024 | Proyectado: Sep 2024+)
   - Quarterly summary table with Real/Proyectado badges
   - Bar chart showing income vs expenses by quarter

2. OPERACIÓN BAU
   - Line chart of bancarizado + no bancarizado flows
   - Quarterly breakdown table
   - Excludes bank and GC debt

3. DEUDA BANCARIA
   - KPI cards: TQN Capital, TQN Interest, LMS Total
   - TQN calendar with capital vs interest separated
   - LMS monthly payment table
   - Stacked bar chart of TQN obligations

4. DEUDA GC (NEW - From Term Sheet Anexo D)
   - KPI cards: Original USD debt, Cedido, Reestructurada, Pagado, Pendiente
   - Editable USD→ARS conversion rate (default 1250)
   - Full cuota schedule: Fecha | USD | ARS (~) | Destino | Estado (Pagada/Pendiente)
   - Summary by destino (GMC vs Inversora GJ)
   - Timeline chart of remaining payments
   - 24-cuota schedule with payment status

5. B / N SPLITS
   - Dual comparison: Bancarizado vs No Bancarizado
   - Separate bar charts and tables for each
   - Income/expense breakdown

6. WHAT-IF SIMULATOR
   - Income change slider (%)
   - Expense change slider (%)
   - Month shift parameter
   - KPI impact analysis
   - Dynamic chart updates

7. CARGA DE DATOS
   - Form to manually add new movements
   - Fields: Fecha, Empresa (TQN/LMS), Tipo, Bancarizado, Monto, Descripción
   - In-memory storage (session-based)
   - CSV export functionality

8. MOVIMIENTOS
   - Searchable/filterable table of all 4,072 movement records
   - Filter by type or concept
   - Sort by date
   - Uses Movimiento_Orig values (full pesos, not divided by 1000)

DATA STRUCTURE:
===============
- Quarters (q): 12 quarters with ib/eb/in/en_ (bancarizado/no-bancarizado, income/expense)
  Plus breakdown by Tipo (9 categories) and Empresa (TQN/LMS)
- Monthly (mes): 32 months with similar structure
- Movements (mov): 4,072 individual transaction records
- Bank Debt TQN (bd): 32 months of capital + interest schedule
- Bank Debt LMS (bdl): 24 months of payment schedule
- GC Term Sheet (gc_term): NEW - 24 cuotas with payment status, destino, USD amounts
  - Original debt composition tracked
  - Summary of paid (8 cuotas) vs pending (16 cuotas)

STYLING:
========
- Dark theme: #0f1117 (bg), #1a1d27 (surface), #4f8cff (accent)
- Color coding: #34d399 (positive/green), #f87171 (negative/red)
- Monospace font for numbers
- Responsive grid layout
- Hover effects and smooth transitions

FORMATTING:
===========
- ARS: $ with thousands separator (.) e.g., $1.234.567
- USD: US$ with 2 decimals e.g., US$1,234.56
- Positive: green text
- Negative: red text
- Quarterly dates include Real/Proyectado badges
- All tables sort-friendly, chart-optimized

USAGE:
======
1. Open Cashflow_Sistema_Taquion.html in any modern browser
2. Navigate using the sidebar (8 pages)
3. View data across quarterly, monthly, and transaction levels
4. Adjust exchange rate in GC debt page for ARS equivalent
5. Run what-if scenarios with income/expense adjustments
6. Add new movements via Carga page (stored in browser session)
7. Export movements as CSV for further analysis

TECHNICAL:
===========
- CDN: Chart.js 4.4.0 (for charts)
- Minified JSON embedded in HTML (file size optimized)
- No build tools required - open HTML file directly
- No external API calls
- Works offline after initial load
- LocalStorage not used (session-only for new movements)

NOTES:
======
- GC debt section completely rebuilt with Anexo D term sheet data
- All amounts use Movimiento_Orig (full pesos) - not divided
- Real data period: March 2024 - August 2024
- Projected data: September 2024 onwards
- Exchange rate default: 1,250 ARS/USD (user-editable)

File: /sessions/sweet-gracious-mayer/mnt/Taquion Financiero/Cashflow_Sistema_Taquion.html
Version: 3.0
Built: 2026-04-04
Data Source: cashflow_v3.json (772 KB)
