# Talkguest Table Downloader

A Tampermonkey userscript that adds a "Download XLSX" button to the Talkguest Owner portal. This script extracts reservation data, including detailed fee breakdowns from tooltips, and automatically calculates Vendas, management commission, and final totals.

## Features

- **One-Click Export**: Adds a floating "Download XLSX" button to the reservations page.
- **Tooltip Extraction**: Automatically hovers over reservation "i" icons to dynamically extract hidden line items (such as `Taxa Municipal Turística`, `Cleaning Fee`, `Comissão do canal`, `Comissão de pagamento`, etc.) into dynamic columns.
- **Checkout Sorting**: Sorts exported reservations by checkout date in ascending order.
- **IVA Calculation**: Computes the included 6% IVA on the Vendas base, excluding city tax and cleaning fees.
- **Vendas & Commission Calculations**:
  - **Vendas**: Calculated from the displayed total minus city tax and cleaning fees, with included 6% IVA removed, then minus `Comissão do canal` and `Comissão de pagamento`.
  - **Comissão de Gestão (30%)**: Automatically calculated as 30% of `Vendas`.
  - **Total Final**: The remaining value after deducting the management commission from `Vendas`.
  - **Despesas Proprietário / Despesas ORM**: Summary expense columns initialized to `0` on the totals row.
  - **Total a Transferir**: Calculated on the totals row as `Total Final - Despesas Proprietário - Despesas ORM`.
- **Totals Row**: Automatically appends a `TOTAL GERAL` row at the very bottom of the XLSX to sum up Vendas, expenses, commission, final total, and transfer columns for the visible page.

## Installation

1. Install the [Tampermonkey](https://www.tampermonkey.net/) extension for your browser (Chrome, Firefox, Safari, Edge).
2. Open the Tampermonkey dashboard and click on **Add a new script**.
3. Copy the entire contents of `talkguest-downloader.user.js` and paste it into the editor, completely replacing any default template code.
4. Save the script (`File` -> `Save`, or `Ctrl+S` / `Cmd+S`).

## Usage

1. Log in to your Talkguest account.
2. Navigate to the Owner Area Bookings page (`https://owner.talkguest.com/Bookings/OwnerArea.aspx*`).
3. Once the page loads, you should see a blue **Download XLSX** button in the bottom right corner of the screen.
4. Click the button. The button text will change to "Loading details..." as it simulates hovering over the tooltips to load the full breakdown data (this may take a few seconds depending on how many reservations are visible).
5. Once complete, a `.xlsx` file will be downloaded to your computer automatically, timestamped with the current date (e.g., `talkguest_reservations_YYYY-MM-DD.xlsx`).

## Technical Details

- **Language**: Vanilla JavaScript (ES6+).
- **DOM Interaction**: The script interacts with OutSystems-generated DOM elements (classes like `.TableRecords`, `.tooltip-widget`, `.OSInline`).
- **Asynchronous Execution**: Because tooltip data is lazy-loaded via an Ajax request (`OsAjax`), the script dispatches a `mouseenter` event and polls the DOM to wait for the data to populate before parsing it.
- **Persistent Injector**: Uses `setInterval` to periodically ensure the button remains injected on the page, bypassing issues with dynamic DOM reloading in Single Page Applications (SPAs).
