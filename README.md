# Talkguest Table Downloader

A Tampermonkey userscript that adds a "Download CSV" button to the Talkguest Owner portal. This script extracts reservation data, including detailed fee breakdowns from tooltips, and automatically calculates the profit, management commission, and final totals.

## Features

- **One-Click Export**: Adds a floating "Download CSV" button to the reservations page.
- **Tooltip Extraction**: Automatically hovers over reservation "i" icons to dynamically extract hidden line items (such as `Taxa Municipal Turística`, `Cleaning Fee`, `Comissão do canal`, `Comissão de pagamento`, etc.) into dynamic columns.
- **IVA Calculation**: Computes the included 6% IVA on the reservation base value.
- **Profit & Commission Calculations**:
  - **Profit / Lucro**: Calculated as the `Valor da Reserva` minus specific expenses (`Comissão do canal` and `Comissão de pagamento`) and minus the calculated 6% IVA. Note: `Taxa Municipal Turística` and `Cleaning Fees` are strictly excluded from the expense deduction.
  - **Comissão de Gestão (30%)**: Automatically calculated as 30% of the Profit.
  - **Total Final**: The remaining value after deducting the management commission from the profit (`Profit` - `Comissão de Gestão`).
- **Totals Row**: Automatically appends a `TOTAL GERAL` row at the very bottom of the CSV to sum up the Final Totals for the visible page.
- **Excel Compatibility**: Adds a UTF-8 BOM so the generated CSV opens cleanly in Microsoft Excel without character encoding issues (preserving accents like 'ç' and 'ã').

## Installation

1. Install the [Tampermonkey](https://www.tampermonkey.net/) extension for your browser (Chrome, Firefox, Safari, Edge).
2. Open the Tampermonkey dashboard and click on **Add a new script**.
3. Copy the entire contents of `talkguest-downloader.user.js` and paste it into the editor, completely replacing any default template code.
4. Save the script (`File` -> `Save`, or `Ctrl+S` / `Cmd+S`).

## Usage

1. Log in to your Talkguest account.
2. Navigate to the Owner Area Bookings page (`https://owner.talkguest.com/Bookings/OwnerArea.aspx*`).
3. Once the page loads, you should see a blue **Download CSV** button in the bottom right corner of the screen.
4. Click the button. The button text will change to "Loading details..." as it simulates hovering over the tooltips to load the full breakdown data (this may take a few seconds depending on how many reservations are visible).
5. Once complete, a `.csv` file will be downloaded to your computer automatically, timestamped with the current date (e.g., `talkguest_reservations_YYYY-MM-DD.csv`).

## Technical Details

- **Language**: Vanilla JavaScript (ES6+).
- **DOM Interaction**: The script interacts with OutSystems-generated DOM elements (classes like `.TableRecords`, `.tooltip-widget`, `.OSInline`).
- **Asynchronous Execution**: Because tooltip data is lazy-loaded via an Ajax request (`OsAjax`), the script dispatches a `mouseenter` event and polls the DOM to wait for the data to populate before parsing it.
- **Persistent Injector**: Uses `setInterval` to periodically ensure the button remains injected on the page, bypassing issues with dynamic DOM reloading in Single Page Applications (SPAs).
