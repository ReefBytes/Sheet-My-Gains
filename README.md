# Sheet-My-Gains

📈 Automatically sync your Robinhood portfolio, trade history, and dividend data to Google Sheets.

---

Tired of the limited analytics in the Robinhood app? Sheet-My-Gains gives you the power to track your entire investment portfolio with the full flexibility of Google Sheets. This Google Apps Script fetches your holdings, order history, and dividend data, allowing you to create powerful, personalized dashboards and automate your financial tracking.

## Core Features:

* **Sync Positions**: Pull all your current stock, ETF, and crypto positions.
* **Fetch Order History**: Keep a complete log of your buy and sell orders.
* **Track Dividends**: Automatically log all dividend payouts.
* **Scheduled Triggers**: Set up the script to run automatically on a daily or hourly basis.
* **Fully Customizable**: Since the data is in your own Google Sheet, you can build any chart, formula, or report you can imagine.

## How to Add the App Script to Your Google Sheet:

To use Sheet-My-Gains, you'll need to add the provided Google Apps Script code to your Google Sheet:

1.  **Open your Google Sheet**: Go to [Google Sheets](https://docs.google.com/spreadsheets/u/0/).
2.  **Access App Script**: From your spreadsheet, click `Extensions` > `Apps Script`.
3.  **Create a New Script**: A new Apps Script project will open in a new tab. If there's any default code (like `function myFunction() {}`), you can delete it.
4.  **Copy the Code**: Copy the entire content of the `sheet_my_gains.gs` file (provided in the `Sheet-My-Gains-initialization/Sheet-My-Gains/` directory) and paste it into the script editor, replacing any existing code.
5.  **Save the Script**: Click the floppy disk icon (Save project) or press `Ctrl + S` (Windows) / `Cmd + S` (Mac). You might be prompted to name your project; you can name it "Sheet-My-Gains" or anything you prefer.
6.  **Refresh Google Sheet**: Go back to your Google Sheet tab and refresh the page. You should now see a new "Robinhood" menu item in your spreadsheet.

## Available Functions:

Once the script is added, you can use the following functions directly in your Google Sheet cells as custom formulas, or interact with them via the "Robinhood" menu.

### Menu Functions:

* **`Robinhood` > `Login / Re-login`**:
    * **Purpose**: Initiates the interactive login process for your Robinhood account. This function clears any old access tokens and guides you through entering your credentials and handling any multi-factor authentication (MFA) challenges.
* **`Robinhood` > `Refresh Data`**:
    * **Purpose**: Manually triggers a refresh of all data points in your sheet that rely on the `LastUpdate` named range. It updates the timestamp in cell A1 of the hidden "Refresh" sheet, prompting functions using `LastUpdate` to recalculate.

### Custom Google Sheet Functions:

These functions are designed to be used directly in your Google Sheet cells. Most require a `LastUpdate` parameter, which is a named range automatically set up by the script to enable automatic refreshing.

* **`ROBINHOOD_GET_ACH_TRANSFERS(LastUpdate)`**
    * **Purpose**: Retrieves a history of all ACH transfers associated with your Robinhood account.
    * **Usage**: `=ROBINHOOD_GET_ACH_TRANSFERS(LastUpdate)`

* **`ROBINHOOD_GET_DIVIDENDS(LastUpdate)`**
    * **Purpose**: Fetches the complete dividend history for your Robinhood account.
    * **Usage**: `=ROBINHOOD_GET_DIVIDENDS(LastUpdate)`

* **`ROBINHOOD_GET_DOCUMENTS(LastUpdate)`**
    * **Purpose**: Retrieves a list of available documents from Robinhood, such as statements and tax forms.
    * **Usage**: `=ROBINHOOD_GET_DOCUMENTS(LastUpdate)`

* **`ROBINHOOD_GET_OPTIONS_ORDERS(LastUpdate)`**
    * **Purpose**: Provides a detailed history of your options orders.
    * **Usage**: `=ROBINHOOD_GET_OPTIONS_ORDERS(LastUpdate)`

* **`ROBINHOOD_GET_OPTIONS_POSITIONS(LastUpdate)`**
    * **Purpose**: Retrieves all current options positions held in your Robinhood account.
    * **Usage**: `=ROBINHOOD_GET_OPTIONS_POSITIONS(LastUpdate)`

* **`ROBINHOOD_GET_ORDERS(days, page_size, LastUpdate)`**
    * **Purpose**: Fetches your stock order history. It can be optionally filtered to look back a specific number of `days` or by `page_size`.
    * **Parameters**:
        * `days` (Optional): Number of days to look back. If `0` or omitted, all orders are returned.
        * `page_size` (Optional): Number of items to return per page.
        * `LastUpdate` (Required): Use the `LastUpdate` named range for automatic refreshing.
    * **Usage**: `=ROBINHOOD_GET_ORDERS(0, 1000, LastUpdate)` or `=ROBINHOOD_GET_ORDERS(30, , LastUpdate)`

* **`ROBINHOOD_GET_PORTFOLIOS(LastUpdate)`**
    * **Purpose**: Retrieves your portfolio data, including account value and historical performance.
    * **Usage**: `=ROBINHOOD_GET_PORTFOLIOS(LastUpdate)`

* **`ROBINHOOD_GET_POSITIONS(LastUpdate)`**
    * **Purpose**: Fetches all current stock positions (equities, ETFs, crypto) held in your Robinhood account.
    * **Usage**: `=ROBINHOOD_GET_POSITIONS(LastUpdate)`

* **`ROBINHOOD_GET_WATCHLIST(LastUpdate)`**
    * **Purpose**: Retrieves the instruments (stocks, ETFs, crypto) from your default Robinhood watchlist.
    * **Usage**: `=ROBINHOOD_GET_WATCHLIST(LastUpdate)`

* **`ROBINHOOD_GET_QUOTE(ticker, includeHeader, LastUpdate)`**
    * **Purpose**: Retrieves the latest quote data for a given stock ticker symbol.
    * **Parameters**:
        * `ticker` (Required): The stock ticker symbol (e.g., "AAPL").
        * `includeHeader` (Optional): Set to `FALSE` to exclude the header row from the output. Defaults to `TRUE`.
        * `LastUpdate` (Required): Use the `LastUpdate` named range for automatic refreshing.
    * **Usage**: `=ROBINHOOD_GET_QUOTE("MSFT", TRUE, LastUpdate)`

* **`ROBINHOOD_GET_HISTORICALS(ticker, interval, span, LastUpdate)`**
    * **Purpose**: Retrieves historical price data for a specified stock ticker.
    * **Parameters**:
        * `ticker` (Required): The stock ticker symbol (e.g., "TSLA").
        * `interval` (Optional): The time interval for data points ('day', 'week', 'month'). Default is 'day'.
        * `span` (Optional): The total time span for the historical data ('week', 'month', '3month', 'year', '5year'). Default is 'year'.
        * `LastUpdate` (Required): Use the `LastUpdate` named range for automatic refreshing.
    * **Usage**: `=ROBINHOOD_GET_HISTORICALS("GOOG", "day", "year", LastUpdate)`

* **`ROBINHOOD_GET_ACCOUNTS(LastUpdate)`**
    * **Purpose**: Retrieves detailed information for all your brokerage accounts within Robinhood.
    * **Usage**: `=ROBINHOOD_GET_ACCOUNTS(LastUpdate)`
