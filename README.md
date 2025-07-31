# Sheet-My-Gains

📈 Automatically sync your Robinhood portfolio, trade history, and dividend data to Google Sheets.

---

Tired of the limited analytics in the Robinhood app? Sheet-My-Gains gives you the power to track your entire investment portfolio with the full flexibility of Google Sheets. This Google Apps Script fetches your holdings, order history, and dividend data, allowing you to create powerful, personalized dashboards and automate your financial tracking.

## Core Features:

* **Sync Positions**: Pull all your current stock and ETF positions.
* **Fetch Order History**: Keep a complete log of your buy and sell orders.
* **Track Dividends**: Automatically log all dividend payouts.
* **Scheduled Triggers**: Set up the script to run automatically on a daily or hourly basis.
* **Fully Customizable**: Since the data is in your own Google Sheet, you can build any chart, formula, or report you can imagine.

## How to Add the App Script to Your Google Sheet:

To use Sheet-My-Gains, you'll need to add the provided Google Apps Script code to your Google Sheet:

1.  **Open Your Google Sheet**: Navigate to [Google Sheets](https://docs.google.com/spreadsheets/u/0/) and open or create a new spreadsheet.
2.  **Access Apps Script**: In the menu bar of your spreadsheet, click `Extensions` > `Apps Script`. This will open the Apps Script editor in a new tab.
3.  **Set Up a New Script**: If the editor contains any default code (like `function myFunction() {}`), delete it to start with a blank script.
4.  **Add the Main Script**: Locate the `sheet_my_gains.gs` file in the `Sheet-My-Gains-initialization/Sheet-My-Gains/` directory. Copy its entire content and paste it into the Apps Script editor, replacing any existing code.
5.  **Save the Script**: Click the save icon (floppy disk) or press `Ctrl + S` (Windows) / `Cmd + S` (Mac). When prompted, name the project "Sheet-My-Gains" or any name of your choice.
6.  **Add the HTML File**: In the Apps Script editor, create a new file by clicking the `+` icon and selecting `HTML`. Name the file `LoginDialog.html`. Then, copy the content of the `LoginDialog.html` file from the `Sheet-My-Gains-initialization/Sheet-My-Gains/` directory and paste it into this new file.
7.  **Save the HTML File**: Save the `LoginDialog.html` file by clicking the save icon or pressing `Ctrl + S` (Windows) / `Cmd + S` (Mac).
8.  **Refresh Your Spreadsheet**: Return to your Google Sheet and refresh the page. You should now see a new "Robinhood" menu item in the spreadsheet toolbar.

## Available Functions:

Once the script is added, you can use the following functions directly in your Google Sheet cells as custom formulas, or interact with them via the "Robinhood" menu.

### Menu Functions:

* **`Robinhood` > `Login / Re-login`**:
    * **Purpose**: Initiates the interactive login process for your Robinhood account. This function clears any old access tokens and guides you through entering your credentials and handling any multi-factor authentication (MFA) challenges.
* **`Robinhood` > `Refresh Data`**:
    * **Purpose**: Manually triggers a refresh of all data points in your sheet that rely on the `LastUpdate` named range. It updates the timestamp in cell A1 of the hidden "Refresh" sheet, prompting functions using `LastUpdate` to recalculate.

### Custom Google Sheet Functions:

These functions are designed to be used directly in your Google Sheet cells. Most require a `LastUpdate` parameter, which is a named range automatically set up by the script to enable automatic refreshing.

* **`ROBINHOOD_GET_LOGIN_STATUS(LastUpdate)`**
    * **Purpose**: Checks the current authentication status and returns "Logged In" or "Logged Out".
    * **Usage**: `=ROBINHOOD_GET_LOGIN_STATUS(LastUpdate)`

* **`ROBINHOOD_GET_URL(url, LastUpdate, includeHeader)`**
    * **Purpose**: Retrieves data from a specific Robinhood API URL. Useful for advanced queries not covered by other functions.
    * **Parameters**:
        * `url` (Required): The full Robinhood API URL.
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
        * `includeHeader` (Optional): Set to `FALSE` to exclude the header row.
    * **Usage**: `=ROBINHOOD_GET_URL("https://api.robinhood.com/...", LastUpdate, FALSE)`

* **`ROBINHOOD_GET_ACCOUNTS(LastUpdate)`**
    * **Purpose**: Retrieves detailed information for all your brokerage accounts within Robinhood.
    * **Usage**: `=ROBINHOOD_GET_ACCOUNTS(LastUpdate)`

* **`ROBINHOOD_GET_ACH_TRANSFERS(accountNumber, LastUpdate)`**
    * **Purpose**: Retrieves a history of all ACH transfers.
    * **Parameters**:
        * `accountNumber` (Optional): Account number to filter results. If omitted, returns data for all accounts.
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Usage**: `=ROBINHOOD_GET_ACH_TRANSFERS(LastUpdate)` or `=ROBINHOOD_GET_ACH_TRANSFERS("5DP12345", LastUpdate)`

* **`ROBINHOOD_GET_DIVIDENDS(accountNumber, LastUpdate)`**
    * **Purpose**: Fetches the complete dividend history.
    * **Parameters**:
        * `accountNumber` (Optional): Account number to filter results. If omitted, returns data for all accounts.
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Usage**: `=ROBINHOOD_GET_DIVIDENDS(LastUpdate)` or `=ROBINHOOD_GET_DIVIDENDS("5DP12345", LastUpdate)`



* **`ROBINHOOD_GET_HISTORICALS(ticker, interval, span, LastUpdate)`**
    * **Purpose**: Retrieves historical price data for a specified stock ticker.
    * **Parameters**:
        * `ticker` (Required): The stock ticker symbol (e.g., "TSLA").
        * `interval` (Optional): Time interval - "5minute", "day" (default), "week"
        * `span` (Optional): Time span - "day", "week", "month", "3month", "year" (default), "5year"
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Usage**: `=ROBINHOOD_GET_HISTORICALS("GOOG", "day", "year", LastUpdate)`

* **`ROBINHOOD_GET_ORDERS(days, page_size, accountNumber, LastUpdate)`**
    * **Purpose**: Fetches your stock order history.
    * **Parameters**:
        * `days` (Optional): Number of days to look back. If `0` or omitted, all orders are returned.
        * `page_size` (Optional): Number of items to return per page.
        * `accountNumber` (Optional): Account number to filter results. If omitted, returns data for all accounts.
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Usage**: `=ROBINHOOD_GET_ORDERS(30, , LastUpdate)` or `=ROBINHOOD_GET_ORDERS(30, 1000, "5DP12345", LastUpdate)`

* **`ROBINHOOD_GET_OPTIONS_ORDERS(days, page_size, accountNumber, LastUpdate)`**
    * **Purpose**: Provides a detailed history of your options orders.
    * **Parameters**:
        * `days` (Optional): Number of days to look back. If `0` or omitted, all orders are returned.
        * `page_size` (Optional): Number of items to return per page (max: 1000). Default: 50.
        * `accountNumber` (Optional): Account number to filter results. If omitted, returns data for all accounts.
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Usage**: `=ROBINHOOD_GET_OPTIONS_ORDERS(30, 100, LastUpdate)` or `=ROBINHOOD_GET_OPTIONS_ORDERS(30, 100, "5DP12345", LastUpdate)`

* **`ROBINHOOD_GET_OPTIONS_POSITIONS(accountNumber, LastUpdate)`**
    * **Purpose**: Retrieves all current options positions held in your Robinhood account.
    * **Parameters**:
        * `accountNumber` (Optional): Account number to filter results. If omitted, returns data for all accounts.
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Usage**: `=ROBINHOOD_GET_OPTIONS_POSITIONS(LastUpdate)` or `=ROBINHOOD_GET_OPTIONS_POSITIONS("5DP12345", LastUpdate)`

* **`ROBINHOOD_GET_PORTFOLIOS(accountNumber, LastUpdate)`**
    * **Purpose**: Retrieves your portfolio data, including account value and historical performance.
    * **Parameters**:
        * `accountNumber` (Optional): Account number to filter results. If omitted, returns data for all accounts.
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Usage**: `=ROBINHOOD_GET_PORTFOLIOS(LastUpdate)` or `=ROBINHOOD_GET_PORTFOLIOS("5DP12345", LastUpdate)`

* **`ROBINHOOD_GET_POSITIONS(accountNumber, LastUpdate)`**
    * **Purpose**: Fetches all current stock positions (equities, ETFs, REITs).
    * **Parameters**:
        * `accountNumber` (Optional): Account number to filter results. If omitted, returns data for all accounts.
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Usage**: `=ROBINHOOD_GET_POSITIONS(LastUpdate)` or `=ROBINHOOD_GET_POSITIONS("5DP12345", LastUpdate)`

* **`ROBINHOOD_GET_QUOTE(ticker, includeHeader, LastUpdate)`**
    * **Purpose**: Retrieves the latest quote data for a given stock ticker symbol.
    * **Parameters**:
        * `ticker` (Required): The stock ticker symbol (e.g., "AAPL").
        * `includeHeader` (Optional): Set to `FALSE` to exclude the header row from the output.
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Usage**: `=ROBINHOOD_GET_QUOTE("MSFT", TRUE, LastUpdate)`

* **`ROBINHOOD_GET_WATCHLISTS(LastUpdate)`**
    * **Purpose**: Lists all your available watchlists with basic information.
    * **Usage**: `=ROBINHOOD_GET_WATCHLISTS(LastUpdate)`

* **`ROBINHOOD_GET_WATCHLIST(watchlistNameOrId, LastUpdate)`**
    * **Purpose**: Retrieves detailed instrument information from a specific watchlist with current market data.
    * **Parameters**:
        * `watchlistNameOrId` (Required): The name or ID of the watchlist (e.g., "General", "Buy Review").
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Returns**: Symbol, name, current price, daily change ($), daily change (%), market cap, state, tradability, date added
    * **Usage**: `=ROBINHOOD_GET_WATCHLIST("General", LastUpdate)`

* **`ROBINHOOD_GET_ALL_WATCHLISTS(LastUpdate)`**
    * **Purpose**: Retrieves instruments from ALL of your watchlists combined into one table with current market data.
    * **Returns**: Watchlist name, symbol, name, current price, daily change ($), daily change (%), market cap, state, tradability, date added
    * **Usage**: `=ROBINHOOD_GET_ALL_WATCHLISTS(LastUpdate)`

## New Enhanced Functions:

### Portfolio Performance & Analytics:

* **`ROBINHOOD_GET_PORTFOLIO_HISTORICALS(span, interval, accountNumber, LastUpdate)`**
    * **Purpose**: Retrieves portfolio historical performance data for tracking value over time.
    * **Parameters**:
        * `span` (Optional): Time span - "day", "week", "month", "3month", "year" (default), "5year"
        * `interval` (Optional): Time interval - "5minute", "day" (default), "week"
        * `accountNumber` (Optional): Account number to filter results. If omitted, returns data for all accounts combined.
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Usage**: `=ROBINHOOD_GET_PORTFOLIO_HISTORICALS("year", "day", LastUpdate)` or `=ROBINHOOD_GET_PORTFOLIO_HISTORICALS("year", "day", "5DP12345", LastUpdate)`




### Enhanced Quote Functions:

* **`ROBINHOOD_GET_QUOTES_BATCH(tickers, LastUpdate)`**
    * **Purpose**: Retrieves quotes for multiple stock tickers in a single efficient call.
    * **Parameters**:
        * `tickers` (Required): Comma-separated list of ticker symbols (e.g., "AAPL,MSFT,GOOGL").
        * `LastUpdate` (Required): Use the `LastUpdate` named range.
    * **Usage**: `=ROBINHOOD_GET_QUOTES_BATCH("AAPL,MSFT,GOOGL", LastUpdate)`

### Tax & Reporting:

Tax reporting functions have been removed due to calculation accuracy issues. Please use official tax documents from Robinhood for tax reporting purposes.

### Utility Functions:

* **`ROBINHOOD_VALIDATE_TICKER(ticker)`**
    * **Purpose**: Validates a stock ticker symbol format.
    * **Usage**: `=ROBINHOOD_VALIDATE_TICKER("AAPL")`

* **`ROBINHOOD_FORMAT_CURRENCY(amount)`**
    * **Purpose**: Formats a number as currency.
    * **Usage**: `=ROBINHOOD_FORMAT_CURRENCY(1234.56)`

* **`ROBINHOOD_LAST_MARKET_DAY()`**
    * **Purpose**: Gets the last market trading day (excluding weekends).
    * **Usage**: `=ROBINHOOD_LAST_MARKET_DAY()`
