# Sheet-My-Gains

📈 Automatically sync your Robinhood portfolio, trade history, and dividend data to Google Sheets.

---

Tired of the limited analytics in the Robinhood app? Sheet-My-Gains gives you the power to track your entire investment portfolio with the full flexibility of Google Sheets. This Google Apps Script fetches your holdings, order history, and dividend data, allowing you to create powerful, personalized dashboards and automate your financial tracking.

## Core Features

*   **Comprehensive Data Sync**: Pull current stock, ETF, and options positions.
*   **Complete History**: Fetch your entire order history and dividend payouts.
*   **Watchlist Management**: Retrieve and view all your custom watchlists.
*   **Portfolio Tracking**: Analyze historical portfolio performance over time.
*   **Real-Time Quotes**: Get real-time price data for single or multiple tickers.
*   **Automatic Refresh**: Set up the script to run on a schedule and keep your data current.
*   **Fully Customizable**: Since the data is in your Google Sheet, you can build any chart, formula, or report you can imagine.

## Setup and Usage

Follow these steps to add the script to your Google Sheet.

### 1. Copy the Script Files

1.  **Open Your Google Sheet**: Navigate to [Google Sheets](https://docs.google.com/spreadsheets/u/0/) and open the spreadsheet you want to use.
2.  **Open the Apps Script Editor**: In the menu, click `Extensions` > `Apps Script`.
3.  **Add `sheet_my_gains.gs`**:
    *   Rename the default `Code.gs` file to `sheet_my_gains.gs`.
    *   Copy the entire content of the `Sheet-My-Gains/sheet_my_gains.gs` file from this repository.
    *   Paste it into the `sheet_my_gains.gs` file in the editor, replacing all existing content.
4.  **Add `LoginDialog.html`**:
    *   In the editor, click the `+` icon next to "Files" and select `HTML`.
    *   Name the file `LoginDialog.html` (without the `.html` extension in the input box).
    *   Copy the entire content of `Sheet-My-Gains/LoginDialog.html` from this repository.
    *   Paste it into the new `LoginDialog.html` file in the editor.
5.  **Configure the Manifest File**:
    *   In the editor, click the `Project Settings` (gear) icon on the left.
    *   Check the box for **`Show "appsscript.json" manifest file in editor`**.
    *   Return to the editor view and click on the `appsscript.json` file.
    *   Copy the content of the `Sheet-My-Gains/appsscript.json` file from this repository and paste it into the editor, replacing all existing content.
6.  **Save Your Project**: Click the `Save project` (floppy disk) icon.

### 2. Grant Permissions and Log In

1.  **Refresh Your Google Sheet**: Go back to your spreadsheet and refresh the page. A new **"Robinhood"** menu will appear.
2.  **Grant Permissions**:
    *   Click `Robinhood` > `Grant Permissions (Run First)`.
    *   Follow the on-screen prompts to authorize the script. This is a **critical step**. You only need to do this once.
3.  **Log In**:
    *   Click `Robinhood` > `Login / Re-login`.
    *   A dialog box will appear. Enter your Robinhood credentials.
    *   If you use Multi-Factor Authentication (MFA), you will be prompted to approve the login on your device.

### 3. Using the Functions

You can now use the custom functions in any cell of your sheet. Most functions require a `LastUpdate` parameter to enable automatic refreshing. The script automatically creates a hidden `Refresh` sheet and a named range called `LastUpdate` to manage this.

To refresh your data, either use `Robinhood` > `Refresh Data` or set up a time-based trigger in the Apps Script editor (`Triggers` > `Add Trigger`).

## Available Functions

### Menu Functions

*   **`Robinhood` > `Grant Permissions (Run First)`**: Authorizes the script to access your spreadsheet and connect to Robinhood's API. **Must be run once before first use.**
*   **`Robinhood` > `Login / Re-login`**: Opens a dialog to securely enter your Robinhood credentials and handle MFA.
*   **`Robinhood` > `Refresh Data`**: Manually triggers a data refresh for all functions using the `LastUpdate` parameter.
*   **`Robinhood` > `Show All Functions`**: Creates a new sheet named "Function Help" with a list of all available functions and their descriptions.

### Custom Sheet Functions

#### Account & Portfolio

*   **`ROBINHOOD_GET_ACCOUNTS(LastUpdate)`**
    *   **Purpose**: Retrieves detailed information for all your brokerage accounts.
    *   **Usage**: `=ROBINHOOD_GET_ACCOUNTS(LastUpdate)`
*   **`ROBINHOOD_GET_PORTFOLIOS(LastUpdate)`**
    *   **Purpose**: Retrieves high-level data for your portfolios, including market value.
    *   **Usage**: `=ROBINHOOD_GET_PORTFOLIOS(LastUpdate)`
*   **`ROBINHOOD_GET_PORTFOLIO_HISTORICALS(span, interval, accountNumber, LastUpdate)`**
    *   **Purpose**: Retrieves historical portfolio value for tracking performance over time.
    *   **Parameters**:
        *   `span` (Optional): "day", "week", "month", "3month", "year" (default), "5year".
        *   `interval` (Optional): "5minute", "day" (default), "week".
        *   `accountNumber` (Optional): Filter by a specific account number.
        *   `LastUpdate` (Required): The `LastUpdate` named range.
    *   **Usage**: `=ROBINHOOD_GET_PORTFOLIO_HISTORICALS("year", "day", LastUpdate)`

#### Positions & Orders

*   **`ROBINHOOD_GET_POSITIONS(LastUpdate)`**
    *   **Purpose**: Fetches all current stock, ETF, and REIT positions.
    *   **Usage**: `=ROBINHOOD_GET_POSITIONS(LastUpdate)`
*   **`ROBINHOOD_GET_ORDERS(days, pageSize, LastUpdate)`**
    *   **Purpose**: Fetches your stock order history.
    *   **Parameters**:
        *   `days` (Optional): Number of days to look back. Default is `30`. Use `0` for all orders.
        *   `pageSize` (Optional): Number of items per page. Default is `1000`.
        *   `LastUpdate` (Required): The `LastUpdate` named range.
    *   **Usage**: `=ROBINHOOD_GET_ORDERS(90, 500, LastUpdate)`
*   **`ROBINHOOD_GET_OPTIONS_POSITIONS(LastUpdate)`**
    *   **Purpose**: Retrieves all current options positions.
    *   **Usage**: `=ROBINHOOD_GET_OPTIONS_POSITIONS(LastUpdate)`
*   **`ROBINHOOD_GET_OPTIONS_ORDERS(days, pageSize, LastUpdate)`**
    *   **Purpose**: Provides a detailed history of your options orders.
    *   **Parameters**:
        *   `days` (Optional): Number of days to look back. Default is `30`.
        *   `pageSize` (Optional): Number of items per page. Default is `100`.
        *   `LastUpdate` (Required): The `LastUpdate` named range.
    *   **Usage**: `=ROBINHOOD_GET_OPTIONS_ORDERS(90, 100, LastUpdate)`

#### Market Data & Quotes

*   **`ROBINHOOD_GET_QUOTE(ticker, includeHeader, LastUpdate)`**
    *   **Purpose**: Retrieves the latest quote for a single stock ticker.
    *   **Usage**: `=ROBINHOOD_GET_QUOTE("AAPL", TRUE, LastUpdate)`
*   **`ROBINHOOD_GET_QUOTES_BATCH(tickers, LastUpdate)`**
    *   **Purpose**: Retrieves quotes for multiple stock tickers in a single call.
    *   **Parameters**:
        *   `tickers` (Required): A comma-separated string of ticker symbols (e.g., "AAPL,TSLA,GOOG").
        *   `LastUpdate` (Required): The `LastUpdate` named range.
    *   **Usage**: `=ROBINHOOD_GET_QUOTES_BATCH("AAPL,TSLA,GOOG", LastUpdate)`
*   **`ROBINHOOD_GET_HISTORICALS(ticker, interval, span, LastUpdate)`**
    *   **Purpose**: Retrieves historical price data for a specified stock ticker.
    *   **Parameters**:
        *   `ticker` (Required): The stock ticker symbol (e.g., "TSLA").
        *   `interval` (Optional): "5minute", "day" (default), "week".
        *   `span` (Optional): "day", "week", "month", "3month", "year" (default), "5year".
        *   `LastUpdate` (Required): The `LastUpdate` named range.
    *   **Usage**: `=ROBINHOOD_GET_HISTORICALS("TSLA", "day", "year", LastUpdate)`

#### Dividends & Transfers

*   **`ROBINHOOD_GET_DIVIDENDS(LastUpdate)`**
    *   **Purpose**: Fetches your complete dividend history.
    *   **Usage**: `=ROBINHOOD_GET_DIVIDENDS(LastUpdate)`
*   **`ROBINHOOD_GET_ACH_TRANSFERS(LastUpdate)`**
    *   **Purpose**: Retrieves a history of all ACH transfers.
    *   **Usage**: `=ROBINHOOD_GET_ACH_TRANSFERS(LastUpdate)`

#### Watchlists

*   **`ROBINHOOD_GET_WATCHLISTS(LastUpdate)`**
    *   **Purpose**: Lists all your available watchlists.
    *   **Usage**: `=ROBINHOOD_GET_WATCHLISTS(LastUpdate)`
*   **`ROBINHOOD_GET_WATCHLIST(watchlistNameOrId, LastUpdate)`**
    *   **Purpose**: Retrieves all instruments from a specific watchlist.
    *   **Usage**: `=ROBINHOOD_GET_WATCHLIST("My Watchlist", LastUpdate)`
*   **`ROBINHOOD_GET_ALL_WATCHLISTS(LastUpdate)`**
    *   **Purpose**: Retrieves all instruments from all of your watchlists combined.
    *   **Usage**: `=ROBINHOOD_GET_ALL_WATCHLISTS(LastUpdate)`

#### Utility Functions

*   **`ROBINHOOD_GET_LOGIN_STATUS(LastUpdate)`**
    *   **Purpose**: Checks and returns the current authentication status ("Logged In" or "Logged Out").
    *   **Usage**: `=ROBINHOOD_GET_LOGIN_STATUS(LastUpdate)`
*   **`ROBINHOOD_GET_URL(url, includeHeader, LastUpdate)`**
    *   **Purpose**: Retrieves data from any valid Robinhood API URL.
    *   **Usage**: `=ROBINHOOD_GET_URL("https://api.robinhood.com/marketdata/quotes/AAPL/", TRUE, LastUpdate)`
*   **`ROBINHOOD_HELP(category)`**
    *   **Purpose**: Displays a list of all available functions and their descriptions.
    *   **Usage**: `=ROBINHOOD_HELP("core")`
*   **`ROBINHOOD_VALIDATE_TICKER(ticker)`**
    *   **Purpose**: Validates if a string is a valid stock ticker symbol format. Returns TRUE or FALSE.
    *   **Usage**: `=ROBINHOOD_VALIDATE_TICKER("AAPL")`
*   **`ROBINHOOD_FORMAT_CURRENCY(amount)`**
    *   **Purpose**: Formats a number into USD currency format (e.g., $1,234.56).
    *   **Usage**: `=ROBINHOOD_FORMAT_CURRENCY(1234.56)`
*   **`ROBINHOOD_LAST_MARKET_DAY()`**
    *   **Purpose**: Returns the most recent market trading day (excludes weekends).
    *   **Usage**: `=ROBINHOOD_LAST_MARKET_DAY()`

---

**Disclaimer**: This tool is not affiliated with Robinhood. Use it at your own risk.
