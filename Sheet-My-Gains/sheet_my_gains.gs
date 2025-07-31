/**
 * A collection of constants for the Robinhood API.
 */
const ROBINHOOD_CONFIG = {
  API_BASE_URL: "https://api.robinhood.com",
  TOKEN_URL: "https://api.robinhood.com/oauth2/token/",
  CLIENT_ID: "c82SH0WZOsabOXGP2sxqcj34FxkvfnWRZBKlBjFS",
  API_URIS: {
    accounts: "/accounts/",
    achTransfers: "/ach/transfers/",
    dividends: "/dividends/",
    documents: "/documents/",
    marketData: "/marketdata/options/?instruments=",
    optionsOrders: "/options/orders/",
    optionsPositions: "/options/positions/",
    orders: "/orders/",
    portfolios: "/portfolios/",
    portfolioHistoricals: "/portfolios/historicals/",
    positions: "/positions/",
    watchlist: "/watchlists/Default/",
    pathfinderUserMachine: "/pathfinder/user_machine/",
    pathfinderInquiries: "/pathfinder/inquiries/",
    challenge: "/challenge/",
    push: "/push/",
    quotes: "/marketdata/quotes/",
    historicals: "/marketdata/historicals/",
    identi: "https://identi.robinhood.com/idl/v1/workflow/",
  },
};

const REFRESH = {
  sheet_name: "Refresh",
  named_range_name: "LastUpdate",
  cell_address: "A1",
};

function validateLastUpdate(LastUpdate) {
  // We add this check to ensure the parameter is always included in the sheet.
  if (LastUpdate === undefined || LastUpdate === null) {
    return [["Error: The LastUpdate parameter is required."]];
  }
}

/**
 * A wrapper for making authenticated requests to the Robinhood API.
 */
const RobinhoodApiClient = (function () {
  const service_ = PropertiesService.getUserProperties();

  function getAuthToken_() {
    const token = service_.getProperty("robinhood_access_token");
    if (!token) {
      throw new Error(
        'Authentication required. Please run from the "Robinhood > Login / Re-login" menu.',
      );
    }
    return token;
  }

  function makeRequest_(url, options, retryCount = 0) {
    const MAX_RETRIES = 5;
    const INITIAL_WAIT_TIME = 5000;

    // Making API request (attempt ${retryCount + 1})
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 429 && retryCount < MAX_RETRIES) {
      const waitTime =
        INITIAL_WAIT_TIME * Math.pow(2, retryCount) + Math.random() * 1000;
      const statusMessage = `Rate limited. Retrying in ${Math.round(waitTime / 1000)}s... (Attempt ${retryCount + 1})`;
      properties.setProperty("robinhood_retry_status", statusMessage);
      // Rate limited, retrying...
      Utilities.sleep(waitTime);
      return makeRequest_(url, options, retryCount + 1);
    }

    // This is the key change: Check if the request is for the token URL
    if (url === ROBINHOOD_CONFIG.TOKEN_URL) {
      // If it is, log a generic message instead of the full response
      Logger.log(
        `Response [${responseCode}] from authentication endpoint received.`,
      );
    } else {
      // Log response code only for non-auth endpoints
      if (responseCode >= 400) {
        Logger.log(`API Error [${responseCode}]: Request failed`);
      }
    }

    if (responseCode >= 200 && responseCode < 400) {
      try {
        return JSON.parse(responseText);
      } catch (e) {
        return responseText;
      }
    } else if (
      responseCode === 401 &&
      options.headers &&
      options.headers.Authorization
    ) {
      Logger.log("Token may have expired (401 Unauthorized). Clearing token.");
      service_.deleteProperty("robinhood_access_token");
      throw new Error(
        "Your session has expired. Please log in again via the menu.",
      );
    } else {
      if (
        url === ROBINHOOD_CONFIG.TOKEN_URL &&
        (responseCode === 400 || responseCode === 401 || responseCode === 403)
      ) {
        try {
          return JSON.parse(responseText);
        } catch (e) {
          throw new Error(
            `API request failed. ${responseCode}: [Sensitive response not shown]`,
          );
        }
      }
      throw new Error(`API request failed. ${responseCode}: ${responseText}`);
    }
  }

  function get(url) {
    const token = getAuthToken_();
    const options = {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        Authorization: "Bearer " + token,
      },
    };
    return makeRequest_(ROBINHOOD_CONFIG.API_BASE_URL + url, options);
  }

  function pagedGet(url, maxPages = 10) {
    let fullUrl = ROBINHOOD_CONFIG.API_BASE_URL + url;
    let responseJson = makeRequest_(fullUrl, {
      method: "get",
      muteHttpExceptions: true,
      headers: { Authorization: "Bearer " + getAuthToken_() },
    });

    let results = responseJson.results || [];
    let nextUrl = responseJson.next;
    let pageCount = 1;

    // Limit pagination to prevent timeouts
    while (nextUrl && pageCount < maxPages) {
      responseJson = makeRequest_(nextUrl, {
        method: "get",
        muteHttpExceptions: true,
        headers: { Authorization: "Bearer " + getAuthToken_() },
      });
      if (responseJson.results) {
        results = results.concat(responseJson.results);
      }
      nextUrl = responseJson.next;
      pageCount++;
    }

    return results;
  }

  return {
    get: get,
    pagedGet: pagedGet,
    makeRequest: makeRequest_,
  };
})();

/**
 * Generates a cryptographically secure-like device token.
 */
function generateDeviceToken_() {
  let token = "";
  const chars = "0123456789abcdef";
  for (let i = 0; i < 36; i++) {
    if (i === 8 || i === 13 || i === 18 || i === 23) {
      token += "-";
    } else {
      token += chars.charAt(Math.floor(Math.random() * chars.length));
    }
  }
  // Device token generated for authentication
  return token;
}

/**
 * Handles the new "Sherrif" verification workflow.
 */
function validateSherrifId_(deviceToken, workflowId) {
  const ui = SpreadsheetApp.getUi();
  // Starting MFA validation workflow

  // Step 1: Trigger the challenge by sending a PATCH request to the identi endpoint
  const identiUrl = ROBINHOOD_CONFIG.API_URIS.identi + workflowId + "/";
  const triggerPayload = {
    clientVersion: "1.0.0",
    id: workflowId,
    entryPointAction: {},
  };

  const triggerOptions = {
    method: "patch",
    contentType: "application/json",
    payload: JSON.stringify(triggerPayload),
    muteHttpExceptions: true,
  };

  // Triggering MFA challenge
  const triggerResponse = RobinhoodApiClient.makeRequest(
    identiUrl,
    triggerOptions,
  );

  if (
    !triggerResponse ||
    !triggerResponse.route ||
    !triggerResponse.route.replace ||
    !triggerResponse.route.replace.screen ||
    !triggerResponse.route.replace.screen.deviceApprovalChallengeScreenParams ||
    !triggerResponse.route.replace.screen.deviceApprovalChallengeScreenParams
      .sheriffChallenge
  ) {
    throw new Error(
      "Failed to trigger the MFA challenge. Invalid response from server.",
    );
  }

  const challenge =
    triggerResponse.route.replace.screen.deviceApprovalChallengeScreenParams
      .sheriffChallenge;

  // MFA challenge received
  const startTime = new Date().getTime();
  const timeout = 120 * 1000; // 2 minutes

  // Step 2: Handle the challenge based on its type
  if (challenge.type === "PROMPT") {
    ui.alert(
      "Login Verification Required",
      "Please check your Robinhood app to approve this login attempt.",
      ui.ButtonSet.OK,
    );

    const promptStatusUrl =
      ROBINHOOD_CONFIG.API_BASE_URL +
      ROBINHOOD_CONFIG.API_URIS.push +
      `${challenge.id}/get_prompts_status/`;

    while (new Date().getTime() - startTime < timeout) {
      // Polling for push notification approval
      const promptStatusResponse = RobinhoodApiClient.makeRequest(
        promptStatusUrl,
        { method: "get", muteHttpExceptions: true },
      );

      if (
        promptStatusResponse &&
        promptStatusResponse.challenge_status === "validated"
      ) {
        // Push notification approved

        // Step 3: Finalize the workflow
        const finalizePayload = {
          clientVersion: "1.0.0",
          screenName: "DEVICE_APPROVAL_CHALLENGE",
          id: workflowId,
          deviceApprovalChallengeAction: { proceed: {} },
        };

        const finalizeOptions = {
          method: "patch",
          contentType: "application/json",
          payload: JSON.stringify(finalizePayload),
          muteHttpExceptions: true,
        };

        // Finalizing MFA workflow
        const finalizeResponse = RobinhoodApiClient.makeRequest(
          identiUrl,
          finalizeOptions,
        );

        if (
          finalizeResponse &&
          finalizeResponse.route &&
          finalizeResponse.route.exit &&
          finalizeResponse.route.exit.status === "WORKFLOW_STATUS_APPROVED"
        ) {
          // MFA workflow completed successfully
          return; // Validation successful
        } else {
          throw new Error(
            `Failed to finalize the workflow. Response: ${JSON.stringify(finalizeResponse)}`,
          );
        }
      }

      Utilities.sleep(5000); // Wait 5 seconds before polling again
    }

    throw new Error(
      "Login approval timed out. You did not approve the push notification on your device within the 2-minute window.",
    );
  }

  // ... (Optional: Add handling for SMS/Email challenges here if needed) ...
  else {
    throw new Error(`Unsupported challenge type: ${challenge.type}`);
  }
}

/**
 * Displays the custom HTML login dialog to the user.
 * This function will be called from the "Robinhood > Login / Re-login" menu.
 */
function showLoginDialog_() {
  const html = HtmlService.createHtmlOutputFromFile("LoginDialog")
    .setWidth(300)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, " ");
}

/**
 * Processes the credentials submitted from the HTML dialog.
 * This function is called by `google.script.run` from the HTML form.
 * @param {object} credentials An object with 'username' and 'password' properties.
 * @return {string} A status message to display back in the dialog.
 */
function processLogin(credentials) {
  // Starting authentication process
  if (!credentials || !credentials.username || !credentials.password) {
    return "Username and password are required.";
  }

  try {
    const deviceToken = generateDeviceToken_();
    const loginPayload = {
      client_id: ROBINHOOD_CONFIG.CLIENT_ID,
      expires_in: 86400,
      grant_type: "password",
      password: credentials.password,
      scope: "internal",
      username: credentials.username,
      device_token: deviceToken,
      try_passkeys: false,
      token_request_path: "/login/",
      create_read_only_secondary_token: true,
    };

    const loginOptions = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(loginPayload),
      muteHttpExceptions: true,
    };

    // Attempting authentication with Robinhood
    let loginResponse = RobinhoodApiClient.makeRequest(
      ROBINHOOD_CONFIG.TOKEN_URL,
      loginOptions,
    );

    let finalTokenResponse;

    if (loginResponse && loginResponse.verification_workflow) {
      // MFA verification required
      validateSherrifId_(deviceToken, loginResponse.verification_workflow.id);

      // MFA complete, finalizing authentication
      finalTokenResponse = RobinhoodApiClient.makeRequest(
        ROBINHOOD_CONFIG.TOKEN_URL,
        loginOptions,
      );
    } else if (loginResponse && loginResponse.access_token) {
      // Authentication successful
      finalTokenResponse = loginResponse;
    } else {
      const errorDetail =
        loginResponse && loginResponse.detail
          ? loginResponse.detail
          : "No additional details provided.";
      throw new Error(`Login failed. Server response: ${errorDetail}`);
    }

    if (finalTokenResponse && finalTokenResponse.access_token) {
      PropertiesService.getUserProperties().setProperty(
        "robinhood_access_token",
        finalTokenResponse.access_token,
      );
      // Authentication completed successfully
      return "Success! You are now logged in. This dialog will close shortly.";
    } else {
      const tokenErrorDetail =
        finalTokenResponse && finalTokenResponse.detail
          ? finalTokenResponse.detail
          : "No additional details provided.";
      throw new Error(
        `Failed to retrieve final access token. Detail: ${tokenErrorDetail}`,
      );
    }
  } catch (e) {
    Logger.log(`Authentication error: ${e.message || "Unknown error"}`);
    // Return the error message to be displayed in the dialog
    return e.toString();
  }
}

/**
 * Recursively unpacks and flattens a result from a Robinhood API endpoint.
 */
function flattenResult_(
  result,
  flattenedResult,
  hyperlinkedFields,
  originalEndpointName,
) {
  for (const key in result) {
    if (Object.prototype.hasOwnProperty.call(result, key)) {
      const value = result[key];
      if (
        hyperlinkedFields.includes(key) &&
        typeof value === "string" &&
        value.startsWith("http")
      ) {
        const responseJson = RobinhoodApiClient.makeRequest(value, {
          method: "get",
          muteHttpExceptions: true,
          headers: {
            Authorization:
              "Bearer " +
              PropertiesService.getUserProperties().getProperty(
                "robinhood_access_token",
              ),
          },
        });
        const nextHyperlinkedFields = hyperlinkedFields.slice();
        nextHyperlinkedFields.splice(nextHyperlinkedFields.indexOf(key), 1);
        flattenResult_(
          responseJson,
          flattenedResult,
          nextHyperlinkedFields,
          key,
        );
      } else if (
        value !== null &&
        typeof value === "object" &&
        !Array.isArray(value)
      ) {
        flattenResult_(
          value,
          flattenedResult,
          hyperlinkedFields,
          originalEndpointName,
        );
      } else if (
        Array.isArray(value) &&
        key !== "executions" &&
        value.length > 0 &&
        typeof value[0] === "object"
      ) {
        flattenResult_(
          value[0],
          flattenedResult,
          hyperlinkedFields,
          originalEndpointName,
        );
      } else {
        const modifiedKey = `${originalEndpointName}_${key}`;
        flattenedResult[modifiedKey] = value;
      }
    }
  }
}

/**
 * Iterates through all results of a Robinhood API endpoint and builds a 2D array for Google Sheets.
 */
function getRobinhoodData_(endpointName, hyperlinkedFields, options = {}) {
  try {
    // Use the custom endpoint from options if it exists, otherwise use the default.
    const endpointUrl =
      options.endpoint || ROBINHOOD_CONFIG.API_URIS[endpointName];
    let results = RobinhoodApiClient.pagedGet(endpointUrl);

    if (endpointName === "positions") {
      results = results.filter((row) => parseFloat(row["quantity"]) > 0);
    }

    if (!results || results.length === 0) {
      return [["No results found for " + endpointName]];
    }

    const allFlattenedResults = [];
    const allKeys = new Set();

    results.forEach((result) => {
      const flattenedResult = {};
      flattenResult_(
        result,
        flattenedResult,
        hyperlinkedFields.slice(),
        endpointName,
      );
      allFlattenedResults.push(flattenedResult);
      Object.keys(flattenedResult).forEach((key) => allKeys.add(key));
    });

    const header = Array.from(allKeys).sort();
    const data = [header];

    allFlattenedResults.forEach((flattenedResult) => {
      const row = header.map((key) =>
        flattenedResult.hasOwnProperty(key) ? flattenedResult[key] : "",
      );
      data.push(row);
    });

    return data;
  } catch (e) {
    return [["Error: " + e.message]];
  }
}

// --- Function Discovery & Help ---

/**
 * Lists all available Robinhood functions with descriptions and usage examples.
 *
 * @param {string} [category] Optional filter by category: 'core', 'analytics', 'options', 'utility', 'all'
 * @return {Array<Array<string>>} A table of available functions with descriptions
 * @customfunction
 */
function ROBINHOOD_HELP(category = "all") {
  const functions = [
    // Core Data Functions
    [
      "ROBINHOOD_GET_POSITIONS",
      "core",
      "Current stock positions",
      "ROBINHOOD_GET_POSITIONS(LastUpdate)",
    ],
    [
      "ROBINHOOD_GET_ORDERS",
      "core",
      "Order history with date filtering",
      "ROBINHOOD_GET_ORDERS(30, 1000, LastUpdate)",
    ],
    [
      "ROBINHOOD_GET_DIVIDENDS",
      "core",
      "Dividend payment history",
      "ROBINHOOD_GET_DIVIDENDS(LastUpdate)",
    ],
    [
      "ROBINHOOD_GET_QUOTE",
      "core",
      "Real-time stock quote",
      'ROBINHOOD_GET_QUOTE("AAPL", TRUE, LastUpdate)',
    ],
    [
      "ROBINHOOD_GET_QUOTES_BATCH",
      "core",
      "Multiple stock quotes efficiently",
      'ROBINHOOD_GET_QUOTES_BATCH("AAPL,MSFT,GOOGL", LastUpdate)',
    ],
    [
      "ROBINHOOD_GET_HISTORICALS",
      "core",
      "Historical price data (supports: day/week/month/3month/year/5year spans)",
      'ROBINHOOD_GET_HISTORICALS("AAPL", "day", "year", LastUpdate)',
    ],
    [
      "ROBINHOOD_GET_PORTFOLIOS",
      "core",
      "Portfolio summary data",
      "ROBINHOOD_GET_PORTFOLIOS(LastUpdate)",
    ],
    [
      "ROBINHOOD_GET_ACCOUNTS",
      "core",
      "Account information",
      "ROBINHOOD_GET_ACCOUNTS(LastUpdate)",
    ],
    [
      "ROBINHOOD_GET_WATCHLISTS",
      "core",
      "List all available watchlists",
      "ROBINHOOD_GET_WATCHLISTS(LastUpdate)",
    ],
    [
      "ROBINHOOD_GET_WATCHLIST",
      "core",
      "Specific watchlist instruments",
      'ROBINHOOD_GET_WATCHLIST("Default", LastUpdate)',
    ],
    [
      "ROBINHOOD_GET_ALL_WATCHLISTS",
      "core",
      "All instruments from all watchlists",
      "ROBINHOOD_GET_ALL_WATCHLISTS(LastUpdate)",
    ],
    [
      "ROBINHOOD_GET_ACH_TRANSFERS",
      "core",
      "ACH transfer history",
      "ROBINHOOD_GET_ACH_TRANSFERS(LastUpdate)",
    ],

    // Analytics Functions
    [
      "ROBINHOOD_GET_PORTFOLIO_HISTORICALS",
      "analytics",
      "Portfolio performance over time (supports: day/week/month/3month/year/5year spans)",
      'ROBINHOOD_GET_PORTFOLIO_HISTORICALS("year", "day", "5DP12345", LastUpdate)',
    ],

    // Options Functions
    [
      "ROBINHOOD_GET_OPTIONS_POSITIONS",
      "options",
      "Current options positions",
      "ROBINHOOD_GET_OPTIONS_POSITIONS(LastUpdate)",
    ],
    [
      "ROBINHOOD_GET_OPTIONS_ORDERS",
      "options",
      "Options order history",
      "ROBINHOOD_GET_OPTIONS_ORDERS(30, 100, LastUpdate)",
    ],

    // Utility Functions
    [
      "ROBINHOOD_VALIDATE_TICKER",
      "utility",
      "Validate ticker symbol format",
      'ROBINHOOD_VALIDATE_TICKER("AAPL")',
    ],
    [
      "ROBINHOOD_FORMAT_CURRENCY",
      "utility",
      "Format number as currency",
      "ROBINHOOD_FORMAT_CURRENCY(1234.56)",
    ],
    [
      "ROBINHOOD_LAST_MARKET_DAY",
      "utility",
      "Get last trading day",
      "ROBINHOOD_LAST_MARKET_DAY()",
    ],
    [
      "ROBINHOOD_GET_LOGIN_STATUS",
      "utility",
      "Check authentication status",
      "ROBINHOOD_GET_LOGIN_STATUS(LastUpdate)",
    ],
    [
      "ROBINHOOD_GET_URL",
      "utility",
      "Get data from any Robinhood API URL",
      'ROBINHOOD_GET_URL("https://api.robinhood.com/accounts/", LastUpdate)',
    ],

    // Help Function
    [
      "ROBINHOOD_HELP",
      "utility",
      "Show all available functions",
      'ROBINHOOD_HELP() or ROBINHOOD_HELP("analytics")',
    ],
  ];

  // Filter by category if specified
  let filteredFunctions = functions;
  if (category !== "all") {
    filteredFunctions = functions.filter(
      (func) => func[1] === category.toLowerCase(),
    );
  }

  // Build the result table - NOTE: Examples are strings, not executable formulas
  const result = [
    [
      "Function Name",
      "Category",
      "Description",
      "Example Usage (copy as =formula)",
    ],
  ];
  filteredFunctions.forEach((func) => {
    result.push([func[0], func[1].toUpperCase(), func[2], func[3]]);
  });

  return result;
}

// --- Utility Helper Functions ---

/**
 * Validates a stock ticker symbol format.
 *
 * @param {string} ticker The stock ticker to validate
 * @return {boolean} True if ticker format is valid
 * @customfunction
 */
function ROBINHOOD_VALIDATE_TICKER(ticker) {
  if (!ticker || typeof ticker !== "string") return false;
  return /^[A-Z]{1,5}$/.test(ticker.toUpperCase());
}

/**
 * Formats a number as currency.
 *
 * @param {number} amount The amount to format
 * @return {string} Formatted currency string
 * @customfunction
 */
function ROBINHOOD_FORMAT_CURRENCY(amount) {
  if (typeof amount !== "number") return "$0.00";
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
  }).format(amount);
}

/**
 * Gets the last market trading day.
 *
 * @return {string} ISO date string of last market day
 * @customfunction
 */
function ROBINHOOD_LAST_MARKET_DAY() {
  const today = new Date();
  let lastMarketDay = new Date(today);

  // Go back to find last weekday (Monday-Friday)
  while (lastMarketDay.getDay() === 0 || lastMarketDay.getDay() === 6) {
    lastMarketDay.setDate(lastMarketDay.getDate() - 1);
  }

  return lastMarketDay.toISOString().split("T")[0];
}

// --- Custom Functions for Sheets ---

/**
 * Retrieves a history of ACH transfers.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range, e.g., `=ROBINHOOD_GET_ACH_TRANSFERS(LastUpdate)`.
 * @return {Array<Array<string>>} A two-dimensional array of ACH transfer data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_ACH_TRANSFERS(LastUpdate) {
  validateLastUpdate(LastUpdate);
  return getRobinhoodData_("achTransfers", ["ach_relationship"]);
}

/**
 * Retrieves dividend history for your account.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range, e.g., `=ROBINHOOD_GET_DIVIDENDS(LastUpdate)`.
 * @return {Array<Array<string>>} A two-dimensional array of dividend data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_DIVIDENDS(LastUpdate) {
  validateLastUpdate(LastUpdate);
  return getRobinhoodData_("dividends", []);
}

/**
 * Retrieves a history of options orders.
 *
 * @param {number} [days=0] Number of days to look back. If 0 or omitted, all orders are returned (limited by page_size).
 * @param {number} [page_size=50] Number of items to return per page (max recommended: 1000).
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range.
 * @return {Array<Array<string>>} A two-dimensional array of options order data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_OPTIONS_ORDERS(days = 0, page_size = 50, LastUpdate) {
  validateLastUpdate(LastUpdate);
  try {
    // Build the endpoint with parameters
    let endpoint = ROBINHOOD_CONFIG.API_URIS.optionsOrders + "?";

    // Set page size (limit to prevent timeouts)
    const maxPageSize = Math.min(page_size || 50, 1000);
    endpoint += `page_size=${maxPageSize}`;

    // Add date filter if days is specified
    if (days && days > 0) {
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - days);
      const isoDate = cutoffDate.toISOString().split("T")[0]; // YYYY-MM-DD format
      endpoint += `&updated_at[gte]=${isoDate}`;
    }

    // Limit pagination to prevent timeouts - max 5 pages for large requests
    const maxPages = Math.min(Math.ceil(1000 / maxPageSize), 5);
    const results = RobinhoodApiClient.pagedGet(endpoint, maxPages);

    if (!results || results.length === 0) {
      return [
        ["No options orders found or options trading not enabled for account"],
      ];
    }

    // Simplified data processing without hyperlinked fields to avoid additional API calls
    const header = Object.keys(results[0]);
    const data = [header];

    // Process results with a reasonable limit to avoid timeout
    const maxResults = Math.min(results.length, 500);
    results.slice(0, maxResults).forEach((order) => {
      data.push(
        header.map((key) => {
          const value = order[key];
          return typeof value === "object" && value !== null
            ? JSON.stringify(value)
            : value;
        }),
      );
    });

    return data;
  } catch (e) {
    return [
      [
        "Error: Options orders may not be available or enabled for your account. " +
          e.message,
      ],
    ];
  }
}

/**
 * Retrieves all current options positions.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range, e.g., `=ROBINHOOD_GET_OPTIONS_POSITIONS(LastUpdate)`.
 * @return {Array<Array<string>>} A two-dimensional array of options position data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_OPTIONS_POSITIONS(LastUpdate) {
  validateLastUpdate(LastUpdate);
  try {
    // Use limited pagination to avoid timeouts
    const results = RobinhoodApiClient.pagedGet(
      ROBINHOOD_CONFIG.API_URIS.optionsPositions,
      3,
    ); // Max 3 pages

    if (!results || results.length === 0) {
      return [
        [
          "No options positions found or options trading not enabled for account",
        ],
      ];
    }

    // Filter for non-zero positions to reduce data size
    const activePositions = results.filter(
      (pos) => parseFloat(pos.quantity || 0) > 0,
    );

    if (activePositions.length === 0) {
      return [["No active options positions found"]];
    }

    // Simplified processing without hyperlinked fields
    const header = Object.keys(activePositions[0]);
    const data = [header];

    activePositions.forEach((position) => {
      data.push(
        header.map((key) => {
          const value = position[key];
          return typeof value === "object" && value !== null
            ? JSON.stringify(value)
            : value;
        }),
      );
    });

    return data;
  } catch (e) {
    return [
      [
        "Error: Options positions may not be available or enabled for your account. " +
          e.message,
      ],
    ];
  }
}

/**
 * Retrieves a history of stock orders, with an optional filter for the last X days.
 *
 * @param {number} [days=0] Optional. Number of days to look back. If 0 or omitted, all orders are returned.
 * @param {number} [page_size=1000] Optional. Number of items to return per a page.
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range.
 * @return {Array<Array<string>>} A two-dimensional array of stock order data.
 * @customfunction
 */
function ROBINHOOD_GET_ORDERS(days = 0, page_size = 1000, LastUpdate) {
  validateLastUpdate(LastUpdate);
  let endpoint = ROBINHOOD_CONFIG.API_URIS.orders + `?`;
  endpoint += `page_size=${page_size}`;

  // If a 'days' filter is applied, modify the endpoint to include a date filter.
  if (days > 0) {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    // Format the date to an ISO string that the API understands.
    const isoDateString = cutoffDate.toISOString();
    // Add the 'updated_at[gte]' parameter to the URL.
    endpoint += `&updated_at[gte]=${isoDateString}`;
  }

  const ordersData = getRobinhoodData_("orders", [], {
    endpoint: endpoint,
  });

  if (
    ordersData[0][0].startsWith("Error:") ||
    ordersData[0].includes("No results found")
  ) {
    if (days > 0) {
      return [["No orders found in the last " + days + " days."]];
    }
    return ordersData;
  }

  return ordersData;
}

/**
 * Retrieves portfolio data, including account value and history.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range, e.g., `=ROBINHOOD_GET_PORTFOLIOS(LastUpdate)`.
 * @return {Array<Array<string>>} A two-dimensional array of portfolio data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_PORTFOLIOS(LastUpdate) {
  validateLastUpdate(LastUpdate);
  return getRobinhoodData_("portfolios", []);
}

/**
 * Retrieves portfolio historical performance data.
 *
 * @param {string} [span="year"] The time span. Supported values:
 *   - "day": Single day data
 *   - "week": One week data  
 *   - "month": One month data
 *   - "3month": Three months data
 *   - "year": One year data (default)
 *   - "5year": Five years data
 * @param {string} [interval="day"] The time interval. Supported values:
 *   - "5minute": 5-minute intervals (best for day/week spans)
 *   - "day": Daily intervals (default, best for month+ spans)
 *   - "week": Weekly intervals (best for 5year span)
 * @param {string} [accountNumber] Optional account number to filter results. If omitted, returns data for all accounts combined.
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range.
 * @return {Array<Array<string>>} A two-dimensional array of portfolio historical data.
 * @customfunction
 */
function ROBINHOOD_GET_PORTFOLIO_HISTORICALS(
  span = "year",
  interval = "day",
  accountNumber,
  LastUpdate,
) {
  // Handle backward compatibility - if 3rd parameter is LastUpdate, shift parameters
  if (arguments.length === 3 && typeof accountNumber !== "string") {
    LastUpdate = accountNumber;
    accountNumber = null;
  }
  
  validateLastUpdate(LastUpdate);

  try {
    // Get all accounts
    const accounts = RobinhoodApiClient.pagedGet(
      ROBINHOOD_CONFIG.API_URIS.accounts,
    );
    if (!accounts || accounts.length === 0) {
      return [["Error: No accounts found"]];
    }

    let targetAccounts = accounts;
    
    // Filter by specific account if provided
    if (accountNumber) {
      targetAccounts = accounts.filter(acc => acc.account_number === accountNumber);
      if (targetAccounts.length === 0) {
        return [["Error: Account number not found"]];
      }
    }

    // If multiple accounts and no specific account requested, combine data
    if (targetAccounts.length === 1) {
      // Single account - return its data directly
      const account = targetAccounts[0];
      const endpoint = `${ROBINHOOD_CONFIG.API_URIS.portfolioHistoricals}${account.account_number}/?span=${span}&interval=${interval}`;
      
      const result = RobinhoodApiClient.get(endpoint);
      if (
        result &&
        result.equity_historicals &&
        result.equity_historicals.length > 0
      ) {
        const header = Object.keys(result.equity_historicals[0]);
        const data = [header];
        result.equity_historicals.forEach((row) => {
          data.push(header.map((key) => row[key]));
        });
        return data;
      }
      return [["Error: No portfolio historical data found"]];
      
    } else {
      // Multiple accounts - combine the data
      const combinedData = [];
      let headerSet = false;
      
      for (const account of targetAccounts) {
        try {
          const endpoint = `${ROBINHOOD_CONFIG.API_URIS.portfolioHistoricals}${account.account_number}/?span=${span}&interval=${interval}`;
          const result = RobinhoodApiClient.get(endpoint);
          
          if (result && result.equity_historicals && result.equity_historicals.length > 0) {
            const header = Object.keys(result.equity_historicals[0]);
            
            // Add header only once
            if (!headerSet) {
              combinedData.push(["account_number", ...header]);
              headerSet = true;
            }
            
            // Add data with account number prefix
            result.equity_historicals.forEach((row) => {
              combinedData.push([account.account_number, ...header.map((key) => row[key])]);
            });
          }
        } catch (accountError) {
          // Skip this account if there's an error but continue with others
          continue;
        }
      }
      
      return combinedData.length > 1 ? combinedData : [["Error: No portfolio historical data found for any account"]];
    }
  } catch (e) {
    return [["Error: " + e.message]];
  }
}

/**
 * Retrieves all current stock positions.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range, e.g., `=ROBINHOOD_GET_POSITIONS(LastUpdate)`.
 * @return {Array<Array<string>>} A two-dimensional array of stock position data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_POSITIONS(LastUpdate) {
  validateLastUpdate(LastUpdate);
  return getRobinhoodData_("positions", []);
}

/**
 * Lists all available watchlists in your account.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range.
 * @return {Array<Array<string>>} A two-dimensional array of available watchlists.
 * @customfunction
 */
function ROBINHOOD_GET_WATCHLISTS(LastUpdate) {
  validateLastUpdate(LastUpdate);

  try {
    const token = PropertiesService.getUserProperties().getProperty(
      "robinhood_access_token",
    );
    if (!token) {
      return [["Error: Authentication required. Please log in first."]];
    }

    const options = {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        Authorization: "Bearer " + token,
      },
    };

    // Use the correct discovery lists endpoint
    const response = UrlFetchApp.fetch(
      ROBINHOOD_CONFIG.API_BASE_URL + "/discovery/lists/default/",
      options,
    );
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      return [["Error: Failed to fetch watchlists. HTTP " + responseCode]];
    }

    const result = JSON.parse(response.getContentText());

    if (!result.results || !Array.isArray(result.results)) {
      return [["Error: Unexpected API response format"]];
    }

    const header = [
      "name",
      "id",
      "created_at",
      "updated_at",
      "item_count",
      "owner_type",
      "icon_emoji",
    ];
    const data = [header];

    result.results.forEach((watchlist) => {
      data.push([
        watchlist.display_name || "N/A",
        watchlist.id || "N/A",
        watchlist.created_at || "N/A",
        watchlist.updated_at || "N/A",
        watchlist.item_count || "0",
        watchlist.owner_type || "N/A",
        watchlist.icon_emoji || "N/A",
      ]);
    });

    return data;
  } catch (e) {
    return [["Error: " + e.message]];
  }
}

/**
 * Retrieves instruments from a specific watchlist with detailed market data.
 *
 * @param {string} [watchlistNameOrId] The name or ID of the watchlist to retrieve (e.g., "General", "Buy Review", or use the ID from ROBINHOOD_GET_WATCHLISTS).
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range.
 * @return {Array<Array<string>>} A two-dimensional array of watchlist data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_WATCHLIST(watchlistNameOrId, LastUpdate) {
  validateLastUpdate(LastUpdate);

  if (!watchlistNameOrId) {
    return [
      [
        "Error: Please provide a watchlist name or ID. Use ROBINHOOD_GET_WATCHLISTS() to see available watchlists.",
      ],
    ];
  }

  try {
    const token = PropertiesService.getUserProperties().getProperty(
      "robinhood_access_token",
    );
    if (!token) {
      return [["Error: Authentication required. Please log in first."]];
    }

    const options = {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        Authorization: "Bearer " + token,
      },
    };

    let watchlistId = watchlistNameOrId;
    let watchlistName = watchlistNameOrId;

    // If it looks like a name rather than a UUID, find the ID first
    if (!watchlistNameOrId.includes("-")) {
      // Get all watchlists to find the ID by name
      const listsResponse = UrlFetchApp.fetch(
        ROBINHOOD_CONFIG.API_BASE_URL + "/discovery/lists/default/",
        options,
      );
      if (listsResponse.getResponseCode() === 200) {
        const listsData = JSON.parse(listsResponse.getContentText());
        if (listsData.results) {
          const matchingList = listsData.results.find(
            (list) =>
              list.display_name &&
              list.display_name.toLowerCase() ===
                watchlistNameOrId.toLowerCase(),
          );
          if (matchingList) {
            watchlistId = matchingList.id;
            watchlistName = matchingList.display_name;
          } else {
            return [
              [
                `Watchlist "${watchlistNameOrId}" not found. Available watchlists: ${listsData.results.map((l) => l.display_name).join(", ")}`,
              ],
            ];
          }
        }
      }
    }

    // Use the richer items endpoint for detailed market data
    const currentTime = new Date();
    const localMidnight = new Date(
      currentTime.getFullYear(),
      currentTime.getMonth(),
      currentTime.getDate(),
    );
    const localMidnightISO = localMidnight.toISOString();

    const itemsEndpoint = `/discovery/lists/items/?list_id=${watchlistId}&local_midnight=${encodeURIComponent(localMidnightISO)}`;
    const itemsResponse = UrlFetchApp.fetch(
      ROBINHOOD_CONFIG.API_BASE_URL + itemsEndpoint,
      options,
    );

    if (itemsResponse.getResponseCode() !== 200) {
      return [
        [
          `Error: Could not fetch watchlist items. HTTP ${itemsResponse.getResponseCode()}`,
        ],
      ];
    }

    const itemsData = JSON.parse(itemsResponse.getContentText());

    if (!itemsData.results || itemsData.results.length === 0) {
      return [
        [
          `Watchlist "${watchlistNameOrId}" is empty or has no accessible items.`,
        ],
      ];
    }

    const header = [
      "symbol",
      "name",
      "price",
      "one_day_change",
      "one_day_percent_change",
      "market_cap",
      "state",
      "us_tradability",
      "added_at",
      "watchlist_name",
    ];
    const data = [header];

    // Process each item in the watchlist
    itemsData.results.forEach((item) => {
      try {
        data.push([
          item.symbol || "N/A",
          item.name || "N/A",
          item.price || "N/A",
          item.one_day_dollar_change || "N/A",
          item.one_day_percent_change
            ? (item.one_day_percent_change * 100).toFixed(2) + "%"
            : "N/A",
          item.market_cap || "N/A",
          item.state || "N/A",
          item.us_tradability || "N/A",
          item.created_at || "N/A",
          watchlistName,
        ]);
      } catch (e) {
        // Add error entry but continue processing other items
        data.push([
          "Error",
          e.message.substring(0, 50),
          "N/A",
          "N/A",
          "N/A",
          "N/A",
          "N/A",
          "N/A",
          "N/A",
          watchlistName,
        ]);
      }
    });

    return data;
  } catch (e) {
    return [["Error: " + e.message]];
  }
}

/**
 * Retrieves instruments from ALL watchlists.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range.
 * @return {Array<Array<string>>} A two-dimensional array of all watchlist data.
 * @customfunction
 */
function ROBINHOOD_GET_ALL_WATCHLISTS(LastUpdate) {
  validateLastUpdate(LastUpdate);

  try {
    const token = PropertiesService.getUserProperties().getProperty(
      "robinhood_access_token",
    );
    if (!token) {
      return [["Error: Authentication required. Please log in first."]];
    }

    const options = {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        Authorization: "Bearer " + token,
      },
    };

    // First get all available watchlists using the discovery API
    const listsResponse = UrlFetchApp.fetch(
      ROBINHOOD_CONFIG.API_BASE_URL + "/discovery/lists/default/",
      options,
    );

    if (listsResponse.getResponseCode() !== 200) {
      return [
        [
          "Error: Failed to fetch watchlists. HTTP " +
            listsResponse.getResponseCode(),
        ],
      ];
    }

    const listsData = JSON.parse(listsResponse.getContentText());

    if (!listsData.results || listsData.results.length === 0) {
      return [["No watchlists found"]];
    }

    const allInstruments = [];
    const header = [
      "watchlist_name",
      "symbol",
      "name",
      "price",
      "one_day_change",
      "one_day_percent_change",
      "market_cap",
      "state",
      "us_tradability",
      "added_at",
    ];

    // Calculate local midnight for the API call
    const currentTime = new Date();
    const localMidnight = new Date(
      currentTime.getFullYear(),
      currentTime.getMonth(),
      currentTime.getDate(),
    );
    const localMidnightISO = localMidnight.toISOString();

    for (const watchlist of listsData.results) {
      const watchlistName = watchlist.display_name || "Unknown";

      // Skip empty watchlists to avoid unnecessary API calls
      if (watchlist.item_count === 0) {
        continue;
      }

      try {
        // Use the richer items endpoint for detailed market data
        const itemsEndpoint = `/discovery/lists/items/?list_id=${watchlist.id}&local_midnight=${encodeURIComponent(localMidnightISO)}`;
        const itemsResponse = UrlFetchApp.fetch(
          ROBINHOOD_CONFIG.API_BASE_URL + itemsEndpoint,
          options,
        );

        if (itemsResponse.getResponseCode() === 200) {
          const itemsData = JSON.parse(itemsResponse.getContentText());

          if (itemsData.results && itemsData.results.length > 0) {
            itemsData.results.forEach((item) => {
              try {
                allInstruments.push([
                  watchlistName,
                  item.symbol || "N/A",
                  item.name || "N/A",
                  item.price || "N/A",
                  item.one_day_dollar_change || "N/A",
                  item.one_day_percent_change
                    ? (item.one_day_percent_change * 100).toFixed(2) + "%"
                    : "N/A",
                  item.market_cap || "N/A",
                  item.state || "N/A",
                  item.us_tradability || "N/A",
                  item.created_at || "N/A",
                ]);
              } catch (e) {
                // Add error entry but continue processing other items
                allInstruments.push([
                  watchlistName,
                  "Error",
                  e.message.substring(0, 50),
                  "N/A",
                  "N/A",
                  "N/A",
                  "N/A",
                  "N/A",
                  "N/A",
                  "N/A",
                ]);
              }
            });
          }
        }

        // Add small delay to prevent rate limiting
        Utilities.sleep(300);
      } catch (e) {
        // Add error entry for this entire watchlist but continue
        allInstruments.push([
          watchlistName,
          "Error fetching watchlist",
          e.message.substring(0, 50),
          "N/A",
          "N/A",
          "N/A",
          "N/A",
          "N/A",
          "N/A",
          "N/A",
        ]);
      }
    }

    if (allInstruments.length === 0) {
      return [["No instruments found in any watchlist"]];
    }

    // Return data with header
    const data = [header];
    data.push(...allInstruments);

    return data;
  } catch (e) {
    return [["Error: " + e.message]];
  }
}

/**
 * Retrieves the latest quote for a given stock ticker.
 *
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @param {boolean} [includeHeader=true] Optional. Set to false to exclude the header row from the output.
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range, e.g., `=ROBINHOOD_GET_QUOTE(LastUpdate)`.
 * @return {Array<Array<string>>} The quote data including price, bid, and ask.
 * @customfunction
 */
function ROBINHOOD_GET_QUOTE(ticker, includeHeader = true, LastUpdate) {
  validateLastUpdate(LastUpdate);
  if (!ticker) {
    return [["Error: Please provide a ticker symbol."]];
  }
  const endpoint =
    ROBINHOOD_CONFIG.API_URIS.quotes + ticker.toUpperCase() + "/";
  try {
    const result = RobinhoodApiClient.get(endpoint);
    if (result && result.last_trade_price) {
      const header = Object.keys(result);
      const values = Object.values(result);
      // Conditionally return the header based on the new parameter
      if (includeHeader) {
        return [header, values];
      } else {
        return [values]; // Return only the values as a single row
      }
    }
    return [["Error: Could not find quote for ticker " + ticker]];
  } catch (e) {
    return [["Error: " + e.message]];
  }
}

/**
 * Retrieves quotes for multiple stock tickers in a single call.
 *
 * @param {string} tickers Comma-separated list of ticker symbols (e.g., "AAPL,MSFT,GOOGL").
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range.
 * @return {Array<Array<string>>} A two-dimensional array with quotes for all tickers.
 * @customfunction
 */
function ROBINHOOD_GET_QUOTES_BATCH(tickers, LastUpdate) {
  validateLastUpdate(LastUpdate);
  if (!tickers) {
    return [["Error: Please provide ticker symbols separated by commas."]];
  }

  try {
    const tickerList = tickers
      .toString()
      .split(",")
      .map((t) => t.trim().toUpperCase());
    const tickerParams = tickerList.join(",");
    const endpoint = `${ROBINHOOD_CONFIG.API_URIS.quotes}?symbols=${tickerParams}`;

    const result = RobinhoodApiClient.get(endpoint);
    if (result && result.results && result.results.length > 0) {
      const header = Object.keys(result.results[0]);
      const data = [header];
      result.results.forEach((quote) => {
        data.push(header.map((key) => quote[key]));
      });
      return data;
    }
    return [["Error: No quotes found for the provided tickers"]];
  } catch (e) {
    return [["Error: " + e.message]];
  }
}

/**
 * Retrieves historical price data for a given stock ticker.
 *
 * @param {string} ticker The stock ticker symbol (e.g., "TSLA").
 * @param {string} [interval="day"] The time interval. Supported values:
 *   - "5minute": 5-minute intervals (best for day/week spans)
 *   - "day": Daily intervals (default, best for month+ spans) 
 *   - "week": Weekly intervals (best for multi-year spans)
 * @param {string} [span="year"] The time span. Supported values:
 *   - "day": Single day data
 *   - "week": One week data
 *   - "month": One month data
 *   - "3month": Three months data
 *   - "year": One year data (default)
 *   - "5year": Five years data
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range.
 * @return {Array<Array<string>>} A 2D array of historical price data.
 * @customfunction
 */
function ROBINHOOD_GET_HISTORICALS(
  ticker,
  interval = "day",
  span = "year",
  LastUpdate,
) {
  validateLastUpdate(LastUpdate);
  if (!ticker) {
    return [["Error: Please provide a ticker symbol."]];
  }
  const endpoint = `${ROBINHOOD_CONFIG.API_URIS.historicals}${ticker.toUpperCase()}/?interval=${interval}&span=${span}`;
  try {
    const result = RobinhoodApiClient.get(endpoint);
    if (result && result.historicals && result.historicals.length > 0) {
      const header = Object.keys(result.historicals[0]);
      const data = [header];
      result.historicals.forEach((row) => {
        data.push(header.map((key) => row[key]));
      });
      return data;
    }
    return [["Error: No historical data found for " + ticker]];
  } catch (e) {
    return [["Error: " + e.message]];
  }
}

/**
 * Retrieves detailed information for all brokerage accounts.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range, e.g., `=ROBINHOOD_GET_ACCOUNTS(LastUpdate)`.
 * @return {Array<Array<string>>} A two-dimensional array of all account data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_ACCOUNTS(LastUpdate) {
  validateLastUpdate(LastUpdate);
  try {
    // This is the specific endpoint you requested
    const endpoint =
      ROBINHOOD_CONFIG.API_URIS.accounts +
      "?default_to_all_accounts=true&include_managed=true&include_multiple_individual=false&is_default=false";
    // Use the existing pagedGet client to fetch the data
    const accounts = RobinhoodApiClient.pagedGet(endpoint);
    if (!accounts || accounts.length === 0) {
      return [["Error: No accounts found."]];
    }

    // Automatically create a header from the keys of the first account
    const header = Object.keys(accounts[0]);
    const data = [header]; // The first row of our output table is the header

    // Iterate through each account object returned by the API
    accounts.forEach((account) => {
      // For each account, create a row of data in the same order as the header
      const row = header.map((key) => {
        const value = account[key];
        // If a value is a nested object (like margin_balances), convert it to a JSON string
        if (typeof value === "object" && value !== null) {
          return JSON.stringify(value);
        }
        return value;
      });
      data.push(row);
    });

    return data;
  } catch (e) {
    return [["Error: " + e.message]];
  }
}

/**
 * Retrieves data from a specific Robinhood API URL.
 *
 * @param {string} url The full Robinhood API URL to fetch data from.
 * @param {boolean} [includeHeader=true] Optional. Set to FALSE to exclude the header row.
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range.
 * @return {Array<Array<string>>} A two-dimensional array of the fetched data.
 * @customfunction
 */
function ROBINHOOD_GET_URL(url, LastUpdate, includeHeader) {
  validateLastUpdate(LastUpdate);
  if (!url) {
    return [["Error: Please provide a URL."]];
  }

  // If includeHeader is explicitly set to FALSE, don't show the header. Otherwise, show it.
  const showHeader = includeHeader !== false;

  // The existing client adds the base URL, so we remove it from the input if present.
  const endpoint = url.replace(ROBINHOOD_CONFIG.API_BASE_URL, "");

  try {
    const result = RobinhoodApiClient.get(endpoint);
    if (result) {
      const header = Object.keys(result);
      const values = Object.values(result).map((value) =>
        typeof value === "object" && value !== null
          ? JSON.stringify(value)
          : value,
      );

      if (showHeader) {
        return [header, values];
      } else {
        return [values];
      }
    }
    return [["Error: Could not retrieve data from the URL: " + url]];
  } catch (e) {
    return [["Error: " + e.message]];
  }
}







// --- Menu Functions ---

function runLoginProcess() {
  PropertiesService.getUserProperties().deleteProperty(
    "robinhood_access_token",
  );
  // Cleared old token to start new login
  showLoginDialog_(); // This is the only change needed here
}

/**
 * Checks the current Robinhood login status by verifying the stored access token.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range.
 * @return {string} The current login status, either "Logged In" or "Logged Out".
 * @customfunction
 */
function ROBINHOOD_GET_LOGIN_STATUS(LastUpdate) {
  validateLastUpdate(LastUpdate);
  const token = PropertiesService.getUserProperties().getProperty(
    "robinhood_access_token",
  );
  if (!token) {
    return "Logged Out";
  }

  try {
    const endpoint = ROBINHOOD_CONFIG.API_URIS.accounts;
    const options = {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        Authorization: "Bearer " + token,
        "X-Robinhood-API-Version": "1.0.0",
      },
      payload: "",
      validateHttpsCertificates: true,
    };

    const response = UrlFetchApp.fetch(
      ROBINHOOD_CONFIG.API_BASE_URL + endpoint,
      options,
    );
    const responseCode = response.getResponseCode();

    if (responseCode >= 200 && responseCode < 300) {
      return "Logged In";
    } else {
      if (responseCode === 401) {
        Logger.log("Token expired, user needs to re-authenticate");
      }
      return "Logged Out";
    }
  } catch (e) {
    Logger.log(`Token validation failed: ${e.message || "Network error"}`);
    return "Logged Out";
  }
}

function refreshLastUpdate_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refresh");
  if (sheet) sheet.getRange("A1").setValue(new Date());
}

function showFunctionHelp() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Check if 'Function Help' sheet exists
  let helpSheet = spreadsheet.getSheetByName("Function Help");
  if (!helpSheet) {
    helpSheet = spreadsheet.insertSheet("Function Help");
  }

  // Clear existing content
  helpSheet.clear();

  // Get help data
  const helpData = ROBINHOOD_HELP();

  // Write data to sheet
  const range = helpSheet.getRange(1, 1, helpData.length, helpData[0].length);
  range.setValues(helpData);

  // Format the header row
  const headerRange = helpSheet.getRange(1, 1, 1, helpData[0].length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4285f4");
  headerRange.setFontColor("white");

  // Auto-resize columns
  helpSheet.autoResizeColumns(1, helpData[0].length);

  // Activate the help sheet
  helpSheet.activate();

  // Show info message
  SpreadsheetApp.getUi().alert(
    "Function Help Created!",
    'A "Function Help" sheet has been created with all available functions. You can also use =ROBINHOOD_HELP() in any cell to get this information.',
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let refreshSheet = spreadsheet.getSheetByName(REFRESH.sheet_name);
  if (refreshSheet === null) {
    refreshSheet = spreadsheet.insertSheet(REFRESH.sheet_name);
    refreshSheet.getRange(REFRESH.cell_address).setValue(new Date());
    refreshSheet.hideSheet();
  }

  // Step 2: Ensure the 'LastUpdate' named range exists.
  if (spreadsheet.getRangeByName(REFRESH.named_range_name) === null) {
    const range = refreshSheet.getRange(REFRESH.cell_address);
    spreadsheet.setNamedRange(REFRESH.named_range_name, range);
    Logger.log(`Named range '${REFRESH.named_range_name}' created.`);
  }

  ui.createMenu("Robinhood")
    .addItem("Login / Re-login", "runLoginProcess")
    .addItem("Refresh Data", "refreshLastUpdate_")
    .addSeparator()
    .addItem("📋 Show All Functions", "showFunctionHelp")
    .addToUi();
}
