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

    Logger.log(`Making request to: ${url} (Attempt ${retryCount + 1})`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 429 && retryCount < MAX_RETRIES) {
      const waitTime =
        INITIAL_WAIT_TIME * Math.pow(2, retryCount) + Math.random() * 1000;
      const statusMessage = `Rate limited. Retrying in ${Math.round(waitTime / 1000)}s... (Attempt ${retryCount + 1})`;
      properties.setProperty("robinhood_retry_status", statusMessage);
      Logger.log(statusMessage);

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
      // Otherwise, log the full response as before for debugging
      Logger.log(`Response [${responseCode}]: ${responseText}`);
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

  function pagedGet(url) {
    let fullUrl = ROBINHOOD_CONFIG.API_BASE_URL + url;
    let responseJson = makeRequest_(fullUrl, {
      method: "get",
      muteHttpExceptions: true,
      headers: { Authorization: "Bearer " + getAuthToken_() },
    });
    let results = responseJson.results;
    let nextUrl = responseJson.next;
    while (nextUrl) {
      responseJson = makeRequest_(nextUrl, {
        method: "get",
        muteHttpExceptions: true,
        headers: { Authorization: "Bearer " + getAuthToken_() },
      });
      if (responseJson.results) {
        results = results.concat(responseJson.results);
      }
      nextUrl = responseJson.next;
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
  Logger.log(`Device Token Generated: ${token}`);
  return token;
}

/**
 * Handles the new "Sherrif" verification workflow.
 */
function validateSherrifId_(deviceToken, workflowId) {
  const ui = SpreadsheetApp.getUi();
  Logger.log(`Starting Sheriff validation for workflow ID: ${workflowId}`);

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
  Logger.log(`Triggering challenge at: ${identiUrl}`);
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
  Logger.log(`Challenge received: ${JSON.stringify(challenge)}`);

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
      Logger.log(
        `Polling for push notification validation at: ${promptStatusUrl}`,
      );
      const promptStatusResponse = RobinhoodApiClient.makeRequest(
        promptStatusUrl,
        { method: "get", muteHttpExceptions: true },
      );

      if (
        promptStatusResponse &&
        promptStatusResponse.challenge_status === "validated"
      ) {
        Logger.log("Push notification successfully validated by user.");

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
        Logger.log(`Finalizing workflow at: ${identiUrl}`);
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
          Logger.log("Workflow successfully approved.");
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
  Logger.log("Starting login process from custom dialog...");

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

    Logger.log("Attempting initial login...");
    let loginResponse = RobinhoodApiClient.makeRequest(
      ROBINHOOD_CONFIG.TOKEN_URL,
      loginOptions,
    );
    let finalTokenResponse;

    if (loginResponse && loginResponse.verification_workflow) {
      Logger.log("Verification required. Starting Sheriff flow...");
      validateSherrifId_(deviceToken, loginResponse.verification_workflow.id);

      Logger.log(
        "Verification complete. Re-attempting login to get final token...",
      );
      finalTokenResponse = RobinhoodApiClient.makeRequest(
        ROBINHOOD_CONFIG.TOKEN_URL,
        loginOptions,
      );
    } else if (loginResponse && loginResponse.access_token) {
      Logger.log("Login successful without MFA.");
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
      Logger.log("Successfully authenticated with Robinhood.");
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
    Logger.log(`An error occurred in the login flow: ${e.toString()}`);
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
 * Retrieves a list of available documents, like statements and tax forms.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range, e.g., `=ROBINHOOD_GET_DOCUMENTS(LastUpdate)`.
 * @return {Array<Array<string>>} A two-dimensional array of document data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_DOCUMENTS(LastUpdate) {
  validateLastUpdate(LastUpdate);
  return getRobinhoodData_("documents", []);
}

/**
 * Retrieves a history of options orders.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range, e.g., `=ROBINHOOD_GET_OPTIONS_ORDERS(LastUpdate)`.
 * @return {Array<Array<string>>} A two-dimensional array of options order data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_OPTIONS_ORDERS(LastUpdate) {
  validateLastUpdate(LastUpdate);
  return getRobinhoodData_("optionsOrders", ["option"]);
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
  return getRobinhoodData_("optionsPositions", ["option"]);
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
 * Retrieves instruments from your default watchlist.
 *
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range, e.g., `=ROBINHOOD_GET_WATCHLIST(LastUpdate)`.
 * @return {Array<Array<string>>} A two-dimensional array of watchlist data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_WATCHLIST(LastUpdate) {
  validateLastUpdate(LastUpdate);
  return getRobinhoodData_("watchlist", []);
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
 * Retrieves historical price data for a given stock ticker.
 *
 * @param {string} ticker The stock ticker symbol (e.g., "TSLA").
 * @param {string} interval The time interval ('day', 'week', 'month'). Default is 'day'.
 * @param {string} span The time span ('week', 'month', '3month', 'year', '5year'). Default is 'year'.
 * @param {any} LastUpdate Required to enable automatic refreshing. Use the `LastUpdate` named range, e.g., `=ROBINHOOD_GET_HISTORICALS(LastUpdate)`.
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
  const showHeader = includeHeader === false ? false : true;

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
  Logger.log("Cleared old token to start new login.");
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
  Logger.log("Running ROBINHOOD_GET_LOGIN_STATUS...");
  validateLastUpdate(LastUpdate);

  const token = PropertiesService.getUserProperties().getProperty(
    "robinhood_access_token",
  );

  if (!token) {
    Logger.log("No access token found. User is logged out.");
    return "Logged Out";
  }
  Logger.log("Access token found. Verifying with API...");

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

    Logger.log("Fetching account details to validate token...");
    const response = UrlFetchApp.fetch(
      ROBINHOOD_CONFIG.API_BASE_URL + endpoint,
      options,
    );
    const responseCode = response.getResponseCode();
    Logger.log(`Token validation API response code: ${responseCode}`);

    if (responseCode >= 200 && responseCode < 300) {
      Logger.log("Token is valid. User is logged in.");
      return "Logged In";
    } else {
      Logger.log(
        `Token is invalid or expired (Response: ${responseCode}). User is logged out.`,
      );
      return "Logged Out";
    }
  } catch (e) {
    Logger.log(`An error occurred during token validation: ${e.toString()}`);
    return "Logged Out";
  }
}

function refreshLastUpdate_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refresh");
  if (sheet) sheet.getRange("A1").setValue(new Date());
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
    .addToUi();
}
