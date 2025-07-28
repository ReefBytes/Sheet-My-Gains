/**
 * A collection of constants for the Robinhood API.
 */
const ROBINHOOD_CONFIG = {
  API_BASE_URL: 'https://api.robinhood.com',
  TOKEN_URL: 'https://api.robinhood.com/oauth2/token/',
  CLIENT_ID: 'c82SH0WZOsabOXGP2sxqcj34FxkvfnWRZBKlBjFS',
  API_URIS: {
    accounts: '/accounts/',
    achTransfers: '/ach/transfers/',
    dividends: '/dividends/',
    documents: '/documents/',
    marketData: '/marketdata/options/?instruments=',
    optionsOrders: '/options/orders/',
    optionsPositions: '/options/positions/',
    orders: '/orders/',
    portfolios: '/portfolios/',
    positions: '/positions/',
    watchlist: '/watchlists/Default/',
    pathfinderUserMachine: '/pathfinder/user_machine/',
    pathfinderInquiries: '/pathfinder/inquiries/',
    challenge: '/challenge/',
    push: '/push/',
    quotes: '/marketdata/quotes/',
    historicals: '/marketdata/historicals/',
    identi: 'https://identi.robinhood.com/idl/v1/workflow/'
  }
};

/**
 * A wrapper for making authenticated requests to the Robinhood API.
 */
const RobinhoodApiClient = (function() {
  const service_ = PropertiesService.getUserProperties();

  function getAuthToken_() {
    const token = service_.getProperty('robinhood_access_token');
    if (!token) {
      throw new Error('Authentication required. Please run from the "Robinhood > Login / Re-login" menu.');
    }
    return token;
  }

  function makeRequest_(url, options, retryCount = 0) {
    const MAX_RETRIES = 3;
    const INITIAL_WAIT_TIME = 5000;

    Logger.log(`Making request to: ${url} (Attempt ${retryCount + 1})`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 429 && retryCount < MAX_RETRIES) {
      const waitTime = INITIAL_WAIT_TIME * Math.pow(2, retryCount) + Math.random() * 1000;
      Logger.log(`Rate limited (429). Waiting ${waitTime / 1000} seconds before retrying...`);
      Utilities.sleep(waitTime);
      return makeRequest_(url, options, retryCount + 1);
    }

    Logger.log(`Response [${responseCode}]: ${responseText}`);

    if (responseCode >= 200 && responseCode < 400) {
      try {
        return JSON.parse(responseText);
      } catch (e) {
        return responseText;
      }
    } else if (responseCode === 401 && options.headers && options.headers.Authorization) {
      Logger.log('Token may have expired (401 Unauthorized). Clearing token.');
      service_.deleteProperty('robinhood_access_token');
      throw new Error('Your session has expired. Please log in again via the menu.');
    } else {
      if (url === ROBINHOOD_CONFIG.TOKEN_URL && (responseCode === 400 || responseCode === 401 || responseCode === 403)) {
        try {
            return JSON.parse(responseText);
        } catch(e) {
            throw new Error(`API request failed. ${responseCode}: ${responseText}`);
        }
      }
      throw new Error(`API request failed. ${responseCode}: ${responseText}`);
    }
  }

  function get(url) {
    const token = getAuthToken_();
    const options = {
      'method': 'get',
      'muteHttpExceptions': true,
      'headers': {
        'Authorization': 'Bearer ' + token
      }
    };
    return makeRequest_(ROBINHOOD_CONFIG.API_BASE_URL + url, options);
  }

  function pagedGet(url) {
    let fullUrl = ROBINHOOD_CONFIG.API_BASE_URL + url;
    let responseJson = makeRequest_(fullUrl, {
        method: 'get',
        muteHttpExceptions: true,
        headers: { 'Authorization': 'Bearer ' + getAuthToken_() }
    });
    let results = responseJson.results;
    let nextUrl = responseJson.next;
    while (nextUrl) {
      responseJson = makeRequest_(nextUrl, {
        method: 'get',
        muteHttpExceptions: true,
        headers: { 'Authorization': 'Bearer ' + getAuthToken_() }
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
    makeRequest: makeRequest_
  };
})();

/**
 * Generates a cryptographically secure-like device token.
 */
function generateDeviceToken_() {
  let token = '';
  const chars = '0123456789abcdef';
  for (let i = 0; i < 36; i++) {
    if (i === 8 || i === 13 || i === 18 || i === 23) {
      token += '-';
    } else {
      token += chars.charAt(Math.floor(Math.random() * chars.length));
    }
  }
  Logger.log(`Device Token Generated: ${token}`)
  return token;
}

/**
 * Handles the new "Sherrif" verification workflow.
 */
function validateSherrifId_(deviceToken, workflowId) {
    const ui = SpreadsheetApp.getUi();
    Logger.log(`Starting Sheriff validation for workflow ID: ${workflowId}`);

    // Step 1: Trigger the challenge by sending a PATCH request to the identi endpoint
    const identiUrl = ROBINHOOD_CONFIG.API_URIS.identi + workflowId + '/';
    const triggerPayload = { "clientVersion": "1.0.0", "id": workflowId, "entryPointAction": {} };
    const triggerOptions = {
        method: 'patch',
        contentType: 'application/json',
        payload: JSON.stringify(triggerPayload),
        muteHttpExceptions: true
    };
    Logger.log(`Triggering challenge at: ${identiUrl}`);
    const triggerResponse = RobinhoodApiClient.makeRequest(identiUrl, triggerOptions);

    if (!triggerResponse || !triggerResponse.route || !triggerResponse.route.replace || !triggerResponse.route.replace.screen || !triggerResponse.route.replace.screen.deviceApprovalChallengeScreenParams || !triggerResponse.route.replace.screen.deviceApprovalChallengeScreenParams.sheriffChallenge) {
        throw new Error("Failed to trigger the MFA challenge. Invalid response from server.");
    }
    const challenge = triggerResponse.route.replace.screen.deviceApprovalChallengeScreenParams.sheriffChallenge;
    Logger.log(`Challenge received: ${JSON.stringify(challenge)}`);

    const startTime = new Date().getTime();
    const timeout = 120 * 1000; // 2 minutes

    // Step 2: Handle the challenge based on its type
    if (challenge.type === 'PROMPT') {
        ui.alert("Login Verification Required", "Please check your Robinhood app to approve this login attempt.", ui.ButtonSet.OK);
        const promptStatusUrl = ROBINHOOD_CONFIG.API_BASE_URL + ROBINHOOD_CONFIG.API_URIS.push + `${challenge.id}/get_prompts_status/`;

        while (new Date().getTime() - startTime < timeout) {
            Logger.log(`Polling for push notification validation at: ${promptStatusUrl}`);
            const promptStatusResponse = RobinhoodApiClient.makeRequest(promptStatusUrl, { method: 'get', muteHttpExceptions: true });

            if (promptStatusResponse && promptStatusResponse.challenge_status === 'validated') {
                Logger.log("Push notification successfully validated by user.");

                // Step 3: Finalize the workflow
                const finalizePayload = { "clientVersion": "1.0.0", "screenName": "DEVICE_APPROVAL_CHALLENGE", "id": workflowId, "deviceApprovalChallengeAction": { "proceed": {} } };
                const finalizeOptions = {
                    method: 'patch',
                    contentType: 'application/json',
                    payload: JSON.stringify(finalizePayload),
                    muteHttpExceptions: true
                };
                Logger.log(`Finalizing workflow at: ${identiUrl}`);
                const finalizeResponse = RobinhoodApiClient.makeRequest(identiUrl, finalizeOptions);

                if (finalizeResponse && finalizeResponse.route && finalizeResponse.route.exit && finalizeResponse.route.exit.status === 'WORKFLOW_STATUS_APPROVED') {
                    Logger.log("Workflow successfully approved.");
                    return; // Validation successful
                } else {
                    throw new Error(`Failed to finalize the workflow. Response: ${JSON.stringify(finalizeResponse)}`);
                }
            }
            Utilities.sleep(5000); // Wait 5 seconds before polling again
        }
        throw new Error("Login approval timed out. You did not approve the push notification on your device within the 2-minute window.");
    }
    // ... (Optional: Add handling for SMS/Email challenges here if needed) ...
    else {
        throw new Error(`Unsupported challenge type: ${challenge.type}`);
    }
}


/**
 * Runs the full interactive authentication flow.
 */
function runInteractiveLoginFlow_() {
  const ui = SpreadsheetApp.getUi();
  Logger.log('Starting interactive login flow...');

  try {
    const username = ui.prompt('Robinhood Login', 'Enter your Robinhood email:', ui.ButtonSet.OK_CANCEL).getResponseText();
    if (!username) { Logger.log('Login cancelled by user.'); return; }
    const password = ui.prompt('Robinhood Login', 'Enter your Robinhood password:', ui.ButtonSet.OK_CANCEL).getResponseText();
    if (!password) { Logger.log('Login cancelled by user.'); return; }

    const deviceToken = generateDeviceToken_();

    const loginPayload = {
      'client_id': ROBINHOOD_CONFIG.CLIENT_ID,
      'expires_in': 86400,
      'grant_type': 'password',
      'password': password,
      'scope': 'internal',
      'username': username,
      'device_token': deviceToken,
      'try_passkeys': false,
      'token_request_path': '/login/',
      'create_read_only_secondary_token': true
    };

    const loginOptions = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(loginPayload),
      muteHttpExceptions: true
    };

    Logger.log('Attempting initial login...');
    let loginResponse = RobinhoodApiClient.makeRequest(ROBINHOOD_CONFIG.TOKEN_URL, loginOptions);
    let finalTokenResponse;

    if (loginResponse && loginResponse.verification_workflow) {
      Logger.log("Verification required. Starting Sheriff flow...");
      validateSherrifId_(deviceToken, loginResponse.verification_workflow.id);

      Logger.log("Verification complete. Re-attempting login to get final token...");
      finalTokenResponse = RobinhoodApiClient.makeRequest(ROBINHOOD_CONFIG.TOKEN_URL, loginOptions);

    } else if (loginResponse && loginResponse.access_token) {
      Logger.log("Login successful without MFA.");
      finalTokenResponse = loginResponse;
    } else {
      throw new Error(`Login failed. Initial response from server: ${JSON.stringify(loginResponse)}`);
    }

    if (finalTokenResponse && finalTokenResponse.access_token) {
      PropertiesService.getUserProperties().setProperty('robinhood_access_token', finalTokenResponse.access_token);
      ui.alert('Success!', 'Successfully authenticated with Robinhood.', ui.ButtonSet.OK);
    } else {
      throw new Error(`Failed to retrieve final access token. Response: ${JSON.stringify(finalTokenResponse)}`);
    }

  } catch (e) {
    Logger.log(`An error occurred in the login flow: ${e.toString()}`);
    ui.alert('Login Error', e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Recursively unpacks and flattens a result from a Robinhood API endpoint.
 */
function flattenResult_(result, flattenedResult, hyperlinkedFields, originalEndpointName) {
  for (const key in result) {
    if (Object.prototype.hasOwnProperty.call(result, key)) {
      const value = result[key];

      if (hyperlinkedFields.includes(key) && typeof value === 'string' && value.startsWith('http')) {
        const responseJson = RobinhoodApiClient.makeRequest(value, {
          'method': 'get',
          'muteHttpExceptions': true,
          'headers': {
            'Authorization': 'Bearer ' + PropertiesService.getUserProperties().getProperty('robinhood_access_token')
          }
        });
        const nextHyperlinkedFields = hyperlinkedFields.slice();
        nextHyperlinkedFields.splice(nextHyperlinkedFields.indexOf(key), 1);
        flattenResult_(responseJson, flattenedResult, nextHyperlinkedFields, key);
      } else if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
        flattenResult_(value, flattenedResult, hyperlinkedFields, originalEndpointName);
      } else if (Array.isArray(value) && key !== 'executions' && value.length > 0 && typeof value[0] === 'object') {
        flattenResult_(value[0], flattenedResult, hyperlinkedFields, originalEndpointName);
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
function getRobinhoodData_(endpoint, hyperlinkedFields) {
  try {
    let results = RobinhoodApiClient.pagedGet(ROBINHOOD_CONFIG.API_URIS[endpoint]);

    if (endpoint === "positions") {
      results = results.filter(row => parseFloat(row['quantity']) > 0);
    }

    if (!results || results.length === 0) {
      return [['No results found for ' + endpoint]];
    }

    const allFlattenedResults = [];
    const allKeys = new Set();

    results.forEach(result => {
      const flattenedResult = {};
      flattenResult_(result, flattenedResult, hyperlinkedFields.slice(), endpoint);
      allFlattenedResults.push(flattenedResult);
      Object.keys(flattenedResult).forEach(key => allKeys.add(key));
    });

    const header = Array.from(allKeys).sort();
    const data = [header];

    allFlattenedResults.forEach(flattenedResult => {
      const row = header.map(key => flattenedResult.hasOwnProperty(key) ? flattenedResult[key] : '');
      data.push(row);
    });

    return data;
  } catch (e) {
    return [['Error: ' + e.message]];
  }
}

// --- Custom Functions for Sheets ---

/**
 * Retrieves a history of ACH transfers.
 *
 * @param {any} datetime A cell reference (e.g., a "Last Refreshed" timestamp) to trigger recalculation.
 * @return {Array<Array<string>>} A two-dimensional array of ACH transfer data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_ACH_TRANSFERS(datetime) {
  return getRobinhoodData_('achTransfers', ['ach_relationship']);
}

/**
 * Retrieves dividend history for your account.
 *
 * @param {any} datetime A cell reference (e.g., a "Last Refreshed" timestamp) to trigger recalculation.
 * @return {Array<Array<string>>} A two-dimensional array of dividend data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_DIVIDENDS(datetime) {
  return getRobinhoodData_('dividends', ['instrument']);
}

/**
 * Retrieves a list of available documents, like statements and tax forms.
 *
 * @param {any} datetime A cell reference (e.g., a "Last Refreshed" timestamp) to trigger recalculation.
 * @return {Array<Array<string>>} A two-dimensional array of document data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_DOCUMENTS(datetime) {
  return getRobinhoodData_('documents', []);
}

/**
 * Retrieves a history of options orders.
 *
 * @param {any} datetime A cell reference (e.g., a "Last Refreshed" timestamp) to trigger recalculation.
 * @return {Array<Array<string>>} A two-dimensional array of options order data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_OPTIONS_ORDERS(datetime) {
  return getRobinhoodData_('optionsOrders', ['option']);
}

/**
 * Retrieves all current options positions.
 *
 * @param {any} datetime A cell reference (e.g., a "Last Refreshed" timestamp) to trigger recalculation.
 * @return {Array<Array<string>>} A two-dimensional array of options position data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_OPTIONS_POSITIONS(datetime) {
  return getRobinhoodData_('optionsPositions', ['option']);
}

/**
 * Retrieves a history of stock orders.
 *
 * @param {any} datetime A cell reference (e.g., a "Last Refreshed" timestamp) to trigger recalculation.
 * @return {Array<Array<string>>} A two-dimensional array of stock order data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_ORDERS(datetime) {
  return getRobinhoodData_('orders', ['instrument', 'position']);
}

/**
 * Retrieves portfolio data, including account value and history.
 *
 * @param {any} datetime A cell reference (e.g., a "Last Refreshed" timestamp) to trigger recalculation.
 * @return {Array<Array<string>>} A two-dimensional array of portfolio data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_PORTFOLIOS(datetime) {
  return getRobinhoodData_('portfolios', []);
}

/**
 * Retrieves all current stock positions.
 *
 * @param {any} datetime A cell reference (e.g., a "Last Refreshed" timestamp) to trigger recalculation.
 * @return {Array<Array<string>>} A two-dimensional array of stock position data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_POSITIONS(datetime) {
  return getRobinhoodData_('positions', ['instrument']);
}

/**
 * Retrieves instruments from your default watchlist.
 *
 * @param {any} datetime A cell reference (e.g., a "Last Refreshed" timestamp) to trigger recalculation.
 * @return {Array<Array<string>>} A two-dimensional array of watchlist data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_WATCHLIST(datetime) {
  return getRobinhoodData_('watchlist', ['instrument']);
}

/**
 * Retrieves the latest quote for a given stock ticker.
 *
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @param {boolean} [includeHeader=true] Optional. Set to false to exclude the header row from the output.
 * @return {Array<Array<string>>} The quote data including price, bid, and ask.
 * @customfunction
 */
function ROBINHOOD_GET_QUOTE(ticker, includeHeader = true) {
  if (!ticker) {
    return [['Error: Please provide a ticker symbol.']];
  }
  const endpoint = ROBINHOOD_CONFIG.API_URIS.quotes + ticker.toUpperCase() + '/';

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
    return [['Error: Could not find quote for ticker ' + ticker]];
  } catch (e) {
    return [['Error: ' + e.message]];
  }
}

/**
 * Retrieves historical price data for a given stock ticker.
 *
 * @param {string} ticker The stock ticker symbol (e.g., "TSLA").
 * @param {string} interval The time interval ('day', 'week', 'month'). Default is 'day'.
 * @param {string} span The time span ('week', 'month', '3month', 'year', '5year'). Default is 'year'.
 * @return {Array<Array<string>>} A 2D array of historical price data.
 * @customfunction
 */
function ROBINHOOD_GET_HISTORICALS(ticker, interval = 'day', span = 'year') {
  if (!ticker) {
    return [['Error: Please provide a ticker symbol.']];
  }
  const endpoint = `${ROBINHOOD_CONFIG.API_URIS.historicals}${ticker.toUpperCase()}/?interval=${interval}&span=${span}`;

  try {
    const result = RobinhoodApiClient.get(endpoint);
    if (result && result.historicals && result.historicals.length > 0) {
      const header = Object.keys(result.historicals[0]);
      const data = [header];
      result.historicals.forEach(row => {
        data.push(header.map(key => row[key]));
      });
      return data;
    }
    return [['Error: No historical data found for ' + ticker]];
  } catch (e) {
    return [['Error: ' + e.message]];
  }
}

/**
 * Retrieves detailed information for all brokerage accounts.
 *
 * @param {any} datetime A cell reference (e.g., a "Last Refreshed" timestamp) to trigger recalculation.
 * @return {Array<Array<string>>} A two-dimensional array of all account data for Google Sheets.
 * @customfunction
 */
function ROBINHOOD_GET_ACCOUNTS(datetime) {
  try {
    // This is the specific endpoint you requested
    const endpoint = ROBINHOOD_CONFIG.API_URIS.accounts + '?default_to_all_accounts=true&include_managed=true&include_multiple_individual=false&is_default=false';

    // Use the existing pagedGet client to fetch the data
    const accounts = RobinhoodApiClient.pagedGet(endpoint);

    if (!accounts || accounts.length === 0) {
      return [['Error: No accounts found.']];
    }

    // Automatically create a header from the keys of the first account
    const header = Object.keys(accounts[0]);
    const data = [header]; // The first row of our output table is the header

    // Iterate through each account object returned by the API
    accounts.forEach(account => {
      // For each account, create a row of data in the same order as the header
      const row = header.map(key => {
        const value = account[key];
        // If a value is a nested object (like margin_balances), convert it to a JSON string
        if (typeof value === 'object' && value !== null) {
          return JSON.stringify(value);
        }
        return value;
      });
      data.push(row);
    });

    return data;

  } catch (e) {
    return [['Error: ' + e.message]];
  }
}

// --- Menu Functions ---
function runLoginProcess() {
  PropertiesService.getUserProperties().deleteProperty('robinhood_access_token');
  Logger.log('Cleared old token to start new login.');
  runInteractiveLoginFlow_();
}

function refreshLastUpdate_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Refresh');
  if (sheet) sheet.getRange('A1').setValue(new Date());
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let refreshSheet = spreadsheet.getSheetByName('Refresh');
  if (refreshSheet === null) {
    refreshSheet = spreadsheet.insertSheet('Refresh');
    refreshSheet.getRange('A1').setValue(new Date());
    refreshSheet.hideSheet();
  }
  ui.createMenu('Robinhood')
    .addItem('Login / Re-login', 'runLoginProcess')
    .addItem('Refresh Data', 'refreshLastUpdate_')
    .addToUi();
}
