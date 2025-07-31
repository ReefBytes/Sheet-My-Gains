/**
 * @OnlyCurrentDoc
 */

// --- Constants ---

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
    watchlist: "/watchlists/Default/", // Legacy, discovery API is now used
    quotes: "/marketdata/quotes/",
    historicals: "/marketdata/historicals/",
    discoveryLists: "/discovery/lists/default/",
    discoveryListItems: "/discovery/lists/items/",
  },
  API_LIMITS: {
    MAX_RETRIES: 5,
    INITIAL_WAIT_TIME: 1000,
    MAX_TOTAL_WAIT_TIME: 90000, // 90 seconds
    DEFAULT_PAGE_SIZE: 1000,
    MAX_PAGE_SIZE: 1000,
    BATCH_CHUNK_SIZE: 5,
    BATCH_DELAY_MS: 200
  },
  CACHE: {
    EXPIRATION_SECONDS: 300, // 5 minutes
    LOGIN_STATUS_KEY: "login_status",
    DEFAULT_TTL: 300
  },
  VALIDATION: {
    TICKER_REGEX: /^[A-Z]{1,5}$/,
    ACCOUNT_REGEX: /^[A-Z0-9]{6,10}$/,
    VALID_TIMESPANS: ['day', 'week', 'month', '3month', 'year', '5year'],
    VALID_INTERVALS: ['5minute', 'day', 'week'],
    MAX_DAYS_LOOKBACK: 365 * 5, // 5 years
    MAX_RESULTS_LIMIT: 50000
  },
  PERFORMANCE: {
    MAX_DEPTH_FLATTEN: 3,
    SLEEP_BETWEEN_WATCHLISTS: 300,
    MFA_TIMEOUT_MS: 180000, // 3 minutes
    MFA_POLL_INTERVAL_MS: 5000
  }
};

const REFRESH = {
  sheet_name: "Refresh",
  named_range_name: "LastUpdate",
  cell_address: "A1",
};

// --- Caching Service ---

const CacheManager = {
  get: function (key) {
    try {
      const cache = CacheService.getScriptCache();
      const cachedValue = cache.get(key);
      return cachedValue ? JSON.parse(cachedValue) : null;
    } catch (e) {
      Logger.log(`CacheManager.get Error: ${e.message}`);
      return null;
    }
  },
  put: function (key, value, ttl = ROBINHOOD_CONFIG.CACHE.DEFAULT_TTL) {
    try {
      const cache = CacheService.getScriptCache();
      cache.put(key, JSON.stringify(value), ttl);
    } catch (e) {
      Logger.log(`CacheManager.put Error: ${e.message}`);
    }
  },
  clear: function (keyPrefix = null) {
    try {
      const cache = CacheService.getScriptCache();
      if (keyPrefix) {
        // Clear specific keys (limited by GAS capabilities)
        cache.remove(keyPrefix);
      } else {
        // Note: GAS doesn't have cache.clear(), so we'd need to track keys
        Logger.log("Cache clear requested - individual key removal required");
      }
    } catch (e) {
      Logger.log(`CacheManager.clear Error: ${e.message}`);
    }
  }
};

// --- Error Handling & Utilities ---

function createErrorOutput_(e, functionName) {
  const errorMessage = `Error in ${functionName}: ${e.message}. Check Logs for details.`;
  Logger.log(`Error in ${functionName}: ${e.message} (Stack: ${e.stack})`);
  return [[errorMessage]];
}

function createContextualError_(error, functionName, context = {}) {
  const commonSolutions = {
    401: "Please re-authenticate via Robinhood menu",
    429: "Rate limited - please wait before retrying",
    404: "Resource not found - check your parameters",
    400: "Invalid request parameters - check your input values"
  };
  
  let message = `Error in ${functionName}: ${error.message}`;
  
  if (context.httpCode && commonSolutions[context.httpCode]) {
    message += `\n💡 ${commonSolutions[context.httpCode]}`;
  }
  
  if (context.suggestion) {
    message += `\n💡 ${context.suggestion}`;
  }
  
  Logger.log(`${message} (Stack: ${error.stack})`);
  return [[message]];
}

function validateLastUpdate_(LastUpdate) {
  if (LastUpdate === undefined || LastUpdate === null) {
    throw new Error("The LastUpdate parameter is required for refreshing.");
  }
}

function checkPermissions_() {
  try {
    // Test permissions by checking if we can access required services
    // This is a lightweight check that doesn't require external URLs
    
    // Check if we can access PropertiesService (script.storage scope)
    PropertiesService.getUserProperties();
    
    // Check if we can access the spreadsheet
    SpreadsheetApp.getActiveSpreadsheet();
    
    // Note: We can't easily test UrlFetchApp without making a real request,
    // so we'll let the actual API calls handle permission errors
    
    return true;
  } catch (e) {
    if (e.message.includes("permission") || e.message.includes("Authorization")) {
      throw new Error("⚠️ Permission required: Please run 'Robinhood > Grant Permissions' from the menu, or check your Apps Script OAuth scopes.");
    }
    throw e;
  }
}

function validateInput_(value, type, options = {}) {
  switch (type) {
    case 'ticker':
      if (!value || typeof value !== 'string') return false;
      return ROBINHOOD_CONFIG.VALIDATION.TICKER_REGEX.test(value.toUpperCase());
    
    case 'accountNumber':
      if (!value || typeof value !== 'string') return false;
      return ROBINHOOD_CONFIG.VALIDATION.ACCOUNT_REGEX.test(value);
    
    case 'timespan':
      return ROBINHOOD_CONFIG.VALIDATION.VALID_TIMESPANS.includes(value?.toLowerCase());
    
    case 'interval':
      return ROBINHOOD_CONFIG.VALIDATION.VALID_INTERVALS.includes(value?.toLowerCase());
    
    case 'number':
      const num = Number(value);
      return !isNaN(num) && num >= (options.min || 0) && num <= (options.max || Infinity);
    
    case 'positiveNumber':
      const posNum = Number(value);
      return !isNaN(posNum) && posNum > 0 && posNum <= (options.max || ROBINHOOD_CONFIG.VALIDATION.MAX_RESULTS_LIMIT);
      
    case 'days':
      const days = Number(value);
      return !isNaN(days) && days >= 0 && days <= ROBINHOOD_CONFIG.VALIDATION.MAX_DAYS_LOOKBACK;
      
    case 'url':
      if (!value || typeof value !== 'string') return false;
      try {
        new URL(value);
        return true;
      } catch {
        return false;
      }
  }
  return true;
}

function normalizeParameters_(args, expectedParams) {
  const params = {};
  let lastUpdateIndex = -1;
  
  // Find LastUpdate parameter (should be a Range object, typically last parameter)
  for (let i = args.length - 1; i >= 0; i--) {
    if (args[i] && (typeof args[i] === 'object' || args[i].constructor?.name === 'Range')) {
      lastUpdateIndex = i;
      break;
    }
  }
  
  // If no Range found, assume LastUpdate is the last parameter
  if (lastUpdateIndex === -1) {
    lastUpdateIndex = args.length - 1;
  }
  
  // Map parameters based on position and expected types
  expectedParams.forEach((param, index) => {
    if (index < lastUpdateIndex) {
      const value = args[index];
      params[param.name] = value !== undefined ? value : param.default;
      
      // Validate parameter if validation type specified
      if (param.validate && !validateInput_(params[param.name], param.validate, param.options)) {
        throw new Error(`Invalid ${param.name}: expected ${param.validate}, got "${value}"`);
      }
    }
  });
  
  params.LastUpdate = args[lastUpdateIndex];
  return params;
}

// --- API Client ---

const RobinhoodApiClient = (function () {
  const service_ = PropertiesService.getUserProperties();

  function getAuthToken_() {
    const token = service_.getProperty("robinhood_access_token");
    if (!token) {
      const refreshed = refreshToken_();
      if (refreshed) {
        return service_.getProperty("robinhood_access_token");
      } else {
        throw new Error('Authentication required. Please run "Robinhood > Login / Re-login".');
      }
    }
    return token;
  }

  function refreshToken_() {
    const refreshToken = service_.getProperty("robinhood_refresh_token");
    if (!refreshToken) {
      return false;
    }

    const payload = {
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      client_id: ROBINHOOD_CONFIG.CLIENT_ID,
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(ROBINHOOD_CONFIG.TOKEN_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) {
      const tokenData = JSON.parse(responseText);
      service_.setProperties({
        robinhood_access_token: tokenData.access_token,
        robinhood_refresh_token: tokenData.refresh_token,
      });
      return true;
    } else {
      service_.deleteProperty("robinhood_access_token");
      service_.deleteProperty("robinhood_refresh_token");
      return false;
    }
  }

  function makeRequest_(url, options, retryCount = 0) {
    const { MAX_RETRIES, INITIAL_WAIT_TIME, MAX_TOTAL_WAIT_TIME } = ROBINHOOD_CONFIG.API_LIMITS;

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 429 && retryCount < MAX_RETRIES) {
      let waitTime = INITIAL_WAIT_TIME * Math.pow(2, retryCount) + Math.random() * 1000;
      if (waitTime > MAX_TOTAL_WAIT_TIME) {
        waitTime = MAX_TOTAL_WAIT_TIME;
      }
      Logger.log(`Rate limited. Retrying in ${Math.round(waitTime / 1000)}s... (attempt ${retryCount + 1}/${MAX_RETRIES})`);
      Utilities.sleep(waitTime);
      return makeRequest_(url, options, retryCount + 1);
    }

    if (url === ROBINHOOD_CONFIG.TOKEN_URL && responseCode >= 400) {
      try {
        return JSON.parse(responseText);
      } catch (e) {
        const error = new Error(`API request failed (${responseCode}): ${responseText}`);
        error.httpCode = responseCode;
        throw error;
      }
    }

    if (responseCode >= 200 && responseCode < 300) {
      try {
        return JSON.parse(responseText);
      } catch (e) {
        return responseText;
      }
    }

    if (responseCode === 401) {
      if (refreshToken_()) {
        return makeRequest_(url, options, retryCount); // Retry the original request
      } else {
        const error = new Error("Session expired. Please log in again via the menu.");
        error.httpCode = 401;
        throw error;
      }
    }

    const error = new Error(`API request failed (${responseCode}): ${responseText}`);
    error.httpCode = responseCode;
    error.responseText = responseText;
    
    // Add rate limit information for 429 errors
    if (responseCode === 429) {
      error.rateLimited = true;
      error.retryAfter = response.getHeaders()['Retry-After'] || "300"; // Default to 5 minutes
    }
    
    throw error;
  }

  function get(url) {
    const fullUrl = url.startsWith("http") ? url : ROBINHOOD_CONFIG.API_BASE_URL + url;
    const cachedResponse = CacheManager.get(fullUrl);
    if (cachedResponse) {
      return cachedResponse;
    }

    const token = getAuthToken_();
    const options = {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        Authorization: "Bearer " + token,
      },
    };

    const response = makeRequest_(fullUrl, options);
    CacheManager.put(fullUrl, response);
    return response;
  }

  function pagedGet(url, maxPages = 10) {
    let fullUrl = ROBINHOOD_CONFIG.API_BASE_URL + url;
    let results = [];
    let pageCount = 1;

    while (fullUrl && pageCount <= maxPages) {
      const responseJson = get(fullUrl);
      if (responseJson.results) {
        results = results.concat(responseJson.results);
      }
      fullUrl = responseJson.next;
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

// --- Core Data Processing ---

function flattenObject_(obj, prefix = "", depth = 0) {
  const MAX_DEPTH = ROBINHOOD_CONFIG.PERFORMANCE.MAX_DEPTH_FLATTEN;
  if (depth > MAX_DEPTH) {
    return { [prefix]: "[Data too deep]" };
  }

  return Object.keys(obj).reduce((acc, k) => {
    const pre = prefix.length ? prefix + "_" : "";
    const value = obj[k];
    if (value && typeof value === "object" && !Array.isArray(value) && Object.keys(value).length > 0) {
      Object.assign(acc, flattenObject_(value, pre + k, depth + 1));
    } else {
      acc[pre + k] = typeof value === "object" ? JSON.stringify(value) : value;
    }
    return acc;
  }, {});
}

function batchApiCalls_(endpoints, chunkSize = ROBINHOOD_CONFIG.API_LIMITS.BATCH_CHUNK_SIZE) {
  const results = [];
  for (let i = 0; i < endpoints.length; i += chunkSize) {
    const chunk = endpoints.slice(i, i + chunkSize);
    const chunkResults = chunk.map(endpoint => {
      try {
        return RobinhoodApiClient.get(endpoint);
      } catch (e) {
        Logger.log(`Batch API call failed for ${endpoint}: ${e.message}`);
        return null;
      }
    });
    results.push(...chunkResults.filter(result => result !== null));
    
    // Small delay between batches to avoid rate limiting
    if (i + chunkSize < endpoints.length) {
      Utilities.sleep(ROBINHOOD_CONFIG.API_LIMITS.BATCH_DELAY_MS);
    }
  }
  return results;
}

function addAccountFilter_(data, accountNumber) {
  if (!accountNumber) return data;
  
  return data.filter(item => {
    // Try different possible account field names
    const accountFields = ['account', 'account_number', 'account_url'];
    return accountFields.some(field => 
      item[field] && (
        item[field] === accountNumber || 
        (typeof item[field] === 'string' && item[field].includes(accountNumber))
      )
    );
  });
}

function formatDataForSheets_(data, options = {}) {
  if (!data || data.length === 0) return [["No data available"]];
  
  const formatted = data.map(item => {
    const flattened = flattenObject_(item);
    
    // Format specific field types
    Object.keys(flattened).forEach(key => {
      const value = flattened[key];
      
      if (options.formatCurrency && key.includes('amount') && !isNaN(Number(value))) {
        flattened[key] = ROBINHOOD_FORMAT_CURRENCY(Number(value));
      }
      if (options.formatDates && (key.includes('date') || key.includes('_at')) && value) {
        try {
          flattened[key] = new Date(value).toLocaleDateString();
        } catch (e) {
          // Keep original value if date parsing fails
        }
      }
      if (options.formatPercentages && key.includes('percent') && !isNaN(Number(value))) {
        const num = Number(value);
        flattened[key] = `${(num * 100).toFixed(2)}%`;
      }
    });
    
    return flattened;
  });
  
  return convertToSheetFormat_(formatted);
}

function convertToSheetFormat_(flattenedData) {
  if (!flattenedData || flattenedData.length === 0) return [["No data available"]];
  
  // Get all unique keys from all objects
  const allKeys = [...new Set(flattenedData.flatMap(obj => Object.keys(obj)))].sort();
  
  // Create header row
  const data = [allKeys];
  
  // Create data rows
  flattenedData.forEach(item => {
    data.push(allKeys.map(key => item[key] || ""));
  });
  
  return data;
}

function getRobinhoodData_(endpointName, options = {}) {
  const endpointUrl = options.endpoint || ROBINHOOD_CONFIG.API_URIS[endpointName];
  let results = RobinhoodApiClient.pagedGet(endpointUrl);

  if (endpointName === "positions") {
    results = results.filter((row) => parseFloat(row["quantity"]) > 0);
  }

  if (!results || results.length === 0) {
    return [["No results found for " + endpointName]];
  }

  const allFlattenedResults = results.map((result) => flattenObject_(result));

  const allKeys = [...allFlattenedResults.reduce((keys, res) => {
    Object.keys(res).forEach((key) => keys.add(key));
    return keys;
  }, new Set())].sort();

  const data = [allKeys];
  allFlattenedResults.forEach((flattenedResult) => {
    data.push(allKeys.map((key) => flattenedResult[key] || ""));
  });

  return data;
}

// --- Custom Functions for Sheets ---

function ROBINHOOD_HELP(category = "all") {
  const functions = [
    ["ROBINHOOD_GET_POSITIONS", "core", "Current stock positions", "ROBINHOOD_GET_POSITIONS(LastUpdate)"],
    ["ROBINHOOD_GET_ORDERS", "core", "Order history with date filtering", "ROBINHOOD_GET_ORDERS(30, 1000, LastUpdate)"],
    ["ROBINHOOD_GET_DIVIDENDS", "core", "Dividend payment history", "ROBINHOOD_GET_DIVIDENDS(LastUpdate)"],
    ["ROBINHOOD_GET_QUOTE", "core", "Real-time stock quote", 'ROBINHOOD_GET_QUOTE("AAPL", TRUE, LastUpdate)'],
    ["ROBINHOOD_GET_QUOTES_BATCH", "core", "Multiple stock quotes efficiently", 'ROBINHOOD_GET_QUOTES_BATCH("AAPL,MSFT", LastUpdate)'],
    ["ROBINHOOD_GET_HISTORICALS", "core", "Historical price data", 'ROBINHOOD_GET_HISTORICALS("AAPL", "day", "year", LastUpdate)'],
    ["ROBINHOOD_GET_PORTFOLIOS", "core", "Portfolio summary data", "ROBINHOOD_GET_PORTFOLIOS(LastUpdate)"],
    ["ROBINHOOD_GET_ACCOUNTS", "core", "Account information", "ROBINHOOD_GET_ACCOUNTS(LastUpdate)"],
    ["ROBINHOOD_GET_WATCHLISTS", "core", "List all available watchlists", "ROBINHOOD_GET_WATCHLISTS(LastUpdate)"],
    ["ROBINHOOD_GET_WATCHLIST", "core", "Instruments in a specific watchlist", 'ROBINHOOD_GET_WATCHLIST("Default", LastUpdate)'],
    ["ROBINHOOD_GET_ALL_WATCHLISTS", "core", "All instruments from all watchlists", "ROBINHOOD_GET_ALL_WATCHLISTS(LastUpdate)"],
    ["ROBINHOOD_GET_ACH_TRANSFERS", "core", "ACH transfer history", "ROBINHOOD_GET_ACH_TRANSFERS(LastUpdate)"],
    ["ROBINHOOD_GET_PORTFOLIO_HISTORICALS", "analytics", "Portfolio performance over time", 'ROBINHOOD_GET_PORTFOLIO_HISTORICALS("year", "day", "5DP12345", LastUpdate)'],
    ["ROBINHOOD_GET_OPTIONS_POSITIONS", "options", "Current options positions", "ROBINHOOD_GET_OPTIONS_POSITIONS(LastUpdate)"],
    ["ROBINHOOD_GET_OPTIONS_ORDERS", "options", "Options order history", "ROBINHOOD_GET_OPTIONS_ORDERS(30, 100, LastUpdate)"],
    ["ROBINHOOD_VALIDATE_TICKER", "utility", "Validate ticker symbol format", 'ROBINHOOD_VALIDATE_TICKER("AAPL")'],
    ["ROBINHOOD_FORMAT_CURRENCY", "utility", "Format number as currency", "ROBINHOOD_FORMAT_CURRENCY(1234.56)"],
    ["ROBINHOOD_LAST_MARKET_DAY", "utility", "Get last trading day", "ROBINHOOD_LAST_MARKET_DAY()"],
    ["ROBINHOOD_GET_LOGIN_STATUS", "utility", "Check authentication status", "ROBINHOOD_GET_LOGIN_STATUS(LastUpdate)"],
    ["ROBINHOOD_GET_URL", "utility", "Get data from any Robinhood API URL", 'ROBINHOOD_GET_URL("https://api.robinhood.com/accounts/", LastUpdate)'],
    ["ROBINHOOD_HELP", "utility", "Show all available functions", 'ROBINHOOD_HELP("core")'],
  ];

  let filteredFunctions = functions;
  if (category && category.toLowerCase() !== "all") {
    filteredFunctions = functions.filter((func) => func[1] === category.toLowerCase());
  }

  const result = [["Function Name", "Category", "Description", "Example Usage"]];
  filteredFunctions.forEach((func) => {
    result.push([func[0], func[1].toUpperCase(), func[2], func[3]]);
  });

  return result;
}

function ROBINHOOD_VALIDATE_TICKER(ticker) {
  return validateInput_(ticker, 'ticker');
}

function ROBINHOOD_FORMAT_CURRENCY(amount) {
  if (typeof amount !== "number") return "$0.00";
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(amount);
}

function ROBINHOOD_LAST_MARKET_DAY() {
  const today = new Date();
  let lastMarketDay = new Date(today);
  while (lastMarketDay.getDay() === 0 || lastMarketDay.getDay() === 6) {
    lastMarketDay.setDate(lastMarketDay.getDate() - 1);
  }
  return lastMarketDay.toISOString().split("T")[0];
}

function ROBINHOOD_GET_ACH_TRANSFERS(LastUpdate) {
  try {
    validateLastUpdate_(LastUpdate);
    return getRobinhoodData_("achTransfers");
  } catch (e) {
    return createErrorOutput_(e, "ROBINHOOD_GET_ACH_TRANSFERS");
  }
}

function ROBINHOOD_GET_DIVIDENDS(LastUpdate) {
  try {
    validateLastUpdate_(LastUpdate);
    return getRobinhoodData_("dividends");
  } catch (e) {
    return createErrorOutput_(e, "ROBINHOOD_GET_DIVIDENDS");
  }
}

function ROBINHOOD_GET_OPTIONS_ORDERS(days = 30, pageSize = 100, LastUpdate) {
  try {
    validateLastUpdate_(LastUpdate);
    let endpoint = `${ROBINHOOD_CONFIG.API_URIS.optionsOrders}?page_size=${Math.min(pageSize, 1000)}`;
    if (days > 0) {
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - days);
      endpoint += `&updated_at[gte]=${cutoffDate.toISOString()}`;
    }
    return getRobinhoodData_("optionsOrders", { endpoint: endpoint });
  } catch (e) {
    return createErrorOutput_(e, "ROBINHOOD_GET_OPTIONS_ORDERS");
  }
}

function ROBINHOOD_GET_OPTIONS_POSITIONS(LastUpdate) {
  try {
    validateLastUpdate_(LastUpdate);
    return getRobinhoodData_("optionsPositions");
  } catch (e) {
    return createErrorOutput_(e, "ROBINHOOD_GET_OPTIONS_POSITIONS");
  }
}

function ROBINHOOD_GET_ORDERS(days = 30, pageSize = 1000, LastUpdate) {
  try {
    const expectedParams = [
      { name: 'days', default: 30, validate: 'days' },
      { name: 'pageSize', default: 1000, validate: 'positiveNumber', options: { max: ROBINHOOD_CONFIG.API_LIMITS.MAX_PAGE_SIZE } }
    ];
    
    const params = normalizeParameters_(arguments, expectedParams);
    validateLastUpdate_(params.LastUpdate);

    let endpoint = `${ROBINHOOD_CONFIG.API_URIS.orders}?page_size=${Math.min(params.pageSize, ROBINHOOD_CONFIG.API_LIMITS.MAX_PAGE_SIZE)}`;
    
    if (params.days > 0) {
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - params.days);
      endpoint += `&updated_at[gte]=${cutoffDate.toISOString()}`;
    }
    
    return getRobinhoodData_("orders", { endpoint: endpoint });
  } catch (e) {
    const context = { 
      httpCode: e.httpCode,
      suggestion: `Days should be 0-${ROBINHOOD_CONFIG.VALIDATION.MAX_DAYS_LOOKBACK}, pageSize should be 1-${ROBINHOOD_CONFIG.API_LIMITS.MAX_PAGE_SIZE}`
    };
    return createContextualError_(e, "ROBINHOOD_GET_ORDERS", context);
  }
}

function ROBINHOOD_GET_PORTFOLIOS(LastUpdate) {
  try {
    validateLastUpdate_(LastUpdate);
    return getRobinhoodData_("portfolios");
  } catch (e) {
    return createErrorOutput_(e, "ROBINHOOD_GET_PORTFOLIOS");
  }
}

function ROBINHOOD_GET_PORTFOLIO_HISTORICALS(span = "year", interval = "day", accountNumber, LastUpdate) {
  try {
    if (typeof accountNumber !== "string") {
      LastUpdate = accountNumber;
      accountNumber = null;
    }
    validateLastUpdate_(LastUpdate);

    const accounts = RobinhoodApiClient.pagedGet(ROBINHOOD_CONFIG.API_URIS.accounts);
    if (!accounts || accounts.length === 0) return [["No accounts found."]];

    const targetAccounts = accountNumber ? accounts.filter((acc) => acc.account_number === accountNumber) : accounts;
    if (targetAccounts.length === 0) return [[`Account ${accountNumber} not found.`]];

    const allHistoricals = [];
    targetAccounts.forEach((account) => {
      const endpoint = `${ROBINHOOD_CONFIG.API_URIS.portfolioHistoricals}${account.account_number}/?span=${span}&interval=${interval}`;
      const result = RobinhoodApiClient.get(endpoint);
      if (result && result.equity_historicals) {
        result.equity_historicals.forEach(h => {
          allHistoricals.push({ account_number: account.account_number, ...h });
        });
      }
    });

    if (allHistoricals.length === 0) return [["No portfolio historicals found."]];
    
    const header = Object.keys(allHistoricals[0]);
    const data = [header];
    allHistoricals.forEach(h => data.push(header.map(key => h[key])));

    return data;
  } catch (e) {
    return createErrorOutput_(e, "ROBINHOOD_GET_PORTFOLIO_HISTORICALS");
  }
}

function ROBINHOOD_GET_POSITIONS(LastUpdate) {
  try {
    validateLastUpdate_(LastUpdate);
    return getRobinhoodData_("positions");
  } catch (e) {
    return createErrorOutput_(e, "ROBINHOOD_GET_POSITIONS");
  }
}

function ROBINHOOD_GET_QUOTE(ticker, includeHeader = true, LastUpdate) {
  try {
    const expectedParams = [
      { name: 'ticker', validate: 'ticker' },
      { name: 'includeHeader', default: true }
    ];
    
    const params = normalizeParameters_(arguments, expectedParams);
    validateLastUpdate_(params.LastUpdate);

    const endpoint = `${ROBINHOOD_CONFIG.API_URIS.quotes}${params.ticker.toUpperCase()}/`;
    const result = RobinhoodApiClient.get(endpoint);

    if (!result || !result.symbol) {
      return [[`Could not find quote for ${params.ticker}`]];
    }
    
    const flattened = flattenObject_(result);
    const header = Object.keys(flattened);
    const values = header.map(key => flattened[key]);

    return params.includeHeader ? [header, values] : [values];
  } catch (e) {
    const context = { 
      httpCode: e.httpCode,
      suggestion: "Check ticker symbol format (1-5 letters, e.g., 'AAPL')"
    };
    return createContextualError_(e, "ROBINHOOD_GET_QUOTE", context);
  }
}

function ROBINHOOD_GET_QUOTES_BATCH(tickers, LastUpdate) {
  try {
    validateLastUpdate_(LastUpdate);
    if (!tickers) throw new Error("Ticker symbols are required.");

    const tickerList = tickers.toString().split(",").map((t) => t.trim().toUpperCase());
    const endpoint = `${ROBINHOOD_CONFIG.API_URIS.quotes}?symbols=${tickerList.join(",")}`;
    const result = RobinhoodApiClient.get(endpoint);

    if (!result || !result.results || result.results.length === 0) return [["No quotes found."]];

    const allFlattened = result.results.map(q => flattenObject_(q));
    const header = [...new Set(allFlattened.flatMap(Object.keys))].sort();
    const data = [header];
    allFlattened.forEach(quote => {
      data.push(header.map(key => quote[key] || ""));
    });

    return data;
  } catch (e) {
    return createErrorOutput_(e, "ROBINHOOD_GET_QUOTES_BATCH");
  }
}

function ROBINHOOD_GET_HISTORICALS(ticker, interval = "day", span = "year", LastUpdate) {
  try {
    const expectedParams = [
      { name: 'ticker', validate: 'ticker' },
      { name: 'interval', default: 'day', validate: 'interval' },
      { name: 'span', default: 'year', validate: 'timespan' }
    ];
    
    const params = normalizeParameters_(arguments, expectedParams);
    validateLastUpdate_(params.LastUpdate);

    const endpoint = `${ROBINHOOD_CONFIG.API_URIS.historicals}${params.ticker.toUpperCase()}/?interval=${params.interval}&span=${params.span}`;
    const result = RobinhoodApiClient.get(endpoint);

    if (!result || !result.historicals || result.historicals.length === 0) {
      return [["No historical data found."]];
    }

    return formatDataForSheets_(result.historicals, { formatDates: true });
  } catch (e) {
    const context = { 
      httpCode: e.httpCode,
      suggestion: "Check ticker symbol format (1-5 letters) and timespan/interval values"
    };
    return createContextualError_(e, "ROBINHOOD_GET_HISTORICALS", context);
  }
}

function ROBINHOOD_GET_ACCOUNTS(LastUpdate) {
  try {
    validateLastUpdate_(LastUpdate);
    const endpoint = `${ROBINHOOD_CONFIG.API_URIS.accounts}?default_to_all_accounts=true&include_managed=true`;
    return getRobinhoodData_("accounts", { endpoint: endpoint });
  } catch (e) {
    return createErrorOutput_(e, "ROBINHOOD_GET_ACCOUNTS");
  }
}

function ROBINHOOD_GET_URL(url, includeHeader = true, LastUpdate) {
  try {
    validateLastUpdate_(LastUpdate);
    if (!url) throw new Error("URL is required.");

    const endpoint = url.replace(ROBINHOOD_CONFIG.API_BASE_URL, "");
    const result = RobinhoodApiClient.get(endpoint);

    if (!result) return [["No data found at URL."]];

    if (result.results) {
        return getRobinhoodData_("", {endpoint: endpoint});
    }

    const flattened = flattenObject_(result);
    const header = Object.keys(flattened);
    const values = header.map(key => flattened[key]);

    return includeHeader ? [header, values] : [values];
  } catch (e) {
    return createErrorOutput_(e, "ROBINHOOD_GET_URL");
  }
}

function ROBINHOOD_GET_WATCHLISTS(LastUpdate) {
    try {
        validateLastUpdate_(LastUpdate);
        const endpoint = ROBINHOOD_CONFIG.API_URIS.discoveryLists;
        const result = RobinhoodApiClient.pagedGet(endpoint);

        if (!result || result.length === 0) return [["No watchlists found."]];

        const header = ["name", "id", "item_count", "owner_type", "icon_emoji", "created_at", "updated_at"];
        const data = [header];
        result.forEach(list => {
            data.push([
                list.display_name || "N/A",
                list.id || "N/A",
                list.item_count || 0,
                list.owner_type || "N/A",
                list.icon_emoji || "N/A",
                list.created_at || "N/A",
                list.updated_at || "N/A",
            ]);
        });
        return data;
    } catch (e) {
        return createErrorOutput_(e, "ROBINHOOD_GET_WATCHLISTS");
    }
}

function ROBINHOOD_GET_WATCHLIST(watchlistNameOrId, LastUpdate) {
    try {
        validateLastUpdate_(LastUpdate);
        if (!watchlistNameOrId) throw new Error("Watchlist name or ID is required.");

        const allLists = RobinhoodApiClient.pagedGet(ROBINHOOD_CONFIG.API_URIS.discoveryLists);
        const list = allLists.find(l => l.id === watchlistNameOrId || (l.display_name && l.display_name.toLowerCase() === watchlistNameOrId.toLowerCase()));

        if (!list) return [[`Watchlist "${watchlistNameOrId}" not found.`]];

        const localMidnightISO = new Date(new Date().setHours(0, 0, 0, 0)).toISOString();
        const itemsEndpoint = `${ROBINHOOD_CONFIG.API_URIS.discoveryListItems}?list_id=${list.id}&local_midnight=${encodeURIComponent(localMidnightISO)}`;
        const items = RobinhoodApiClient.pagedGet(itemsEndpoint);

        if (!items || items.length === 0) return [[`Watchlist "${list.display_name}" is empty.`]];

        const header = ["symbol", "name", "price", "one_day_change", "one_day_percent_change", "market_cap", "state", "added_at", "watchlist_name"];
        const data = [header];
        items.forEach(item => {
            data.push([
                item.symbol || "N/A",
                item.name || "N/A",
                item.price || "N/A",
                item.one_day_dollar_change || "N/A",
                item.one_day_percent_change ? `${(item.one_day_percent_change * 100).toFixed(2)}%` : "N/A",
                item.market_cap || "N/A",
                item.state || "N/A",
                item.created_at || "N/A",
                list.display_name,
            ]);
        });
        return data;
    } catch (e) {
        return createErrorOutput_(e, "ROBINHOOD_GET_WATCHLIST");
    }
}

function ROBINHOOD_GET_ALL_WATCHLISTS(LastUpdate) {
    try {
        validateLastUpdate_(LastUpdate);
        const allLists = RobinhoodApiClient.pagedGet(ROBINHOOD_CONFIG.API_URIS.discoveryLists);
        if (!allLists || allLists.length === 0) return [["No watchlists found."]];

        const allInstruments = [];
        const localMidnightISO = new Date(new Date().setHours(0, 0, 0, 0)).toISOString();

        allLists.forEach(list => {
            if (list.item_count === 0) return;
            const itemsEndpoint = `${ROBINHOOD_CONFIG.API_URIS.discoveryListItems}?list_id=${list.id}&local_midnight=${encodeURIComponent(localMidnightISO)}`;
            const items = RobinhoodApiClient.pagedGet(itemsEndpoint);
            if (items) {
                items.forEach(item => allInstruments.push({ ...item, watchlist_name: list.display_name }));
            }
            Utilities.sleep(ROBINHOOD_CONFIG.PERFORMANCE.SLEEP_BETWEEN_WATCHLISTS);
        });

        if (allInstruments.length === 0) return [["No instruments found in any watchlist."]];

        const header = ["watchlist_name", "symbol", "name", "price", "one_day_change", "one_day_percent_change", "market_cap", "state"];
        const data = [header];
        allInstruments.forEach(item => {
            data.push([
                item.watchlist_name,
                item.symbol || "N/A",
                item.name || "N/A",
                item.price || "N/A",
                item.one_day_dollar_change || "N/A",
                item.one_day_percent_change ? `${(item.one_day_percent_change * 100).toFixed(2)}%` : "N/A",
                item.market_cap || "N/A",
                item.state || "N/A",
            ]);
        });
        return data;
    } catch (e) {
        return createErrorOutput_(e, "ROBINHOOD_GET_ALL_WATCHLISTS");
    }
}

// --- Menu & Login Functions ---

/**
 * Requests necessary permissions for the script to function.
 * This function should be run once when first setting up the script.
 */
function requestPermissions() {
  try {
    // Test spreadsheet access
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // Test script properties access
    const userProps = PropertiesService.getUserProperties();
    const scriptProps = PropertiesService.getScriptProperties();
    
    // Test by making a simple Robinhood API call (this will trigger permission request for external URLs)
    try {
      // This will either work or trigger the permission dialog
      UrlFetchApp.fetch("https://api.robinhood.com/", {
        method: "GET",
        muteHttpExceptions: true
      });
      
      ui.alert(
        "✅ Permissions Granted", 
        "All required permissions have been granted successfully!\n\n" +
        "You can now:\n" +
        "• Use all Robinhood functions\n" +
        "• Login to your Robinhood account\n" +
        "• Fetch portfolio data\n\n" +
        "Next step: Use 'Robinhood > Login / Re-login' to authenticate.",
        ui.ButtonSet.OK
      );
      
      return true;
      
    } catch (urlError) {
      if (urlError.message.includes("permission") || urlError.message.includes("whitelisted")) {
        ui.alert(
          "⚠️ External Request Permission Needed", 
          "The script needs permission to make external API requests to Robinhood.\n\n" +
          "Please:\n" +
          "1. Check that your appsscript.json includes 'script.external_request' scope\n" +
          "2. Make sure Robinhood URLs are whitelisted\n" +
          "3. Try running a Robinhood function to trigger permission dialog\n\n" +
          "Error: " + urlError.message,
          ui.ButtonSet.OK
        );
        return false;
      } else {
        // Other error, but basic permissions seem to work
        ui.alert(
          "✅ Basic Permissions OK", 
          "Spreadsheet and storage permissions are working.\n\n" +
          "External API permissions will be requested when you first use a Robinhood function.\n\n" +
          "Next step: Try using 'Robinhood > Login / Re-login'",
          ui.ButtonSet.OK
        );
        return true;
      }
    }
    
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "❌ Permission Required", 
      `Permission error: ${e.message}\n\n` +
      `Please authorize the script to:\n` +
      `• Access your spreadsheet\n` +
      `• Make external API requests\n` +
      `• Store authentication data\n\n` +
      `You may need to run this function again after granting permissions.`,
      ui.ButtonSet.OK
    );
    return false;
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let refreshSheet = spreadsheet.getSheetByName(REFRESH.sheet_name);
  if (!refreshSheet) {
    refreshSheet = spreadsheet.insertSheet(REFRESH.sheet_name);
    refreshSheet.getRange(REFRESH.cell_address).setValue(new Date());
    refreshSheet.hideSheet();
  }
  if (!spreadsheet.getRangeByName(REFRESH.named_range_name)) {
    spreadsheet.setNamedRange(REFRESH.named_range_name, refreshSheet.getRange(REFRESH.cell_address));
  }

  ui.createMenu("Robinhood")
    .addItem("🔑 Grant Permissions (Run First)", "requestPermissions")
    .addSeparator()
    .addItem("Login / Re-login", "runLoginProcess")
    .addItem("Refresh Data", "refreshLastUpdate")
    .addSeparator()
    .addItem("📋 Show All Functions", "showFunctionHelp")
    .addToUi();
}

function runLoginProcess() {
  PropertiesService.getUserProperties().deleteProperty("robinhood_access_token");
  showLoginDialog();
}

function showLoginDialog() {
  const html = HtmlService.createHtmlOutputFromFile("LoginDialog").setWidth(300).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, "Robinhood Login");
}

function showFunctionHelp() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let helpSheet = ss.getSheetByName("Function Help");
    if (!helpSheet) {
        helpSheet = ss.insertSheet("Function Help");
    }
    helpSheet.clear();
    const helpData = ROBINHOOD_HELP("all");
    const range = helpSheet.getRange(1, 1, helpData.length, helpData[0].length);
    range.setValues(helpData);
    range.breakApart();
    helpSheet.getRange(1, 1, 1, helpData[0].length).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
    helpSheet.autoResizeColumns(1, helpData[0].length);
    helpSheet.activate();
    SpreadsheetApp.getUi().alert("A 'Function Help' sheet has been created with usage details.");
}

function refreshLastUpdate() {
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName(REFRESH.named_range_name).setValue(new Date());
}

function ROBINHOOD_GET_LOGIN_STATUS(LastUpdate) {
  try {
    validateLastUpdate_(LastUpdate);
    
    const token = PropertiesService.getUserProperties().getProperty("robinhood_access_token");
    if (!token) {
      return "Logged Out";
    }
    
    // Check cache first to avoid unnecessary API calls
    const cachedStatus = CacheManager.get(ROBINHOOD_CONFIG.CACHE.LOGIN_STATUS_KEY);
    if (cachedStatus) {
      return cachedStatus;
    }
    
    // Verify token is still valid
    RobinhoodApiClient.get(ROBINHOOD_CONFIG.API_URIS.accounts);
    CacheManager.put(ROBINHOOD_CONFIG.CACHE.LOGIN_STATUS_KEY, "Logged In", ROBINHOOD_CONFIG.CACHE.EXPIRATION_SECONDS);
    return "Logged In";
  } catch (e) {
    // Clear invalid cached status
    CacheManager.clear(ROBINHOOD_CONFIG.CACHE.LOGIN_STATUS_KEY);
    return "Logged Out";
  }
}

// --- Authentication Flow (Called from Dialog) ---

function generateDeviceToken_() {
  let token = "";
  const chars = "0123456789abcdef";
  for (let i = 0; i < 36; i++) {
    if (i === 8 || i === 13 || i === 18 || i === 23) token += "-";
    else token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

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
          : "Unable to retrieve access token after successful verification.";
      throw new Error(`Failed to retrieve final access token: ${tokenErrorDetail}`);
    }
  } catch (e) {
    Logger.log(`Login error: ${e.message}`);
    return `Login failed: ${e.message}`;
  }
}


function validateSherrifId_(deviceToken, workflowId) {
  Logger.log(`Starting MFA validation for workflow: ${workflowId}`);
  const identiUrl = `https://identi.robinhood.com/idl/v1/workflow/${workflowId}/`;

  try {
    // Trigger the MFA challenge
    const triggerPayload = { clientVersion: "1.0.0", id: workflowId, entryPointAction: {} };
    const triggerOptions = { 
      method: "patch", 
      contentType: "application/json", 
      payload: JSON.stringify(triggerPayload), 
      muteHttpExceptions: true 
    };
    
    Logger.log(`Triggering MFA challenge...`);
    const triggerResponse = RobinhoodApiClient.makeRequest(identiUrl, triggerOptions);
    
    if (!triggerResponse) {
      throw new Error("No response from MFA trigger request");
    }

    const challenge = triggerResponse?.route?.replace?.screen?.deviceApprovalChallengeScreenParams?.sheriffChallenge;
    if (!challenge) {
      Logger.log(`Trigger response: ${JSON.stringify(triggerResponse)}`);
      throw new Error("Failed to trigger MFA challenge - no challenge object found");
    }

    Logger.log(`Challenge type: ${challenge.type}, Challenge ID: ${challenge.id}`);

    if (challenge.type === "PROMPT") {
      const promptStatusUrl = `${ROBINHOOD_CONFIG.API_BASE_URL}/push/${challenge.id}/get_prompts_status/`;
      const startTime = new Date().getTime();
      const timeout = ROBINHOOD_CONFIG.PERFORMANCE.MFA_TIMEOUT_MS;
      let attempts = 0;

      Logger.log(`Waiting for device approval... (timeout: ${timeout/1000}s)`);

      while (new Date().getTime() - startTime < timeout) {
        attempts++;
        
        try {
          const statusRes = RobinhoodApiClient.makeRequest(promptStatusUrl, { 
            method: "get", 
            muteHttpExceptions: true 
          });
          
          Logger.log(`Status check attempt ${attempts}: ${statusRes?.challenge_status || 'no status'}`);
          
          if (statusRes && statusRes.challenge_status === "validated") {
            Logger.log(`MFA approved! Finalizing workflow...`);
            
            // Finalize the workflow
            const finalizePayload = { 
              clientVersion: "1.0.0", 
              id: workflowId, 
              deviceApprovalChallengeAction: { proceed: {} } 
            };
            const finalizeOptions = { 
              method: "patch", 
              contentType: "application/json", 
              payload: JSON.stringify(finalizePayload), 
              muteHttpExceptions: true 
            };
            
            const finalizeResponse = RobinhoodApiClient.makeRequest(identiUrl, finalizeOptions);
            
            if (finalizeResponse?.route?.exit?.status === "WORKFLOW_STATUS_APPROVED") {
              Logger.log(`MFA workflow successfully completed`);
              return;
            } else {
              Logger.log(`Finalize response: ${JSON.stringify(finalizeResponse)}`);
              throw new Error("MFA workflow finalization failed - unexpected response");
            }
          } else if (statusRes && statusRes.challenge_status === "failed") {
            throw new Error("MFA challenge was denied or failed");
          }
          
        } catch (statusError) {
          Logger.log(`Status check error: ${statusError.message}`);
          // Continue trying unless it's a fatal error
          if (statusError.message.includes("denied") || statusError.message.includes("failed")) {
            throw statusError;
          }
        }
        
        Utilities.sleep(ROBINHOOD_CONFIG.PERFORMANCE.MFA_POLL_INTERVAL_MS);
      }
      
      throw new Error(`Login approval timed out after ${attempts} attempts. Please try again and approve more quickly.`);
      
    } else {
      throw new Error(`Unsupported MFA challenge type: ${challenge.type}. Expected 'PROMPT'.`);
    }
    
  } catch (e) {
    Logger.log(`MFA validation error: ${e.message}`);
    throw e;
  }
}
