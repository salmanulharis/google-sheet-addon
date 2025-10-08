/**
 * Sheets API Add-on for WordPress WooCommerce Products
 * @OnlyCurrentDoc
 */

// WordPress API configuration
var WP_API_CONFIG = {
  baseUrl: '', // Will be set by user
  namespace: 'sheets-api/v1',
  endpoints: {
    getProducts: '/get_products',
    updateProducts: '/update_products',
    testConnection: '/test_connection'
  }
};

// Encryption configuration
var ENCRYPTION_CONFIG = {
  secretKey: '', // Will be set by user
  algorithm: 'XOR' // Simple but effective
};

// Sheet configuration
var SHEET_CONFIG = {
  sheetName: 'Products',
  headers: [
    'Product ID', 'Type', 'Parent ID', 'Name', 'SKU', 'Attributes', 
    'Regular Price', 'Sale Price', 'Stock', 'Status'
  ],
  keyColumn: 0, // Product ID column
  lastUpdatedColumn: 9 // Status column (no longer using Last Updated)
};

/**
 * onInstall trigger
 */
function onInstall() {
  onOpen();
}

/**
 * onOpen trigger - creates menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('WordPress Products')
    .addItem('Open Sidebar', 'showSidebar')
    .addItem('Settings', 'showSettings')
    .addToUi();
}

/**
 * Shows the sidebar
 */
function showSidebar() {
  var html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('WordPress Products')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows settings dialog
 */
function showSettings() {
  var html = HtmlService.createTemplateFromFile('Settings')
    .evaluate()
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

/**
 * Saves API configuration
 */
function saveConfig(baseUrl, secretKey) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('WP_BASE_URL', baseUrl);
  if (secretKey) {
    userProperties.setProperty('WP_SECRET_KEY', secretKey);
    ENCRYPTION_CONFIG.secretKey = secretKey;
  }
  WP_API_CONFIG.baseUrl = baseUrl;
  return true;
}

/**
 * Gets API configuration
 */
function getConfig() {
  var userProperties = PropertiesService.getUserProperties();
  var baseUrl = userProperties.getProperty('WP_BASE_URL');
  var secretKey = userProperties.getProperty('WP_SECRET_KEY');
  
  // Update the global config
  if (secretKey) {
    ENCRYPTION_CONFIG.secretKey = secretKey;
  }
  
  return {
    baseUrl: baseUrl || '',
    secretKey: secretKey || '',
    isConfigured: !!baseUrl && !!secretKey
  };
}

/**
 * Validates that the secret key is configured
 */
function validateSecretKey() {
  var config = getConfig();
  if (!config.secretKey) {
    throw new Error('Secret key is not configured. Please set up your secret key in Settings before using this feature.');
  }
  return true;
}

/**
 * Fetches products from WordPress API
 */
function fetchProducts() {
  try {
    var config = getConfig();
    if (!config.isConfigured) {
      throw new Error('Please configure the WordPress API URL and secret key first in Settings.');
    }
    
    // Validate secret key is present
    validateSecretKey();
    
    // Validate secret key is present
    validateSecretKey();
    
    // Use the currently active sheet instead of creating a specific "Products" sheet
    var sheet = SpreadsheetApp.getActiveSheet();
    
    // Set up headers if the sheet is empty or doesn't have proper headers
    setupSheetHeaders(sheet);
    
    var existingData = getExistingProductData(sheet);
    
    // Generate encrypted sheet token for secure authentication
    var sheetToken = generateSheetToken();
    
    var fetchUrl = config.baseUrl + '/wp-json/' + WP_API_CONFIG.namespace + WP_API_CONFIG.endpoints.getProducts;
    
    var response = UrlFetchApp.fetch(
      fetchUrl,
      {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          "ngrok-skip-browser-warning": "true",
          "User-Agent": "GoogleAppsScript",
          "X-Sheet-Token": sheetToken,
          "Authorization": "Bearer " + sheetToken
        },
        muteHttpExceptions: true
      }
    );
    
    if (response.getResponseCode() !== 200) {
      throw new Error('API Error: ' + response.getContentText() + ' (URL: ' + fetchUrl + ')');
    }
    
    var result = JSON.parse(response.getContentText());
    var products = result.data;
    
    if (!products || !Array.isArray(products)) {
      throw new Error('Invalid response format from API');
    }
    
    // Reverse the order to show products in reverse order (newest first or reverse of API response)
    products = products.reverse();
    
    updateSheetWithProducts(sheet, products, existingData);
    
    // Store current product IDs for deletion tracking
    storeCurrentProductIds();

    return {
      success: true,
      message: 'Successfully fetched ' + products.length + ' products',
      count: products.length,
      url: fetchUrl
    };
    
  } catch (error) {
    Logger.log('Error fetching products: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Sets up headers in the current sheet if needed
 */
function setupSheetHeaders(sheet) {
  // Check if sheet is empty or first row doesn't match our headers
  var firstRow = sheet.getRange(1, 1, 1, SHEET_CONFIG.headers.length).getValues()[0];
  var hasCorrectHeaders = true;
  
  // Check if headers match
  for (var i = 0; i < SHEET_CONFIG.headers.length; i++) {
    if (firstRow[i] !== SHEET_CONFIG.headers[i]) {
      hasCorrectHeaders = false;
      break;
    }
  }
  
  // Set headers if they don't exist or don't match
  if (!hasCorrectHeaders) {
    sheet.getRange(1, 1, 1, SHEET_CONFIG.headers.length).setValues([SHEET_CONFIG.headers]);
    // Format header row
    sheet.getRange(1, 1, 1, SHEET_CONFIG.headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

/**
 * Gets or creates the products sheet (legacy function - kept for compatibility)
 */
function getOrCreateProductsSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SHEET_CONFIG.sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_CONFIG.sheetName);
    // Set headers
    sheet.getRange(1, 1, 1, SHEET_CONFIG.headers.length).setValues([SHEET_CONFIG.headers]);
    // Format header row
    sheet.getRange(1, 1, 1, SHEET_CONFIG.headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * Gets existing product data from sheet
 */
function getExistingProductData(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};
  
  var existingData = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var productId = row[SHEET_CONFIG.keyColumn]; // Product ID column
    if (productId) {
      existingData[productId] = {
        rowIndex: i + 1,
        status: row[9] || '' // Status column
      };
    }
  }
  
  return existingData;
}

/**
 * Updates sheet with products data
 */
function updateSheetWithProducts(sheet, products, existingData) {
  var newData = [];
  var updatedCount = 0;
  var newCount = 0;
  
  // Clear existing data but keep headers
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, SHEET_CONFIG.headers.length).clearContent();
  }
  
  products.forEach(function(product) {
    var rowData = [
      product.id,                                                    // Product ID
      product.type || product.product_type || 'simple',            // Type
      product.parent_id || '',                                      // Parent ID
      product.name,                                                 // Name
      product.sku,                                                  // SKU
      product.attributes,                                          // Attributes
      product.regular_price,                                        // Regular Price
      product.sale_price,                                           // Sale Price
      product.stock_quantity || product.stock,                     // Stock
      product.status || product.stock_status || 'instock'          // Status
    ];
    
    newData.push(rowData);
    
    if (existingData[product.id]) {
      updatedCount++;
    } else {
      newCount++;
    }
  });
  
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, SHEET_CONFIG.headers.length).setValues(newData);
    
    // Auto-resize columns for better readability
    for (var i = 1; i <= SHEET_CONFIG.headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    // Set specific widths for certain columns
    var setSizeColumns = [4, 6]; // Name and Attributes columns
    for (var i = 0; i < setSizeColumns.length; i++) {
      sheet.setColumnWidth(setSizeColumns[i], 200);
    }
  }
  
  return {
    updated: updatedCount,
    new: newCount,
    total: newData.length
  };
}

/**
 * Updates selected products back to WordPress (with deletion detection)
 */
function updateSelectedProducts() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var selectedRows = getSelectedProductRows(range, sheet);
  
  if (selectedRows.length === 0) {
    throw new Error('Please select at least one product row to update.');
  }
  
  // Get deleted product IDs
  var deletedIds = getDeletedProductIds();
  
  return updateProductsToWordPress(selectedRows, deletedIds);
}

/**
 * Updates all products back to WordPress (with deletion detection)
 */
function updateAllProducts() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var allRows = getAllProductRows(sheet);
  
  if (allRows.length === 0) {
    throw new Error('No products found in the sheet. Please fetch products first.');
  }
  
  // Get deleted product IDs
  var deletedIds = getDeletedProductIds();
  
  return updateProductsToWordPress(allRows, deletedIds);
}

/**
 * Gets selected product rows
 */
function getSelectedProductRows(range, sheet) {
  var selectedRows = [];
  var data = sheet.getDataRange().getValues();
  
  for (var i = range.getRow(); i <= range.getLastRow(); i++) {
    if (i > 1 && i <= data.length) { // Skip header row
      var rowData = data[i-1];
      selectedRows.push({
        rowIndex: i,
        data: rowData
      });
    }
  }
  
  return selectedRows;
}

/**
 * Gets all product rows
 */
function getAllProductRows(sheet) {
  var allRows = [];
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var rowData = data[i];
    allRows.push({
      rowIndex: i + 1,
      data: rowData
    });
  }
  
  return allRows;
}

/**
 * Updates products to WordPress via API (including deletions)
 */
function updateProductsToWordPress(productRows, deletedIds) {
  try {
    var config = getConfig();
    if (!config.isConfigured) {
      throw new Error('Please configure the WordPress API URL and secret key first in Settings.');
    }
    
    // Validate secret key is present
    validateSecretKey();
    
    var productsData = productRows.map(function(row) {
      return {
        id: row.data[0],                    // Product ID
        type: row.data[1],                  // Type
        parent_id: row.data[2],             // Parent ID
        name: row.data[3],                  // Name
        sku: row.data[4],                   // SKU
        attributes: row.data[5],            // Attributes
        regular_price: row.data[6],         // Regular Price
        sale_price: row.data[7],            // Sale Price
        stock_quantity: row.data[8],        // Stock
        status: row.data[9]                 // Status
      };
    });
    
    // Generate encrypted sheet token for secure authentication
    var sheetToken = generateSheetToken();
    
    var payload = {
      products: productsData,
      deleted_ids: deletedIds || []  // Include deleted product IDs
    };
    
    var response = UrlFetchApp.fetch(
      config.baseUrl + '/wp-json/' + WP_API_CONFIG.namespace + WP_API_CONFIG.endpoints.updateProducts,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          "ngrok-skip-browser-warning": "true",
          "User-Agent": "GoogleAppsScript",
          "X-Sheet-Token": sheetToken,
          "Authorization": "Bearer " + sheetToken
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );
    
    if (response.getResponseCode() !== 200) {
      var errorResponse = JSON.parse(response.getContentText());
      throw new Error('API Error: ' + (errorResponse.message || response.getContentText()));
    } else {
      // Clear stored IDs after successful update
      PropertiesService.getDocumentProperties().deleteProperty('STORED_PRODUCT_IDS');
      // trigger fetch products to refresh the sheet data
      fetchProducts();
    }
    
    var result = JSON.parse(response.getContentText());
    
    return {
      success: true,
      message: result.message || 'Products updated successfully',
      updatedCount: result.data.updated ? result.data.updated.length : 0,
      deletedCount: result.data.deleted ? result.data.deleted.length : 0
    };
    
  } catch (error) {
    Logger.log('Error updating products: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * XOR encryption for sheet ID authentication
 */
function xorEncrypt(data, secretKey) {
  if (!secretKey || secretKey.length === 0) {
    throw new Error('Secret key is required for encryption');
  }
  
  var encrypted = [];
  var keyLength = secretKey.length;
  
  for (var i = 0; i < data.length; i++) {
    var keyChar = secretKey.charCodeAt(i % keyLength);
    var dataChar = data.charCodeAt(i);
    encrypted.push(String.fromCharCode(dataChar ^ keyChar));
  }
  
  // Encode to Base64 to make it safe for URL transmission
  return Utilities.base64Encode(encrypted.join(''));
}

function testConnection() {
  var config = getConfig();
  if (!config.isConfigured) {
    throw new Error('Please configure the WordPress API URL and secret key first in Settings.');
  }
  
  // Validate secret key is present
  validateSecretKey();
  
  // Generate encrypted sheet token for secure authentication
  var sheetToken = generateSheetToken();

  var response = UrlFetchApp.fetch(
    config.baseUrl + '/wp-json/' + WP_API_CONFIG.namespace + WP_API_CONFIG.endpoints.testConnection,
    {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        "ngrok-skip-browser-warning": "true",
        "User-Agent": "GoogleAppsScript",
        "X-Sheet-Token": sheetToken,
        "Authorization": "Bearer " + sheetToken
      },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200) {
    var errorResponse = JSON.parse(response.getContentText());
    throw new Error('API Error: ' + (errorResponse.message || response.getContentText()));
  }

  return { success: true, message: 'Connection successful' };
}

/**
 * XOR decryption for sheet ID authentication
 */
function xorDecrypt(encryptedData, secretKey) {
  try {
    // Decode from Base64 first
    var decoded = Utilities.base64Decode(encryptedData);
    var decodedString = Utilities.newBlob(decoded).getDataAsString();
    
    var decrypted = [];
    var keyLength = secretKey.length;
    
    for (var i = 0; i < decodedString.length; i++) {
      var keyChar = secretKey.charCodeAt(i % keyLength);
      var encryptedChar = decodedString.charCodeAt(i);
      decrypted.push(String.fromCharCode(encryptedChar ^ keyChar));
    }
    
    return decrypted.join('');
  } catch (error) {
    throw new Error('Decryption failed: Invalid token');
  }
}

/**
 * Generate encrypted sheet token for API authentication
 */
function generateSheetToken() {
  // Validate secret key is present
  if (!ENCRYPTION_CONFIG.secretKey) {
    throw new Error('Secret key is not configured. Cannot generate authentication token.');
  }
  
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var timestamp = new Date().getTime();
  var tokenData = sheetId + '|' + timestamp;
  
  return xorEncrypt(tokenData, ENCRYPTION_CONFIG.secretKey);
}

/**
 * Validate and extract sheet ID from encrypted token
 */
function validateSheetToken(encryptedToken) {
  try {
    // Validate secret key is present
    if (!ENCRYPTION_CONFIG.secretKey) {
      throw new Error('Secret key is not configured. Cannot validate token.');
    }
    
    var decryptedData = xorDecrypt(encryptedToken, ENCRYPTION_CONFIG.secretKey);
    var parts = decryptedData.split('|');
    
    if (parts.length !== 2) {
      throw new Error('Invalid token format');
    }
    
    var sheetId = parts[0];
    var timestamp = parseInt(parts[1]);
    var currentTime = new Date().getTime();
    
    // Token expires after 24 hours (86400000 ms)
    if (currentTime - timestamp > 86400000) {
      throw new Error('Token expired');
    }
    
    return {
      isValid: true,
      sheetId: sheetId,
      timestamp: timestamp
    };
  } catch (error) {
    return {
      isValid: false,
      error: error.message
    };
  }
}

/**
 * Includes HTML file content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Stores current product IDs for deletion tracking
 */
function storeCurrentProductIds() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var currentIds = [];
  
  // Skip header row and collect all product IDs
  for (var i = 1; i < data.length; i++) {
    var productId = data[i][SHEET_CONFIG.keyColumn]; // Product ID column (0)
    if (productId) {
      currentIds.push(productId.toString());
    }
  }
  
  // Store in PropertiesService for temporary storage
  PropertiesService.getDocumentProperties().setProperty('STORED_PRODUCT_IDS', JSON.stringify(currentIds));
  
  return currentIds;
}

/**
 * Detects deleted product IDs by comparing stored IDs with current sheet data
 */
function getDeletedProductIds() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var currentIds = [];
  
  // Get current product IDs from sheet
  for (var i = 1; i < data.length; i++) {
    var productId = data[i][SHEET_CONFIG.keyColumn]; // Product ID column (0)
    if (productId) {
      currentIds.push(productId.toString());
    }
  }
  
  // Get stored product IDs
  var storedIdsJson = PropertiesService.getDocumentProperties().getProperty('STORED_PRODUCT_IDS');
  if (!storedIdsJson) {
    return []; // No stored IDs, no deletions to detect
  }
  
  var storedIds = JSON.parse(storedIdsJson);
  
  // Find deleted IDs (IDs that were stored but are no longer in current data)
  var deletedIds = storedIds.filter(function(id) {
    return currentIds.indexOf(id) === -1;
  });
  
  return deletedIds;
}