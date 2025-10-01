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
    'ID', 'Name', 'Description', 'Short Description', 'Product Type', 
    'Price', 'Regular Price', 'Sale Price', 'Stock Quantity', 
    'Stock Status', 'SKU', 'Categories', 'Tags', 'Featured Image', 'Last Updated'
  ],
  keyColumn: 0, // ID column
  lastUpdatedColumn: 14 // Last Updated column
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
    var productId = row[SHEET_CONFIG.keyColumn];
    if (productId) {
      existingData[productId] = {
        rowIndex: i + 1,
        lastUpdated: row[SHEET_CONFIG.lastUpdatedColumn] || ''
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
      product.id,
      product.name,
      product.description,
      product.short_description,
      product.product_type,
      product.price,
      product.regular_price,
      product.sale_price,
      product.stock_quantity,
      product.stock_status,
      product.sku,
      Array.isArray(product.categories) ? product.categories.join(', ') : '',
      Array.isArray(product.tags) ? product.tags.join(', ') : '',
      product.featured_image,
      new Date().toISOString()
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

    var setSizeColumns = [3, 4]; // Name, Description, Short Description, Categories, Tags, Featured Image
    for (var i = 0; i < setSizeColumns.length; i++) {
      sheet.setColumnWidth(setSizeColumns[i], 150);
    }
  }
  
  return {
    updated: updatedCount,
    new: newCount,
    total: newData.length
  };
}

/**
 * Updates selected products back to WordPress
 */
function updateSelectedProducts() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var selectedRows = getSelectedProductRows(range, sheet);
  
  if (selectedRows.length === 0) {
    throw new Error('Please select at least one product row to update.');
  }
  
  return updateProductsToWordPress(selectedRows);
}

/**
 * Updates all products back to WordPress
 */
function updateAllProducts() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var allRows = getAllProductRows(sheet);
  
  if (allRows.length === 0) {
    throw new Error('No products found in the sheet. Please fetch products first.');
  }
  
  return updateProductsToWordPress(allRows);
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
 * Updates products to WordPress via API
 */
function updateProductsToWordPress(productRows) {
  try {
    var config = getConfig();
    if (!config.isConfigured) {
      throw new Error('Please configure the WordPress API URL and secret key first in Settings.');
    }
    
    // Validate secret key is present
    validateSecretKey();
    
    var productsData = productRows.map(function(row) {
      return {
        id: row.data[0],
        name: row.data[1],
        description: row.data[2],
        short_description: row.data[3],
        price: row.data[5],
        regular_price: row.data[6],
        sale_price: row.data[7],
        stock_quantity: row.data[8],
        stock_status: row.data[9],
        sku: row.data[10]
      };
    });
    
    // Generate encrypted sheet token for secure authentication
    var sheetToken = generateSheetToken();
    
    var payload = {
      products: productsData
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
    }
    
    var result = JSON.parse(response.getContentText());
    
    // Update last updated timestamp for successfully updated rows
    if (result.data && result.data.length > 0) {
      var sheet = SpreadsheetApp.getActiveSheet();
      var now = new Date().toISOString();
      
      productRows.forEach(function(row) {
        if (result.data.includes(row.data[0])) {
          sheet.getRange(row.rowIndex, SHEET_CONFIG.lastUpdatedColumn + 1).setValue(now);
        }
      });
    }
    
    return {
      success: true,
      message: result.message || 'Products updated successfully',
      updatedCount: result.data ? result.data.length : 0
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