/**
 * Gift Card Balance Checker with Advanced CAPTCHA Alert & Optimizations
 * Automatically checks card balances and opens browser when CAPTCHA is needed
 * Features: Window reuse, CAPTCHA queueing, progress notifications, error handling
 * Supports: Standard gift cards (card/CVV/expiry) and David Jones Reward cards (number/PIN)
 */

// Global variables for window management and queue
var globalWindow = null;
var processingQueue = [];
var currentIndex = 0;
var totalCards = 0;

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🧧 Card Balance')
    .addItem('Check All Balances', 'checkAllBalances')
    .addItem('Check Selected Row', 'checkSelectedRow')
    .addItem('Resume After CAPTCHA', 'resumeQueue')
    .addToUi();
}

function checkAllBalances() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  // Reset queue
  processingQueue = [];
  currentIndex = 0;
  
  // Build queue (skip header row)
  for (var i = 1; i < data.length; i++) {
    if (data[i][3]) { // If card number exists (Column D)
      processingQueue.push(i + 1); // Store 1-based row numbers
    }
  }
  
  totalCards = processingQueue.length;
  
  if (totalCards === 0) {
    Browser.msgBox('No Cards Found', 'No valid card numbers found in the sheet.', Browser.Buttons.OK);
    return;
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Starting batch check for ' + totalCards + ' cards...', 
    'Balance Checker', 
    5
  );
  
  // Start processing
  processNextInQueue();
}

function checkSelectedRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  
  // Reset queue for single card
  processingQueue = [row];
  currentIndex = 0;
  totalCards = 1;
  
  processNextInQueue();
}

function processNextInQueue() {
  if (currentIndex >= processingQueue.length) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'All ' + totalCards + ' cards processed!', 
      'Complete ✓', 
      5
    );
    if (globalWindow) {
      globalWindow.close();
      globalWindow = null;
    }
    return;
  }
  
  var row = processingQueue[currentIndex];
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Checking card ' + (currentIndex + 1) + ' of ' + totalCards + '...', 
    'Progress', 
    3
  );
  
  checkBalance(row);
}

function resumeQueue() {
  if (processingQueue.length === 0) {
    Browser.msgBox('No Queue', 'No cards in queue. Start a new check first.', Browser.Buttons.OK);
    return;
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Resuming from card ' + (currentIndex + 1) + ' of ' + totalCards + '...', 
    'Resuming', 
    3
  );
  
  processNextInQueue();
}

function checkBalance(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var retailer = sheet.getRange(row, 2).getValue(); // Column B - Retailer
  var cardNumber = sheet.getRange(row, 4).getValue(); // Column D
  var cvv = sheet.getRange(row, 6).getValue(); // Column F
  var expiry = sheet.getRange(row, 8).getValue(); // Column H
  var balanceUrl = sheet.getRange(row, 9).getValue(); // Column I
  
  if (!cardNumber || !cvv || !expiry) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Skipping row ' + row + ' - missing card details', 
      'Warning', 
      3
    );
    currentIndex++;
    Utilities.sleep(500);
    processNextInQueue();
    return;
  }
  
  // Check if this is a David Jones card
  var isDavidJones = false;
  if (retailer && retailer.toString().toLowerCase().includes('david jones')) {
    isDavidJones = true;
  }
  
  // Parse expiry date
  var expiryDate = new Date(expiry);
  var month = String(expiryDate.getMonth() + 1).padStart(2, '0');
  var year = String(expiryDate.getFullYear()).slice(-2);
  
  // Create HTML for CAPTCHA handling
  var html = createBalanceCheckHTML(cardNumber, month, year, cvv, balanceUrl, row, isDavidJones);
  
  // Reuse window or create new one
  if (!globalWindow || globalWindow.closed) {
    var htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(900)
      .setHeight(700);
    globalWindow = SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔍 Checking Balance - Card ' + (currentIndex + 1) + '/' + totalCards);
  } else {
    // Update existing window content
    var htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(900)
      .setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔍 Checking Balance - Card ' + (currentIndex + 1) + '/' + totalCards);
  }
}

function createBalanceCheckHTML(cardNumber, month, year, cvv, url, row, isDavidJones) {
  // Use default URL if not provided
  if (!url) {
    if (isDavidJones) {
      url = 'https://www.davidjones.com/rewards/balance-check';
    } else {
      url = 'https://cardbalance.com.au/';
    }
  }
  
  var cardTypeLabel = isDavidJones ? 'Reward Number' : 'Card Number';
  var cvvLabel = isDavidJones ? 'PIN' : 'CVV';
  
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_blank">
  <style>
    body { 
      font-family: Arial, sans-serif; 
      margin: 0; 
      padding: 20px;
      background: #f5f5f5;
    }
    #status {
      background: #4285f4;
      color: white;
      padding: 15px;
      border-radius: 8px;
      margin-bottom: 15px;
      text-align: center;
      font-size: 16px;
      font-weight: bold;
    }
    #status.warning {
      background: #fbbc04;
      color: #000;
    }
    #status.success {
      background: #34a853;
    }
    #checkFrame {
      width: 100%;
      height: 550px;
      border: 2px solid #ddd;
      border-radius: 8px;
      background: white;
    }
    .info {
      background: white;
      padding: 10px;
      border-radius: 8px;
      margin-top: 10px;
      font-size: 13px;
      color: #666;
    }
    .balance-input {
      margin-top: 15px;
      padding: 15px;
      background: white;
      border-radius: 8px;
      border: 2px solid #4285f4;
    }
    .balance-input input {
      width: 200px;
      padding: 8px;
      font-size: 16px;
      border: 1px solid #ddd;
      border-radius: 4px;
      margin-right: 10px;
    }
    .balance-input button {
      padding: 8px 20px;
      font-size: 16px;
      background: #4285f4;
      color: white;
      border: none;
      border-radius: 4px;
      cursor: pointer;
    }
    .balance-input button:hover {
      background: #357ae8;
    }
    .card-details {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 8px;
      margin-bottom: 15px;
      font-size: 14px;
    }
    .card-details strong {
      color: #1a73e8;
    }
  </style>
</head>
<body>
  <div id="status">⏳ Form loaded. Fill in details and complete CAPTCHA if needed.</div>
  
  <div class="card-details">
    <strong>${cardTypeLabel}:</strong> ${cardNumber}<br>
    <strong>${cvvLabel}:</strong> ${cvv}<br>
    <strong>Expiry:</strong> ${month}/${year}
  </div>
  
  <iframe id="checkFrame" src="about:blank"></iframe>
  
  <div class="balance-input">
    <strong>After CAPTCHA:</strong> Enter balance here: 
    <input type="text" id="balanceInput" placeholder="e.g., 45.50" />
    <button onclick="submitBalance()">✓ Save & Continue</button>
  </div>
  
  <div class="info">
    💡 <strong>Tip:</strong> If the form is hard to see, click "Open in Full Browser"<br>
    🔄 <strong>After CAPTCHA:</strong> Enter the balance above and click "Save & Continue" to process the next card
  </div>
  
  <script>
    const cardNumber = '${cardNumber}';
    const month = '${month}';
    const year = '${year}';
    const cvv = '${cvv}';
    const url = '${url}';
    const row = ${row};
    const isDavidJones = ${isDavidJones};
    
    function openFullWindow() {
      window.open(url, '_blank');
      document.getElementById('status').innerHTML = 
        '🌐 Opened in new window. Please complete the balance check there, then click Refresh.';
    }
    
    function refreshCheck() {
      document.getElementById('checkFrame').src = url;
    }
    
    function submitBalance() {
      const balance = document.getElementById('balanceInput').value.trim();
      
      if (!balance) {
        alert('Please enter a balance amount');
        return;
      }
      
      document.getElementById('status').className = 'success';
      document.getElementById('status').innerHTML = '✓ Saving balance and moving to next card...';
      
      google.script.run
        .withSuccessHandler(function() {
          document.getElementById('status').innerHTML = '✓ Saved! Moving to next card...';
        })
        .withFailureHandler(function(error) {
          alert('Error saving balance: ' + error);
        })
        .updateBalanceAndContinue(row, balance);
    }
    
    // Auto-load on start
    setTimeout(() => {
      document.getElementById('checkFrame').src = url;
      document.getElementById('status').innerHTML = 
        '✅ Form loaded. Fill in details and complete CAPTCHA if needed.';
      document.getElementById('status').className = 'warning';
    }, 1000);
    
    // Instructions
    setTimeout(() => {
      document.getElementById('status').innerHTML += 
        '<br><br>⚠️ Tip: If the form is hard to see, click "Open in Full Browser"';
    }, 3000);
  </script>
</body>
</html>
  `;
}

// Function to manually update balance after CAPTCHA
function updateBalanceAndContinue(row, balance) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Add a new column for "Current Balance" if it doesn't exist
  var lastCol = sheet.getLastColumn();
  if (sheet.getRange(1, lastCol + 1).getValue() !== 'Current Balance') {
    sheet.getRange(1, lastCol + 1).setValue('Current Balance');
  }
  
  // Update the balance
  sheet.getRange(row, lastCol + 1).setValue(balance);
  sheet.getRange(row, lastCol + 2).setValue(new Date()); // Last checked
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Balance $' + balance + ' saved for row ' + row, 
    'Saved ✓', 
    2
  );
  
  // Move to next card in queue
  currentIndex++;
  
  // Add delay before next check
  Utilities.sleep(1500);
  
  processNextInQueue();
}

// Legacy function for manual updates (kept for compatibility)
function updateBalance(row, balance) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Add a new column for "Current Balance" if it doesn't exist
  var lastCol = sheet.getLastColumn();
  if (sheet.getRange(1, lastCol + 1).getValue() !== 'Current Balance') {
    sheet.getRange(1, lastCol + 1).setValue('Current Balance');
  }
  
  // Update the balance
  sheet.getRange(row, lastCol + 1).setValue(balance);
  sheet.getRange(row, lastCol + 2).setValue(new Date()); // Last checked
  
  Browser.msgBox('Balance Updated', 'Balance for row ' + row + ' set to: $' + balance, Browser.Buttons.OK);
}
