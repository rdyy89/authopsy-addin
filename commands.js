/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.onReady() can be called here to perform
  // initialization that is required after the Office.js library is loaded.
  console.log('Authopsy commands loaded');
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Authopsy add-in command executed successfully.",
    icon: "Icon.80x80",
    persistent: true
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Quick analysis function for dropdown menu
 * @param event {Office.AddinCommands.Event}
 */
function quickAnalysis(event) {
  try {
    // Get the current item (email)
    const item = Office.context.mailbox.item;
    
    if (!item) {
      showNotification("Error: No email selected", "error", event);
      return;
    }

    // Show loading notification
    showNotification("Analyzing email headers...", "informational", event, false);

    // Get internet headers for quick analysis
    item.internetHeaders.getAsync(
      [
        'Authentication-Results',
        'ARC-Authentication-Results', 
        'Received-SPF',
        'DKIM-Signature',
        'X-MS-Exchange-Authentication-Results'
      ],
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const headers = asyncResult.value;
          const results = quickAnalyzeHeaders(headers);
          showNotification(results, "informational", event);
        } else {
          showNotification("Quick Analysis: Headers not accessible. Use Full Analysis for complete results.", "informational", event);
        }
      }
    );

  } catch (error) {
    console.error('Error in quickAnalysis function:', error);
    showNotification("Error: Unable to analyze email headers", "error", event);
  }
}

function quickAnalyzeHeaders(headers) {
  let results = "📧 Header Analysis:\n";
  let authPassed = 0;
  let totalChecks = 0;

  // Parse Authentication-Results header
  const authResultsHeader = headers['Authentication-Results'] || headers['ARC-Authentication-Results'] || headers['X-MS-Exchange-Authentication-Results'];
  
  if (authResultsHeader) {
    const authText = authResultsHeader.toLowerCase();
    
    // Check DMARC
    totalChecks++;
    if (authText.includes('dmarc=pass')) {
      results += "✅ DMARC: Pass\n";
      authPassed++;
    } else if (authText.includes('dmarc=fail')) {
      results += "❌ DMARC: Fail\n";
    } else if (authText.includes('dmarc=')) {
      results += "⚠️ DMARC: Unknown\n";
    } else {
      results += "⚠️ DMARC: Not found\n";
    }

    // Check DKIM
    totalChecks++;
    if (authText.includes('dkim=pass')) {
      results += "✅ DKIM: Pass\n";
      authPassed++;
    } else if (authText.includes('dkim=fail')) {
      results += "❌ DKIM: Fail\n";
    } else if (authText.includes('dkim=')) {
      results += "⚠️ DKIM: Unknown\n";
    } else if (headers['DKIM-Signature']) {
      results += "⚠️ DKIM: Signature present\n";
    } else {
      results += "⚠️ DKIM: Not found\n";
    }

    // Check SPF
    totalChecks++;
    if (authText.includes('spf=pass')) {
      results += "✅ SPF: Pass\n";
      authPassed++;
    } else if (authText.includes('spf=fail')) {
      results += "❌ SPF: Fail\n";
    } else if (authText.includes('spf=')) {
      results += "⚠️ SPF: Unknown\n";
    } else {
      results += "⚠️ SPF: Not found\n";
    }
  } else {
    results += "⚠️ No authentication headers found\n";
  }

  // Add summary
  if (totalChecks > 0) {
    results += `\n📊 Score: ${authPassed}/${totalChecks} passed`;
  }
  
  results += "\n\n💡 Use 'Full Analysis' for detailed explanations";

  return results;
}

function showNotification(message, type, event, complete = true) {
  const notificationType = type === "error" ? 
    Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage :
    Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;

  const notification = {
    type: notificationType,
    message: message,
    icon: "Icon.80x80",
    persistent: true
  };

  // Show the notification
  Office.context.mailbox.item.notificationMessages.replaceAsync("quickAnalysis", notification);

  // Complete the event if requested
  if (complete && event) {
    event.completed();
  }
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
g.quickAnalysis = quickAnalysis;
