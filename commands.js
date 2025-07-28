(function () {
  "use strict";

  // Prevent multiple initialization
  if (window.authopsyCommandsInitialized) {
    return;
  }
  window.authopsyCommandsInitialized = true;

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    console.log("Commands initialized with reason: " + reason);
  };

  // Helper function to parse email headers
  function parseEmailHeaders(callback) {
    try {
      Office.context.mailbox.item.getAllInternetHeadersAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const headers = result.value;
          
          // Parse DMARC results
          const dmarcResult = parseDmarcResult(headers);
          
          // Parse DKIM results
          const dkimResult = parseDkimResult(headers);
          
          // Parse SPF results
          const spfResult = parseSpfResult(headers);
          
          const results = {
            dmarc: dmarcResult,
            dkim: dkimResult,
            spf: spfResult
          };
          
          callback(results);
        } else {
          console.error("Failed to get headers: " + result.error.message);
          callback({
            dmarc: { status: "unknown", details: "Error retrieving headers: " + result.error.message },
            dkim: { status: "unknown", details: "Error retrieving headers: " + result.error.message },
            spf: { status: "unknown", details: "Error retrieving headers: " + result.error.message }
          });
        }
      });
    } catch (error) {
      console.error("Error in parseEmailHeaders: " + error.message);
      callback({
        dmarc: { status: "unknown", details: "Error parsing headers: " + error.message },
        dkim: { status: "unknown", details: "Error parsing headers: " + error.message },
        spf: { status: "unknown", details: "Error parsing headers: " + error.message }
      });
    }
  }
  
  // Parse DMARC result from headers
  function parseDmarcResult(headers) {
    try {
      // Look for Authentication-Results header with dmarc
      const dmarcRegex = /Authentication-Results:.*?dmarc=([^;\s]+)/i;
      const dmarcMatch = headers.match(dmarcRegex);
      
      if (dmarcMatch && dmarcMatch[1]) {
        const result = dmarcMatch[1].toLowerCase().trim();
        if (result.includes("pass")) {
          return { status: "pass", details: "DMARC authentication passed: " + result };
        } else if (result.includes("fail") || result.includes("none")) {
          return { status: "fail", details: "DMARC authentication failed: " + result };
        } else {
          return { status: "unknown", details: "DMARC result unclear: " + result };
        }
      }
      return { status: "unknown", details: "No DMARC results found in headers" };
    } catch (error) {
      console.error("Error parsing DMARC:", error);
      return { status: "unknown", details: "Error parsing DMARC results: " + error.message };
    }
  }
  
  // Parse DKIM result from headers
  function parseDkimResult(headers) {
    try {
      // Look for Authentication-Results header with dkim
      const dkimRegex = /Authentication-Results:.*?dkim=([^;\s]+)/i;
      const dkimMatch = headers.match(dkimRegex);
      
      if (dkimMatch && dkimMatch[1]) {
        const result = dkimMatch[1].toLowerCase().trim();
        if (result.includes("pass")) {
          return { status: "pass", details: "DKIM signature verified: " + result };
        } else if (result.includes("fail") || result.includes("none")) {
          return { status: "fail", details: "DKIM signature verification failed: " + result };
        } else {
          return { status: "unknown", details: "DKIM result unclear: " + result };
        }
      }
      return { status: "unknown", details: "No DKIM results found in headers" };
    } catch (error) {
      console.error("Error parsing DKIM:", error);
      return { status: "unknown", details: "Error parsing DKIM results: " + error.message };
    }
  }
  
  // Parse SPF result from headers
  function parseSpfResult(headers) {
    try {
      // Look for Authentication-Results header with spf
      const spfRegex = /Authentication-Results:.*?spf=([^;\s]+)/i;
      const spfMatch = headers.match(spfRegex);
      
      if (spfMatch && spfMatch[1]) {
        const result = spfMatch[1].toLowerCase().trim();
        if (result.includes("pass")) {
          return { status: "pass", details: "SPF check passed: " + result };
        } else if (result.includes("fail") || result.includes("none")) {
          return { status: "fail", details: "SPF check failed: " + result };
        } else {
          return { status: "unknown", details: "SPF result unclear: " + result };
        }
      }
      return { status: "unknown", details: "No SPF results found in headers" };
    } catch (error) {
      console.error("Error parsing SPF:", error);
      return { status: "unknown", details: "Error parsing SPF results: " + error.message };
    }
  }
  
  // Show results - DEBUG VERSION
  function showResult(title, content) {
    console.log("=== " + title + " ===");
    console.log(content);
    console.log("===============");
    
    // DEBUG: Check what APIs are available
    console.log("Office.context available:", !!Office.context);
    console.log("Office.context.mailbox available:", !!Office.context.mailbox);
    console.log("Office.context.mailbox.item available:", !!Office.context.mailbox.item);
    console.log("notificationMessages available:", !!Office.context.mailbox.item.notificationMessages);
    
    // Try multiple notification approaches
    try {
      if (Office.context.mailbox.item.notificationMessages) {
        const simpleKey = "authopsy_debug_" + Date.now();
        console.log("Attempting notification with key:", simpleKey);
        
        Office.context.mailbox.item.notificationMessages.addAsync(simpleKey, {
          type: "informationalMessage",
          message: title + ": " + content.substring(0, 80) + "...",
          persistent: true  // Make it stick around
        }, function(result) {
          console.log("Notification result:", result);
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Notification failed:", result.error);
            console.error("Error name:", result.error.name);
            console.error("Error message:", result.error.message);
            console.error("Error code:", result.error.code);
          } else {
            console.log("✅ Notification added successfully!");
            
            // Auto-remove after 8 seconds for testing
            setTimeout(function() {
              Office.context.mailbox.item.notificationMessages.removeAsync(simpleKey, function(removeResult) {
                console.log("Notification removal result:", removeResult);
              });
            }, 8000);
          }
        });
      } else {
        console.error("❌ notificationMessages API not available");
        console.log("Available APIs:", Object.keys(Office.context.mailbox.item));
      }
    } catch (error) {
      console.error("❌ Exception in showResult:", error);
      console.error("Error name:", error.name);
      console.error("Error message:", error.message);
      console.error("Error stack:", error.stack);
    }
  }
  
  // Handler for DMARC details - SIMPLIFIED
  function showDmarcDetails(event) {
    console.log("DMARC details requested");
    
    parseEmailHeaders(function(results) {
      showResult("DMARC Analysis", results.dmarc.details);
    });
    
    // ALWAYS complete the event immediately
    if (event && event.completed) {
      event.completed();
    }
  }
  
  // Handler for DKIM details - SIMPLIFIED
  function showDkimDetails(event) {
    console.log("DKIM details requested");
    
    parseEmailHeaders(function(results) {
      showResult("DKIM Analysis", results.dkim.details);
    });
    
    // ALWAYS complete the event immediately
    if (event && event.completed) {
      event.completed();
    }
  }
  
  // Handler for SPF details - SIMPLIFIED
  function showSpfDetails(event) {
    console.log("SPF details requested");
    
    parseEmailHeaders(function(results) {
      showResult("SPF Analysis", results.spf.details);
    });
    
    // ALWAYS complete the event immediately
    if (event && event.completed) {
      event.completed();
    }
  }

  // Register functions with Office.actions - only once
  function registerActions() {
    try {
      if (Office.actions && !window.authopsyActionsRegistered) {
        Office.actions.associate("showDmarcDetails", showDmarcDetails);
        Office.actions.associate("showDkimDetails", showDkimDetails);
        Office.actions.associate("showSpfDetails", showSpfDetails);
        window.authopsyActionsRegistered = true;
        console.log("Actions registered successfully");
      } else if (window.authopsyActionsRegistered) {
        console.log("Actions already registered");
      } else {
        console.warn("Office.actions not available");
      }
    } catch (error) {
      console.error("Failed to register actions: " + error.message);
    }
  }

  // Register actions when Office is ready
  if (Office.context && Office.actions) {
    registerActions();
  } else {
    Office.onReady(function() {
      registerActions();
    });
  }

  // Also expose globally as fallback
  if (!window.showDmarcDetails) {
    window.showDmarcDetails = showDmarcDetails;
    window.showDkimDetails = showDkimDetails;
    window.showSpfDetails = showSpfDetails;
  }
})();