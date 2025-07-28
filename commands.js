(function () {
  "use strict";

  // Prevent multiple initialization
  if (window.authopsyCommandsInitialized) {
    return;
  }
  window.authopsyCommandsInitialized = true;

  // Enhanced debugging
  console.log("🚀 AUTHOPSY COMMANDS LOADING 🚀");
  console.log("Timestamp:", new Date().toISOString());
  console.log("User Agent:", navigator.userAgent);

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    console.log("🎯 Commands initialized with reason:", reason);
    console.log("Office context:", Office.context);
    console.log("Office version:", Office.context ? Office.context.requirements : "N/A");
  };

  // Helper function to parse email headers
  function parseEmailHeaders(callback) {
    console.log("📧 === PARSING EMAIL HEADERS ===");
    console.log("🔍 Office.context.mailbox.item:", Office.context.mailbox.item);
    console.log("🔍 getAllInternetHeadersAsync available:", !!Office.context.mailbox.item.getAllInternetHeadersAsync);
    
    try {
      console.log("🚀 Calling getAllInternetHeadersAsync...");
      
      Office.context.mailbox.item.getAllInternetHeadersAsync(function (result) {
        console.log("📬 getAllInternetHeadersAsync callback triggered");
        console.log("📊 Result status:", result.status);
        console.log("📊 Result object:", result);
        
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("✅ Headers retrieved successfully");
          const headers = result.value;
          console.log("📄 Headers length:", headers ? headers.length : "null");
          console.log("📄 Headers preview:", headers ? headers.substring(0, 200) + "..." : "null");
          
          // Parse DMARC results
          const dmarcResult = parseDmarcResult(headers);
          console.log("🛡️ DMARC parsed:", dmarcResult);
          
          // Parse DKIM results
          const dkimResult = parseDkimResult(headers);
          console.log("🔐 DKIM parsed:", dkimResult);
          
          // Parse SPF results
          const spfResult = parseSpfResult(headers);
          console.log("📮 SPF parsed:", spfResult);
          
          const results = {
            dmarc: dmarcResult,
            dkim: dkimResult,
            spf: spfResult
          };
          
          console.log("📋 Final results:", results);
          callback(results);
        } else {
          console.error("❌ Failed to get headers");
          console.error("🚨 Error:", result.error);
          console.error("📛 Error name:", result.error.name);
          console.error("💬 Error message:", result.error.message);
          console.error("🔢 Error code:", result.error.code);
          
          callback({
            dmarc: { status: "unknown", details: "Error retrieving headers: " + result.error.message },
            dkim: { status: "unknown", details: "Error retrieving headers: " + result.error.message },
            spf: { status: "unknown", details: "Error retrieving headers: " + result.error.message }
          });
        }
      });
    } catch (error) {
      console.error("💥 Exception in parseEmailHeaders:", error);
      console.error("📛 Error name:", error.name);
      console.error("💬 Error message:", error.message);
      console.error("📚 Error stack:", error.stack);
      
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
  
  // Show results - ENHANCED DEBUG VERSION
  function showResult(title, content) {
    console.log("🔍 === SHOWING RESULT ===");
    console.log("📋 Title:", title);
    console.log("📄 Content:", content);
    console.log("🕐 Timestamp:", new Date().toLocaleTimeString());
    
    // Comprehensive API availability check
    const apiCheck = {
      "Office.context": !!Office.context,
      "Office.context.mailbox": !!(Office.context && Office.context.mailbox),
      "Office.context.mailbox.item": !!(Office.context && Office.context.mailbox && Office.context.mailbox.item),
      "notificationMessages": !!(Office.context && Office.context.mailbox && Office.context.mailbox.item && Office.context.mailbox.item.notificationMessages),
      "displayDialogAsync": !!(Office.context && Office.context.ui && Office.context.ui.displayDialogAsync)
    };
    
    console.table(apiCheck);
    
    // Try notification with extensive error handling
    if (Office.context && Office.context.mailbox && Office.context.mailbox.item && Office.context.mailbox.item.notificationMessages) {
      const debugKey = "authopsy_debug_" + Date.now();
      console.log("🔔 Attempting notification with key:", debugKey);
      
      const notificationData = {
        type: "informationalMessage",
        message: title + ": " + content.substring(0, 100) + (content.length > 100 ? "..." : ""),
        persistent: true
      };
      
      console.log("📤 Notification data:", notificationData);
      
      Office.context.mailbox.item.notificationMessages.addAsync(debugKey, notificationData, function(result) {
        console.log("🎯 Notification callback triggered");
        console.log("📊 Result object:", result);
        console.log("✅ Status:", result.status);
        console.log("🔄 AsyncContext:", result.asyncContext);
        
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("❌ NOTIFICATION FAILED");
          console.error("🚨 Error object:", result.error);
          console.error("📛 Error name:", result.error.name);
          console.error("💬 Error message:", result.error.message);
          console.error("🔢 Error code:", result.error.code);
          
          // Try to get more error details
          try {
            console.error("🔍 Full error:", JSON.stringify(result.error, null, 2));
          } catch (e) {
            console.error("🔍 Error serialization failed:", e.message);
          }
        } else {
          console.log("🎉 NOTIFICATION SUCCESS!");
          console.log("⏰ Auto-removing in 10 seconds...");
          
          // Auto-remove with debug info
          setTimeout(function() {
            console.log("🗑️ Removing notification:", debugKey);
            Office.context.mailbox.item.notificationMessages.removeAsync(debugKey, function(removeResult) {
              console.log("🗑️ Removal result:", removeResult);
            });
          }, 10000);
        }
      });
    } else {
      console.error("❌ NOTIFICATION API NOT AVAILABLE");
      console.log("🔍 Available item APIs:", Object.keys(Office.context?.mailbox?.item || {}));
    }
  }
  
  // Enhanced handler functions with more debugging
  function showDmarcDetails(event) {
    console.log("🛡️ === DMARC ANALYSIS STARTED ===");
    console.log("📧 Event object:", event);
    console.log("⚡ Event type:", typeof event);
    console.log("🔧 Event properties:", Object.keys(event || {}));
    
    parseEmailHeaders(function(results) {
      console.log("📊 DMARC Results:", results.dmarc);
      showResult("DMARC Analysis", results.dmarc.details);
    });
    
    // Signal completion with debug
    if (event && event.completed) {
      console.log("✅ Completing DMARC event");
      event.completed();
    } else {
      console.warn("⚠️ No completion callback available");
    }
  }
  
  function showDkimDetails(event) {
    console.log("🔐 === DKIM ANALYSIS STARTED ===");
    console.log("📧 Event object:", event);
    
    parseEmailHeaders(function(results) {
      console.log("📊 DKIM Results:", results.dkim);
      showResult("DKIM Analysis", results.dkim.details);
    });
    
    if (event && event.completed) {
      console.log("✅ Completing DKIM event");
      event.completed();
    } else {
      console.warn("⚠️ No completion callback available");
    }
  }
  
  function showSpfDetails(event) {
    console.log("📮 === SPF ANALYSIS STARTED ===");
    console.log("📧 Event object:", event);
    
    parseEmailHeaders(function(results) {
      console.log("📊 SPF Results:", results.spf);
      showResult("SPF Analysis", results.spf.details);
    });
    
    if (event && event.completed) {
      console.log("✅ Completing SPF event");
      event.completed();
    } else {
      console.warn("⚠️ No completion callback available");
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

  console.log("🏁 AUTHOPSY COMMANDS LOADED SUCCESSFULLY 🏁");
})();