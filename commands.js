(function () {
  "use strict";

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
  
  // Show notification instead of dialog for better compatibility
  function showNotification(title, content) {
    try {
      Office.context.mailbox.item.notificationMessages.addAsync("authopsyResult", {
        type: "informationalMessage",
        message: title + ": " + content,
        icon: "iconid",
        persistent: true
      }, function(result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Notification failed: " + result.error.message);
        }
      });
    } catch (error) {
      console.error("Notification error: " + error.message);
    }
  }
  
  // Handler for DMARC details
  function showDmarcDetails(event) {
    console.log("DMARC details requested");
    try {
      parseEmailHeaders(function(results) {
        showNotification("DMARC Analysis", results.dmarc.details);
        if (event && event.completed) {
          event.completed();
        }
      });
    } catch (error) {
      console.error("Error in showDmarcDetails: " + error.message);
      showNotification("DMARC Error", "Failed to analyze DMARC: " + error.message);
      if (event && event.completed) {
        event.completed();
      }
    }
  }
  
  // Handler for DKIM details
  function showDkimDetails(event) {
    console.log("DKIM details requested");
    try {
      parseEmailHeaders(function(results) {
        showNotification("DKIM Analysis", results.dkim.details);
        if (event && event.completed) {
          event.completed();
        }
      });
    } catch (error) {
      console.error("Error in showDkimDetails: " + error.message);
      showNotification("DKIM Error", "Failed to analyze DKIM: " + error.message);
      if (event && event.completed) {
        event.completed();
      }
    }
  }
  
  // Handler for SPF details
  function showSpfDetails(event) {
    console.log("SPF details requested");
    try {
      parseEmailHeaders(function(results) {
        showNotification("SPF Analysis", results.spf.details);
        if (event && event.completed) {
          event.completed();
        }
      });
    } catch (error) {
      console.error("Error in showSpfDetails: " + error.message);
      showNotification("SPF Error", "Failed to analyze SPF: " + error.message);
      if (event && event.completed) {
        event.completed();
      }
    }
  }

  // Register functions with Office.actions
  function registerActions() {
    try {
      if (Office.actions) {
        Office.actions.associate("showDmarcDetails", showDmarcDetails);
        Office.actions.associate("showDkimDetails", showDkimDetails);
        Office.actions.associate("showSpfDetails", showSpfDetails);
        console.log("Actions registered successfully");
      } else {
        console.warn("Office.actions not available");
      }
    } catch (error) {
      console.error("Failed to register actions: " + error.message);
    }
  }

  // Register actions when Office is ready
  if (Office.context) {
    registerActions();
  } else {
    Office.onReady(function() {
      registerActions();
    });
  }

  // Also expose globally as fallback
  window.showDmarcDetails = showDmarcDetails;
  window.showDkimDetails = showDkimDetails;
  window.showSpfDetails = showSpfDetails;
})();