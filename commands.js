(function () {
  "use strict";

  let _headerResults = {};
  let _messageId = "";
  
  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    console.log("Commands initialized with reason: " + reason);
  };

  // Helper function to parse email headers
  function parseEmailHeaders(callback) {
    Office.context.mailbox.item.getAllInternetHeadersAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const headers = result.value;
        _messageId = Office.context.mailbox.item.itemId;
        
        // Parse DMARC results
        const dmarcResult = parseDmarcResult(headers);
        
        // Parse DKIM results
        const dkimResult = parseDkimResult(headers);
        
        // Parse SPF results
        const spfResult = parseSpfResult(headers);
        
        _headerResults = {
          dmarc: dmarcResult,
          dkim: dkimResult,
          spf: spfResult
        };
        
        callback(_headerResults);
      } else {
        console.error("Failed to get headers: " + result.error.message);
        callback({
          dmarc: { status: "unknown", details: "Error retrieving headers" },
          dkim: { status: "unknown", details: "Error retrieving headers" },
          spf: { status: "unknown", details: "Error retrieving headers" }
        });
      }
    });
  }
  
  // Parse DMARC result from headers
  function parseDmarcResult(headers) {
    try {
      // Look for Authentication-Results header with dmarc
      const dmarcRegex = /Authentication-Results:.*dmarc=([^;]+)/i;
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
      return { status: "unknown", details: "Error parsing DMARC results" };
    }
  }
  
  // Parse DKIM result from headers
  function parseDkimResult(headers) {
    try {
      // Look for Authentication-Results header with dkim
      const dkimRegex = /Authentication-Results:.*dkim=([^;]+)/i;
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
      return { status: "unknown", details: "Error parsing DKIM results" };
    }
  }
  
  // Parse SPF result from headers
  function parseSpfResult(headers) {
    try {
      // Look for Authentication-Results header with spf
      const spfRegex = /Authentication-Results:.*spf=([^;]+)/i;
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
      return { status: "unknown", details: "Error parsing SPF results" };
    }
  }
  
  // Show dialog with details - improved error handling
  function showDialog(title, content) {
    try {
      Office.context.ui.displayDialogAsync(
        "https://rdyy89.github.io/authopsy-addin/dialog.html?title=" + 
        encodeURIComponent(title) + 
        "&content=" + 
        encodeURIComponent(content),
        { height: 40, width: 30, displayInIframe: true },
        function (result) {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Dialog creation failed: " + result.error.message);
            // Show notification as fallback
            Office.context.mailbox.item.notificationMessages.addAsync("authopsyError", {
              type: "informationalMessage",
              message: title + ": " + content,
              icon: "iconid",
              persistent: false
            });
          }
        }
      );
    } catch (error) {
      console.error("Dialog error: " + error.message);
      // Show notification as fallback
      try {
        Office.context.mailbox.item.notificationMessages.addAsync("authopsyError", {
          type: "informationalMessage",
          message: title + ": " + content,
          icon: "iconid",
          persistent: false
        });
      } catch (notifError) {
        console.error("Notification also failed: " + notifError.message);
      }
    }
  }
  
  // Handler for DMARC details - improved
  function showDmarcDetails(event) {
    console.log("DMARC details requested");
    parseEmailHeaders(function(results) {
      showDialog("DMARC Details", results.dmarc.details);
      if (event && event.completed) {
        event.completed();
      }
    });
  }
  
  // Handler for DKIM details - improved
  function showDkimDetails(event) {
    console.log("DKIM details requested");
    parseEmailHeaders(function(results) {
      showDialog("DKIM Details", results.dkim.details);
      if (event && event.completed) {
        event.completed();
      }
    });
  }
  
  // Handler for SPF details - improved
  function showSpfDetails(event) {
    console.log("SPF details requested");
    parseEmailHeaders(function(results) {
      showDialog("SPF Details", results.spf.details);
      if (event && event.completed) {
        event.completed();
      }
    });
  }

  // Register functions with error handling
  try {
    if (Office.actions) {
      Office.actions.associate("showDmarcDetails", showDmarcDetails);
      Office.actions.associate("showDkimDetails", showDkimDetails);
      Office.actions.associate("showSpfDetails", showSpfDetails);
    }
  } catch (error) {
    console.error("Failed to register actions: " + error.message);
  }

  // Also expose globally as fallback
  window.showDmarcDetails = showDmarcDetails;
  window.showDkimDetails = showDkimDetails;
  window.showSpfDetails = showSpfDetails;
})();