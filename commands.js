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
  
  // Show results using notification messages (more reliable than dialogs)
  function showResult(title, content) {
    console.log(title + ": " + content);
    
    try {
      // Try notification first (more reliable in Outlook Web)
      if (Office.context.mailbox.item.notificationMessages) {
        Office.context.mailbox.item.notificationMessages.addAsync("authopsyResult", {
          type: "informationalMessage",
          message: title + ": " + content.substring(0, 150) + (content.length > 150 ? "..." : ""),
          icon: "iconid",
          persistent: false
        }, function(result) {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Notification failed, trying task pane: " + result.error.message);
            tryTaskPane(title, content);
          } else {
            console.log("Notification shown successfully");
          }
        });
      } else {
        console.log("Notifications not available, trying task pane");
        tryTaskPane(title, content);
      }
    } catch (error) {
      console.error("Error showing result: " + error.message);
      tryTaskPane(title, content);
    }
  }
  
  // Fallback to try opening task pane
  function tryTaskPane(title, content) {
    try {
      // For Outlook Web, try opening the task pane
      if (Office.context.mailbox.item.body) {
        Office.context.ui.displayDialogAsync(
          "https://rdyy89.github.io/authopsy-addin/results.html?title=" + 
          encodeURIComponent(title) + "&content=" + encodeURIComponent(content),
          { 
            height: 60, 
            width: 80, 
            displayInIframe: false  // Don't use iframe for better compatibility
          },
          function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.error("Dialog failed: " + result.error.message);
              // Final fallback - log to console
              console.log("RESULT: " + title + " - " + content);
            } else {
              console.log("Dialog opened successfully");
              const dialog = result.value;
              
              // Handle dialog events
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                if (arg.message === "dialogClosed") {
                  dialog.close();
                }
              });
            }
          }
        );
      }
    } catch (error) {
      console.error("Task pane fallback failed: " + error.message);
      console.log("FINAL RESULT: " + title + " - " + content);
    }
  }
  
  // Handler for DMARC details
  function showDmarcDetails(event) {
    console.log("DMARC details requested");
    try {
      parseEmailHeaders(function(results) {
        showResult("DMARC Analysis", results.dmarc.details);
        
        // Always signal completion
        if (event && event.completed) {
          setTimeout(function() {
            event.completed();
          }, 100);
        }
      });
    } catch (error) {
      console.error("Error in showDmarcDetails: " + error.message);
      showResult("DMARC Error", "Failed to analyze DMARC: " + error.message);
      
      if (event && event.completed) {
        setTimeout(function() {
          event.completed();
        }, 100);
      }
    }
  }
  
  // Handler for DKIM details
  function showDkimDetails(event) {
    console.log("DKIM details requested");
    try {
      parseEmailHeaders(function(results) {
        showResult("DKIM Analysis", results.dkim.details);
        
        // Always signal completion
        if (event && event.completed) {
          setTimeout(function() {
            event.completed();
          }, 100);
        }
      });
    } catch (error) {
      console.error("Error in showDkimDetails: " + error.message);
      showResult("DKIM Error", "Failed to analyze DKIM: " + error.message);
      
      if (event && event.completed) {
        setTimeout(function() {
          event.completed();
        }, 100);
      }
    }
  }
  
  // Handler for SPF details
  function showSpfDetails(event) {
    console.log("SPF details requested");
    try {
      parseEmailHeaders(function(results) {
        showResult("SPF Analysis", results.spf.details);
        
        // Always signal completion
        if (event && event.completed) {
          setTimeout(function() {
            event.completed();
          }, 100);
        }
      });
    } catch (error) {
      console.error("Error in showSpfDetails: " + error.message);
      showResult("SPF Error", "Failed to analyze SPF: " + error.message);
      
      if (event && event.completed) {
        setTimeout(function() {
          event.completed();
        }, 100);
      }
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