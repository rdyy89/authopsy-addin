(function () {
  "use strict";

  let _headerResults = {};
  let _messageId = "";

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      // Set up event handlers
      $("#dmarcDetails").on("click", showDmarcDetails);
      $("#dkimDetails").on("click", showDkimDetails);
      $("#spfDetails").on("click", showSpfDetails);
      $("#pinButton").on("click", handlePinning);
      
      // Start the analysis
      analyzeEmailHeaders();
    });
  };

  // Main function to analyze email headers
  function analyzeEmailHeaders() {
    // Set status to loading
    updateUIStatus("loading");
    
    // Get email headers
    Office.context.mailbox.item.getAllInternetHeadersAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const headers = result.value;
        _messageId = Office.context.mailbox.item.itemId;
        
        // Parse results
        const dmarcResult = parseDmarcResult(headers);
        const dkimResult = parseDkimResult(headers);
        const spfResult = parseSpfResult(headers);
        
        // Store results
        _headerResults = {
          dmarc: dmarcResult,
          dkim: dkimResult,
          spf: spfResult
        };
        
        // Update UI with results
        updateUIWithResults(_headerResults);
      } else {
        console.error("Failed to get headers: " + result.error.message);
        handleError("Could not retrieve email headers.");
      }
    });
  }
  
  // Update UI based on loading status
  function updateUIStatus(status) {
    if (status === "loading") {
      $("#dmarcStatus").text("Loading...").removeClass().addClass("result-status loading");
      $("#dkimStatus").text("Loading...").removeClass().addClass("result-status loading");
      $("#spfStatus").text("Loading...").removeClass().addClass("result-status loading");
      
      $("#dmarcIcon").text("❓");
      $("#dkimIcon").text("❓");
      $("#spfIcon").text("❓");
    }
  }
  
  // Update UI with analysis results
  function updateUIWithResults(results) {
    // Update DMARC
    updateStatusElement("dmarc", results.dmarc.status, getStatusText(results.dmarc.status));
    
    // Update DKIM
    updateStatusElement("dkim", results.dkim.status, getStatusText(results.dkim.status));
    
    // Update SPF
    updateStatusElement("spf", results.spf.status, getStatusText(results.spf.status));
  }
  
  // Helper function to update status element
  function updateStatusElement(type, status, text) {
    const $statusElement = $("#" + type + "Status");
    const $iconElement = $("#" + type + "Icon");
    
    // Remove all classes and add the status class
    $statusElement.removeClass().addClass("result-status " + status);
    
    // Set the text
    $statusElement.text(text);
    
    // Update icon with emoji
    const iconMap = {
      "pass": "✅",
      "fail": "❌", 
      "unknown": "❓"
    };
    $iconElement.text(iconMap[status] || "❓");
  }
  
  // Get status text from status code
  function getStatusText(status) {
    switch(status) {
      case "pass": return "Pass";
      case "fail": return "Fail";
      case "unknown": return "Unknown";
      default: return "Unknown";
    }
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
  
  // Handle errors
  function handleError(message) {
    $("#dmarcStatus").text("Error").removeClass().addClass("result-status unknown");
    $("#dkimStatus").text("Error").removeClass().addClass("result-status unknown");
    $("#spfStatus").text("Error").removeClass().addClass("result-status unknown");
    
    console.error("Error: " + message);
  }
  
  // Show dialog with details
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
            // Fallback to alert if dialog fails
            alert(title + "\n\n" + content);
          }
        }
      );
    } catch (error) {
      console.error("Dialog error: " + error.message);
      // Fallback to alert if dialog completely fails
      alert(title + "\n\n" + content);
    }
  }
  
  // Show DMARC details
  function showDmarcDetails() {
    if (_headerResults.dmarc) {
      showDialog("DMARC Details", _headerResults.dmarc.details);
    } else {
      showDialog("DMARC Details", "No DMARC information available");
    }
  }
  
  // Show DKIM details
  function showDkimDetails() {
    if (_headerResults.dkim) {
      showDialog("DKIM Details", _headerResults.dkim.details);
    } else {
      showDialog("DKIM Details", "No DKIM information available");
    }
  }
  
  // Show SPF details
  function showSpfDetails() {
    if (_headerResults.spf) {
      showDialog("SPF Details", _headerResults.spf.details);
    } else {
      showDialog("SPF Details", "No SPF information available");
    }
  }
  
  // Handle pinning behavior
  function handlePinning() {
    try {
      // Try to pin the add-in
      Office.context.ui.displayDialogAsync(
        "https://rdyy89.github.io/authopsy-addin/dialog.html?title=Pin%20Authopsy&content=This%20add-in%20is%20now%20pinned%20for%20quick%20access.",
        { height: 30, width: 25, displayInIframe: true },
        function (result) {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Pin dialog failed: " + result.error.message);
          }
        }
      );
    } catch (error) {
      console.error("Pin error: " + error.message);
    }
  }

  // Global functions for command access
  window.showDmarcDetails = showDmarcDetails;
  window.showDkimDetails = showDkimDetails;
  window.showSpfDetails = showSpfDetails;
})();