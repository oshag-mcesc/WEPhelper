/**
 * Email functionality for WEPhelper library
 * Consolidates all email sending logic and moves it to the library
 */

// Configuration for link time tabs
const EMAIL_CONFIG = {
  linkTimeTab: {
    initial: "forGoalLinks",
    mid: "forMidYear",
    final: "forFinalEvaluation"
  }
};

/**
 * Main email namespace - contains all email functionality
 */
const nsEmailer = (function () {

  /**
   * Main function to handle the email sending process
   * Gets data from sheet, confirms with user, and sends emails
   */
  function emailLinks() {
    let done = false; // to let calling function know if we are good or not

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // Get settings using cSettings
      const cSettings = getSettingsInstance('config');
      const linktime = cSettings.getSetting('linktime') || 'initial';

      // Get the correct tab name based on linktime setting
      const tab = EMAIL_CONFIG.linkTimeTab[linktime];

      if (!tab) {
        throw new Error(`Invalid linktime setting: ${linktime}`);
      }

      const ws = ss.getSheetByName(tab);

      if (!ws) {
        throw new Error(`Sheet "${tab}" not found. Please check your linktime setting.`);
      }

      const [header, ...allData] = ws.getDataRange().getValues();

      // Create a new array of lowercase headers for reliable, case-insensitive matching
      const lowerCaseHeaders = header.map(h => h.toLowerCase());

      // Get the index of each column using the new lowercase header array
      const teacherEmailIndex = lowerCaseHeaders.indexOf("teacheremail");
      const preFilledURLIndex = lowerCaseHeaders.indexOf("prefilledurl");
      const linksIndex = lowerCaseHeaders.indexOf("links");

      if (teacherEmailIndex === -1 || preFilledURLIndex === -1 || linksIndex === -1) {
        throw new Error("Required columns (TeacherEmail, PreFilledURL, Links) not found in sheet");
      }

      const teachEmail = [...new Set(allData.map(row => row[teacherEmailIndex]))];

      // Map the data we need
      const info = allData.map(row => [
        row[teacherEmailIndex],
        row[preFilledURLIndex],
        row[linksIndex]
      ]);

      const ui = SpreadsheetApp.getUi();
      const result = ui.alert(
        'Please confirm',
        `You are sending a total of ${info.length} links for "${tab}". \nAre you sure you want to continue?`,
        ui.ButtonSet.YES_NO
      );

      // Process the user's response
      if (result == ui.Button.YES) {
        // Get email settings
        const emailSettings = {
          subject: cSettings.getSetting('subject') || 'Email Links',
          greeting: cSettings.getSetting('greeting') || 'Hello,',
          closing: cSettings.getSetting('closing') || 'Thank you.'
        };

        logIt({
          level: "info",
          theMsg: `Starting to send ${teachEmail.length} emails for ${tab}`
        });

        // Send progress notification
        showProgress(true);
        // Variable for bad email counting
        const failures = [];

        for (let i = 0; i < teachEmail.length; i++) {
          const data = info.filter(row => teachEmail[i] === row[0]);

          const emailInfo = {
            recipient: teachEmail[i],
            data: data,
            subject: emailSettings.subject,
            greeting: emailSettings.greeting,
            closing: emailSettings.closing
          };
          //ADDED 4/230/2026 to catch bad emails BUT to keep sending
          try {
            sendTheMail(emailInfo);
          } catch (err) {
            logIt({
              level: "severe",
              theMsg: `Failed to send email to ${teachEmail[i]} — skipping`,
              error: err
            });
            failures.push({
              email: teachEmail[i],
              error: err.message
            });
          }
        }

        showProgress(false);
        done = true;

        if (failures.length > 0) {
          writeEmailErrors(failures);
          logIt({
            level: "severe",
            theMsg: `Email sending complete. ${teachEmail.length - failures.length} sent, ${failures.length} failed.`
          });
          ui.alert(
            'Completed with errors',
            `Emails sent: ${teachEmail.length - failures.length}\nFailed: ${failures.length}\n\nSee the "Email Errors" tab for details.`,
            ui.ButtonSet.OK
          );
        } else {
          logIt({
            level: "info",
            theMsg: `Successfully sent emails to ${teachEmail.length} recipients`
          });
          ui.alert('Success!', `Emails sent successfully to ${teachEmail.length} recipients.`, ui.ButtonSet.OK);
        }

      } else {
        // User clicked "No" or X in the title bar
        ui.alert('Email aborted.');
        logIt({
          level: "info",
          theMsg: "Email sending aborted by user"
        });
      }

    } catch (err) {
      showProgress(false);
      logIt({
        level: "severe",
        theMsg: "Error in email links process",
        error: err
      });

      SpreadsheetApp.getUi().alert('Error!', 'Something went wrong emailing links: ' + err.message, SpreadsheetApp.getUi().ButtonSet.OK);
      throw err;
    }

    return done;
  }

  /**
   * Send individual email with links
   * @param {Object} emailInfo - The emailInfo object
   * @param {Array} emailInfo.data - The data of all the sheet data
   * @param {string} emailInfo.recipient - The email of the recipient
   * @param {string} emailInfo.subject - The subject for the email
   * @param {string} emailInfo.greeting - The opening of the message
   * @param {string} emailInfo.closing - The closing of the message
   */
  function sendTheMail(emailInfo) {
    try {
      // Get all the links into HTML rows   
      const theRows = emailInfo.data.map(row => {
        return `<tr><td><a href="${row[1]}">${row[2]}</a></td></tr>`;
      }).join('');

      const temp = HtmlService.createTemplateFromFile("emailTemplate");
      temp.theRows = theRows;
      temp.greeting = emailInfo.greeting;
      temp.closing = emailInfo.closing;

      const html = temp.evaluate().getContent().toString();

      GmailApp.sendEmail(
        emailInfo.recipient,
        emailInfo.subject,
        "Body",
        { htmlBody: html }
      );

      logIt({
        level: 'info',
        theMsg: `Email sent to ${emailInfo.recipient}`
      });

    } catch (err) {
      logIt({
        level: "severe",
        theMsg: `Error sending email to ${emailInfo.recipient}`,
        error: err
      });
      throw err;
    }
  }

  /**
   * Show/hide progress indicator
   * @param {boolean} show - Whether to show or hide progress
   */
  function showProgress(show) {
    try {
      // You can implement your progress function here
      // For now, just log the progress state
      if (show) {
        logIt({
          level: "info",
          theMsg: "Email sending process started - showing progress"
        });
      } else {
        logIt({
          level: "info",
          theMsg: "Email sending process completed - hiding progress"
        });
      }
    } catch (err) {
      logIt({
        level: "warning",
        theMsg: "Error updating progress indicator",
        error: err
      });
    }
  }

  /**
   * Get email settings for preview or testing
   * @returns {Object} Current email settings
   */
  function getEmailSettings() {
    try {
      const cSettings = getSettingsInstance('config');
      return {
        subject: cSettings.getSetting('subject') || 'Email Links',
        greeting: cSettings.getSetting('greeting') || 'Hello,',
        closing: cSettings.getSetting('closing') || 'Thank you.',
        linktime: cSettings.getSetting('linktime') || 'initial'
      };
    } catch (err) {
      logIt({
        level: "warning",
        theMsg: "Error getting email settings",
        error: err
      });
      return {
        subject: 'Email Links',
        greeting: 'Hello,',
        closing: 'Thank you.',
        linktime: 'initial'
      };
    }
  }

/**
 * Writes email sending failures to a dedicated "Email Errors" tab in the active spreadsheet.
 *
 * If the tab does not exist, it is created automatically. If it already exists,
 * its contents are cleared before writing new data, ensuring only the most
 * recent run's failures are displayed.
 *
 * Each failure is logged with a timestamp, the recipient's email address, and
 * the associated error message. The header row is bolded, frozen, and all
 * columns are auto-resized for readability.
 *
 * @function writeEmailErrors
 * @param {Object[]} failures - Array of failure objects from the email sending process.
 * @param {string}   failures[].email - The recipient's email address that failed.
 * @param {string}   failures[].error - The error message describing why the send failed.
 * @returns {void}
 *
 * @example
 * const failures = [
 *   { email: "user@example.com", error: "Invalid address" },
 *   { email: "other@example.com", error: "Quota exceeded" }
 * ];
 * writeEmailErrors(failures);
 */
function writeEmailErrors(failures) {
  try {
    // Get a reference to the currently active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const tabName = "Email Errors";

    // Attempt to find an existing "Email Errors" sheet
    let ws = ss.getSheetByName(tabName);

    if (!ws) {
      // Sheet doesn't exist yet — create it fresh
      ws = ss.insertSheet(tabName);
    } else {
      // Sheet already exists — wipe prior run's data before writing new errors
      ws.clearContents();
    }

    // Write the header row as the first row in the sheet
    ws.appendRow(["Timestamp", "Recipient Email", "Error Message"]);

    // Capture a single timestamp to apply uniformly to all failures in this run
    const timestamp = new Date().toLocaleString();

    // Iterate over each failure object and write it as its own row
    failures.forEach(f => {
      ws.appendRow([timestamp, f.email, f.error]);
    });

    // --- Formatting ---

    // Bold the header row (row 1, columns 1–3) to visually distinguish it from data rows
    ws.getRange(1, 1, 1, 3).setFontWeight("bold");

    // Freeze the header row so it stays visible when scrolling through many errors
    ws.setFrozenRows(1);

    // Auto-resize all three columns to fit their content without manual adjustment
    ws.autoResizeColumns(1, 3);

    // Log a success message indicating how many errors were written
    logIt({
      level: "info",
      theMsg: `Wrote ${failures.length} email error(s) to "${tabName}" tab`
    });

  } catch (err) {
    // If anything above throws, log it as a severe error rather than letting it fail silently
    logIt({
      level: "severe",
      theMsg: "Failed to write email errors to sheet",
      error: err
    });
  }
}
  // Return public interface
  return {
    emailLinks: emailLinks,
    sendTheMail: sendTheMail,
    getEmailSettings: getEmailSettings,
    showProgress: showProgress,
    writeEmailErrors: writeEmailErrors
  };

})();

/**
 * Namespace for email-related helper functions
 * These functions handle the interaction between bound script and library functions
 */
var nsPropsEmailer = (function () {

  /**
   * Helper function to send email links
   * Called from bound script via callLibraryNS
   */
  function helperEmailLinks() {
    try {
      return nsEmailer.emailLinks();
    } catch (error) {
      logIt({
        level: "severe",
        theMsg: "Error in helper email links function",
        error: error
      });
      throw new Error("Error sending email links: " + error.message);
    }
  }

  /**
   * Helper function to get email settings
   * Called from bound script via callLibraryNS
   */
  function helperGetEmailSettings() {
    try {
      return nsEmailer.getEmailSettings();
    } catch (error) {
      logIt({
        level: "warning",
        theMsg: "Error getting email settings in helper",
        error: error
      });
      return {};
    }
  }

  /**
   * Helper function to send individual email (for testing)
   * Called from bound script via callLibraryNS
   */
  function helperSendTestEmail(emailInfo) {
    try {
      return nsEmailer.sendTheMail(emailInfo);
    } catch (error) {
      logIt({
        level: "severe",
        theMsg: "Error sending test email",
        error: error
      });
      throw new Error("Error sending test email: " + error.message);
    }
  }

  // Return public interface
  return {
    helperEmailLinks: helperEmailLinks,
    helperGetEmailSettings: helperGetEmailSettings,
    helperSendTestEmail: helperSendTestEmail
  };

})();

// Make nsEmailer available globally for backward compatibility if needed
var nsWEPemail = nsEmailer;