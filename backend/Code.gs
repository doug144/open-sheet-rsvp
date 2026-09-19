// =================================================================
// SCRIPT CONFIGURATION
// =================================================================

/**
 * Retrieves script properties. It's recommended to set these in the Apps Script editor:
 * File > Project Properties > Script Properties.
 *
 * @property {string} websiteUrl - The full URL to your deployed web app's index.html.
 * @property {boolean} sendEmails - 'true' to enable sending confirmation/reminder emails.
 * @returns {Object} The script configuration properties.
 */
function getScriptConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    websiteUrl: props.getProperty('WEBSITE_URL') || ScriptApp.getService().getUrl(),
    sendEmails: props.getProperty('SEND_EMAILS') === 'true'
  };
}

// =================================================================
// WEB APP ENTRY POINTS (doGet, doPost)
// =================================================================

/**
 * Handles GET requests. This is used for the initial lookup of a guest group.
 * The frontend will call this with a group ID to get invitation details.
 *
 * Query Params: ?id=GRP-101 (or ?groupId=GRP-101)
 */
function doGet(e) {
  try {
    var searchGroupId = e.parameter.id || e.parameter.groupId || "";
    searchGroupId = searchGroupId.toString().toLowerCase().trim();

    if (!searchGroupId) {
      return createJsonResponse({ found: false, error: "No Group ID provided." });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var guestsSheet = ss.getSheetByName("Guests");
    var eventsSheet = ss.getSheetByName("Events");
    var configSheet = ss.getSheetByName("Config");

    if (!guestsSheet || !eventsSheet || !configSheet) {
      throw new Error("Required sheets ('Guests', 'Events', 'Config') not found.");
    }

    var guestsList = sheetToObjects(guestsSheet);
    var eventsList = sheetToObjects(eventsSheet);
    var globalConfig = getGlobalConfig(ss);
    delete globalConfig.admin_password; // Remove sensitive data before sending to client

    var groupGuests = guestsList.filter(function(guest) {
      return guest.Group_ID && guest.Group_ID.toString().toLowerCase().trim() === searchGroupId;
    });

    if (groupGuests.length === 0) {
      return createJsonResponse({ found: false, config: globalConfig });
    }

    // 1. Find all unique event IDs this group is invited to.
    var allInvitedEventIds = {}; // Using an object as a set for efficient lookups.
    groupGuests.forEach(function(guest) {
      (guest.Allowed_Events || "").split(',').forEach(function(id) {
        var trimmedId = id.trim();
        if (trimmedId) {
          allInvitedEventIds[trimmedId] = true;
        }
      });
    });

    // 2. Build the eventsMap containing ONLY the events the group is invited to.
    var eventsMap = eventsList.reduce(function(map, event) {
      if (event.Event_ID && allInvitedEventIds[event.Event_ID]) {
        map[event.Event_ID] = {
          id: event.Event_ID,
          title: event.Title,
          dateTime: event.DateTime,
          location: event.Location,
          description: event.Description
        };
      }
      return map;
    }, {});

    var responseGuests = groupGuests.map(function(guest) {
      var allowedEventIds = (guest.Allowed_Events || "").split(',').map(function(id) {
        return id.trim();
      }).filter(id => id && eventsMap[id]); // Ensure IDs are valid and exist in the events map

      var existingRsvps = {};
      try {
        existingRsvps = JSON.parse(guest.RSVPs || '{}');
      } catch (e) {}

      var additionalGuests = parseInt(guest.Additional_Guests, 10);
      return {
        row: guest.rowIndex,
        guestId: guest.Guest_ID,
        fullName: guest.Full_Name,
        isPlusOne: guest.Is_Plus_One === true || guest.Is_Plus_One === 'TRUE',
        additionalGuestsAllowed: !isNaN(additionalGuests) ? additionalGuests : 0,
        allowedEventIds: allowedEventIds,
        existingRsvps: existingRsvps
      };
    });

    var groupNotes = groupGuests[0] ? groupGuests[0].Notes : "";

    var responsePayload = {
      found: true,
      groupId: searchGroupId,
      guests: responseGuests,
      eventsMap: eventsMap,
      groupNotes: groupNotes,
      config: globalConfig
    };

    return createJsonResponse(responsePayload);

  } catch (err) {
    Logger.log("doGet Error: " + err.message + "\n" + err.stack);
    return createJsonResponse({ found: false, error: "An internal server error occurred." });
  }
}

/**
 * Handles POST requests. This is used for submitting RSVPs from the main form
 * and for all actions from the admin panel.
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var guestsSheet = ss.getSheetByName("Guests");

    // --- Admin Action Routing ---
    if (payload.action) {
      // Public action, no auth needed
      if (payload.action === 'adminLogin') {
        var isValid = checkAdminPassword(payload.password);
        if (isValid) return createJsonResponse({ status: 'success', data: { login: true } });
        return createJsonResponse({ status: 'unauthorized', error: 'Incorrect password.' });
      }

      // All other admin actions require authentication
      if (!payload.auth || !checkAdminPassword(payload.auth.password)) {
        return createJsonResponse({ status: 'unauthorized', error: 'Authentication failed.' });
      }

      if (payload.action === 'getDashboard') {
        return createJsonResponse({ status: 'success', data: getAdminDashboardData() });
      }
      if (payload.action === 'sendReminders') {
        return createJsonResponse({ status: 'success', data: { message: sendReminderEmails(payload.groupIds) } });
      }
    }
    
    // --- RSVP Submission Logic ---
    if (!payload.guestResponses || !Array.isArray(payload.guestResponses)) {
      return createJsonResponse({ result: "error", message: "Invalid payload format." });
    }

    var guestData = guestsSheet.getDataRange().getValues();
    var headers = guestData[0];
    var rsvpCol = headers.indexOf("RSVPs");
    var notesCol = headers.indexOf("Notes");
    var fullNameCol = headers.indexOf("Full_Name");

    var primaryGuestEmail = "";
    var primaryGuestName = "";

    payload.guestResponses.forEach(function(response) {
      var rowIndex = response.row;
      if (rowIndex > 0 && rowIndex <= guestData.length) {
        var sheetRowIndex = parseInt(rowIndex) + 1; // Convert 0-based to 1-based for sheet
        
        // Update RSVPs
        guestsSheet.getRange(sheetRowIndex, rsvpCol + 1).setValue(JSON.stringify(response.rsvps));

        // Update Notes for the entire group (applied to each member)
        if (notesCol !== -1 && payload.notes !== undefined) {
          guestsSheet.getRange(sheetRowIndex, notesCol + 1).setValue(payload.notes);
        }

        // Update Plus One name if provided
        if (response.isPlusOne && response.plusOneName && fullNameCol !== -1) {
          guestsSheet.getRange(sheetRowIndex, fullNameCol + 1).setValue(response.plusOneName);
        }

        // Find primary guest email for confirmation
        var emailCol = headers.indexOf("Email");
        if (emailCol !== -1 && guestData[rowIndex][emailCol]) {
          primaryGuestEmail = guestData[rowIndex][emailCol];
          primaryGuestName = guestData[rowIndex][fullNameCol];
        }
      }
    });

    SpreadsheetApp.flush(); // Ensure all writes are committed

    // Send confirmation email if enabled and email exists
    var config = getScriptConfig();
    if (config.sendEmails && primaryGuestEmail) {
      sendConfirmationEmail(primaryGuestEmail, primaryGuestName, payload);
    }

    return createJsonResponse({ result: "success" });

  } catch (err) {
    Logger.log("doPost Error: " + err.message + "\n" + err.stack);
    return createJsonResponse({ result: "error", message: "An internal server error occurred." });
  }
}

// =================================================================
// HELPER & UTILITY FUNCTIONS
// =================================================================

/**
 * Converts a Google Sheet to an array of objects.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to convert.
 * @returns {Array<Object>} An array of objects, where each object represents a row.
 */
function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  var headers = data.shift().map(function(header) {
    return String(header || '').replace(/\s+/g, '_'); // Sanitize headers for object keys
  });
  return data.map(function(row, index) {
    var obj = {};
    headers.forEach(function(header, i) {
      obj[header] = row[i];
    });
    obj.rowIndex = index + 1; // Add 0-based row index for easy lookup
    return obj;
  });
}

/**
 * Creates a JSON response for the web app.
 * @param {Object} data The data to be stringified.
 * @returns {GoogleAppsScript.Content.TextOutput} The JSON response.
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Retrieves key-value pairs from the 'Config' sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @returns {Object} An object containing the configuration.
 */
function getGlobalConfig(ss) {
  var configSheet = ss.getSheetByName("Config");
  if (!configSheet) return {};
  var data = configSheet.getDataRange().getValues();
  return data.reduce(function(obj, row) {
    if (row[0]) obj[row[0]] = row[1];
    return obj;
  }, {});
}

/**
 * Sends a confirmation email to the guest after they RSVP.
 * @param {string} email The recipient's email address.
 * @param {string} name The recipient's name.
 * @param {Object} payload The submitted RSVP data.
 */
function sendConfirmationEmail(email, name, payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var globalConfig = getGlobalConfig(ss);
    var eventTitle = globalConfig.event_title || "Your Celebration";
    var emailSignature = globalConfig.email_signature || "The Family";

    var subject = "RSVP Confirmation for " + eventTitle;
    var htmlBody = 
      "<div style='font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;'>" +
        "<h2 style='color: #2c3e50;'>" + (globalConfig.email_salutation || "Hello") + " " + name + ",</h2>" +
        "<p>Thank you for your RSVP for <strong>" + eventTitle + "</strong>. Your response has been recorded.</p>" +
        "<p>If you need to make any changes, please use your original invitation link.</p>" +
        "<br>" +
        "<p>" + emailSignature + "</p>" +
      "</div>";

    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log("Confirmation email sent to " + email);
  } catch (err) {
    Logger.log("Failed to send confirmation email to " + email + ". Error: " + err.message);
  }
}

// =================================================================
// ADMIN PANEL FUNCTIONS
// =================================================================

/**
 * Checks if the provided password matches the one in the 'Config' sheet.
 * @param {string} password The password submitted by the user.
 * @returns {boolean} True if the password is correct.
 */
function checkAdminPassword(password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var globalConfig = getGlobalConfig(ss);
  var correctPassword = globalConfig.admin_password || "";
  // Ensure password is not empty and matches
  return password && correctPassword && password === correctPassword;
}

/**
 * Returns the raw data for the admin dashboard. The frontend will process this.
 * @returns {Object} An object containing raw 'guests' and 'events' data.
 */
function getAdminDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var guestsSheet = ss.getSheetByName("Guests");
  var eventsSheet = ss.getSheetByName("Events");

  if (!guestsSheet) throw new Error("Sheet 'Guests' not found.");
  if (!eventsSheet) throw new Error("Sheet 'Events' not found.");

  var guests = sheetToObjects(guestsSheet);
  var events = sheetToObjects(eventsSheet);

  return {
    guests: guests,
    events: events
  };
}

/**
 * Sends reminder emails to groups who have not yet RSVP'd.
 * @param {Array<string>} [groupIdsToSend] Optional array of Group IDs to send reminders to. If not provided, sends to all non-responders.
 * @returns {string} A summary of the action taken.
 */
function sendReminderEmails(groupIdsToSend) {
  var config = getScriptConfig();
  if (!config.sendEmails) {
    var logMsg = "Email sending is disabled in Script Properties. No reminders sent.";
    Logger.log(logMsg);
    return logMsg;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var globalConfig = getGlobalConfig(ss);
  var eventTitle = globalConfig.event_title || "Your Celebration";
  var emailSignature = globalConfig.email_signature || "The Family";
  var websiteUrl = config.websiteUrl;

  var guestsSheet = ss.getSheetByName("Guests");
  if (!guestsSheet) throw new Error("'Guests' sheet not found.");

  var guestsList = sheetToObjects(guestsSheet);
  
  // Find groups that have not responded at all
  var groupsToRemind = {}; // { groupId: { name: 'Guest Name', email: 'guest@email.com' } }

  guestsList.forEach(function(guest) {
    var groupId = guest.Group_ID ? guest.Group_ID.toString().trim() : "";
    var rsvps = guest.RSVPs ? guest.RSVPs.toString().trim() : "";
    var email = guest.Email ? guest.Email.toString().trim() : "";

    if (groupId && email) {
      // If this group isn't already in our list
      if (!groupsToRemind[groupId]) {
        groupsToRemind[groupId] = {
          name: guest.Full_Name || "Guest",
          email: email,
          hasResponded: false
        };
      }
      // If any guest in the group has an RSVP, mark the group as responded
      if (rsvps && rsvps !== "{}") {
        groupsToRemind[groupId].hasResponded = true;
      }
    }
  });

  var countSent = 0;
  for (var groupId in groupsToRemind) {
    var groupInfo = groupsToRemind[groupId];
    
    // Check if this group should receive a reminder
    var shouldSend = !groupInfo.hasResponded && 
                     (!groupIdsToSend || groupIdsToSend.includes(groupId));

    if (shouldSend) {
      var personalizedUrl = websiteUrl + (websiteUrl.indexOf('?') === -1 ? '?' : '&') + 'id=' + encodeURIComponent(groupId);
      var subject = "Reminder: Please RSVP for " + eventTitle;
      var htmlBody = 
        "<div style='font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;'>" +
          "<h2 style='color: #2c3e50;'>" + (globalConfig.email_salutation || "Hello") + " " + groupInfo.name + ",</h2>" +
          "<p>This is a friendly reminder to RSVP for <strong>" + eventTitle + "</strong>.</p>" +
          "<p>Please click the link below to let us know if your family can make it. We can't wait to celebrate with you!</p>" +
          "<p style='text-align: center; margin: 30px 0;'>" +
            "<a href='" + personalizedUrl + "' style='background-color: #4a90e2; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold;'>View Invitation & RSVP</a>" +
          "</p>" +
          "<br>" +
          "<p>" + emailSignature + "</p>" +
        "</div>";

      MailApp.sendEmail({
        to: groupInfo.email,
        subject: subject,
        htmlBody: htmlBody
      });
      countSent++;
      Logger.log("Sent reminder to " + groupInfo.email + " for Group ID: " + groupId);
      Utilities.sleep(500); // Pause to avoid exceeding email quotas
    }
  }

  return "Successfully sent " + countSent + " reminder email(s).";
}
