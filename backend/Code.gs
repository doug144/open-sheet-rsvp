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
      
      if (payload.action === 'sendInvitations') {
        return createJsonResponse({ status: 'success', data: { message: sendInvitations(payload.groupIds) } });
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

    var globalConfig = getGlobalConfig(ss);
    // Use the configured contact email, falling back to the script owner if it's missing
    var adminEmail = globalConfig.contact_email || Session.getEffectiveUser().getEmail(); 
    sendAdminNotification(adminEmail, primaryGuestName, payload, ss, config.websiteUrl);

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
 * Sends a confirmation email to the guest after they RSVP, including a summary of their response.
 * @param {string} email The recipient's email address.
 * @param {string} name The recipient's name.
 * @param {Object} payload The submitted RSVP data.
 */
function sendConfirmationEmail(email, name, payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var globalConfig = getGlobalConfig(ss);
    var config = getScriptConfig(); 
    
    var eventTitle = globalConfig.event_title || "Your Celebration";
    var emailSignature = globalConfig.email_signature || "The Family";
    var subject = "RSVP Confirmation for " + eventTitle;
    
    // Construct the personalized URL
    var websiteUrl = config.websiteUrl;
    var personalizedUrl = websiteUrl + (websiteUrl.indexOf('?') === -1 ? '?' : '&') + 'id=' + encodeURIComponent(payload.groupId);

    // 1. Fetch Event Titles for the summary
    var eventsSheet = ss.getSheetByName("Events");
    var eventMap = {};
    if (eventsSheet) {
      var eventsData = eventsSheet.getDataRange().getValues();
      var headers = eventsData.shift();
      var idCol = headers.indexOf("Event_ID");
      var titleCol = headers.indexOf("Title");
      
      if (idCol !== -1 && titleCol !== -1) {
        eventsData.forEach(function(row) {
          eventMap[row[idCol]] = row[titleCol];
        });
      }
    }

    // 2. Fetch Guests data to map row indexes back to their names
    var guestsSheet = ss.getSheetByName("Guests");
    var guestsData = guestsSheet ? guestsSheet.getDataRange().getValues() : [];
    var fullNameCol = guestsData.length > 0 ? guestsData[0].indexOf("Full_Name") : -1;

    // 3. Group attendance by event to avoid duplicate event headers for large families
    var eventBreakdown = {};
    
    payload.guestResponses.forEach(function(response) {
       var primaryName = "Guest";
       // Retrieve the primary guest's real name using their sheet row index
       if (fullNameCol !== -1 && response.row > 0 && response.row < guestsData.length) {
         primaryName = guestsData[response.row][fullNameCol] || "Guest";
       }

       for (var eventId in response.rsvps) {
         if (!eventBreakdown[eventId]) {
           eventBreakdown[eventId] = [];
         }
         
         var rsvpData = response.rsvps[eventId];
         
         // Add primary guest
         eventBreakdown[eventId].push({
           name: primaryName,
           status: rsvpData.status
         });
         
         // Add plus ones
         if (rsvpData.additionalGuests && rsvpData.additionalGuests.length > 0) {
           rsvpData.additionalGuests.forEach(function(guest) {
             eventBreakdown[eventId].push({
               name: guest.name || "Guest",
               status: guest.status
             });
           });
         }
       }
    });

    // 4. Build the Summary HTML
    var summaryHtml = "<div style='background-color: #f8f9fa; padding: 15px 20px; border-radius: 6px; margin: 20px 0;'>";
    summaryHtml += "<h3 style='margin-top: 0; color: #2c3e50; font-size: 1.1em;'>Your Response Summary:</h3>";
    
    for (var eventId in eventBreakdown) {
       var eTitle = eventMap[eventId] || eventId;
       
       summaryHtml += "<div style='margin-bottom: 15px;'>";
       summaryHtml += "<h4 style='margin: 0 0 5px 0; color: #4a90e2; border-bottom: 1px solid #e0e0e0; padding-bottom: 4px;'>" + eTitle + "</h4>";
       summaryHtml += "<ul style='margin: 5px 0 0 0; padding-left: 20px; font-size: 0.95em; color: #333;'>";
       
       eventBreakdown[eventId].forEach(function(person) {
         var statusColor = person.status === "Attending" ? "#28a745" : "#dc3545";
         summaryHtml += "<li style='margin-bottom: 4px;'><strong>" + person.name + ":</strong> <span style='color: " + statusColor + "; font-weight: bold;'>" + person.status + "</span></li>";
       });
       
       summaryHtml += "</ul></div>";
    }
    
    if (payload.notes) {
      summaryHtml += "<p style='margin-top: 15px; padding-top: 10px; border-top: 1px solid #e0e0e0; font-size: 0.9em;'><strong>Notes left:</strong> " + payload.notes + "</p>";
    }
    summaryHtml += "</div>";

    // 5. Construct the final email body
    var htmlBody = 
        "<div style='font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;'>" +
        "<h2 style='color: #2c3e50;'>" + (globalConfig.email_salutation || "Hello") + " " + name + ",</h2>" +
        "<p>Thank you for your RSVP for <strong>" + eventTitle + "</strong>. Your response has been securely recorded.</p>" +
        summaryHtml + 
        "<p>If you need to make any changes, you can <a href='" + personalizedUrl + "' style='color: #4a90e2; font-weight: bold;'>click here to return to your personalized invitation</a>.</p>" +
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

/**
 * Sends an HTML notification email to the admin with full RSVP details grouped by event.
 */
function sendAdminNotification(adminEmail, guestName, payload, ss, websiteUrl) {
  try {
    var subject = "New RSVP Received: " + (guestName || "A Guest");
    
    // 1. Get Event Titles from the spreadsheet for readable emails
    var eventsSheet = ss.getSheetByName("Events");
    var eventMap = {};
    if (eventsSheet) {
      var eventsData = eventsSheet.getDataRange().getValues();
      var headers = eventsData.shift();
      var idCol = headers.indexOf("Event_ID");
      var titleCol = headers.indexOf("Title");
      
      if (idCol !== -1 && titleCol !== -1) {
        eventsData.forEach(function(row) {
          eventMap[row[idCol]] = row[titleCol];
        });
      }
    }

    // 2. Fetch Guests data to map row indexes back to their names
    var guestsSheet = ss.getSheetByName("Guests");
    var guestsData = guestsSheet ? guestsSheet.getDataRange().getValues() : [];
    var fullNameCol = guestsData.length > 0 ? guestsData[0].indexOf("Full_Name") : -1;

    // 3. Group attendance by event
    var eventBreakdown = {};
    
    payload.guestResponses.forEach(function(response) {
       var primaryName = "Guest";
       // Retrieve the primary guest's real name using their sheet row index
       if (fullNameCol !== -1 && response.row > 0 && response.row < guestsData.length) {
         primaryName = guestsData[response.row][fullNameCol] || "Guest";
       }

       for (var eventId in response.rsvps) {
         if (!eventBreakdown[eventId]) {
           eventBreakdown[eventId] = [];
         }
         
         var rsvpData = response.rsvps[eventId];
         
         // Add primary guest
         eventBreakdown[eventId].push({
           name: primaryName,
           status: rsvpData.status
         });
         
         // Add plus ones
         if (rsvpData.additionalGuests && rsvpData.additionalGuests.length > 0) {
           rsvpData.additionalGuests.forEach(function(guest) {
             eventBreakdown[eventId].push({
               name: guest.name || "Guest",
               status: guest.status
             });
           });
         }
       }
    });

    // 4. Build the Email HTML Body
    var htmlBody = "<div style='font-family: Arial, sans-serif; max-width: 600px; color: #333;'>";
    htmlBody += "<h2 style='color: #2c3e50;'>New RSVP Submission</h2>";
    htmlBody += "<p><strong>" + (guestName || "A guest") + "</strong> has just submitted an RSVP.</p>";
    htmlBody += "<hr style='border: none; border-top: 1px solid #eee; margin: 20px 0;'>";
    htmlBody += "<h3 style='color: #4a90e2;'>Attendance Details:</h3>";
    
    for (var eventId in eventBreakdown) {
       var eTitle = eventMap[eventId] || eventId;
       
       htmlBody += "<div style='margin-bottom: 15px;'>";
       htmlBody += "<h4 style='margin: 0 0 5px 0; color: #2c3e50; border-bottom: 1px solid #e0e0e0; padding-bottom: 4px;'>" + eTitle + "</h4>";
       htmlBody += "<ul style='margin: 5px 0 0 0; padding-left: 20px; font-size: 0.95em;'>";
       
       eventBreakdown[eventId].forEach(function(person) {
         var statusColor = person.status === "Attending" ? "#28a745" : "#dc3545";
         htmlBody += "<li style='margin-bottom: 4px;'><strong>" + person.name + ":</strong> <span style='color: " + statusColor + "; font-weight: bold;'>" + person.status + "</span></li>";
       });
       
       htmlBody += "</ul></div>";
    }

    // Append notes if they exist
    if (payload.notes) {
      htmlBody += "<h3 style='color: #4a90e2; margin-top: 20px;'>Guest Notes/Dietary Restrictions:</h3>";
      htmlBody += "<p style='background-color: #f8f9fa; padding: 12px; border-left: 4px solid #f39c12; border-radius: 4px;'>" + payload.notes + "</p>";
    }
    
    // 5. Admin Portal Link
    var adminUrl = websiteUrl ? websiteUrl.replace("index.html", "admin.html") : "YOUR_ADMIN_URL_HERE";
    if (adminUrl.indexOf("admin.html") === -1 && adminUrl !== "YOUR_ADMIN_URL_HERE") {
       adminUrl = adminUrl.substring(0, adminUrl.lastIndexOf('/')) + "/admin.html";
    }
    
    htmlBody += "<p style='margin-top: 30px; text-align: center;'>";
    htmlBody += "<a href='" + adminUrl + "' style='background-color: #4a90e2; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;'>Open Admin Panel</a>";
    htmlBody += "</p></div>";
    
    MailApp.sendEmail({
      to: adminEmail,
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log("Detailed admin notification sent to " + adminEmail);
    
  } catch (err) {
    Logger.log("Failed to send detailed admin notification. Error: " + err.message);
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
 * Sends initial invitation emails to groups.
 * @param {Array<string>} [groupIdsToSend] Optional array of Group IDs to send invites to. If not provided, sends to all groups.
 * @returns {string} A summary of the action taken.
 */
function sendInvitations(groupIdsToSend) {
  var config = getScriptConfig();
  if (!config.sendEmails) {
    var logMsg = "Email sending is disabled in Script Properties. No invitations sent.";
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
  var headers = guestsSheet.getDataRange().getValues()[0];
  var inviteSentCol = headers.indexOf("Invite_Sent");

  // Group guests and collect ALL names, emails, and sheet rows
  var groupsToInvite = {}; 
  guestsList.forEach(function(guest) {
    var groupId = guest.Group_ID ? guest.Group_ID.toString().trim() : "";
    if (!groupIdsToSend || !groupIdsToSend.includes(groupId)) return;
    
    var email = guest.Email ? guest.Email.toString().trim() : "";
    var fullName = guest.Full_Name ? guest.Full_Name.toString().trim() : "Guest";
    
    if (groupId) {
      if (!groupsToInvite[groupId]) {
        groupsToInvite[groupId] = {
          names: [],
          emails: [],
          rows: []
        };
      }
      
      // Add every guest's name to the list
      groupsToInvite[groupId].names.push(fullName);
      
      // Add email to the list if it exists and isn't a duplicate
      if (email && groupsToInvite[groupId].emails.indexOf(email) === -1) {
        groupsToInvite[groupId].emails.push(email);
      }
      
      // Store the actual spreadsheet row
      groupsToInvite[groupId].rows.push(parseInt(guest.rowIndex) + 1);
    }
  });

  var countSent = 0;
  var countFailed = 0;
  var errors = "";
  for (var groupId in groupsToInvite) {
    var groupInfo = groupsToInvite[groupId];

    try {
      // Only send if the group has at least one email address
      if (groupInfo.emails.length > 0) {
        var personalizedUrl = websiteUrl + (websiteUrl.indexOf('?') === -1 ? '?' : '&') + 'id=' + encodeURIComponent(groupId);
        var subject = "You're Invited: " + eventTitle;
        
        // Join multiple emails with a comma
        var toEmails = groupInfo.emails.join(",");
        
        // Format the names into a natural list (e.g. "A", "A and B", or "A, B, and C")
        var formattedNames = "";
        var namesList = groupInfo.names.slice(); // Copy the array so we can manipulate it
        if (namesList.length === 1) {
          formattedNames = namesList[0];
        } else if (namesList.length === 2) {
          formattedNames = namesList.join(" and ");
        } else if (namesList.length > 2) {
          var last = namesList.pop();
          formattedNames = namesList.join(", ") + ", and " + last;
        } else {
          formattedNames = "Guest";
        }
        
        var htmlBody = 
            "<div style='font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;'>" +
            "<h2 style='color: #2c3e50;'>" + (globalConfig.email_salutation || "Hello") + " " + formattedNames + ",</h2>" +
            "<p>You are warmly invited to <strong>" + eventTitle + "</strong>!</p>" +
            "<p>We would love for you to join us. Please click the link below to view the event details and let us know if you can make it.</p>" +
            "<p style='text-align: center; margin: 30px 0;'>" +
              "<a href='" + personalizedUrl + "' style='background-color: #4a90e2; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;'>View Invitation & RSVP</a>" +
            "</p>" +
            "<br>" +
            "<p>" + emailSignature + "</p>" +
          "</div>";

        MailApp.sendEmail({
          to: toEmails,
          subject: subject,
          htmlBody: htmlBody
        });
        
        countSent++;
        Logger.log("Sent invitation to " + toEmails + " for Group ID: " + groupId);
        
        // Update the "Invite_Sent" column to "Yes" for EVERY member in the group
        if (inviteSentCol !== -1) {
          groupInfo.rows.forEach(function(rowNum) {
            guestsSheet.getRange(rowNum, inviteSentCol + 1).setValue("Yes");
          });
        }
        
        Utilities.sleep(500); 
      }
    } catch (e) {
      countFailed++;
      const errMsg = "Faild to send invitation to Group ID: " + groupId + ". Error: " + e.message;
      Logger.log(errMsg);
      errors += "\n" + errMsg;
    }
  }

  SpreadsheetApp.flush();
  var successMsg = "Successfully sent " + countSent + " invitation email(s).";
  return errors == "" ? successMsg : successMsg + errors;
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
