/**
 * Dynamic Group-Based Multi-Event RSVP Backend with Email Engine
 * Uses Script Properties for environment configuration.
 */

/**
 * Retrieves environment properties configured under Project Settings -> Script Properties
 */
function getScriptConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    websiteUrl: props.getProperty("WEBSITE_URL") || "https://example.com",
    adminEmail: props.getProperty("ADMIN_EMAIL") || "",
    sendEmails: props.getProperty("SEND_EMAILS") === "true"
  };
}

/**
 * Helper to convert sheet rows into array of objects with 1-based row numbers
 */
function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  var headers = data[0].map(function(h) { return h.toString().trim(); });
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var obj = { _row: i + 1 };
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

/**
 * Loads key-value pairs from the 'Config' sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @returns {Object} An object containing the configuration key-value pairs.
 */
function getGlobalConfig(ss) {
  var configSheet = ss.getSheetByName("Config");
  var globalConfig = {};
  if (configSheet) {
    var configData = configSheet.getDataRange().getValues();
    for (var c = 1; c < configData.length; c++) {
      var key = configData[c][0] ? configData[c][0].toString().trim() : "";
      var val = configData[c][1] ? configData[c][1].toString().trim() : "";
      if (key) globalConfig[key] = val;
    }
  }
  return globalConfig;
}

/**
 * GET Endpoint: Fetch group details + global page config
 * Query Params: ?id=GRP-101 (or ?groupId=GRP-101)
 */
function doGet(e) {
  try {
    var searchGroupId = e.parameter.id || e.parameter.groupId || "";
    searchGroupId = searchGroupId.toString().toLowerCase().trim();
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Load Global Config Tab
    var globalConfig = getGlobalConfig(ss);
    
    // 2. Load Events Tab Configuration
    var eventsSheet = ss.getSheetByName("Events");
    if (!eventsSheet) return createJsonResponse({ error: "Sheet 'Events' not found." });
    
    var eventsList = sheetToObjects(eventsSheet);
    var eventsMap = {};
    eventsList.forEach(function(evt) {
      if (evt.Event_ID) {
        eventsMap[evt.Event_ID.toString().trim()] = {
          id: evt.Event_ID,
          title: evt.Title || "",
          dateTime: evt.Date_Time || "",
          location: evt.Location || "",
          description: evt.Description || ""
        };
      }
    });

    // 3. Load Guests matching Group_ID
    var guestsSheet = ss.getSheetByName("Guests");
    if (!guestsSheet) return createJsonResponse({ error: "Sheet 'Guests' not found." });
    
    var guestsList = sheetToObjects(guestsSheet);
    var groupGuests = [];
    var groupNotes = "";

    for (var i = 0; i < guestsList.length; i++) {
      var guest = guestsList[i];
      var groupId = guest.Group_ID ? guest.Group_ID.toString().toLowerCase().trim() : "";
      
      if (groupId === searchGroupId && searchGroupId !== "") {
        var allowedIds = guest.Allowed_Events ? guest.Allowed_Events.toString().split(",") : [];
        var allowedEvents = [];
        
        allowedIds.forEach(function(rawId) {
          var cleanId = rawId.trim();
          if (eventsMap[cleanId]) {
            allowedEvents.push(eventsMap[cleanId]);
          }
        });

        var existingRsvps = {};
        if (guest.RSVPs) {
          try { existingRsvps = JSON.parse(guest.RSVPs); } catch (err) {}
        }

        if (guest.Notes && !groupNotes) {
          groupNotes = guest.Notes;
        }

        groupGuests.push({
          row: guest._row,
          guestId: guest.Guest_ID,
          fullName: guest.Full_Name || "Guest",
          email: guest.Email || "",
          allowedEvents: allowedEvents,
          existingRsvps: existingRsvps
        });

        // Check for and create a "Plus One" virtual guest if allowed
        var plusOneFlag = guest.Plus_One_Allowed ? guest.Plus_One_Allowed.toString().toUpperCase().trim() : "";
        if (plusOneFlag === "TRUE" || plusOneFlag === "YES" || plusOneFlag === "1") {
          var plusOneGuestId = guest.Guest_ID + "_plusone";
          var plusOneRsvps = {};
          try {
            // Attempt to parse existing RSVPs for the plus one, if they exist
            if (guest.RSVPs) plusOneRsvps = JSON.parse(guest.RSVPs)[plusOneGuestId] || {};
          } catch(err) {}

          groupGuests.push({
            row: guest._row, // The plus one's data is tied to the primary guest's row
            guestId: plusOneGuestId,
            fullName: guest.Full_Name + "'s Guest",
            isPlusOne: true,
            allowedEvents: allowedEvents,
            existingRsvps: plusOneRsvps
          });
        }
      }
    }

    if (groupGuests.length > 0) {
      return createJsonResponse({
        found: true,
        groupId: searchGroupId,
        config: globalConfig,
        eventsMap: eventsMap,
        guests: groupGuests,
        groupNotes: groupNotes
      });
    }

    return createJsonResponse({ found: false, config: globalConfig });

  } catch (err) {
    return createJsonResponse({ error: err.toString() });
  }
}

/**
 * POST Endpoint: Updates multiple guest rows in a single batch
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var guestsSheet = ss.getSheetByName("Guests");
    
    if (!payload.guestResponses || !Array.isArray(payload.guestResponses)) {
      return createJsonResponse({ result: "error", message: "Invalid payload format." });
    }

    var headers = guestsSheet.getRange(1, 1, 1, guestsSheet.getLastColumn()).getValues()[0];
    var rsvpsColIndex = -1;
    var notesColIndex = -1;
    var nameColIndex = -1;
    var emailColIndex = -1;

    for (var j = 0; j < headers.length; j++) {
      var headerName = headers[j].toString().trim();
      if (headerName === "RSVPs") rsvpsColIndex = j + 1;
      if (headerName === "Notes") notesColIndex = j + 1;
      if (headerName === "Full_Name") nameColIndex = j + 1;
      if (headerName === "Email") emailColIndex = j + 1;
    }

    var primaryGuestName = "";
    var primaryGuestEmail = "";

    // Save responses for each family member row
    payload.guestResponses.forEach(function(resp, index) {
      if (rsvpsColIndex > 0 && resp.rsvps) {
        // For plus ones, we merge their RSVP into the primary guest's JSON object
        if (!resp.isPlusOne) {
          var fullRsvpObject = resp.rsvps;
          // Find any plus one associated with this guest and merge their rsvp
          var plusOneResp = payload.guestResponses.find(function(r) { return r.isPlusOne && r.guestId.startsWith(resp.guestId); });
          if (plusOneResp) {
            fullRsvpObject[plusOneResp.guestId] = plusOneResp.rsvps;
          }
          guestsSheet.getRange(resp.row, rsvpsColIndex).setValue(JSON.stringify(fullRsvpObject));
        }
      }
      if (notesColIndex > 0 && index === 0) {
        var finalNotes = payload.notes || "";
        if (payload.plusOneName) {
          finalNotes += (finalNotes ? "\n" : "") + "Plus One Guest Name: " + payload.plusOneName;
        }
        guestsSheet.getRange(resp.row, notesColIndex).setValue(finalNotes);
      }

      if (index === 0) {
        if (nameColIndex > 0) primaryGuestName = guestsSheet.getRange(resp.row, nameColIndex).getValue();
        if (emailColIndex > 0) primaryGuestEmail = guestsSheet.getRange(resp.row, emailColIndex).getValue();
      }
    });

    var config = getScriptConfig();
    var eventsSheet = ss.getSheetByName("Events");
    var eventsList = sheetToObjects(eventsSheet);

    sendConfirmationEmail(primaryGuestName, primaryGuestEmail, payload.guestResponses, payload.notes, eventsList, config.adminEmail, ss, payload.plusOneName);

    return createJsonResponse({ result: "success" });

  } catch (err) {
    return createJsonResponse({ result: "error", message: err.toString() });
  }
}

/**
 * Send Email Invitations to EVERY guest row with a defined Email.
 * Generates personalized ?id=GRP-101 links.
 */
function sendInvitations() {
  var config = getScriptConfig();
  var websiteUrl = config.websiteUrl;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Load Global Config Tab for event title
  var globalConfig = getGlobalConfig(ss);
  var eventTitle = globalConfig.event_title || "Your Celebration";
  var emailSignature = globalConfig.email_signature || "The Family";
  var guestsSheet = ss.getSheetByName("Guests");
  if (!guestsSheet) {
    Logger.log("Error: 'Guests' sheet missing.");
    return;
  }

  var guestsData = guestsSheet.getDataRange().getValues();
  if (guestsData.length < 2) {
    Logger.log("No guest data found.");
    return;
  }

  var headers = guestsData[0].map(function(h) { return h.toString().trim(); });
  
  var groupCol = headers.indexOf("Group_ID");
  var nameCol = headers.indexOf("Full_Name");
  var emailCol = headers.indexOf("Email");
  var sentCol = headers.indexOf("Invite_Sent");

  if (groupCol === -1 || emailCol === -1 || nameCol === -1) {
    Logger.log("Error: 'Group_ID', 'Full_Name', or 'Email' column missing.");
    return;
  }

  if (sentCol === -1) {
    sentCol = headers.length;
    guestsSheet.getRange(1, sentCol + 1).setValue("Invite_Sent");
  }

  var countSent = 0;

  for (var i = 1; i < guestsData.length; i++) {
    var row = guestsData[i];
    var groupId = row[groupCol] ? row[groupCol].toString().trim() : "";
    var guestName = row[nameCol] ? row[nameCol].toString().trim() : "Guest";
    var guestEmail = row[emailCol] ? row[emailCol].toString().trim() : "";
    var isSent = (sentCol < row.length) ? row[sentCol].toString().trim().toLowerCase() : "";

    if (groupId !== "" && guestEmail !== "" && isSent !== "yes" && isSent !== "invited") {
      
      var personalizedUrl = websiteUrl + "?id=" + encodeURIComponent(groupId);
      var subject = "You're Invited: " + eventTitle + "!";
      
      if (!config.sendEmails) {
        Logger.log("Email sending disabled. Skipping invitation for " + guestEmail + " (Group ID: " + groupId + ")");
        continue; // Skip sending email for this guest
      }

      var htmlBody = 
        "<div style='font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;'>" +
          "<h2 style='color: #2c3e50;'>" + (globalConfig.email_salutation || "Hello") + " " + guestName + ",</h2>" +
          "<p>You are invited to <strong>" + eventTitle + "</strong>!</p>" +
          "<p>Please click below to view your family's dynamic invitation and let us know if you can make it:</p>" +
          "<p style='text-align: center; margin: 30px 0;'>" +
            "<a href='" + personalizedUrl + "' style='background-color: #4a90e2; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold;'>View Invitation & RSVP</a>" +
          "</p>" +
          "<p style='font-size: 0.85em; color: #777; text-align: center;'>" +
            "Group Code: <strong>" + groupId + "</strong>" +
          "</p>" +
          "<br>" +
          "<p>" + emailSignature + "</p>" +
        "</div>";

      MailApp.sendEmail({
        to: guestEmail,
        subject: subject,
        htmlBody: htmlBody
      });
      
      Logger.log("Sent invitation to " + guestEmail + " for Group ID: " + groupId);

      guestsSheet.getRange(i + 1, sentCol + 1).setValue("Yes");
      countSent++;
      Utilities.sleep(500); 
    }
  }

  Logger.log("Successfully sent " + countSent + " individual invitation email(s).");
}

/**
 * Sends event-first HTML confirmation emails to guests, 
 * and a separate concise summary email to the admin.
 */
function sendConfirmationEmail(primaryName, recipientEmail, guestResponses, notes, eventsList, adminEmail, ss, plusOneName) {
  var config = getScriptConfig();

  // Load Global Config Tab for event title
  var globalConfig = getGlobalConfig(ss);
  var eventTitle = globalConfig.event_title || "Your Celebration";
  var eventsMap = {};
  eventsList.forEach(function(evt) { 
    eventsMap[evt.Event_ID] = evt; 
  });

  // Map out attendance per sub-event across all family members
  var eventAttendance = {}; // { evtId: { attending: [names], declined: [names] } }

  guestResponses.forEach(function(g) {
    var name = g.fullName;
    // If it's a plus one and a name was provided, use that name.
    if (g.isPlusOne && plusOneName) {
      name = plusOneName;
    }

    for (var evtId in g.rsvps) {
      if (!eventAttendance[evtId]) {
        eventAttendance[evtId] = { attending: [], declined: [] };
      }
      if (g.rsvps[evtId] === "Attending") {
        eventAttendance[evtId].attending.push(name);
      } else {
        eventAttendance[evtId].declined.push(name);
      }
    }
  });

  // -------------------------------------------------------------
  // 1. BUILD GUEST EMAIL (Event-First Layout)
  // -------------------------------------------------------------
  var guestSummaryHtml = "<div style='line-height: 1.6; font-family: Arial, sans-serif;'>";

  for (var evtId in eventAttendance) {
    var evt = eventsMap[evtId] || { Title: evtId };
    var attendees = eventAttendance[evtId].attending;

    if (attendees.length > 0) {
      // Generate Google Calendar date parameter if possible
      var calendarDates = "";
      if (evt.Date_Time) {
        var eventDate = new Date(evt.Date_Time);
        if (!isNaN(eventDate.getTime())) {
          var duration = parseInt(evt.Duration_Minutes, 10) || 60; // Default to 60 mins if not specified
          calendarDates = "&dates=" + formatToGoogleCalendarDate(eventDate, duration);
        }
      }

      // Dynamic Google Calendar link
      var gCalUrl = "https://calendar.google.com/calendar/render?action=TEMPLATE" +
        "&text=" + encodeURIComponent(evt.Title || "Event") +
        "&location=" + encodeURIComponent(evt.Location || "") +
        calendarDates +
        "&details=" + encodeURIComponent((evt.Description || "") + "\n\nAttending: " + formatNameList(attendees.slice()) + "\n\nWe look forward to celebrating with you!").replace(/'/g, '%27');

      guestSummaryHtml += "<div style='margin-bottom: 20px; padding: 16px; background-color: #f8f9fa; border-left: 4px solid #2e7d32; border-radius: 6px;'>" +
        "<h3 style='margin: 0 0 6px 0; color: #2e7d32; font-size: 1.15em;'>✓ " + (evt.Title || evtId) + "</h3>" +
        (evt.Date_Time ? "<div style='color: #444;'>📅 <strong>When:</strong> " + evt.Date_Time + "</div>" : "") +
        (evt.Location ? "<div style='color: #444;'>📍 <strong>Where:</strong> " + evt.Location + "</div>" : "") +
        (evt.Description ? "<div style='color: #666; font-style: italic; margin-top: 6px; font-size: 0.95em;'>" + evt.Description + "</div>" : "") +
        "<div style='margin-top: 10px; padding-top: 10px; border-top: 1px dashed #ddd;'>" +
          "<strong>Attending (" + attendees.length + "):</strong> " + formatNameList(attendees.slice()) +
        "</div>" +
        "<div style='margin-top: 10px;'>" +
          "<a href='" + gCalUrl + "' target='_blank' style='display: inline-block; background-color: #4a90e2; color: #ffffff; padding: 6px 12px; text-decoration: none; border-radius: 4px; font-size: 0.85em; font-weight: bold;'>📅 Add to Google Calendar</a>" +
        "</div>" +
      "</div>";
    } else {
      guestSummaryHtml += "<div style='margin-bottom: 12px; padding: 12px; background-color: #fafafa; border-left: 4px solid #c62828; border-radius: 4px; color: #777;'>" +
        "<strong>✕ " + (evt.Title || evtId) + ":</strong> Not Attending" +
      "</div>";
    }
  }
  guestSummaryHtml += "</div>";

  var guestBodyHtml = 
    "<div style='max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; font-family: Arial, sans-serif;'>" +
      "<h2 style='color: #2c3e50; text-align: center;'>" + (globalConfig.email_salutation || "Hello") + " " + primaryName + ",</h2>" +
      "<p style='text-align: center; color: #555;'>Thank you for submitting your RSVP! Here is your confirmed event schedule:</p>" +
      "<hr style='border: none; border-top: 1px solid #eee; margin: 20px 0;'>" +
      guestSummaryHtml +
      (notes ? "<div style='margin-top: 20px; padding: 12px; background-color: #fffde7; border: 1px solid #fff59d; border-radius: 4px;'><strong>Notes / Dietary Restrictions:</strong><br>" + notes.replace(/\n/g, '<br>') + "</div>" : "") +
      (plusOneName ? "<div style='margin-top: 10px; padding: 12px; background-color: #e3f2fd; border: 1px solid #bbdefb; border-radius: 4px;'><strong>Plus One Guest:</strong> " + plusOneName + "</div>" : "") +
      "<br><p style='text-align: center; color: #555;'>We look forward to celebrating with you!</p>" +
    "</div>";

  // Send email to Guest
  if (recipientEmail) {
    if (config.sendEmails) {
      MailApp.sendEmail({ 
        to: recipientEmail, 
        subject: "RSVP Confirmation - " + eventTitle, 
        htmlBody: guestBodyHtml 
      });
      Logger.log("Sent confirmation email to guest: " + recipientEmail);
    } else {
      Logger.log("Email sending disabled. Skipping confirmation email for guest: " + recipientEmail);
    }
  }

  // -------------------------------------------------------------
  // 2. BUILD DEDICATED ADMIN EMAIL
  // -------------------------------------------------------------
  if (adminEmail) {
    var adminSummaryHtml = "<div style='line-height: 1.6; font-family: Arial, sans-serif;'>";

    for (var evtIdAdmin in eventAttendance) {
      var evtObj = eventsMap[evtIdAdmin] || { Title: evtIdAdmin };
      var attList = eventAttendance[evtIdAdmin].attending;
      var decList = eventAttendance[evtIdAdmin].declined;

      adminSummaryHtml += "<div style='margin-bottom: 15px; padding: 12px; background-color: #f8f9fa; border: 1px solid #e2e8f0; border-radius: 6px;'>" +
        "<h4 style='margin: 0 0 6px 0; color: #2c3e50;'>" + (evtObj.Title || evtIdAdmin) + "</h4>" +
        "<div><strong style='color: #2e7d32;'>Attending (" + attList.length + "):</strong> " + (attList.length > 0 ? formatNameList(attList.slice()) : "<em>None</em>") + "</div>" +
        "<div><strong style='color: #c62828;'>Declined (" + decList.length + "):</strong> " + (decList.length > 0 ? formatNameList(decList.slice()) : "<em>None</em>") + "</div>" +
      "</div>";
    }
    adminSummaryHtml += "</div>";

    var adminBodyHtml = 
      "<div style='max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; font-family: Arial, sans-serif;'>" +
        "<h3 style='color: #2c3e50; margin-top: 0;'>RSVP Submission: " + primaryName + "</h3>" +
        "<p><strong>Primary Email:</strong> " + (recipientEmail || "N/A") + "</p>" +
        "<hr style='border: none; border-top: 1px solid #eee; margin: 15px 0;'>" +
        "<h4>Event Breakdown:</h4>" +
        adminSummaryHtml +
        (notes ? "<div style='margin-top: 15px; padding: 10px; background-color: #fffde7; border: 1px solid #fff59d; border-radius: 4px;'><strong>Dietary / Notes:</strong><br>" + notes.replace(/\n/g, '<br>') + "</div>" : "") +
        (plusOneName ? "<div style='margin-top: 10px; padding: 12px; background-color: #e3f2fd; border: 1px solid #bbdefb; border-radius: 4px;'><strong>Plus One Guest:</strong> " + plusOneName + "</div>" : "") +
      "</div>";

    if (config.sendEmails) {
      MailApp.sendEmail({ 
        to: adminEmail, 
        subject: "[Admin Notification] RSVP Update: " + primaryName, 
        htmlBody: adminBodyHtml 
      });
      Logger.log("Sent admin notification email to: " + adminEmail);
    } else {
      Logger.log("Email sending disabled. Skipping admin notification email for: " + adminEmail);
    }
  }
}

/**
 * Formats an array of names into a natural language string.
 * e.g., ["A", "B", "C"] -> "A, B, and C"
 * @param {string[]} names The array of names.
 * @returns {string} The formatted string.
 */
function formatNameList(names) {
  if (!names || names.length === 0) return "";
  if (names.length === 1) return names[0];
  if (names.length === 2) return names.join(" and ");
  var last = names.pop();
  return names.join(", ") + ", and " + last;
}

/**
 * Helper to format a Date object into a Google Calendar-compatible date string.
 * @param {Date} date The start date and time of the event.
 * @param {number} durationMinutes The duration of the event in minutes.
 * @returns {string} A formatted string like '20240525T100000/20240525T110000'.
 */
function formatToGoogleCalendarDate(date, durationMinutes) {
  var pad = function(num) { return num < 10 ? '0' + num : '' + num; };
  
  var startDate = date;
  var durationMs = (durationMinutes || 60) * 60 * 1000; // Default to 60 minutes in milliseconds
  var endDate = new Date(startDate.getTime() + durationMs);

  var formatDate = function(d) {
    return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) +
           'T' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
  };

  return formatDate(startDate) + '/' + formatDate(endDate);
}

/**
 * Helper to build JSON HTTP responses
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function EVENT_SUMMARY() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var guestsSheet = ss.getSheetByName("Guests");
  var eventsSheet = ss.getSheetByName("Events");
  
  // 1. Load dynamic event definitions from the "Events" tab
  var eventData = eventsSheet.getDataRange().getValues();
  var eventHeaders = eventData[0];
  var eventIdCol = eventHeaders.indexOf("Event_ID");
  var titleCol = eventHeaders.indexOf("Title");
  
  var events = {};
  // Loop through events tab (skipping header)
  for (var e = 1; e < eventData.length; e++) {
    var evId = String(eventData[e][eventIdCol] || "").trim();
    var evTitle = String(eventData[e][titleCol] || "").trim();
    if (evId) {
      events[evId] = {
        name: evTitle || evId,
        invited: 0,
        attending: 0,
        noResponse: 0
      };
    }
  }
  
  // 2. Load guest data from the "Guests" tab
  var guestData = guestsSheet.getDataRange().getValues();
  var guestHeaders = guestData[0];
  var allowedCol = guestHeaders.indexOf("Allowed_Events");
  var rsvpCol = guestHeaders.indexOf("RSVPs");
  
  // 3. Loop through all guest rows to calculate counts
  for (var i = 1; i < guestData.length; i++) {
    var allowedStr = String(guestData[i][allowedCol] || "");
    var rsvpStr = String(guestData[i][rsvpCol] || "{}");
    
    var rsvpJson = {};
    try {
      rsvpJson = JSON.parse(rsvpStr);
    } catch (err) {
      rsvpJson = {};
    }
    
    // Check each dynamically loaded event
    for (var eventKey in events) {
      // Check if the guest is allowed/invited to this event
      if (allowedStr.indexOf(eventKey) !== -1) {
        events[eventKey].invited++;
        
        var status = rsvpJson[eventKey];
        if (status === "Attending") {
          events[eventKey].attending++;
        } else {
          events[eventKey].noResponse++;
        }
      }
    }
  }
  
  // 4. Format the output table
  var output = [["Event", "Invited", "Attending", "Not Responded"]];
  for (var key in events) {
    output.push([
      events[key].name,
      events[key].invited,
      events[key].attending,
      events[key].noResponse
    ]);
  }
  
  return output;
}