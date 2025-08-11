/**
 * @OnlyCurrentDoc
 * The above comment directs Apps Script to limit the scope of file access for this add-on. It specifies that this add-on will only have access to the current document.
 */

// =================================================================
// SECTION 1: ADD-ON USER INTERFACE (UI)
// =================================================================

/**
 * This is the main entry point for the add-on, which runs when the user opens it.
 * It reads the current state and then calls the UI builder.
 * @param {Object} e The event object.
 * @return {Card}
 */
function onHomepage(e) {
  const userProperties = PropertiesService.getUserProperties();
  // We determine the "running" state based on the existence of the briefing trigger
  const triggerId = userProperties.getProperty('briefingTriggerId');
  const savedEmail = userProperties.getProperty('recipientEmail') || '';
  const savedFrequency = userProperties.getProperty('frequencyHours') || '12';
  
  let isRunning = false;
  if (triggerId) {
    const allTriggers = ScriptApp.getProjectTriggers();
    isRunning = allTriggers.some(t => t.getUniqueId() === triggerId);
    if (!isRunning) {
      // If the main trigger is gone, something is wrong. Clean up all related triggers.
      deleteManagedTriggers();
    }
  }
  
  // Pass the current state to the UI builder function
  return buildHomepageCard(isRunning, savedEmail, savedFrequency);
}


/**
 * This is a "pure" UI builder function that displays the interface based on the passed parameters.
 * @param {boolean} isRunning Whether the service is running.
 * @param {string} email The currently set email address.
 * @param {string} frequency The currently set frequency.
 * @return {Card}
 */
function buildHomepageCard(isRunning, email, frequency) {
  let statusMessage = "Service is currently stopped.";
  if (isRunning) {
    statusMessage = `Service is running. Briefings are updated approximately every ${frequency} hours. The email fetch feature is also active. Will send to: ${email}`;
  }

  const statusWidget = CardService.newTextParagraph().setText(statusMessage);

  const emailInput = CardService.newTextInput()
    .setFieldName("recipient_email_input")
    .setTitle("Email address for briefings and fetched emails")
    .setValue(email);

  const frequencyDropdown = CardService.newSelectionInput()
    .setFieldName("frequency_input")
    .setTitle("Briefing update frequency")
    .setType(CardService.SelectionInputType.DROPDOWN)
    .addItem("Every hour", "1", frequency === "1")
    .addItem("Every 3 hours", "3", frequency === "3")
    .addItem("Every 6 hours", "6", frequency === "6")
    .addItem("Every 12 hours (default)", "12", frequency === "12")
    .addItem("Every 24 hours", "24", frequency === "24");

  const startButton = CardService.newTextButton()
    .setText("Start / Update Service")
    .setOnClickAction(CardService.newAction().setFunctionName("handleStartService"));

  const stopButton = CardService.newTextButton()
    .setText("Stop Service")
    .setOnClickAction(CardService.newAction().setFunctionName("handleStopService"));
    
  const runNowButton = CardService.newTextButton()
    .setText("Run Briefing Manually Now")
    .setOnClickAction(CardService.newAction().setFunctionName("handleRunNow"));

  const cardSection = CardService.newCardSection()
    .setHeader("Status & Settings")
    .addWidget(statusWidget)
    .addWidget(emailInput)
    .addWidget(frequencyDropdown)
    .addWidget(startButton)
    .addWidget(stopButton)
    .addWidget(runNowButton);

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("AI Assistant Settings"))
    .addSection(cardSection)
    .build();

  return card;
}


// =================================================================
// SECTION 2: USER ACTIONS & TRIGGER MANAGEMENT
// =================================================================

/**
 * Handle "Run Now" button click for briefing. 
 */
function handleRunNow(e) {
  const userProperties = PropertiesService.getUserProperties();
  const recipientEmail = userProperties.getProperty('recipientEmail');

  if (!recipientEmail) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Please set up and start the service first."))
      .build();
  }
  
  forwardAllEmails();
  
  return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Manual trigger successful! The briefing has been sent to your email."))
      .build();
}

/**
 * Handles the click event for the "Start/Update Service" button.
 */
function handleStartService(e) {
  const userProperties = PropertiesService.getUserProperties();
  const recipientEmail = e.formInput.recipient_email_input;
  const frequencyHours = parseInt(e.formInput.frequency_input, 10);

  if (!recipientEmail || !recipientEmail.includes('@')) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Please enter a valid email address."))
      .build();
  }

  userProperties.setProperty('recipientEmail', recipientEmail);
  userProperties.setProperty('frequencyHours', String(frequencyHours));
  
  // Delete old managed triggers to create new ones
  deleteManagedTriggers();

  // Add a 10-second delay to allow the platform to process the deletions
  console.log("Waiting for 10 seconds before creating new triggers to avoid race conditions...");
  Utilities.sleep(10000);
  console.log("Wait finished. Creating new triggers.");

  // Create a trigger for the AI briefing
  const briefingTrigger = ScriptApp.newTrigger('forwardAllEmails')
      .timeBased().everyHours(frequencyHours).create();
  userProperties.setProperty('briefingTriggerId', briefingTrigger.getUniqueId());
  console.log(`Created briefing trigger: ${briefingTrigger.getUniqueId()}`);

  // Create a trigger for the email fetch feature (fixed at every 1 hour, the minimum frequency for add-ons)
  const requestTrigger = ScriptApp.newTrigger('processEmailRequest')
      .timeBased().everyHours(1).create();
  userProperties.setProperty('requestTriggerId', requestTrigger.getUniqueId());
  console.log(`Created email fetch trigger: ${requestTrigger.getUniqueId()}`);

  const updatedCard = buildHomepageCard(true, recipientEmail, String(frequencyHours));

  return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(updatedCard))
      .setNotification(CardService.newNotification().setText("Service started! Briefing and email fetch features are now active."))
      .build();
}


/**
 * Handles the click event for the "Stop Service" button.
 */
function handleStopService(e) {
  deleteManagedTriggers();
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
  
  const updatedCard = buildHomepageCard(false, '', '12');

  return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(updatedCard))
      .setNotification(CardService.newNotification().setText("Service has been successfully stopped."))
      .build();
}


/**
 * Helper function: Deletes all triggers managed by this add-on, including leftovers from previous versions.
 */
function deleteManagedTriggers() {
  const userProperties = PropertiesService.getUserProperties();
  const briefingTriggerId = userProperties.getProperty('briefingTriggerId');
  const requestTriggerId = userProperties.getProperty('requestTriggerId');
  const oldTriggerId = userProperties.getProperty('triggerId'); // Check for the old property

  const allTriggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;

  allTriggers.forEach(trigger => {
    const triggerUid = trigger.getUniqueId();
    const handlerFunction = trigger.getHandlerFunction();

    // Delete if the trigger's ID matches one of our stored IDs, or if its handler function is one we manage.
    if (triggerUid === briefingTriggerId || 
        triggerUid === requestTriggerId || 
        triggerUid === oldTriggerId ||
        handlerFunction === 'forwardAllEmails' ||
        handlerFunction === 'processEmailRequest') 
    {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });
  
  if (deletedCount > 0) {
    console.log(`Deleted ${deletedCount} project triggers.`);
  }

  // Clean up all possible trigger properties
  userProperties.deleteProperty('briefingTriggerId');
  userProperties.deleteProperty('requestTriggerId');
  userProperties.deleteProperty('triggerId'); // Delete old property
}


// =================================================================
// SECTION 3: CORE LOGIC - EMAIL PROCESSING & AI
// =================================================================

/**
 * New feature: Fetches and sends the full content of a specified email based on a user's email request.
 * Supports fetching by unique Gmail Thread ID or by a general search query.
 */
function processEmailRequest() {
  const userProperties = PropertiesService.getUserProperties();
  const authorizedEmail = userProperties.getProperty('recipientEmail'); 

  if (!authorizedEmail) {
    console.log("Authorized email not configured, cannot process remote requests.");
    return;
  }

  const searchQuery = `is:inbox is:unread from:(${authorizedEmail}) subject:("Fetch Email Body:")`;
  const threads = GmailApp.search(searchQuery);

  if (threads.length === 0) {
    return;
  }

  console.log(`Found ${threads.length} new email fetch requests.`);

  threads.forEach(thread => {
    const message = thread.getMessages()[0];
    if (message.isUnread()) {
      const requestSubject = message.getSubject();
      const requestContent = requestSubject.replace("Fetch Email Body:", "").trim();

      if (requestContent) {
        let targetThread = null;

        // Priority 1: Check if the request is for a specific Thread ID
        if (requestContent.toLowerCase().startsWith("id:")) {
          const threadId = requestContent.substring(3).trim();
          console.log(`Fetching email by Thread ID: "${threadId}"`);
          try {
            targetThread = GmailApp.getThreadById(threadId);
          } catch (e) {
            console.error(`Error fetching thread by ID: ${e.toString()}`);
            targetThread = null;
          }
        } else {
          // Fallback to general search if not an ID-based request
          console.log(`Executing custom search query: "${requestContent}"`);
          const finalQuery = `${requestContent} -from:me`;
          const targetThreads = GmailApp.search(finalQuery, 0, 1);
          if (targetThreads.length > 0) {
            targetThread = targetThreads[0];
          }
        }

        if (targetThread) {
          const messages = targetThread.getMessages();
          const targetMessage = messages[messages.length - 1]; // Get the last message in the thread
          const originalSubject = targetMessage.getSubject();
          const originalBody = targetMessage.getBody();
          const originalSender = targetMessage.getFrom();

          const replySubject = `Email Body Reply: ${originalSubject}`;
          const replyBody = `
            <div style="font-family: Arial, sans-serif; padding: 20px;">
              <p>Hello, here is the full content of the email you requested:</p>
              <div style="border: 1px solid #ccc; border-radius: 8px; margin-top: 15px; padding: 15px; background-color: #f9f9f9;">
                <p><b>From:</b> ${originalSender}</p>
                <p><b>Subject:</b> ${originalSubject}</p>
                <hr>
                ${originalBody}
              </div>
            </div>
          `;

          try {
            MailApp.sendEmail(authorizedEmail, replySubject, "", { htmlBody: replyBody });
            console.log(`Successfully sent the content of email from request "${requestContent}" to ${authorizedEmail}`);
          } catch (e) {
            console.error(`Failed to send email to ${authorizedEmail}. Error: ${e.toString()}`);
          }

        } else {
          console.log(`No results found for request: "${requestContent}"`);
          try {
            MailApp.sendEmail(authorizedEmail, `Could not find email for request: ${requestContent}`, `Sorry, no email was found in your inbox for the request "${requestContent}". Please try a different query or ID.`);
          } catch (e) {
            console.error(`Failed to send "not found" notification. Error: ${e.toString()}`);
          }
        }
      }
      
      thread.markRead();
    }
  });
}


/**
 * Main function (AI Briefing), called periodically by a trigger.
 */
/**
 * Main function (AI Briefing), called periodically by a trigger.
 */
function forwardAllEmails() {
  const userProperties = PropertiesService.getUserProperties();
  const recipientEmail = userProperties.getProperty('recipientEmail');
  const scriptUserEmail = Session.getActiveUser().getEmail(); // Get the email of the user running the script

  if (!recipientEmail) {
    console.log("Recipient email not found, stopping the briefing function.");
    return;
  }
  
  const quotaLeft = MailApp.getRemainingDailyQuota();
  if (quotaLeft < 1) {
    console.warn("Mail quota is insufficient. The AI briefing function will be skipped for this run.");
    return;
  }
  
  const searchQuery = 'is:inbox is:unread';
  const threads = GmailApp.search(searchQuery);

  if (threads.length === 0) {
    return;
  }

  console.log(`Found ${threads.length} new email threads, processing for briefing...`);

  const senderGroups = new Map();
  const processedThreads = [];

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      if (message.isUnread() && message.getFrom().indexOf(recipientEmail) === -1) {
        const fromString = message.getFrom();
        const senderEmailMatch = fromString.match(/<(.*)>/);
        const senderEmail = senderEmailMatch ? senderEmailMatch[1] : fromString;

        if (!senderGroups.has(senderEmail)) {
          senderGroups.set(senderEmail, []);
        }

        const threadId = message.getThread().getId();
        const messageLink = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;

        const emailData = {
          date: message.getDate(),
          subject: message.getSubject(),
          summary: getGeminiSummary(message.getPlainBody()),
          link: messageLink,
          threadId: threadId // Store threadId for the new button
        };
        
        senderGroups.get(senderEmail).push(emailData);
      }
    });
    processedThreads.push(thread);
  });

  if (Array.from(senderGroups.keys()).length === 0) {
    console.log("All unread emails are self-sent or already processed. No briefing needed.");
    processedThreads.forEach(thread => { thread.markRead(); });
    return;
  }

  senderGroups.forEach(emails => {
    emails.sort((a, b) => a.date - b.date);
  });
  
  const rankedSenders = getSenderRankingFromGemini(senderGroups);
  
  let emailBlocks = '';
  rankedSenders.forEach(senderEmail => {
    const emails = senderGroups.get(senderEmail);
    if (!emails) return;

    emailBlocks += `<h2 style="padding-bottom: 10px; border-bottom: 2px solid #1A73E8; color: #1A73E8;">From: ${senderEmail}</h2>`;
    
    emails.forEach(email => {
      const requestSubject = `Fetch Email Body: id:${email.threadId}`;
      const encodedSubject = encodeURIComponent(requestSubject);
      const mailtoLink = `mailto:${scriptUserEmail}?subject=${encodedSubject}`;

      emailBlocks += `
        <div style="border: 1px solid #ccc; border-radius: 8px; margin-bottom: 20px; padding: 15px; background-color: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
           <p style="margin:0 0 10px 0;"><b>Subject:</b> ${email.subject}</p>
           <div style="background-color: #E8F0FE; border-left: 4px solid #1A73E8; padding: 12px; font-size: 14px; color: #1C3A5A;">
             <p style="margin:0;">✨ <b>AI Summary:</b> ${email.summary}</p>
           </div>
           <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 10px;">
             <p style="margin:0; font-size: 12px; color: #777;">Time: ${email.date.toLocaleString("en-US", { timeZone: "America/New_York" })}</p>
             <div style="display: flex; gap: 10px;">
               <a href="${mailtoLink}" target="_blank" style="font-size: 12px; font-weight: bold; color: #ffffff; background-color: #185ABC; padding: 5px 12px; border-radius: 4px; text-decoration: none;">Request Full Body</a>
               <a href="${email.link}" target="_blank" style="font-size: 12px; font-weight: bold; color: #ffffff; background-color: #4285F4; padding: 5px 12px; border-radius: 4px; text-decoration: none;">Open in Gmail</a>
             </div>
           </div>
         </div>`;
    });
  });

  const summarySubject = `✨ AI Smart Briefing - ${new Date().toLocaleString("en-US", { timeZone: "America/New_York" })}`;
  const summaryBody = `
    <div style="font-family: Arial, sans-serif; background-color: #f4f4f9; padding: 20px;">
      <h1 style="color: #333; text-align: center;">AI Smart Briefing</h1>
      <p style="text-align: center; color: #555;">Your AI assistant has sorted and summarized your new emails by sender importance:</p>
      <hr style="border:none; border-top: 1px solid #ddd; margin: 20px 0;">
      ${emailBlocks}
    </div>
  `;

  try {
    MailApp.sendEmail(recipientEmail, summarySubject, "", { htmlBody: summaryBody });
    console.log("AI Smart Briefing sent successfully.");
    processedThreads.forEach(thread => { thread.markRead(); });
    console.log(`${processedThreads.length} email threads have been marked as read.`);
  } catch (e) {
    console.error(`Failed to send AI briefing: ${e.toString()}`);
  }
}


/**
 * @description Calls the Gemini API to get a summary of a given text.
 * @param {string} text The email body text to summarize.
 * @return {string} The summary from Gemini, or an error message.
 */
function getGeminiSummary(text) {
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!API_KEY) {
    return "Error: GEMINI_API_KEY not found. Please check the script properties in the project settings.";
  }

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + API_KEY;
  
  const payload = {
    "contents": [{
      "parts": [{
        "text": "Please briefly summarize the core points of the following email in one or two Chinese sentences, so the recipient can quickly judge its importance. Do not add any extra explanations or introductory phrases; just provide the summary directly. Here is the email content:\n\n" + text
      }]
    }],
    "generationConfig": {
      "temperature": 0.3,
      "maxOutputTokens": 100
    }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const data = JSON.parse(responseBody);
      return data.candidates[0].content.parts[0].text.trim();
    } else {
      console.error("Gemini API call failed. Response code: " + responseCode + " | Response body: " + responseBody);
      return "[AI summary failed: API call error. Please check the logs.]";
    }
  } catch (e) {
    console.error("An exception occurred when calling UrlFetchApp: " + e.toString());
    return "[AI summary failed: Script execution error. Please check the logs.]";
  }
}

/**
 * @description Takes grouped email summaries and asks Gemini to rank ALL senders by importance.
 * @param {Map<string, Array<Object>>} senderGroups A Map where keys are sender emails and values are arrays of their email data.
 * @return {Array<string>} A ranked array of sender email addresses.
 */
function getSenderRankingFromGemini(senderGroups) {
  let promptText = "I have a list of email summaries, grouped by sender. Please act as my executive assistant. Based on the content of each group of email summaries, determine which senders are more urgent or important, and then sort these senders. Your response must include all the senders I provide, without missing any. Please only return a comma-separated list of sender email addresses, sorted from most to least important, without any other text, explanations, or numbering. For example: 'boss@example.com,client@example.com,team@example.com'. Here is the data for you to analyze:\n\n";

  senderGroups.forEach((emails, sender) => {
    promptText += `Sender: ${sender}\n`;
    emails.forEach(email => {
      promptText += `- Summary: ${email.summary}\n`;
    });
    promptText += "\n";
  });
  
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!API_KEY) {
    console.error("API key not found, cannot perform ranking.");
    return Array.from(senderGroups.keys());
  }
  
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + API_KEY;
  const payload = {
    "contents": [{"parts": [{"text": promptText}]}],
    "generationConfig": { "temperature": 0.1 }
  };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      const rankedListString = data.candidates[0].content.parts[0].text.trim();
      return rankedListString.split(',').map(email => email.trim());
    } else {
      console.error("AI ranking API call failed: " + response.getContentText());
      return Array.from(senderGroups.keys());
    }
  } catch (e) {
    console.error("An exception occurred during AI ranking: " + e.toString());
    return Array.from(senderGroups.keys());
  }
}