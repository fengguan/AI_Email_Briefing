// =================================================================
// SECTION 1: ADD-ON USER INTERFACE (UI)
// =================================================================

/**
 * 这是插件的“主入口”，当用户打开插件时首先运行。
 * 它的作用是读取当前状态，然后调用UI构建器。
 * @param {Object} e 事件对象. 
 * @return {Card}
 */
function onHomepage(e) {
  const userProperties = PropertiesService.getUserProperties();
  // 我们根据简报触发器的存在来判断“运行中”状态
  const triggerId = userProperties.getProperty('briefingTriggerId');
  const savedEmail = userProperties.getProperty('recipientEmail') || '';
  const savedFrequency = userProperties.getProperty('frequencyHours') || '12';
  
  let isRunning = false;
  if (triggerId) {
    const allTriggers = ScriptApp.getProjectTriggers();
    isRunning = allTriggers.some(t => t.getUniqueId() === triggerId);
    if (!isRunning) {
      // 如果主触发器丢失，说明状态异常，清理所有相关触发器
      deleteManagedTriggers();
    }
  }
  
  // 将读取到的当前状态，作为参数传递给UI构建函数
  return buildHomepageCard(isRunning, savedEmail, savedFrequency);
}


/**
 * 这是一个“纯粹的”UI构建函数，它根据传入的参数来显示界面。
 * @param {boolean} isRunning 服务是否在运行. 
 * @param {string} email 当前设置的邮箱地址. 
 * @param {string} frequency 当前设置的频率. 
 * @return {Card}
 */
function buildHomepageCard(isRunning, email, frequency) {
  let statusMessage = "服务当前状态：已停止。";
  if (isRunning) {
    statusMessage = `服务运行中，简报大约每 ${frequency} 小时更新一次。邮件提取功能已激活。将发送至: ${email}`;
  }

  const statusWidget = CardService.newTextParagraph().setText(statusMessage);

  const emailInput = CardService.newTextInput()
    .setFieldName("recipient_email_input")
    .setTitle("接收简报和邮件的邮箱地址")
    .setValue(email);

  const frequencyDropdown = CardService.newSelectionInput()
    .setFieldName("frequency_input")
    .setTitle("简报更新频率")
    .setType(CardService.SelectionInputType.DROPDOWN)
    .addItem("每小时", "1", frequency === "1")
    .addItem("每 3 小时", "3", frequency === "3")
    .addItem("每 6 小时", "6", frequency === "6")
    .addItem("每 12 小时 (默认)", "12", frequency === "12")
    .addItem("每 24 小时", "24", frequency === "24");

  const startButton = CardService.newTextButton()
    .setText("启动 / 更新服务")
    .setOnClickAction(CardService.newAction().setFunctionName("handleStartService"));

  const stopButton = CardService.newTextButton()
    .setText("停止服务")
    .setOnClickAction(CardService.newAction().setFunctionName("handleStopService"));
    
  const runNowButton = CardService.newTextButton()
    .setText("立即手动触发一次简报")
    .setOnClickAction(CardService.newAction().setFunctionName("handleRunNow"));

  const cardSection = CardService.newCardSection()
    .setHeader("状态与设置")
    .addWidget(statusWidget)
    .addWidget(emailInput)
    .addWidget(frequencyDropdown)
    .addWidget(startButton)
    .addWidget(stopButton)
    .addWidget(runNowButton);

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("AI助手设置"))
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
      .setNotification(CardService.newNotification().setText("请先设置并启动服务。"))
      .build();
  }
  
  forwardAllEmails();
  
  return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("手动触发成功！简报已发送至您的邮箱。"))
      .build();
}

/**
 * 处理“启动/更新服务”按钮的点击事件
 */
function handleStartService(e) {
  const userProperties = PropertiesService.getUserProperties();
  const recipientEmail = e.formInput.recipient_email_input;
  const frequencyHours = parseInt(e.formInput.frequency_input, 10);

  if (!recipientEmail || !recipientEmail.includes('@')) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("请输入一个有效的邮箱地址。"))
      .build();
  }

  userProperties.setProperty('recipientEmail', recipientEmail);
  userProperties.setProperty('frequencyHours', String(frequencyHours));
  
  // 删除旧的受管触发器，以便创建新的
  deleteManagedTriggers();

  // 为AI简报创建触发器
  const briefingTrigger = ScriptApp.newTrigger('forwardAllEmails')
      .timeBased().everyHours(frequencyHours).create();
  userProperties.setProperty('briefingTriggerId', briefingTrigger.getUniqueId());
  console.log(`已创建简报触发器: ${briefingTrigger.getUniqueId()}`);

  // 为邮件正文提取功能创建触发器 (固定每1小时，此为插件最低频率限制)
  const requestTrigger = ScriptApp.newTrigger('processEmailRequest')
      .timeBased().everyHours(1).create();
  userProperties.setProperty('requestTriggerId', requestTrigger.getUniqueId());
  console.log(`已创建邮件提取触发器: ${requestTrigger.getUniqueId()}`);

  const updatedCard = buildHomepageCard(true, recipientEmail, String(frequencyHours));

  return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(updatedCard))
      .setNotification(CardService.newNotification().setText("服务已启动！简报和邮件提取功能均已激活。"))
      .build();
}


/**
 * 处理“停止服务”按钮的点击事件
 */
function handleStopService(e) {
  deleteManagedTriggers();
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
  
  const updatedCard = buildHomepageCard(false, '', '12');

  return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(updatedCard))
      .setNotification(CardService.newNotification().setText("服务已成功停止。"))
      .build();
}


/**
 * 辅助函数：删除由本插件管理的所有触发器，包括旧版本的残留触发器。
 */
function deleteManagedTriggers() {
  console.log("--- 开始执行 deleteManagedTriggers ---");
  const userProperties = PropertiesService.getUserProperties();
  const briefingTriggerId = userProperties.getProperty('briefingTriggerId');
  const requestTriggerId = userProperties.getProperty('requestTriggerId');
  const oldTriggerId = userProperties.getProperty('triggerId');

  console.log(`存储的触发器ID: briefingTriggerId=${briefingTriggerId}, requestTriggerId=${requestTriggerId}, oldTriggerId=${oldTriggerId}`);

  const allTriggers = ScriptApp.getProjectTriggers();
  console.log(`ScriptApp.getProjectTriggers() 找到了 ${allTriggers.length} 个触发器。`);
  let deletedCount = 0;

  if (allTriggers.length > 0) {
    allTriggers.forEach((trigger, index) => {
      const triggerUid = trigger.getUniqueId();
      const handlerFunction = trigger.getHandlerFunction();
      console.log(`检查第 ${index + 1} 个触发器: ID=${triggerUid}, 函数=${handlerFunction}`);

      let shouldDelete = false;
      if (triggerUid === briefingTriggerId) {
        console.log(` -> 匹配到 briefingTriggerId，准备删除。`);
        shouldDelete = true;
      } else if (triggerUid === requestTriggerId) {
        console.log(` -> 匹配到 requestTriggerId，准备删除。`);
        shouldDelete = true;
      } else if (triggerUid === oldTriggerId) {
        console.log(` -> 匹配到旧的 triggerId，准备删除。`);
        shouldDelete = true;
      } else if (handlerFunction === 'forwardAllEmails' || handlerFunction === 'processEmailRequest') {
        console.log(` -> 函数名匹配，准备删除。`);
        shouldDelete = true;
      } else {
        console.log(` -> 无匹配，将保留此触发器。`);
      }

      if (shouldDelete) {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
        console.log(`   --> 触发器 ${triggerUid} 已删除。`);
      }
    });
  }
  
  if (deletedCount > 0) {
    console.log(`总共删除了 ${deletedCount} 个项目触发器。`);
  }

  userProperties.deleteProperty('briefingTriggerId');
  userProperties.deleteProperty('requestTriggerId');
  userProperties.deleteProperty('triggerId');
  console.log("--- 结束执行 deleteManagedTriggers ---");
}


// =================================================================
// SECTION 3: CORE LOGIC - EMAIL PROCESSING & AI
// =================================================================

/**
 * 新功能：根据用户邮件请求，提取并发送指定邮件的全文。
 */
function processEmailRequest() {
  const userProperties = PropertiesService.getUserProperties();
  const authorizedEmail = userProperties.getProperty('recipientEmail'); 

  if (!authorizedEmail) {
    console.log("未配置授权邮箱，无法处理远程请求。");
    return;
  }

  const searchQuery = `is:inbox is:unread from:(${authorizedEmail}) subject:("获取邮件正文:")`;
  const threads = GmailApp.search(searchQuery);

  if (threads.length === 0) {
    return;
  }

  console.log(`发现 ${threads.length} 个新的邮件提取请求。`);

  threads.forEach(thread => {
    const message = thread.getMessages()[0];
    if (message.isUnread()) {
      const requestSubject = message.getSubject();
      const targetSubject = requestSubject.replace("获取邮件正文:", "").trim();

      if (targetSubject) {
        console.log(`正在搜索主题为: "${targetSubject}" 的邮件`);
        const targetThreads = GmailApp.search(`subject:("${targetSubject}") -from:me`, 0, 1);

        if (targetThreads.length > 0) {
          const targetMessage = targetThreads[0].getMessages()[0];
          const originalSubject = targetMessage.getSubject();
          const originalBody = targetMessage.getBody();
          const originalSender = targetMessage.getFrom();

          const replySubject = `邮件正文回复: ${originalSubject}`;
          const replyBody = `
            <div style="font-family: Arial, sans-serif; padding: 20px;">
              <p>您好，您请求的邮件全文如下:</p>
              <div style="border: 1px solid #ccc; border-radius: 8px; margin-top: 15px; padding: 15px; background-color: #f9f9f9;">
                <p><b>发件人:</b> ${originalSender}</p>
                <p><b>主题:</b> ${originalSubject}</p>
                <hr>
                ${originalBody}
              </div>
            </div>
          `;

          try {
            MailApp.sendEmail(authorizedEmail, replySubject, "", { htmlBody: replyBody });
            console.log(`已成功将主题为 "${targetSubject}" 的邮件内容发送至 ${authorizedEmail}`);
          } catch (e) {
            console.error(`发送邮件至 ${authorizedEmail} 失败。错误: ${e.toString()}`);
          }

        } else {
          console.log(`未能找到主题为 "${targetSubject}" 的邮件。`);
          try {
            MailApp.sendEmail(authorizedEmail, `未能找到邮件: ${targetSubject}`, `抱歉，您的收件箱中没有找到主题为 "${targetSubject}" 的邮件。请检查主题是否完全匹配。`);
          } catch (e) {
            console.error(`发送“未找到”通知邮件失败。错误: ${e.toString()}`);
          }
        }
      }
      
      thread.markRead();
    }
  });
}


/**
 * 主功能函数 (AI简报)，由触发器定时调用
 */
function forwardAllEmails() {
  const userProperties = PropertiesService.getUserProperties();
  const recipientEmail = userProperties.getProperty('recipientEmail');

  if (!recipientEmail) {
    console.log("未找到目标邮箱配置，简报功能执行停止。");
    return;
  }
  
  const quotaLeft = MailApp.getRemainingDailyQuota();
  if (quotaLeft < 1) {
    console.warn("邮件配额不足，AI简报功能本次运行已自动停止。");
    return;
  }
  
  const searchQuery = 'is:inbox is:unread';
  const threads = GmailApp.search(searchQuery);

  if (threads.length === 0) {
    return;
  }

  console.log(`发现 ${threads.length} 个新的邮件对话，正在为简报进行处理...`);

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
          link: messageLink
        };
        
        senderGroups.get(senderEmail).push(emailData);
      }
    });
    processedThreads.push(thread);
  });

  if (Array.from(senderGroups.keys()).length === 0) {
    console.log("所有未读邮件均为自己发送或已处理，无需生成简报。");
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

    emailBlocks += `<h2 style="padding-bottom: 10px; border-bottom: 2px solid #1A73E8; color: #1A73E8;">发件人: ${senderEmail}</h2>`;
    
    emails.forEach(email => {
      emailBlocks += `
        <div style="border: 1px solid #ccc; border-radius: 8px; margin-bottom: 20px; padding: 15px; background-color: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
           <p style="margin:0 0 10px 0;"><b>主题:</b> ${email.subject}</p>
           <div style="background-color: #E8F0FE; border-left: 4px solid #1A73E8; padding: 12px; font-size: 14px; color: #1C3A5A;">
             <p style="margin:0;">✨ <b>AI 智能摘要:</b> ${email.summary}</p>
           </div>
           <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 10px;">
             <p style="margin:0; font-size: 12px; color: #777;">时间: ${email.date.toLocaleString("zh-CN", { timeZone: "America/New_York" })}</p>
             <a href="${email.link}" target="_blank" style="font-size: 12px; font-weight: bold; color: #ffffff; background-color: #4285F4; padding: 5px 12px; border-radius: 4px; text-decoration: none;">在Gmail中打开</a>
           </div>
         </div>`;
    });
  });

  const summarySubject = `✨ AI 智能简报 - ${new Date().toLocaleString("zh-CN", { timeZone: "America/New_York" })}`;
  const summaryBody = `
    <div style="font-family: Arial, sans-serif; background-color: #f4f4f9; padding: 20px;">
      <h1 style="color: #333; text-align: center;">AI 智能简报</h1>
      <p style="text-align: center; color: #555;">您的AI助手已将新邮件按发件人重要性排序并总结如下：</p>
      <hr style="border:none; border-top: 1px solid #ddd; margin: 20px 0;">
      ${emailBlocks}
    </div>
  `;

  try {
    MailApp.sendEmail(recipientEmail, summarySubject, "", { htmlBody: summaryBody });
    console.log("AI智能简报已成功发送。");
    processedThreads.forEach(thread => { thread.markRead(); });
    console.log(`${processedThreads.length} 个邮件对话已被成功标记为已读。`);
  } catch (e) {
    console.error(`发送AI简报失败: ${e.toString()}`);
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
    return "错误：未能找到GEMINI_API_KEY，请检查项目设置中的脚本属性。";
  }

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + API_KEY;
  
  const payload = {
    "contents": [{
      "parts": [{
        "text": "请用一两句中文简要总结以下邮件内容的核心要点，以便收件人能快速判断其重要性。不要添加任何多余的解释或开头语，直接给出摘要。邮件内容如下：\n\n" + text
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
      console.error("Gemini API 调用失败。响应码: " + responseCode + " | 响应内容: " + responseBody);
      return "【AI摘要失败：API调用出错，请检查日志】";
    }
  } catch (e) {
    console.error("调用UrlFetchApp时发生异常: " + e.toString());
    return "【AI摘要失败：脚本执行出错，请检查日志】";
  }
}

/**
 * @description Takes grouped email summaries and asks Gemini to rank ALL senders by importance.
 * @param {Map<string, Array<Object>>} senderGroups A Map where keys are sender emails and values are arrays of their email data.
 * @return {Array<string>} A ranked array of sender email addresses.
 */
function getSenderRankingFromGemini(senderGroups) {
  let promptText = "我有一个邮件摘要列表，按发件人分组。请你扮演我的行政助理，根据每组邮件摘要的内容，判断哪些发件人的信息更紧急或更重要，然后对这些发件人进行排序。你的回答必须包含所有我给出的发件人，一个都不能少。请只返回一个按重要性从高到低排序的、用逗号分隔的发件人邮箱地址列表，不要添加任何其他文字、解释或编号。例如：'boss@example.com,client@example.com,team@example.com'。这是需要你分析的数据：\n\n";

  senderGroups.forEach((emails, sender) => {
    promptText += `发件人: ${sender}\n`;
    emails.forEach(email => {
      promptText += `- 摘要: ${email.summary}\n`;
    });
    promptText += "\n";
  });
  
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!API_KEY) {
    console.error("未能找到API密钥，无法进行排序。");
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
      console.error("AI排序API调用失败: " + response.getContentText());
      return Array.from(senderGroups.keys());
    }
  } catch (e) {
    console.error("AI排序时发生异常: " + e.toString());
    return Array.from(senderGroups.keys());
  }
}
