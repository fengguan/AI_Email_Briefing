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
  const triggerId = userProperties.getProperty('triggerId');
  const savedEmail = userProperties.getProperty('recipientEmail') || '';
  const savedFrequency = userProperties.getProperty('frequencyHours') || '12';
  
  let isRunning = false;
  if (triggerId) {
    const allTriggers = ScriptApp.getProjectTriggers();
    isRunning = allTriggers.some(t => t.getUniqueId() === triggerId);
    if (!isRunning) {
      userProperties.deleteProperty('triggerId'); // 清理无效ID
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
    statusMessage = `服务运行中，大约每 ${frequency} 小时更新一次，将发送至: ${email}`;
  }

  const statusWidget = CardService.newTextParagraph().setText(statusMessage);

  const emailInput = CardService.newTextInput()
    .setFieldName("recipient_email_input")
    .setTitle("接收简报的邮箱地址")
    .setValue(email);

  const frequencyDropdown = CardService.newSelectionInput()
    .setFieldName("frequency_input")
    .setTitle("更新频率")
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
    
  const cardSection = CardService.newCardSection()
    .setHeader("状态与设置")
    .addWidget(statusWidget)
    .addWidget(emailInput)
    .addWidget(frequencyDropdown)
    .addWidget(startButton)
    .addWidget(stopButton);

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("AI简报设置"))
    .addSection(cardSection)
    .build();

  return card;
}


// =================================================================
// SECTION 2: USER ACTIONS & TRIGGER MANAGEMENT
// =================================================================

/**
 * 处理“启动/更新服务”按钮的点击事件
 */
function handleStartService(e) {
  const userProperties = PropertiesService.getUserProperties();
  const recipientEmail = e.formInput.recipient_email_input;
  const frequencyHours = parseInt(e.formInput.frequency_input, 10);

  if (!recipientEmail || !recipientEmail.includes('@')) {
    // ... (错误处理部分不变)
  }

  userProperties.setProperty('recipientEmail', recipientEmail);
  userProperties.setProperty('frequencyHours', String(frequencyHours));
  deleteUserTriggers();

  const newTrigger = ScriptApp.newTrigger('forwardAllEmails')
      .timeBased().everyHours(frequencyHours).create();
      
  userProperties.setProperty('triggerId', newTrigger.getUniqueId());

  // --- 这是修改过的部分：明确地传递“新状态”来更新UI ---
  const updatedCard = buildHomepageCard(true, recipientEmail, String(frequencyHours));

  return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(updatedCard))
      .setNotification(CardService.newNotification().setText("服务已启动！"))
      .build();
}


/**
 * 处理“停止服务”按钮的点击事件
 */
function handleStopService(e) {
  deleteUserTriggers();
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties(); // 更彻底地清除所有设置
  
  // --- 这是修改过的部分：明确地传递“已停止”状态来更新UI ---
  const updatedCard = buildHomepageCard(false, '', '12'); // 传递“已停止”的参数

  return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(updatedCard))
      .setNotification(CardService.newNotification().setText("服务已成功停止。"))
      .build();
}


/**
 * 辅助函数：删除当前用户的所有项目触发器
 */
function deleteUserTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}


// =================================================================
// SECTION 3: CORE LOGIC - EMAIL PROCESSING & AI
// 这是我们之前写好的核心功能代码，经过改造以适应插件环境
// =================================================================

/**
 * 主功能函数，由触发器定时调用
 */
function forwardAllEmails() {
  const userProperties = PropertiesService.getUserProperties();
  const recipientEmail = userProperties.getProperty('recipientEmail'); // <-- 属性名修改

  if (!recipientEmail) {
    console.log("未找到目标邮箱配置，函数执行停止。可能用户已停止服务。");
    deleteUserTriggers();
    return;
  }
  
  // --- 后面的所有代码，都与我们之前的最终版本完全相同 ---
  // (包括配额检查、摘要、排序、发送邮件等)
  
  const quotaLeft = MailApp.getRemainingDailyQuota();
  console.log(`运行前检查: 当前剩余邮件配额为 ${quotaLeft}。`);
  if (quotaLeft < 1) {
    console.warn("邮件配额不足 (剩余: 0)。为避免不必要的API调用和潜在费用，本次运行已自动停止。");
    return;
  }
  
  const searchQuery = 'is:inbox is:unread'; // 我们不再需要日期限制
  const threads = GmailApp.search(searchQuery);

  if (threads.length === 0) {
    console.log("配额充足，但没有发现需要处理的新邮件。");
    return;
  }

  console.log(`配额充足。发现 ${threads.length} 个新的邮件对话，正在进行两步AI处理...`);

  const senderGroups = new Map();
  const processedThreads = [];

  console.log("AI处理第1步：为每封邮件生成摘要和链接...");
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

        // --- 新增：获取邮件的唯一链接 ---
        const threadId = message.getThread().getId();
        const messageLink = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;
        // --- 新增结束 ---

        const emailData = {
          date: message.getDate(),
          subject: message.getSubject(),
          summary: getGeminiSummary(message.getPlainBody()),
          link: messageLink // <-- 将链接存入数据对象
        };
        
        senderGroups.get(senderEmail).push(emailData);
      }
    });
    processedThreads.push(thread);
  });

  senderGroups.forEach(emails => {
    emails.sort((a, b) => a.date - b.date);
  });
  
  console.log("AI处理第2步：调用AI对发件人进行全局重要性排序...");
  const rankedSenders = getSenderRankingFromGemini(senderGroups);
  
  console.log("构建最终的智能简报...");
  let emailBlocks = '';
  rankedSenders.forEach(senderEmail => {
    const emails = senderGroups.get(senderEmail);
    if (!emails) return;

    emailBlocks += `<h2 style="padding-bottom: 10px; border-bottom: 2px solid #1A73E8; color: #1A73E8;">发件人: ${senderEmail}</h2>`;
    
    emails.forEach(email => {
      // --- 新增：在HTML中加入链接按钮 ---
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
       // --- 新增结束 ---
    });
  });

  // ... (此函数剩余的构建和发送邮件部分，与上一版完全相同) ...
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
  // 从脚本属性中获取API密钥
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!API_KEY) {
    return "错误：未能找到GEMINI_API_KEY，请检查项目设置中的脚本属性。";
  }

  // 准备API请求
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
    'muteHttpExceptions': true // 发生错误时不抛出异常，而是返回响应，方便我们处理
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const data = JSON.parse(responseBody);
      // 安全地访问返回的文本
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
  // --- 这是修改过的部分：优化了给AI的指令(Prompt) ---
  let promptText = "我有一个邮件摘要列表，按发件人分组。请你扮演我的行政助理，根据每组邮件摘要的内容，判断哪些发件人的信息更紧急或更重要，然后对这些发件人进行排序。你的回答必须包含所有我给出的发件人，一个都不能少。请只返回一个按重要性从高到低排序的、用逗号分隔的发件人邮箱地址列表，不要添加任何其他文字、解释或编号。例如：'boss@example.com,client@example.com,team@example.com'。这是需要你分析的数据：\n\n";
  // --- 修改结束 ---

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
