/*
 * Word 学术写作助手 - taskpane.js
 */

// 默认提示词模板
const defaultPolishTemplate = "请对以下汉语学术文本进行补全和润色，要求术语准确、逻辑清晰、符合学术写作规范，保留原文核心含义：{text}";
const defaultTranslateTemplate = "请将以下润色后的汉语学术文本翻译成学术英语，要求语法正确、表达地道、符合英文论文写作规范：{text}";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // 加载保存的配置
    loadConfig();

    // 绑定按钮事件
    document.getElementById("btn-save-config").onclick = saveConfig;
    document.getElementById("btn-polish-execute").onclick = executePolish;
    document.getElementById("btn-polish-apply").onclick = () => applyResult("polish");
    document.getElementById("btn-translate-execute").onclick = executeTranslate;
    document.getElementById("btn-translate-apply").onclick = () => applyResult("translate");

    // 绑定输入框和文本域事件，实时更新按钮状态
    document.getElementById("doubao-api-key").oninput = updateButtonStates;
    document.getElementById("endpoint-polish").oninput = updateButtonStates;
    document.getElementById("endpoint-translate").oninput = updateButtonStates;
    document.getElementById("result-polish").oninput = updateButtonStates;

    // 初始状态检查
    updateButtonStates();
  }
});

/**
 * 从 localStorage 加载配置
 */
function loadConfig() {
  const apiKey = localStorage.getItem("doubao-api-key") || "";
  const endpointId = localStorage.getItem("endpoint-id") || "";
  const promptPolish = localStorage.getItem("prompt-polish-template") || defaultPolishTemplate;
  const promptTranslate = localStorage.getItem("prompt-translate-template") || defaultTranslateTemplate;
  document.getElementById("doubao-api-key").value = apiKey;
  if (document.getElementById("endpoint-id")) document.getElementById("endpoint-id").value = endpointId;
  if (document.getElementById("prompt-polish-template")) document.getElementById("prompt-polish-template").value = promptPolish;
  if (document.getElementById("prompt-translate-template")) document.getElementById("prompt-translate-template").value = promptTranslate;
}

/**
 * 保存配置到 localStorage
 */
function saveConfig() {
  const apiKey = document.getElementById("doubao-api-key").value.trim();
  const endpointId = (document.getElementById("endpoint-id")?.value || "").trim();
  const promptPolish = (document.getElementById("prompt-polish-template")?.value || defaultPolishTemplate).trim();
  const promptTranslate = (document.getElementById("prompt-translate-template")?.value || defaultTranslateTemplate).trim();
  
  localStorage.setItem("doubao-api-key", apiKey);
  localStorage.setItem("endpoint-id", endpointId);
  localStorage.setItem("prompt-polish-template", promptPolish);
  localStorage.setItem("prompt-translate-template", promptTranslate);
  
  showConfigStatus("配置已保存", "success");
  updateButtonStates();
}

/**
 * 显示配置保存状态
 */
function showConfigStatus(message, category) {
  const element = document.getElementById("status-config");
  element.textContent = message;
  element.className = "status-hint " + (category === "error" ? "status-error" : "status-success");
  setTimeout(() => { element.textContent = ""; }, 2000);
}

/**
 * 更新按钮的启用/禁用状态
 */
async function updateButtonStates() {
  const elApiKey = document.getElementById("doubao-api-key");
  const elEndpointId = document.getElementById("endpoint-id");
  const elPolishResult = document.getElementById("result-polish");
  const elTranslateResult = document.getElementById("result-translate");

  // 如果元素不存在（比如在错误的页面或 HTML 未更新），直接返回避免崩溃
  if (!elApiKey || !elEndpointId || !elPolishResult) return;

  const apiKey = elApiKey.value.trim();
  const endpointId = elEndpointId.value.trim();
  const polishResult = elPolishResult.value.trim();
  
  // 检查是否有选中文本
  let hasSelection = false;
  try {
    const selectedText = await getSelectedText();
    hasSelection = selectedText.length > 0;
  } catch (e) {
    hasSelection = false;
  }

  // 1. 补全润色执行按钮：API Key 和 润色 Endpoint 有效 且 有选中文本
  document.getElementById("btn-polish-execute").disabled = !apiKey || !endpointId || !hasSelection;
  
  // 2. 译英执行按钮：API Key 和 翻译 Endpoint 有效 且 补全润色结果非空
  document.getElementById("btn-translate-execute").disabled = !apiKey || !endpointId || !polishResult;

  // 3. 采纳按钮：结果非空时启用
  document.getElementById("btn-polish-apply").disabled = !polishResult;
  document.getElementById("btn-translate-apply").disabled = !document.getElementById("result-translate").value.trim();
}

// 轮询选中状态以动态更新按钮 (Office.js 暂时没有 SelectionChanged 事件的可靠触发，简单轮询或依赖用户操作)
setInterval(updateButtonStates, 2000);

/**
 * 读取 Word 选中文本
 */
async function getSelectedText() {
  return await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();
    return selection.text ? selection.text.trim() : "";
  });
}

/**
 * 向 Word 写入文本
 * @param {string} text 要写入的文本
 * @param {string} mode "replace" (替换选中) 或 "insert" (插入到光标位置)
 */
async function writeTextToDocument(text, mode) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    if (mode === "replace") {
      selection.insertText(text, Word.InsertLocation.replace);
    } else if (mode === "insert") {
      // 在 Word 中，insertText 到 end 实际上是在当前范围之后插入
      // 如果要插入到光标位置且保留换行，可以使用 Word.InsertLocation.end
      selection.insertText(text, Word.InsertLocation.end);
    }
    await context.sync();
  });
}

/**
 * 调用豆包 AI 接口
 */
async function callDoubaoAPI(apiKey, endpointId, prompt) {
  const doubaoApiUrl = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"; 
  const requestData = {
    model: endpointId,
    messages: [
      { role: "system", content: "你是一个专业的学术写作助手，精通汉语润色和学术汉译英。" },
      { role: "user", content: prompt }
    ],
    temperature: 0.2,
    max_tokens: 2048
  };

  try {
    const response = await fetch(doubaoApiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`
      },
      body: JSON.stringify(requestData)
    });

    if (!response.ok) {
      if (response.status === 401) return "API Key 错误，请核对后重新输入";
      if (response.status === 404) return "模型 Endpoint 未找到，请检查配置";
      const errorDetail = await response.text();
      return `API 请求失败：${response.status}`;
    }

    const result = await response.json();
    if (!result || !result.choices || !result.choices[0] || !result.choices[0].message) {
      return "AI 返回结果格式错误，请重试";
    }
    return result.choices[0].message.content || "AI 返回结果为空";
  } catch (error) {
    if (error.name === "TypeError" || error.message.includes("fetch")) {
      return "网络异常，请检查网络连接";
    }
    return "AI 调用失败：" + error.message;
  }
}

/**
 * 专门针对种子翻译模型的 API 调用 (v3/responses)
 */
async function callTranslationAPI(apiKey, endpointId, text) {
  const url = "https://ark.cn-beijing.volces.com/api/v3/responses";
  const items = [{ type: "input_text", text }];
  const requestData = { model: endpointId, input: [{ role: "user", content: items }] };

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`
      },
      body: JSON.stringify(requestData)
    });

    if (!response.ok) {
      if (response.status === 401) return "API Key 错误";
      if (response.status === 404) return "翻译接入点未找到";
      return `翻译请求失败：${response.status}`;
    }

    const result = await response.json();
    // v3/responses 返回结构解析
    const content = result?.choices?.[0]?.message?.content;
    if (Array.isArray(content)) {
      const firstText = content.find(c => c.text)?.text || content[0]?.text || "";
      if (firstText) return firstText;
    } else if (typeof content === "string") {
      return content;
    }

    // 兜底：尝试使用 chat/completions 做翻译
    const fallbackPrompt = `请将以下汉语文本翻译为学术英文：\n\n${text}`;
    const fallback = await callDoubaoAPI(apiKey, endpointId, fallbackPrompt);
    if (fallback && typeof fallback === "string") return fallback;

    return "翻译失败：接口返回格式异常";
  } catch (error) {
    return "翻译异常：" + error.message;
  }
}

/**
 * 执行汉语补全润色
 */
async function executePolish() {
  const apiKey = document.getElementById("doubao-api-key").value.trim();
  const endpointId = (document.getElementById("endpoint-id")?.value || "").trim();
  const selectedText = await getSelectedText();
  
  if (!selectedText) {
    showStatus("polish", "未选中任何文本，无法执行", "error");
    return;
  }

  const tplPolish = (document.getElementById("prompt-polish-template")?.value || defaultPolishTemplate);
  const prompt = tplPolish.replace("{text}", selectedText);
  showStatus("polish", "AI 处理中...", "success");

  const url = "https://ark.cn-beijing.volces.com/api/v3/responses";
  const requestData = { model: endpointId, input: [{ role: "user", content: [{ type: "input_text", text: prompt }] }] };
  let aiResult = "";
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`
      },
      body: JSON.stringify(requestData)
    });
    if (!response.ok) {
      aiResult = `请求失败：${response.status}`;
    } else {
      const result = await response.json();
      const content = result?.choices?.[0]?.message?.content;
      if (Array.isArray(content)) {
        aiResult = content.find(c => c.text)?.text || content[0]?.text || "结果为空";
      } else if (typeof content === "string") {
        aiResult = content;
      } else {
        aiResult = "结果格式异常";
      }
    }
  } catch (e) {
    aiResult = "网络异常或接口错误";
  }
  
  if (aiResult.includes("失败") || aiResult.includes("异常") || aiResult.includes("错误") || aiResult.includes("无效")) {
    showStatus("polish", aiResult, "error");
  } else {
    document.getElementById("result-polish").value = aiResult;
    showStatus("polish", "处理完成（可编辑后点击采纳）", "success");
    updateButtonStates();
  }
}

/**
 * 执行润色后汉语译英
 */
async function executeTranslate() {
  const apiKey = document.getElementById("doubao-api-key").value.trim();
  const endpointId = (document.getElementById("endpoint-id")?.value || "").trim();
  const polishText = document.getElementById("result-polish").value.trim();
  
  if (!polishText) {
    showStatus("translate", "请先执行并生成润色结果", "error");
    return;
  }

  const tplTranslate = (document.getElementById("prompt-translate-template")?.value || defaultTranslateTemplate);
  const prompt = tplTranslate.replace("{text}", polishText);
  showStatus("translate", "AI 翻译中...", "success");

  const aiResult = await callTranslationAPI(apiKey, endpointId, prompt);
  
  if (aiResult.includes("失败") || aiResult.includes("异常") || aiResult.includes("错误") || aiResult.includes("无效")) {
    showStatus("translate", aiResult, "error");
  } else {
    document.getElementById("result-translate").value = aiResult;
    showStatus("translate", "翻译完成（可编辑后点击采纳）", "success");
    updateButtonStates();
  }
}

/**
 * 采纳结果并写入 Word
 */
async function applyResult(type) {
  const textareaId = type === "polish" ? "result-polish" : "result-translate";
  const mode = type === "polish" ? "replace" : "insert";
  const text = document.getElementById(textareaId).value;

  if (!text) return;

  try {
    await writeTextToDocument(text, mode);
    showStatus(type, "已成功应用到文档", "success");
  } catch (error) {
    showStatus(type, "应用失败：" + error.message, "error");
  }
}

/**
 * 显示状态提示
 */
function showStatus(type, message, category) {
  const element = document.getElementById(`status-${type}`);
  element.textContent = message;
  element.className = "status-hint " + (category === "error" ? "status-error" : "status-success");
  
  // 3秒后自动消失
  setTimeout(() => {
    if (element.textContent === message) {
      element.textContent = "";
    }
  }, 3000);
}
