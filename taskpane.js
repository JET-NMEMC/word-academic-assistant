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
    if (document.getElementById("endpoint-id")) document.getElementById("endpoint-id").oninput = updateButtonStates;
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
 * 通用豆包 Ark API 调用 (智能自动切换 chat/completions 和 responses)
 * @param {string} apiKey API Key
 * @param {string} endpointId 模型 Endpoint ID
 * @param {string} text 替换模板后的完整提示词文本
 */
async function callArkAPI(apiKey, endpointId, text) {
  const headers = {
    "Content-Type": "application/json",
    "Authorization": `Bearer ${apiKey}`
  };

  // 1. 标准 chat/completions 接口 (兼容 OpenAI 格式)
  const doChat = async () => {
    const url = "https://ark.cn-beijing.volces.com/api/v3/chat/completions";
    const body = {
      model: endpointId,
      messages: [
        { role: "system", content: "你是一个专业的学术写作助手。" },
        { role: "user", content: text }
      ],
      temperature: 0.2
    };
    
    const start = Date.now();
    try {
      const res = await fetch(url, { method: "POST", headers, body: JSON.stringify(body) });
      console.log(`[ArkAPI] Chat API cost: ${Date.now() - start}ms, status: ${res.status}`);
      if (!res.ok) throw new Error(`Status ${res.status}`);
      const data = await res.json();
      return data?.choices?.[0]?.message?.content || "结果为空";
    } catch (e) {
      console.warn(`[ArkAPI] Chat API failed: ${e.message}`);
      throw e;
    }
  };

  // 2. 火山引擎特有 responses 接口 (使用 input 和 input_text)
  const doResponses = async () => {
    const url = "https://ark.cn-beijing.volces.com/api/v3/responses";
    const body = {
      model: endpointId,
      input: [{ 
        role: "user", 
        content: [{ type: "input_text", text: text }] 
      }]
    };

    const start = Date.now();
    try {
      const res = await fetch(url, { method: "POST", headers, body: JSON.stringify(body) });
      console.log(`[ArkAPI] Responses API cost: ${Date.now() - start}ms, status: ${res.status}`);
      if (!res.ok) throw new Error(`Status ${res.status}`);
      const data = await res.json();
      
      // 解析 responses 接口特有的返回结构
      const content = data?.choices?.[0]?.message?.content;
      if (Array.isArray(content)) {
        return content.find(c => c.text)?.text || content[0]?.text || "结果为空";
      }
      return typeof content === "string" ? content : "结果格式异常";
    } catch (e) {
      console.warn(`[ArkAPI] Responses API failed: ${e.message}`);
      throw e;
    }
  };

  // --- 智能调度逻辑 ---

  // 读取上次成功的 API 类型，默认优先尝试 chat
  let preferredType = localStorage.getItem("preferred-api-type") || "chat";
  
  try {
    if (preferredType === "chat") {
      try {
        return await doChat();
      } catch (err) {
        // 如果 chat 失败，尝试 responses
        console.log("Switching to Responses API...");
        const result = await doResponses();
        localStorage.setItem("preferred-api-type", "responses"); // 记住这次成功的类型
        return result;
      }
    } else {
      try {
        return await doResponses();
      } catch (err) {
        // 如果 responses 失败，尝试 chat
        console.log("Switching to Chat API...");
        const result = await doChat();
        localStorage.setItem("preferred-api-type", "chat"); // 记住这次成功的类型
        return result;
      }
    }
  } catch (finalError) {
    // 两次都失败
    const errMsg = finalError.message || "";
    if (errMsg.includes("401")) return "API Key 错误";
    if (errMsg.includes("404")) return "模型 Endpoint ID 错误或不支持该接口";
    return `接口调用失败 (${errMsg})`;
  }
}

/**
 * 执行汉语补全润色
 */
async function executePolish() {
  const apiKey = document.getElementById("doubao-api-key").value.trim();
  const endpointId = (document.getElementById("endpoint-id")?.value || "").trim();
  const selectedText = await getSelectedText();
  const polishTemplate = document.getElementById("prompt-polish-template").value;
  
  if (!selectedText) {
    showStatus("polish", "未选中任何文本", "error");
    return;
  }

  const prompt = polishTemplate.replace("{text}", selectedText);
  showStatus("polish", "正在润色...", "success");

  const aiResult = await callArkAPI(apiKey, endpointId, prompt);
  
  if (aiResult.includes("失败") || aiResult.includes("异常") || aiResult.includes("错误") || aiResult.includes("无效")) {
    showStatus("polish", aiResult, "error");
  } else {
    document.getElementById("result-polish").value = aiResult;
    showStatus("polish", "润色完成", "success");
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
  const translateTemplate = document.getElementById("prompt-translate-template").value;
  
  if (!polishText) {
    showStatus("translate", "润色结果为空", "error");
    return;
  }

  const prompt = translateTemplate.replace("{text}", polishText);
  showStatus("translate", "正在翻译...", "success");

  const aiResult = await callArkAPI(apiKey, endpointId, prompt);
  
  if (aiResult.includes("失败") || aiResult.includes("异常") || aiResult.includes("错误") || aiResult.includes("无效")) {
    showStatus("translate", aiResult, "error");
  } else {
    document.getElementById("result-translate").value = aiResult;
    showStatus("translate", "翻译完成", "success");
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
