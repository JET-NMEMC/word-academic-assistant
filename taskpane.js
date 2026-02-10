import { storage } from "./storage.js";
import { ui } from "./ui.js";
import { callArkAPI } from "./api.js";
import { wordHelper } from "./word.js";

/**
 * Word 学术写作助手 - 核心入口
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    init();
  }
});

/**
 * 初始化插件
 */
function init() {
  // 1. 加载并展示配置
  const config = storage.loadConfig();
  ui.elements.apiKey.value = config.apiKey;
  ui.elements.endpointId.value = config.endpointId;
  ui.elements.polishTemplate.value = config.promptPolish;
  ui.elements.translateTemplate.value = config.promptTranslate;

  // 2. 绑定事件
  bindEvents();

  // 3. 初始状态检查
  ui.updateButtonStates();

  // 4. 定时检查选中状态 (Office.js 缺乏可靠的 selectionChanged 事件触发)
  setInterval(() => ui.updateButtonStates(), 2000);
}

/**
 * 绑定 UI 事件
 */
function bindEvents() {
  const { 
    btnSaveConfig, btnPolishExecute, btnPolishApply, 
    btnTranslateExecute, btnTranslateApply,
    apiKey, endpointId, resultPolish, resultTranslate
  } = ui.elements;

  // 保存配置
  btnSaveConfig.onclick = () => {
    const newConfig = {
      apiKey: apiKey.value.trim(),
      endpointId: endpointId.value.trim(),
      promptPolish: ui.elements.polishTemplate.value.trim(),
      promptTranslate: ui.elements.translateTemplate.value.trim()
    };
    storage.saveConfig(newConfig);
    ui.showStatus("config", "配置已保存", "success");
    ui.updateButtonStates();
  };

  // 执行润色
  btnPolishExecute.onclick = async () => {
    const apiKeyValue = apiKey.value.trim();
    const endpointIdValue = endpointId.value.trim();
    const polishTemplate = ui.elements.polishTemplate.value;
    
    const selectedText = await wordHelper.getSelectedText();
    if (!selectedText) {
      ui.showStatus("polish", "未选中任何文本", "error");
      return;
    }

    const prompt = polishTemplate.replace("{text}", selectedText);
    ui.showStatus("polish", "正在润色...", "success");
    ui.setBusy(true, "polish");
    resultPolish.value = ""; // 清空上次结果

    try {
      await callArkAPI(apiKeyValue, endpointIdValue, prompt, (chunk) => {
        // 当收到第一个数据块时，恢复按钮状态，让用户看到流式输出
        ui.setBusy(false, "polish");
        
        // 使用 requestAnimationFrame 确保在浏览器渲染帧中更新 UI，提升流式顺滑度
        requestAnimationFrame(() => {
          resultPolish.value = chunk;
          resultPolish.scrollTop = resultPolish.scrollHeight;
        });
      });
      
      ui.showStatus("polish", "润色完成", "success");
      ui.updateButtonStates();
    } catch (err) {
      ui.showStatus("polish", err.message, "error");
    } finally {
      // 确保无论成功还是失败，状态最终都会恢复
      ui.setBusy(false, "polish");
    }
  };

  // 执行翻译
  btnTranslateExecute.onclick = async () => {
    const apiKeyValue = apiKey.value.trim();
    const endpointIdValue = endpointId.value.trim();
    const polishText = resultPolish.value.trim();
    const translateTemplate = ui.elements.translateTemplate.value;
    
    if (!polishText) {
      ui.showStatus("translate", "润色结果为空", "error");
      return;
    }

    const prompt = translateTemplate.replace("{text}", polishText);
    ui.showStatus("translate", "正在翻译...", "success");
    ui.setBusy(true, "translate");
    resultTranslate.value = ""; // 清空上次结果

    try {
      await callArkAPI(apiKeyValue, endpointIdValue, prompt, (chunk) => {
        // 当收到第一个数据块时，恢复按钮状态，让用户看到流式输出
        ui.setBusy(false, "translate");

        requestAnimationFrame(() => {
          resultTranslate.value = chunk;
          resultTranslate.scrollTop = resultTranslate.scrollHeight;
        });
      });
      
      ui.showStatus("translate", "翻译完成", "success");
      ui.updateButtonStates();
    } catch (err) {
      ui.showStatus("translate", err.message, "error");
    } finally {
      // 确保无论成功还是失败，状态最终都会恢复
      ui.setBusy(false, "translate");
    }
  };

  // 采纳结果
  btnPolishApply.onclick = () => applyResultToDoc("polish");
  btnTranslateApply.onclick = () => applyResultToDoc("translate");

  // 输入变化实时更新状态
  [apiKey, endpointId, resultPolish].forEach(el => {
    el.oninput = () => ui.updateButtonStates();
  });
}

/**
 * 采纳结果并写入文档
 */
async function applyResultToDoc(type) {
  const textarea = type === "polish" ? ui.elements.resultPolish : ui.elements.resultTranslate;
  const mode = type === "polish" ? "replace" : "insert";
  const text = textarea.value.trim();

  if (!text) return;

  try {
    await wordHelper.writeTextToDocument(text, mode);
    ui.showStatus(type, "已成功应用到文档", "success");
  } catch (error) {
    ui.showStatus(type, "应用失败：" + error.message, "error");
  }
}
