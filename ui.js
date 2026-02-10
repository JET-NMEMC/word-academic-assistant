import { wordHelper } from "./word.js";

/**
 * UI 管理模块
 */
export const ui = {
  elements: {
    apiKey: document.getElementById("doubao-api-key"),
    endpointId: document.getElementById("endpoint-id"),
    polishTemplate: document.getElementById("prompt-polish-template"),
    translateTemplate: document.getElementById("prompt-translate-template"),
    resultPolish: document.getElementById("result-polish"),
    resultTranslate: document.getElementById("result-translate"),
    btnSaveConfig: document.getElementById("btn-save-config"),
    btnPolishExecute: document.getElementById("btn-polish-execute"),
    btnPolishApply: document.getElementById("btn-polish-apply"),
    btnTranslateExecute: document.getElementById("btn-translate-execute"),
    btnTranslateApply: document.getElementById("btn-translate-apply"),
    statusConfig: document.getElementById("status-config"),
    statusPolish: document.getElementById("status-polish"),
    statusTranslate: document.getElementById("status-translate"),
    globalStatus: document.getElementById("global-status")
  },

  setBusy(busy, type = null) {
    const { btnPolishExecute, btnTranslateExecute, btnSaveConfig } = this.elements;
    
    // 针对特定按钮设置加载状态
    if (type) {
      const btn = type === "polish" ? btnPolishExecute : btnTranslateExecute;
      const icon = btn.querySelector(".ms-Icon");
      const text = btn.querySelector(".btn-text");

      if (busy) {
        btn.classList.add("is-loading");
        if (icon) {
          icon.dataset.oldClass = icon.className;
          icon.className = "ms-Icon ms-Icon--ProgressRingDots"; // 切换为加载图标
        }
        if (text) text.textContent = type === "polish" ? "润色中..." : "翻译中...";
      } else {
        btn.classList.remove("is-loading");
        if (icon && icon.dataset.oldClass) {
          icon.className = icon.dataset.oldClass; // 恢复原图标
        }
        if (text) text.textContent = "执行";
      }
    }

    // 通用状态处理
    [btnPolishExecute, btnTranslateExecute, btnSaveConfig].forEach(btn => {
      if (btn) btn.disabled = busy;
    });
    
    document.body.style.cursor = busy ? "wait" : "default";
  },

  showStatus(type, message, category) {
    const element = this.elements[`status${type.charAt(0).toUpperCase() + type.slice(1)}`];
    if (!element) return;
    
    element.textContent = message;
    element.className = "status-hint " + (category === "error" ? "status-error" : "status-success");
    
    setTimeout(() => {
      if (element.textContent === message) {
        element.textContent = "";
      }
    }, 3000);
  },

  async updateButtonStates() {
    const { apiKey, endpointId, resultPolish, resultTranslate, btnPolishExecute, btnTranslateExecute, btnPolishApply, btnTranslateApply } = this.elements;

    if (!apiKey || !endpointId || !resultPolish) return;

    const apiKeyValue = apiKey.value.trim();
    const endpointIdValue = endpointId.value.trim();
    const polishResultValue = resultPolish.value.trim();
    const translateResultValue = resultTranslate.value.trim();
    
    let hasSelection = false;
    try {
      const selectedText = await wordHelper.getSelectedText();
      hasSelection = selectedText.length > 0;
    } catch (e) {
      hasSelection = false;
    }

    if (btnPolishExecute) btnPolishExecute.disabled = !apiKeyValue || !endpointIdValue || !hasSelection;
    if (btnTranslateExecute) btnTranslateExecute.disabled = !apiKeyValue || !endpointIdValue || !polishResultValue;
    if (btnPolishApply) btnPolishApply.disabled = !polishResultValue;
    if (btnTranslateApply) btnTranslateApply.disabled = !translateResultValue;
  }
};
