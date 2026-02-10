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

  setBusy(busy) {
    const { btnPolishExecute, btnTranslateExecute, btnSaveConfig } = this.elements;
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

    btnPolishExecute.disabled = !apiKeyValue || !endpointIdValue || !hasSelection;
    btnTranslateExecute.disabled = !apiKeyValue || !endpointIdValue || !polishResultValue;
    btnPolishApply.disabled = !polishResultValue;
    btnTranslateApply.disabled = !translateResultValue;
  }
};
