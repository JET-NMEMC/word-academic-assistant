import { STORAGE_KEYS, DEFAULT_POLISH_TEMPLATE, DEFAULT_TRANSLATE_TEMPLATE } from "./constants.js";

/**
 * 配置管理模块
 */
export const storage = {
  loadConfig() {
    return {
      apiKey: localStorage.getItem(STORAGE_KEYS.API_KEY) || "",
      endpointId: localStorage.getItem(STORAGE_KEYS.ENDPOINT_ID) || "",
      promptPolish: localStorage.getItem(STORAGE_KEYS.POLISH_TEMPLATE) || DEFAULT_POLISH_TEMPLATE,
      promptTranslate: localStorage.getItem(STORAGE_KEYS.TRANSLATE_TEMPLATE) || DEFAULT_TRANSLATE_TEMPLATE,
    };
  },

  saveConfig(config) {
    localStorage.setItem(STORAGE_KEYS.API_KEY, config.apiKey || "");
    localStorage.setItem(STORAGE_KEYS.ENDPOINT_ID, config.endpointId || "");
    localStorage.setItem(STORAGE_KEYS.POLISH_TEMPLATE, config.promptPolish || DEFAULT_POLISH_TEMPLATE);
    localStorage.setItem(STORAGE_KEYS.TRANSLATE_TEMPLATE, config.promptTranslate || DEFAULT_TRANSLATE_TEMPLATE);
  },

  getPreferredApiType() {
    return localStorage.getItem(STORAGE_KEYS.PREFERRED_API_TYPE) || "chat";
  },

  setPreferredApiType(type) {
    localStorage.setItem(STORAGE_KEYS.PREFERRED_API_TYPE, type);
  }
};
