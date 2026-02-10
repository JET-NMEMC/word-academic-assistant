/**
 * 默认提示词模板和常量
 */
export const DEFAULT_POLISH_TEMPLATE = "请对以下汉语学术文本进行补全和润色，要求术语准确、逻辑清晰、符合学术写作规范，保留原文核心含义：{text}";
export const DEFAULT_TRANSLATE_TEMPLATE = "请将以下润色后的汉语学术文本翻译成学术英语，要求语法正确、表达地道、符合英文论文写作规范：{text}";

export const API_URLS = {
  CHAT: "https://ark.cn-beijing.volces.com/api/v3/chat/completions",
  RESPONSES: "https://ark.cn-beijing.volces.com/api/v3/responses"
};

export const STORAGE_KEYS = {
  API_KEY: "doubao-api-key",
  ENDPOINT_ID: "endpoint-id",
  POLISH_TEMPLATE: "prompt-polish-template",
  TRANSLATE_TEMPLATE: "prompt-translate-template",
  PREFERRED_API_TYPE: "preferred-api-type"
};
