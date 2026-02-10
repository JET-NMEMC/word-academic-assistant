import { API_URLS } from "./constants.js";

/**
 * 豆包 Ark Chat API 调用模块 (专注 Chat 接口并优化流式解析)
 */
export async function callArkAPI(apiKey, endpointId, text, onChunk) {
  const headers = {
    "Content-Type": "application/json",
    "Authorization": `Bearer ${apiKey}`
  };

  const body = {
    model: endpointId,
    messages: [
      { role: "system", content: "你是一个专业的学术写作助手。" },
      { role: "user", content: text }
    ],
    temperature: 0.6,
    stream: true // 强制开启流式输出
  };

  console.log("[ArkAPI] 发送 Chat 请求:", API_URLS.CHAT);
  
  const res = await fetch(API_URLS.CHAT, { 
    method: "POST", 
    headers, 
    body: JSON.stringify(body),
    mode: 'cors'
  });
  
  if (!res.ok) {
    const errorText = await res.text();
    console.error("[ArkAPI] 请求失败:", res.status, errorText);
    try {
      const errorData = JSON.parse(errorText);
      throw new Error(errorData?.error?.message || `HTTP ${res.status}`);
    } catch (e) {
      throw new Error(`请求失败 (HTTP ${res.status}): ${errorText.substring(0, 100)}`);
    }
  }

  // 获取可读流
  const reader = res.body.getReader();
  const decoder = new TextDecoder("utf-8");
  let fullText = "";
  let buffer = "";

  console.log("[ArkAPI] 开始解析流式响应...");

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      // 实时解码并拼接到缓冲区
      buffer += decoder.decode(value, { stream: true });
      
      // 处理缓冲区中的每一行
      let lines = buffer.split("\n");
      // 保持最后一行（可能不完整）在缓冲区中
      buffer = lines.pop() || "";

      for (const line of lines) {
        const trimmedLine = line.trim();
        
        // 忽略空行
        if (!trimmedLine) continue;
        
        // 检查流结束标志
        if (trimmedLine === "data: [DONE]") {
          console.log("[ArkAPI] 接收到 [DONE] 标志");
          continue;
        }

        // 必须以 data: 开头
        if (trimmedLine.startsWith("data: ")) {
          const jsonStr = trimmedLine.substring(6);
          try {
            const data = JSON.parse(jsonStr);
            // 按照 Chat API 标准格式提取 delta.content
            const content = data.choices?.[0]?.delta?.content || "";
            if (content) {
              fullText += content;
              if (onChunk) onChunk(fullText);
            }
          } catch (e) {
            // 记录解析失败的行，但不中断流
            console.warn("[ArkAPI] JSON 解析跳过:", jsonStr);
          }
        }
      }
    }
  } catch (readError) {
    console.error("[ArkAPI] 流读取过程中断:", readError);
    throw new Error(`流传输中断: ${readError.message}`);
  }

  return fullText;
}
