/**
 * Word 文档交互模块
 */
export const wordHelper = {
  /**
   * 读取 Word 选中文本
   */
  async getSelectedText() {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return selection.text ? selection.text.trim() : "";
    });
  },

  /**
   * 向 Word 写入文本
   * @param {string} text 要写入的文本
   * @param {string} mode "replace" (替换选中) 或 "insert" (插入到光标位置)
   */
  async writeTextToDocument(text, mode) {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      if (mode === "replace") {
        selection.insertText(text, Word.InsertLocation.replace);
      } else if (mode === "insert") {
        selection.insertText(text, Word.InsertLocation.end);
      }
      await context.sync();
    });
  }
};
