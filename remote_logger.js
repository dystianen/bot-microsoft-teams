const axios = require("axios");
const os = require("os");
const config = require("./config");

class RemoteLogger {
  constructor() {
    this.token = config.telegram?.token;
    this.chatId = config.telegram?.logChatId;
    this.sessionMap = new Map(); // Store message_id per account (email)
  }

  async send(text, parse_mode = "HTML") {
    if (!this.token || !this.chatId || !this.chatId.trim()) return;
    try {
      await axios.post(`https://api.telegram.org/bot${this.token}/sendMessage`, {
        chat_id: this.chatId,
        text: text.substring(0, 4000),
        parse_mode,
      });
    } catch (err) {
      console.error(`[RemoteLogger] Send failed: ${err.message}`);
    }
  }

  async _sendOrEdit(email, text, isFinal = false) {
    if (!this.token || !this.chatId || !this.chatId.trim()) return;

    const messageId = this.sessionMap.get(email);
    
    try {
      if (messageId) {
        await axios.post(`https://api.telegram.org/bot${this.token}/editMessageText`, {
          chat_id: this.chatId,
          message_id: messageId,
          text: text.substring(0, 4000),
          parse_mode: "HTML",
        });
      } else {
        const resp = await axios.post(`https://api.telegram.org/bot${this.token}/sendMessage`, {
          chat_id: this.chatId,
          text: text.substring(0, 4000),
          parse_mode: "HTML",
        });
        if (resp.data?.result?.message_id) {
          this.sessionMap.set(email, resp.data.result.message_id);
        }
      }
    } catch (err) {
      if (messageId) {
        // Fallback: If edit fails, send new message
        this.sessionMap.delete(email);
        return this._sendOrEdit(email, text, isFinal);
      }
      console.error(`[RemoteLogger] Failed: ${err.message}`);
    } finally {
      if (isFinal) {
        this.sessionMap.set(email, null); // Clear ID after final status
        this.sessionMap.delete(email);
      }
    }
  }

  getProgressBar(current, total = 28) {
    const size = 10;
    const progress = Math.min(Math.max(current / total, 0), 1);
    const filled = Math.round(size * progress);
    const empty = size - filled;
    const bar = "█".repeat(filled) + "░".repeat(empty);
    return `<code>[${bar}] ${Math.round(progress * 100)}%</code>`;
  }

  escapeHTML(text) {
    if (!text) return "";
    return text.toString()
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  async info(msg) {
    console.log(`[INFO] ${msg}`);
    // Info messages don't need editing, they are general
    if (!this.token || !this.chatId) return;
    await axios.post(`https://api.telegram.org/bot${this.token}/sendMessage`, {
      chat_id: this.chatId,
      text: `ℹ️ <b>[INFO]</b> ${this.escapeHTML(msg)}`,
      parse_mode: "HTML",
    }).catch(() => {});
  }

  async logStep(email, stepNum, msg) {
    const user = email ? email.split("@")[0] : "unknown";
    const identifier = `🚀 <b>Processing:</b> <code>${this.escapeHTML(user)}</code>`;
    
    console.log(`[${user}] [STEP ${stepNum}] ${msg}`);
    
    let text = `${identifier}\n`;
    text += `📍 <b>Current:</b> Step ${stepNum}/28\n`;
    text += `📝 <b>Status:</b> ${this.escapeHTML(msg)}\n`;
    text += `${this.getProgressBar(stepNum)}`;

    await this._sendOrEdit(email, text, false);
  }

  async logError(email, msg, details = "") {
    const user = email ? email.split("@")[0] : "unknown";
    const identifier = `❌ <b>Failed:</b> <code>${this.escapeHTML(user)}</code>`;
    
    console.error(`[${user}] [ERROR] ${msg} ${details}`);
    
    let formattedMsg = `${identifier}\n\n`;
    formattedMsg += `<b>Issue:</b> ${this.escapeHTML(msg)}\n`;
    if (details) {
      formattedMsg += `\n<b>Technical Details:</b>\n<pre>${this.escapeHTML(details.substring(0, 1000))}</pre>`;
    }
    
    await this._sendOrEdit(email, formattedMsg, true);
  }

  async logSuccess(email, msg) {
    const user = email ? email.split("@")[0] : "unknown";
    const identifier = `✅ <b>Success:</b> <code>${this.escapeHTML(user)}</code>`;
    
    console.log(`[${user}] [SUCCESS] ${msg}`);
    
    let text = `${identifier}\n`;
    text += `🏁 <b>Status:</b> ${this.escapeHTML(msg)}\n`;
    text += `${this.getProgressBar(28, 28)}`;

    await this._sendOrEdit(email, text, true);
  }

  async reportSystemStatus(prefix = "") {
    const memory = process.memoryUsage();
    const freeMem = os.freemem() / (1024 * 1024 * 1024);
    const totalMem = os.totalmem() / (1024 * 1024 * 1024);
    const loadAvg = os.loadavg();

    const status = `🖥 <b>System Status ${this.escapeHTML(prefix)}</b>:
      - CPU Load (1m): <code>${loadAvg[0].toFixed(2)}</code>
      - CPU Load (5m): <code>${loadAvg[1].toFixed(2)}</code>
      - RAM: <code>${freeMem.toFixed(2)} GB Free / ${totalMem.toFixed(2)} GB Total</code>
      - Process RSS: <code>${(memory.rss / (1024 * 1024)).toFixed(2)} MB</code>`;

    console.log(`[SYSTEM] ${status.replace(/<[^>]*>/g, "")}`);
    if (!this.token || !this.chatId) return;
    await axios.post(`https://api.telegram.org/bot${this.token}/sendMessage`, {
      chat_id: this.chatId,
      text: status,
      parse_mode: "HTML",
    }).catch(() => {});
  }
}

module.exports = new RemoteLogger();
