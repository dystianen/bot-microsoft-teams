const axios = require("axios");
const os = require("os");
const config = require("./config");

class RemoteLogger {
  constructor() {
    this.token = config.telegram?.token;
    this.chatId = config.telegram?.logChatId;
  }

  async send(text, parse_mode = "HTML") {
    if (!this.token || !this.chatId || !this.chatId.trim()) {
      return; // Silently skip if not configured
    }

    try {
      await axios.post(`https://api.telegram.org/bot${this.token}/sendMessage`, {
        chat_id: this.chatId,
        text,
        parse_mode,
      });
    } catch (err) {
      console.error(`[RemoteLogger] Failed to send log to Telegram: ${err.message}`);
    }
  }

  async info(msg) {
    console.log(`[INFO] ${msg}`);
    await this.send(`ℹ️ <b>[INFO]</b> ${msg}`);
  }

  async logStep(email, stepNum, msg) {
    const identifier = email ? `[<code>${email.split("@")[0]}</code>]` : "";
    console.log(`${identifier} [STEP ${stepNum}] ${msg}`);
    await this.send(`${identifier} 🚀 <b>STEP ${stepNum}</b>: ${msg}`);
  }

  async logError(email, msg, details = "") {
    const identifier = email ? `[<code>${email.split("@")[0]}</code>]` : "";
    console.error(`${identifier} [ERROR] ${msg} ${details}`);
    await this.send(`${identifier} ❌ <b>ERROR</b>: ${msg}\n<pre>${details.substring(0, 500)}</pre>`);
  }

  async logSuccess(email, msg) {
    const identifier = email ? `[<code>${email.split("@")[0]}</code>]` : "";
    console.log(`${identifier} [SUCCESS] ${msg}`);
    await this.send(`${identifier} ✅ <b>SUCCESS</b>: ${msg}`);
  }

  async reportSystemStatus(prefix = "") {
    const memory = process.memoryUsage();
    const freeMem = os.freemem() / (1024 * 1024 * 1024);
    const totalMem = os.totalmem() / (1024 * 1024 * 1024);
    const loadAvg = os.loadavg();
    
    const status = `🖥 <b>System Status ${prefix}</b>:
- CPU Load (1m): <code>${loadAvg[0].toFixed(2)}</code>
- CPU Load (5m): <code>${loadAvg[1].toFixed(2)}</code>
- RAM: <code>${freeMem.toFixed(2)} GB Free / ${totalMem.toFixed(2)} GB Total</code>
- Process RSS: <code>${(memory.rss / (1024 * 1024)).toFixed(2)} MB</code>`;
    
    console.log(`[SYSTEM] ${status.replace(/<[^>]*>/g, "")}`);
    await this.send(status);
  }
}

module.exports = new RemoteLogger();
