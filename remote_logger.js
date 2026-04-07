const axios = require("axios");
const os = require("os");
const config = require("./config");

class RemoteLogger {
  constructor() {
    this.token = config.telegram?.token;
    this.chatId = config.telegram?.logChatId;
    this.sessionMap = new Map();
    this.queue = Promise.resolve();
  }

  async _enqueue(action) {
    // FIX: chain dengan benar, jangan replace queue sebelum waktunya
    const promise = this.queue.then(async () => {
      try {
        await action();
      } catch (err) {
        console.error(`[RemoteLogger] Action failed: ${err.message}`);
      }
      await new Promise((r) => setTimeout(r, 450)); // sedikit naikin buffer
    });

    // FIX: assign dulu baru return, dan pastikan chain tidak putus
    this.queue = promise;
    return promise;
  }

  async _post(endpoint, payload, retries = 3) {
    for (let attempt = 1; attempt <= retries; attempt++) {
      try {
        const res = await axios.post(
          `https://api.telegram.org/bot${this.token}/${endpoint}`,
          payload,
        );
        return res;
      } catch (err) {
        const status = err.response?.status;
        const desc = err.response?.data?.description || "";

        // Rate limited — tunggu sesuai retry_after dari Telegram
        if (status === 429) {
          const retryAfter =
            (err.response?.data?.parameters?.retry_after || 5) * 1000;
          console.warn(
            `[RemoteLogger] Rate limited, retry after ${retryAfter}ms (attempt ${attempt})`,
          );
          await new Promise((r) => setTimeout(r, retryAfter));
          continue;
        }

        // HTML parse error — strip tag dan coba lagi tanpa parse_mode
        if (desc.includes("can't parse entities")) {
          payload = {
            ...payload,
            text: (payload.text || "").replace(/<[^>]*>?/gm, ""),
            parse_mode: "",
          };
          continue;
        }

        // Message not modified — bukan error, skip
        if (desc.includes("message is not modified")) return null;

        // Error lain — lempar kalau sudah attempt terakhir
        if (attempt === retries) throw err;

        // Backoff sebelum retry
        await new Promise((r) => setTimeout(r, 1000 * attempt));
      }
    }
  }

  async send(text, parse_mode = "HTML") {
    if (!this.token || !this.chatId?.trim()) return;

    const CHUNK_SIZE = 4000;
    const chunks = [];

    if (text.length <= CHUNK_SIZE) {
      chunks.push(text);
    } else {
      let currentIdx = 0;
      while (currentIdx < text.length) {
        let chunk = text.substring(currentIdx, currentIdx + CHUNK_SIZE);
        const lastNewline = chunk.lastIndexOf("\n");
        if (lastNewline > 500 && currentIdx + CHUNK_SIZE < text.length) {
          chunk = text.substring(currentIdx, currentIdx + lastNewline);
          currentIdx += lastNewline + 1;
        } else {
          currentIdx += CHUNK_SIZE;
        }
        chunks.push(chunk);
      }
    }

    for (const chunk of chunks) {
      await this._enqueue(() =>
        this._post("sendMessage", {
          chat_id: this.chatId,
          text: chunk,
          parse_mode,
        }),
      );
    }
  }

  async _sendOrEdit(email, text, isFinal = false) {
    if (!this.token || !this.chatId?.trim()) return;

    await this._enqueue(async () => {
      const messageId = this.sessionMap.get(email);
      const truncated = text.substring(0, 4000);

      try {
        if (messageId) {
          // Sudah ada message sebelumnya — edit saja
          await this._post("editMessageText", {
            chat_id: this.chatId,
            message_id: messageId,
            text: truncated,
            parse_mode: "HTML",
          });
        } else {
          // Belum ada — kirim baru, tapi set dulu ke "pending"
          // supaya kalau ada step berikutnya yang masuk queue
          // mereka tunggu message_id ini selesai
          this.sessionMap.set(email, "pending");

          const resp = await this._post("sendMessage", {
            chat_id: this.chatId,
            text: truncated,
            parse_mode: "HTML",
          });

          if (resp?.data?.result?.message_id) {
            this.sessionMap.set(email, resp.data.result.message_id);
          } else {
            this.sessionMap.delete(email);
          }
        }
      } catch (err) {
        // Edit gagal (message dihapus manual dll) — reset dan kirim baru
        this.sessionMap.delete(email);
        const resp = await this._post("sendMessage", {
          chat_id: this.chatId,
          text: truncated,
          parse_mode: "HTML",
        });
        if (resp?.data?.result?.message_id) {
          this.sessionMap.set(email, resp.data.result.message_id);
        }
      } finally {
        if (isFinal) this.sessionMap.delete(email);
      }
    });
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
    return text
      .toString()
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  async info(msg) {
    console.log(`[INFO] ${msg}`);
    if (!this.token || !this.chatId) return;
    await this._enqueue(async () => {
      await axios.post(
        `https://api.telegram.org/bot${this.token}/sendMessage`,
        {
          chat_id: this.chatId,
          text: `ℹ️ <b>[INFO]</b> ${this.escapeHTML(msg)}`,
          parse_mode: "HTML",
        },
      );
    });
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
    const isWindows = os.platform() === "win32";

    let cpuLine = "";
    if (isWindows) {
      // Windows tidak support loadavg — tampilkan CPU count saja
      cpuLine = `- CPU Cores: <code>${os.cpus().length} cores</code>`;
    } else {
      const loadAvg = os.loadavg();
      cpuLine = `- CPU Load (1m/5m/15m): <code>${loadAvg[0].toFixed(2)} / ${loadAvg[1].toFixed(2)} / ${loadAvg[2].toFixed(2)}</code>`;
    }

    const usedMem = totalMem - freeMem;
    const memPercent = ((usedMem / totalMem) * 100).toFixed(1);

    const status = `🖥 <b>System Status ${this.escapeHTML(prefix)}</b>
      ${cpuLine}
      - RAM: <code>${freeMem.toFixed(2)} GB Free / ${totalMem.toFixed(2)} GB Total (${memPercent}% used)</code>
      - Process RSS: <code>${(memory.rss / (1024 * 1024)).toFixed(2)} MB</code>
      - Heap Used: <code>${(memory.heapUsed / (1024 * 1024)).toFixed(2)} MB</code>`;

    console.log(`[SYSTEM] ${status.replace(/<[^>]*>/g, "")}`);
    await this.send(status, "HTML");
  }
}

module.exports = new RemoteLogger();
