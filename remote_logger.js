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
          await this._post("editMessageText", {
            chat_id: this.chatId,
            message_id: messageId,
            text: truncated,
            parse_mode: "HTML",
          });
        } else {
          const resp = await this._post("sendMessage", {
            chat_id: this.chatId,
            text: truncated,
            parse_mode: "HTML",
          });
          if (resp?.data?.result?.message_id) {
            this.sessionMap.set(email, resp.data.result.message_id);
          }
        }
      } catch (err) {
        // Edit gagal (message dihapus dll) — fallback ke send baru
        if (messageId) {
          this.sessionMap.delete(email);
          const resp = await this._post("sendMessage", {
            chat_id: this.chatId,
            text: truncated,
            parse_mode: "HTML",
          });
          if (resp?.data?.result?.message_id) {
            this.sessionMap.set(email, resp.data.result.message_id);
          }
        } else {
          throw err;
        }
      } finally {
        if (isFinal) this.sessionMap.delete(email);
      }
    });
  }

  // ... method lain (getProgressBar, escapeHTML, logStep, dll) tidak perlu diubah

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
    // FIX: pakai send() supaya masuk queue yang sama, bukan _enqueue() baru
    await this.send(status, "HTML");
  }
}

module.exports = new RemoteLogger();
