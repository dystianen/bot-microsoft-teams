const TelegramBot = require("node-telegram-bot-api");
const config = require("./config");
const { processSingleAccount } = require("./index");
const connectDB = require("./db");
const { UserConfig } = require("./models");
const date = require("date-and-time");

async function startBot() {
  try {
    await connectDB();
    const token = config.telegram.token;
    if (!token) {
      console.error("Please set TELEGRAM_BOT_TOKEN in .env and restart.");
      process.exit(1);
    }
    const bot = new TelegramBot(token, { polling: true });
    initializeBotHandlers(bot);
    console.log("Teams Bot started.");
  } catch (err) {
    console.error("Failed to start bot:", err.message);
    process.exit(1);
  }
}

function initializeBotHandlers(bot) {
  let _msgQueue = Promise.resolve();
  async function safeSendMessage(chatId, text, options = {}) {
    _msgQueue = _msgQueue.then(async () => {
      try {
        await bot.sendMessage(chatId, text, options);
      } catch (err) {
        console.error("[Telegram] Error sending message:", err.message);
      }
      await new Promise((r) => setTimeout(r, 500));
    });
    return _msgQueue;
  }

  const sessions = {};
  async function getUserConfig(telegram_id) {
    let userConf = await UserConfig.findOne({ telegram_id: telegram_id.toString() });
    if (!userConf) {
      userConf = new UserConfig({
        telegram_id: telegram_id.toString(),
        concurrencyLimit: 1, // default 1 for Teams bot
      });
      await userConf.save();
    }
    return userConf;
  }

  const mainMenu = {
    reply_markup: {
      keyboard: [
        [{ text: "➕ Add Account" }],
        [{ text: "🚀 Generate" }, { text: "🛑 Stop Queue" }],
        [{ text: "⚙️ Config" }, { text: "🧹 Reset Session" }],
      ],
      resize_keyboard: true,
    },
  };

  bot.onText(/\/start/, async (msg) => {
    const chatId = msg.chat.id;
    await getUserConfig(chatId);
    sessions[chatId] = { accounts: [], step: "IDLE", running: false };
    bot.sendMessage(chatId, "Welcome to Microsoft Teams Bot! 🤖\n\nAdd accounts as email|password and they will be queued for processing.", mainMenu);
  });

  bot.onText(/➕ Add Account/, (msg) => {
    const chatId = msg.chat.id;
    sessions[chatId] = sessions[chatId] || { accounts: [], step: "IDLE" };
    sessions[chatId].step = "WAIT_ACCOUNT";
    bot.sendMessage(chatId, "Send email|password (one per line):", { parse_mode: "Markdown" });
  });

  bot.onText(/🚀 Generate/, async (msg) => {
    const chatId = msg.chat.id;
    const session = sessions[chatId] || { accounts: [], running: false };

    if (session.running) return bot.sendMessage(chatId, "⚠️ Automation is already running!");
    if (session.accounts.length === 0) return bot.sendMessage(chatId, "Please add accounts first.");

    const userConf = await getUserConfig(chatId);
    session.running = true;

    bot.sendMessage(chatId, `🚀 Starting batch for ${session.accounts.length} accounts... (Concurrency: ${userConf.concurrencyLimit})`);

    const runQueue = async () => {
      let activeWorkers = 0;
      const maxWorkers = userConf.concurrencyLimit || 1;
      let globalIdx = 0;
      let total = session.accounts.length;
      const pendingPromises = new Set();
      
      const processAccount = async (accountData, currentIdx) => {
        await safeSendMessage(chatId, `⏳ [${currentIdx}/${total}] Processing: ${accountData.email}`);
        
        try {
          const pairedData = { microsoftAccount: accountData, telegram_id: chatId };
          const result = await processSingleAccount(pairedData, currentIdx - 1, total);

          if (result.status === "SUCCESS") {
            await safeSendMessage(chatId, `✅ <b>Success [${currentIdx}/${total}]</b>\nEmail: <code>${accountData.email}</code>`, { parse_mode: "HTML" });
          } else {
            await safeSendMessage(chatId, `❌ <b>Failed [${currentIdx}/${total}]</b>\nEmail: <code>${accountData.email}</code>\nLog: ${result.log}`, { parse_mode: "HTML" });
          }
        } catch (err) {
          await safeSendMessage(chatId, `❌ Error: ${err.message}`);
        }
      };

      try {
        while (true) {
          if (activeWorkers < maxWorkers && session.accounts.length > 0 && !session.forceStop) {
            const accountData = session.accounts.shift();
            globalIdx++;
            const currentIdx = globalIdx;

            activeWorkers++;
            const promise = processAccount(accountData, currentIdx).finally(() => {
              activeWorkers--;
              pendingPromises.delete(promise);
            });
            pendingPromises.add(promise);

            if (session.accounts.length > 0 || activeWorkers < maxWorkers) {
              await new Promise((r) => setTimeout(r, 10000)); // Stagger launches (10s)
            }
          } else if (activeWorkers === 0 && (session.accounts.length === 0 || session.forceStop)) {
            break;
          } else {
            await new Promise((r) => setTimeout(r, 1000));
          }
        }

        if (pendingPromises.size > 0) {
          await Promise.all(Array.from(pendingPromises));
        }
      } catch (err) {
        console.error("[Queue Error]", err);
        await safeSendMessage(chatId, `⚠️ Queue error: ${err.message}`);
      } finally {
        session.running = false;
        if (session.forceStop) {
          session.forceStop = false;
          bot.sendMessage(chatId, "🛑 Queue processing stopped successfully.", mainMenu);
        } else {
          bot.sendMessage(chatId, "🏁 Finished processing session accounts.", mainMenu);
        }
      }
    };

    runQueue();
  });

  bot.onText(/🛑 Stop Queue/, (msg) => {
    const chatId = msg.chat.id;
    if (sessions[chatId]) {
      sessions[chatId].accounts = [];
      sessions[chatId].forceStop = true;
    }
    bot.sendMessage(chatId, "🛑 Stopping queue...");
  });

  bot.onText(/🧹 Reset Session/, (msg) => {
    const chatId = msg.chat.id;
    sessions[chatId] = { accounts: [], step: "IDLE", running: false };
    bot.sendMessage(chatId, "Session cleared.", mainMenu);
  });

  bot.onText(/⚙️ Config/, async (msg) => {
    const chatId = msg.chat.id;
    const userConf = await getUserConfig(chatId);

    const options = {
      reply_markup: {
        inline_keyboard: [
          [{ text: `🚀 Concurrency: ${userConf.concurrencyLimit}`, callback_data: `set_concurrency` }]
        ]
      }
    };
    bot.sendMessage(chatId, `⚙️ <b>Current Configuration:</b>\nConcurrency: ${userConf.concurrencyLimit}`, { parse_mode: "HTML", ...options });
  });

  bot.on("callback_query", async (callbackQuery) => {
    const message = callbackQuery.message;
    const chatId = message.chat.id;
    const data = callbackQuery.data;

    sessions[chatId] = sessions[chatId] || { accounts: [], step: "IDLE", running: false };

    if (data === "set_concurrency") {
      sessions[chatId].step = "SET_CONCURRENCY";
      bot.sendMessage(chatId, "Please send the new concurrency limit (number).");
    }
  });

  bot.on("message", async (msg) => {
    const chatId = msg.chat.id;
    const text = msg.text;
    if (!text || text.startsWith("/")) return;
    const session = sessions[chatId];
    if (!session || session.step === "IDLE") return;

    if (session.step === "WAIT_ACCOUNT") {
      const lines = text.split("\n");
      let added = 0;
      for (const line of lines) {
        const parts = line.split("|").map((s) => s.trim());
        if (parts.length >= 2) {
          session.accounts.push({
            email: parts[0],
            password: parts[1],
            // Filler placeholders as MicrosoftBot might expect them but for Teams bot it might be simplified.
            // Currently, teams_bot.js only uses email and password (for now).
          });
          added++;
        }
      }
      if (added > 0) {
        bot.sendMessage(chatId, `Successfully added ${added} accounts.`, mainMenu);
        session.step = "IDLE";
      } else {
        bot.sendMessage(chatId, "Invalid format. Use email|password");
      }
    } else if (session.step === "SET_CONCURRENCY") {
      const num = parseInt(text);
      if (!isNaN(num) && num > 0) {
        const userConf = await getUserConfig(chatId);
        userConf.concurrencyLimit = num;
        await userConf.save();
        bot.sendMessage(chatId, `Concurrency limit updated to ${num}.`, mainMenu);
        session.step = "IDLE";
      } else {
        bot.sendMessage(chatId, "Invalid number. Update cancelled.", mainMenu);
        session.step = "IDLE";
      }
    }
  });
}

startBot();
