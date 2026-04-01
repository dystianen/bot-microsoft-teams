const TelegramBot = require("node-telegram-bot-api");
const config = require("./config");
const { processSingleAccount } = require("./index");
const connectDB = require("./db");
const { SuccessAccount, UserConfig } = require("./models");
const date = require("date-and-time");

// Wrap in an async function to allow awaiting connection
async function startBot() {
  try {
    // Connect to MongoDB
    await connectDB();

    const token = config.telegram.token;

    if (!token) {
      console.error("Please set TELEGRAM_BOT_TOKEN in .env and restart.");
      process.exit(1);
    }

    const bot = new TelegramBot(token, { polling: true });

    // Move the initialization of everything that depends on 'bot' inside
    initializeBotHandlers(bot);

    console.log("Teams Bot (Playwright Local) started.");
  } catch (err) {
    console.error("Failed to start bot:", err.message);
    process.exit(1);
  }
}

// Separate handlers to keep it clean
function initializeBotHandlers(bot) {
  // Sequential message queue for Telegram
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

  function escapeHTML(str) {
    if (!str) return "";
    return str
      .toString()
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  // Memory storage for temporary interactive steps and pending accounts
  const sessions = {};

  async function getUserConfig(telegram_id) {
    let userConf = await UserConfig.findOne({
      telegram_id: telegram_id.toString(),
    });
    if (!userConf) {
      userConf = new UserConfig({
        telegram_id: telegram_id.toString(),
        concurrencyLimit: 5, // Default set to 5 for "barengan" launch
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

    bot.sendMessage(
      chatId,
      "Welcome to Microsoft Teams Bot! 🤖 (Playwright Local)\n\nAdd accounts as email|password and they will be launched concurrently with a 5s stagger.",
      mainMenu,
    );
  });

  bot.onText(/➕ Add Account/, (msg) => {
    const chatId = msg.chat.id;
    sessions[chatId] = sessions[chatId] || { accounts: [], step: "IDLE" };
    sessions[chatId].step = "WAIT_ACCOUNT";
    bot.sendMessage(
      chatId,
      "Send email|password (one per line):",
      { parse_mode: "Markdown" },
    );
  });

  bot.onText(/🚀 Generate/, async (msg) => {
    const chatId = msg.chat.id;
    const session = sessions[chatId] || { accounts: [], running: false };

    if (session.running) {
      return bot.sendMessage(chatId, "⚠️ Automation is already running!");
    }

    if (session.accounts.length === 0) {
      return bot.sendMessage(
        chatId,
        "Please add accounts first.",
      );
    }

    const userConf = await getUserConfig(chatId);
    session.running = true;
    sessions[chatId] = session;

    bot.sendMessage(
      chatId,
      `🚀 Starting batch for ${session.accounts.length} accounts...\n(Launching every 5 seconds, max ${userConf.concurrencyLimit} active windows)`,
    );

    const runQueue = async () => {
      let activeWorkers = 0;
      const maxWorkers = userConf.concurrencyLimit;
      let globalIdx = 0;
      const originalTotal = session.accounts.length;
      const pendingPromises = new Set();

      const processAccount = async (accountData, currentIdx) => {
        await safeSendMessage(
          chatId,
          `⏳ [${currentIdx}/${originalTotal}] Processing: ${escapeHTML(accountData.email)}`,
        );

        const pairedData = {
          microsoftAccount: accountData,
          telegram_id: chatId,
          productUrl: userConf.microsoftUrl,
          headless: userConf.headless, // Pass the user's preference
        };

        try {
          const result = await processSingleAccount(pairedData, currentIdx - 1, originalTotal);

          if (result.status === "SUCCESS") {
            let message = `✅ <b>Success [${currentIdx}/${originalTotal}]</b>\n`;
            message += `Time: <code>${date.format(new Date(), "DD MMM YYYY HH:mm", true)}</code>\n`;
            message += `Email: <code>${escapeHTML(accountData.email)}</code>\n`;
            await safeSendMessage(chatId, message, { parse_mode: "HTML" });
            
            const successAcc = new SuccessAccount({
                email: accountData.email,
                password: accountData.password,
                telegram_id: chatId.toString(),
            });
            await successAcc.save().catch(e => {});
          } else {
            let message = `❌ <b>Failed [${currentIdx}/${originalTotal}] for ${escapeHTML(accountData.email)}</b>\n`;
            message += `Log: ${escapeHTML(result.log || "Unknown error")}`;
            await safeSendMessage(chatId, message, { parse_mode: "HTML" });
          }
        } catch (err) {
          await safeSendMessage(chatId, `❌ Error: ${escapeHTML(err.message)}`);
        }
      };

      try {
        while (true) {
          // If we can start a new worker (below concurrency limit AND have accounts left)
          if (activeWorkers < maxWorkers && session.accounts.length > 0 && !session.forceStop) {
            const accountData = session.accounts.shift();
            globalIdx++;
            const currentIdx = globalIdx;

            activeWorkers++;
            // Launch in background
            const promise = processAccount(accountData, currentIdx).finally(() => {
              activeWorkers--;
              pendingPromises.delete(promise);
            });
            pendingPromises.add(promise);

            // STAGGER: ALWAYS wait 5s after starting each new window, 
            // unless we've reached the concurrency limit or no accounts left.
            if (session.accounts.length > 0 && activeWorkers < maxWorkers) {
              await new Promise((r) => setTimeout(r, 5000));
            }
          } 
          // If the queue is finished or forced to stop, wait for remaining workers to finish then break
          else if (activeWorkers === 0 && (session.accounts.length === 0 || session.forceStop)) {
            break;
          } 
          // If we are at concurrency limit, wait 1s and check again
          else {
            await new Promise((r) => setTimeout(r, 1000));
          }
        }

        if (pendingPromises.size > 0) {
          await Promise.all(Array.from(pendingPromises));
        }
      } catch (err) {
        console.error("[Queue Error]", err);
        await safeSendMessage(chatId, `⚠️ Queue error: ${escapeHTML(err.message)}`);
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
    bot.sendMessage(
      chatId,
      "Session cleared.",
      mainMenu,
    );
  });

  bot.onText(/⚙️ Config/, async (msg) => {
    const chatId = msg.chat.id;
    const userConf = await getUserConfig(chatId);

    const options = {
      reply_markup: {
        inline_keyboard: [
          [
            {
              text: `🚀 Concurrency: ${userConf.concurrencyLimit}`,
              callback_data: `set_concurrency`,
            },
            {
              text: `👁️ Headless: ${userConf.headless ? "Active" : "Inactive"}`,
              callback_data: `toggle_headless`,
            },
          ],
          [
            {
              text: `📞 Select Microsoft Copilot`,
              callback_data: `select_copilot`,
            },
            {
              text: `📺 Select Microsoft Teams Room`,
              callback_data: `select_teams`,
            },
          ],
        ],
      },
    };

    bot.sendMessage(
      chatId,
      `⚙️ <b>Current Configuration:</b>\n\n` +
      `<b>Concurrency:</b> ${userConf.concurrencyLimit}\n` +
      `<b>Headless Mode:</b> <code>${userConf.headless ? "Active (No window)" : "Inactive (Visible window)"}</code>\n` +
      `<b>Active URL:</b> <code>${escapeHTML(userConf.microsoftUrl)}</code>\n\n` +
      `Click buttons below to change settings.`,
      { parse_mode: "HTML", ...options },
    );
  });

  bot.on("callback_query", async (callbackQuery) => {
    const message = callbackQuery.message;
    const chatId = message.chat.id;
    const data = callbackQuery.data;

    sessions[chatId] = sessions[chatId] || {
      accounts: [],
      step: "IDLE",
      running: false,
    };

    if (data === "set_concurrency") {
      sessions[chatId].step = "SET_CONCURRENCY";
      bot.sendMessage(
        chatId,
        "How many accounts should run together? (Enter a number, e.g., 5)",
      );
    } else if (data === "toggle_headless") {
      const userConf = await getUserConfig(chatId);
      userConf.headless = !userConf.headless;
      userConf.updatedAt = new Date();
      await userConf.save();
      bot.answerCallbackQuery(callbackQuery.id, { text: `Browser window now: ${userConf.headless ? 'Hidden' : 'Visible'}` });
      bot.sendMessage(chatId, `✅ <b>Headless Mode:</b> ${userConf.headless ? "Active" : "Inactive"}`, { parse_mode: "HTML" });
    } else if (data === "select_copilot") {
      const userConf = await getUserConfig(chatId);
      userConf.microsoftUrl = "https://admin.cloud.microsoft/?#/catalog/m/offer-details/microsoft-365-copilot/CFQ7TTC0MM8R";
      userConf.updatedAt = new Date();
      await userConf.save();
      bot.answerCallbackQuery(callbackQuery.id, { text: "📞 Microsoft Copilot Selected" });
      bot.sendMessage(chatId, "✅ <b>Product Set to:</b> Microsoft 365 Copilot", { parse_mode: "HTML" });
    } else if (data === "select_teams") {
      const userConf = await getUserConfig(chatId);
      userConf.microsoftUrl = "https://admin.cloud.microsoft/?#/catalog/m/offer-details/microsoft-teams-rooms-basic/CFQ7TTC0QW5P";
      userConf.updatedAt = new Date();
      await userConf.save();
      bot.answerCallbackQuery(callbackQuery.id, { text: "📺 Microsoft Teams Room Selected" });
      bot.sendMessage(chatId, "✅ <b>Product Set to:</b> Microsoft Teams Rooms Basic", { parse_mode: "HTML" });
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
            password: parts[parts.length - 1],
          });
          added++;
        }
      }
      if (added > 0) {
        bot.sendMessage(chatId, `Successfully added ${added} accounts.`, mainMenu);
        session.step = "IDLE";
      }
    } else if (session.step === "SET_CONCURRENCY") {
      const num = parseInt(text);
      if (!isNaN(num) && num > 0) {
        const userConf = await getUserConfig(chatId);
        userConf.concurrencyLimit = num;
        await userConf.save();
        bot.sendMessage(chatId, `Concurrency updated to ${num}. Bot will now open up to ${num} windows together with 5s delay.`, mainMenu);
        session.step = "IDLE";
      }
    }
  });
}

startBot();
