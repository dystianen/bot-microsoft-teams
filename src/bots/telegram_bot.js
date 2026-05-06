const TelegramBot = require('node-telegram-bot-api');
const config = require('../config');
const { processSingleAccount } = require('../index');
const connectDB = require('../db/connection');
const { AccountHistory, UserConfig } = require('../db/models');
const remoteLogger = require('../utils/logger');

// Global state for graceful shutdown
let isShuttingDown = false;
let activeAccountsCount = 0;

// Wrap in an async function to allow awaiting connection
async function startBot() {
  try {
    // Connect to MongoDB
    await connectDB();

    const token = config.telegram.token;

    if (!token) {
      console.error('Please set TELEGRAM_BOT_TOKEN in .env and restart.');
      process.exit(1);
    }

    const bot = new TelegramBot(token, {
      polling: {
        interval: 300,
        autoStart: true,
        params: {
          timeout: 50,
        },
      },
    });

    // Listen for polling errors to avoid unhandled crashes and log them properly
    bot.on('polling_error', (err) => {
      if (err.code === 'EFATAL' || err.message.includes('ETIMEDOUT')) {
        console.warn('[Telegram] Polling timeout detected. Bot will automatically retry...');
      } else {
        console.error('[Telegram] Polling Error:', err.code, err.message);
      }
    });

    bot.on('error', (err) => {
      console.error('[Telegram] General Error:', err);
    });

    // Move the initialization of everything that depends on 'bot' inside
    initializeBotHandlers(bot);

    console.log('Teams Bot (Playwright Local) started.');

    // Graceful Shutdown Handler
    const shutdown = () => {
      if (isShuttingDown) return;
      console.log('\n[Graceful Shutdown] Signal received. Finishing current tasks...');
      isShuttingDown = true;

      // Stop receiving new commands from Telegram
      bot.stopPolling();

      if (activeAccountsCount === 0) {
        console.log('[Graceful Shutdown] No active tasks. Exiting now.');
        process.exit(0);
      } else {
        console.log(
          `[Graceful Shutdown] Waiting for ${activeAccountsCount} active accounts to finish...`
        );
        // Safety timeout to prevent hanging forever
        setTimeout(() => {
          console.log('[Graceful Shutdown] Timeout reached. Force exiting.');
          process.exit(1);
        }, 290000); // 4.8 minutes (matches kill_timeout in pm2)
      }
    };

    process.on('SIGINT', shutdown);
    process.on('SIGTERM', shutdown);
  } catch (err) {
    console.error('Failed to start bot:', err.message);
    process.exit(1);
  }
}

// Separate handlers to keep it clean
function initializeBotHandlers(bot) {
  // Sequential message queue for Telegram
  let _msgQueue = Promise.resolve();
  async function safeSendMessage(chatId, text, options = {}) {
    _msgQueue = _msgQueue
      .then(async () => {
        try {
          await bot.sendMessage(chatId, text, options);
        } catch (err) {
          console.error('[Telegram] Error sending message:', err.message);
        }
        await new Promise((r) => setTimeout(r, 500));
      })
      .catch(() => {})
      .then(() => Promise.resolve()); // reset ke fresh promise supaya chain tidak numpuk
    return _msgQueue;
  }

  function escapeHTML(str) {
    if (!str) return '';
    return str.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  // Memory storage for temporary interactive steps and pending accounts
  const sessions = {};
  const SESSION_TTL = 30 * 60 * 1000; // 30 menit

  // Helper: ambil atau buat session, sekaligus update lastActivity
  function getSession(chatId) {
    if (!sessions[chatId]) {
      sessions[chatId] = {
        accounts: [],
        step: 'IDLE',
        running: false,
        lastActivity: Date.now(),
      };
    }
    sessions[chatId].lastActivity = Date.now();
    return sessions[chatId];
  }

  // Cleanup session yang tidak aktif tiap 10 menit
  setInterval(
    () => {
      const now = Date.now();
      for (const chatId in sessions) {
        if (now - sessions[chatId].lastActivity > SESSION_TTL) {
          console.log(`[Session] Cleaning up inactive session for chatId: ${chatId}`);
          delete sessions[chatId];
        }
      }
    },
    10 * 60 * 1000
  );

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
        [{ text: '➕ Add Account' }],
        [{ text: '🚀 Generate' }, { text: '🛑 Stop Queue' }],
        [{ text: '⚙️ Config' }, { text: '🧹 Reset Session' }],
        [{ text: '📜 History' }, { text: `🗑️ Delete History` }],
      ],
      resize_keyboard: true,
    },
  };

  bot.onText(/\/start/, async (msg) => {
    const chatId = msg.chat.id;
    await getUserConfig(chatId);
    const session = getSession(chatId);
    session.accounts = [];
    session.step = 'IDLE';
    session.running = false;

    bot.sendMessage(
      chatId,
      'Welcome to Microsoft Teams Bot! 🤖 (Playwright Local)\n\nAdd accounts as email|password and they will be launched concurrently with a 5s stagger.',
      mainMenu
    );
  });

  bot.onText(/➕ Add Account/, (msg) => {
    const chatId = msg.chat.id;
    const session = getSession(chatId);
    session.step = 'WAIT_ACCOUNT';
    bot.sendMessage(chatId, 'Send email|password (one per line):', {
      parse_mode: 'Markdown',
    });
  });

  bot.onText(/🚀 Generate/, async (msg) => {
    const chatId = msg.chat.id;
    const session = getSession(chatId);

    if (session.running) {
      return bot.sendMessage(chatId, '⚠️ Automation is already running!');
    }

    if (session.accounts.length === 0) {
      return bot.sendMessage(chatId, 'Please add accounts first.');
    }

    const userConf = await getUserConfig(chatId);
    session.running = true;

    await remoteLogger.reportSystemStatus(`(Queue Start - ${session.accounts.length} accts)`);

    bot.sendMessage(
      chatId,
      `🚀 Starting batch for ${session.accounts.length} accounts...\n(Max ${userConf.concurrencyLimit} active windows)`
    );

    const runQueue = async () => {
      let activeWorkers = 0;
      const maxWorkers = userConf.concurrencyLimit;
      let globalIdx = 0;

      // Snapshot: only process what was here when we started
      const accountsToProcess = [...session.accounts];
      session.accounts = []; // Clear queue to avoid re-processing same items

      const originalTotal = accountsToProcess.length;
      const pendingPromises = new Set();
      const queueResults = {
        all: [],
        success: [],
        failed: [],
      };

      const processAccount = async (accountData, currentIdx) => {
        await safeSendMessage(
          chatId,
          `⏳ [${currentIdx}/${originalTotal}] Processing: ${escapeHTML(accountData.email)}`
        );

        const pairedData = {
          microsoftAccount: accountData,
          telegram_id: chatId,
          productUrl: userConf.microsoftUrl,
          headless: userConf.headless,
        };

        activeAccountsCount++;
        try {
          let lastResult = null;
          let attempts = 0;
          const maxAttempts = 2; // Total attempt (1 initial + 1 retry)

          while (attempts < maxAttempts) {
            attempts++;
            const result = await processSingleAccount(pairedData, currentIdx - 1, originalTotal);
            lastResult = result;

            if (result.status === 'SUCCESS') break;

            // Detect if the error is retryable (system-related)
            const errMsg = (result.log || '').toLowerCase();
            const isRetryable =
              (errMsg.includes('something went wrong') ||
                errMsg.includes('terjadi kesalahan') ||
                errMsg.includes('une erreur s\'est produite') ||
                errMsg.includes('microsoft_error') ||
                errMsg.includes('system_error')) &&
              !errMsg.includes('something happened') &&
              !errMsg.includes('terjadi sesuatu') &&
              !errMsg.includes('quelque chose s\'est passé') &&
              !errMsg.includes('password') &&
              !errMsg.includes('sandi') &&
              !errMsg.includes('mot de passe') &&
              !errMsg.includes('incorrect') &&
              !errMsg.includes('recognized') &&
              !errMsg.includes('reconnu') &&
              !errMsg.includes('dikenali');

            if (!isRetryable || attempts >= maxAttempts) break;

            // Inform user about the retry
            await safeSendMessage(
              chatId,
              `🔄 <b>Retry [${attempts}/${maxAttempts - 1}]</b> for <code>${escapeHTML(accountData.email)}</code>\nError sistem terdeteksi, mencoba ulang otomatis dalam 5 detik...`,
              { parse_mode: 'HTML' }
            );

            // Wait before starting over with a fresh browser
            await new Promise((r) => setTimeout(r, 5000));
          }

          const result = lastResult;

          const historyRecord = new AccountHistory({
            email: accountData.email,
            password: accountData.password,
            telegram_id: chatId.toString(),
            status: result.status,
            log: result.log,
          });
          await historyRecord.save().catch((e) => {});

          if (result.status === 'SUCCESS') {
            const resItem = {
              email: accountData.email,
              password: accountData.password,
              status: 'SUCCESS',
              finishedAt: new Date(),
            };
            queueResults.success.push(resItem);
            queueResults.all.push(resItem);

            let message = `✅ <b>Success [${currentIdx}/${originalTotal}]</b>\n`;
            // Calculate WIB (UTC+7) manually and format without the 'true' flag
            const timeWib = new Date()
              .toLocaleString('en-GB', {
                timeZone: 'Asia/Jakarta',
                day: '2-digit',
                month: 'short',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
                hour12: false,
              })
              .replace(',', '');
            message += `Time: <code>${timeWib}</code>\n`;
            message += `Email: <code>${escapeHTML(accountData.email)}</code>\n`;
            await safeSendMessage(chatId, message, { parse_mode: 'HTML' });
          } else {
            const resItem = {
              email: accountData.email,
              password: accountData.password,
              status: 'FAILED',
              log: result.log,
              finishedAt: new Date(),
            };
            queueResults.failed.push(resItem);
            queueResults.all.push(resItem);

            let message = `❌ <b>Failed [${currentIdx}/${originalTotal}] for ${escapeHTML(accountData.email)}</b>\n`;
            message += `Log: ${escapeHTML(result.log || 'Unknown error')}`;
            await safeSendMessage(chatId, message, { parse_mode: 'HTML' });
          }
        } catch (err) {
          const resItem = {
            email: accountData.email,
            password: accountData.password,
            status: 'FAILED',
            log: err.message,
            finishedAt: new Date(),
          };
          queueResults.failed.push(resItem);
          queueResults.all.push(resItem);

          await safeSendMessage(chatId, `❌ Error: ${escapeHTML(err.message)}`);
          const failRecord = new AccountHistory({
            email: accountData.email,
            password: accountData.password,
            telegram_id: chatId.toString(),
            status: 'FAILED',
            log: err.message,
          });
          await failRecord.save().catch(() => {});
        } finally {
          activeAccountsCount--;
          if (isShuttingDown && activeAccountsCount === 0) {
            console.log('[Graceful Shutdown] Last active task finished. Exiting process.');
            process.exit(0);
          }
        }
      };

      try {
        while (true) {
          if (
            activeWorkers < maxWorkers &&
            accountsToProcess.length > 0 &&
            !session.forceStop &&
            !isShuttingDown
          ) {
            const accountData = accountsToProcess.shift();
            globalIdx++;
            const currentIdx = globalIdx;

            activeWorkers++;
            const promise = processAccount(accountData, currentIdx).finally(() => {
              activeWorkers--;
              pendingPromises.delete(promise);
            });
            pendingPromises.add(promise);

            if (accountsToProcess.length > 0 && activeWorkers < maxWorkers) {
              await new Promise((r) => setTimeout(r, 2500));
            }
          } else if (activeWorkers === 0 && (accountsToProcess.length === 0 || session.forceStop)) {
            break;
          } else {
            await new Promise((r) => setTimeout(r, 1000));
          }
        }

        if (pendingPromises.size > 0) {
          await Promise.all(Array.from(pendingPromises));
        }
      } catch (err) {
        console.error('[Queue Error]', err);
        await safeSendMessage(chatId, `⚠️ Queue error: ${escapeHTML(err.message)}`);
      } finally {
        session.running = false;
        const processedCount = queueResults.success.length + queueResults.failed.length;
        let summaryMsg = session.forceStop
          ? `🛑 <b>Batch Queue Stopped Manually</b>\n`
          : `🏁 <b>Batch Queue Finished</b>\n`;

        summaryMsg += `🔢 Total Queue: <code>${originalTotal}</code>\n`;
        summaryMsg += `✅ Success: <code>${queueResults.success.length}</code>\n`;
        summaryMsg += `❌ Failed: <code>${queueResults.failed.length}</code>\n`;

        if (processedCount < originalTotal) {
          summaryMsg += `🛑 Stopped: <code>${originalTotal - processedCount}</code> accounts skipped\n`;
        }
        summaryMsg += `\n`;

        if (queueResults.success.length > 0) {
          summaryMsg += `🟢 <b>SUCCESS LIST:</b>\n`;
          const items = [...queueResults.success].reverse();
          items.forEach((r, i) => {
            const timeStr = r.finishedAt.toLocaleString('en-GB', {
              timeZone: 'Asia/Jakarta',
              day: '2-digit',
              month: '2-digit',
              hour: '2-digit',
              minute: '2-digit',
              hour12: false,
            });
            const [datePart, timePart] = timeStr.split(', ');
            const [d, m] = datePart.split('/');
            const shortTime = `${d}/${m}, ${timePart}`;

            summaryMsg += `${i + 1}. ${shortTime}\n`;
            summaryMsg += `- <code>${escapeHTML(r.email)}</code>\n`;
            summaryMsg += `- <code>${escapeHTML(r.password)}</code>\n`;
            summaryMsg += `────────────────\n`;
          });
          summaryMsg += `\n`;
        }

        if (queueResults.failed.length > 0) {
          summaryMsg += `🔴 <b>FAILED LIST:</b>\n`;
          const items = [...queueResults.failed].reverse();
          items.forEach((r, i) => {
            const timeStr = r.finishedAt.toLocaleString('en-GB', {
              timeZone: 'Asia/Jakarta',
              day: '2-digit',
              month: '2-digit',
              hour: '2-digit',
              minute: '2-digit',
              hour12: false,
            });
            const [datePart, timePart] = timeStr.split(', ');
            const [d, m] = datePart.split('/');
            const shortTime = `${d}/${m}, ${timePart}`;

            summaryMsg += `${i + 1}. ${shortTime}\n`;
            summaryMsg += `- <code>${escapeHTML(r.email)}</code>\n`;
            summaryMsg += `- <code>${escapeHTML(r.password)}</code>\n`;
            summaryMsg += `⚠️ Log: <i>${escapeHTML(r.log || 'No log')}</i>\n`;
            summaryMsg += `────────────────\n`;
          });
        }

        const isManual = session.forceStop;
        if (isManual) session.forceStop = false;

        // Send summary to remoteLogger
        await remoteLogger.send(summaryMsg);
        await remoteLogger.reportSystemStatus(isManual ? '(Queue Stopped)' : '(Queue Finished)');

        // Also send detailed report to user DM directly (chunked)
        const CHUNK_SIZE = 4000;
        for (let i = 0; i < summaryMsg.length; i += CHUNK_SIZE) {
          const chunk = summaryMsg.substring(i, i + CHUNK_SIZE);
          await safeSendMessage(chatId, chunk, {
            parse_mode: 'HTML',
            reply_markup: mainMenu,
          });
        }
      }
    };

    runQueue();
  });

  bot.onText(/🛑 Stop Queue/, (msg) => {
    const chatId = msg.chat.id;
    const session = getSession(chatId);
    session.accounts = [];
    session.forceStop = true;
    bot.sendMessage(chatId, '🛑 Stopping queue...');
  });

  bot.onText(/📜 History/, async (msg) => {
    const chatId = msg.chat.id;
    try {
      const records = await AccountHistory.find({
        telegram_id: chatId.toString(),
      })
        .sort({ createdAt: -1 })
        .limit(100);

      if (records.length === 0) {
        return bot.sendMessage(chatId, '📭 No history found.');
      }

      const successRecords = records.filter((r) => r.status === 'SUCCESS');
      const failedRecords = records.filter((r) => r.status !== 'SUCCESS');

      let message = '📜 <b>Recent Account History (Last 100):</b>\n\n';

      let counter = 1;

      if (successRecords.length > 0) {
        message += '🟢 <b>SUCCESS LIST:</b>\n';
        successRecords.forEach((rec) => {
          const timeStr = rec.createdAt.toLocaleString('en-GB', {
            timeZone: 'Asia/Jakarta',
            day: '2-digit',
            month: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            hour12: false,
          });
          const [datePart, timePart] = timeStr.split(', ');
          const [d, m] = datePart.split('/');
          const shortTime = `${d}/${m}, ${timePart}`;

          message += `${counter++}. ${shortTime}\n`;
          message += `- <code>${escapeHTML(rec.email)}</code>\n`;
          message += `- <code>${escapeHTML(rec.password)}</code>\n`;
          message += '────────────────\n';
        });
        message += '\n';
      }

      if (failedRecords.length > 0) {
        message += '🔴 <b>FAILED LIST:</b>\n';
        failedRecords.forEach((rec) => {
          const timeStr = rec.createdAt.toLocaleString('en-GB', {
            timeZone: 'Asia/Jakarta',
            day: '2-digit',
            month: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            hour12: false,
          });
          const [datePart, timePart] = timeStr.split(', ');
          const [d, m] = datePart.split('/');
          const shortTime = `${d}/${m}, ${timePart}`;

          message += `${counter++}. ${shortTime}\n`;
          message += `- <code>${escapeHTML(rec.email)}</code>\n`;
          message += `- <code>${escapeHTML(rec.password)}</code>\n`;
          message += `⚠️ Log: ${escapeHTML(rec.log || 'No log')}\n`;
          message += '────────────────\n';
        });
      }

      bot.sendMessage(chatId, message, { parse_mode: 'HTML' });
    } catch (err) {
      bot.sendMessage(chatId, `❌ Error fetching history: ${err.message}`);
    }
  });

  bot.onText(/🗑️ Delete History/, async (msg) => {
    const chatId = msg.chat.id;
    try {
      await AccountHistory.deleteMany({ telegram_id: chatId.toString() });
      bot.sendMessage(chatId, '✅ <b>Your account history has been cleared.</b>', {
        parse_mode: 'HTML',
      });
    } catch (err) {
      bot.sendMessage(chatId, `❌ Error deleting history: ${err.message}`);
    }
  });

  bot.onText(/🧹 Reset Session/, (msg) => {
    const chatId = msg.chat.id;
    const session = getSession(chatId);
    session.accounts = [];
    session.step = 'IDLE';
    session.running = false;
    bot.sendMessage(chatId, 'Session cleared.', mainMenu);
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
              text: `👁️ Headless: ${userConf.headless ? 'Active' : 'Inactive'}`,
              callback_data: `toggle_headless`,
            },
          ],
          [
            {
              text: `📦 Change Product / URL`,
              callback_data: `show_product_menu`,
            },
          ],
        ],
      },
    };

    bot.sendMessage(
      chatId,
      `⚙️ <b>Current Configuration:</b>\n\n` +
        `<b>Concurrency:</b> ${userConf.concurrencyLimit}\n` +
        `<b>Headless Mode:</b> <code>${userConf.headless ? 'Active (No window)' : 'Inactive (Visible window)'}</code>\n` +
        `<b>Active URL:</b> <code>${escapeHTML(userConf.microsoftUrl)}</code>\n\n` +
        `Click buttons below to change settings.`,
      { parse_mode: 'HTML', ...options }
    );
  });

  bot.on('callback_query', async (callbackQuery) => {
    const message = callbackQuery.message;
    const chatId = message.chat.id;
    const data = callbackQuery.data;

    const session = getSession(chatId);

    if (data === 'set_concurrency') {
      session.step = 'SET_CONCURRENCY';
      bot.sendMessage(chatId, 'How many accounts should run together? (Enter a number, e.g., 5)');
    } else if (data === 'toggle_headless') {
      const userConf = await getUserConfig(chatId);
      userConf.headless = !userConf.headless;
      userConf.updatedAt = new Date();
      await userConf.save();
      bot.answerCallbackQuery(callbackQuery.id, {
        text: `Browser window now: ${userConf.headless ? 'Hidden' : 'Visible'}`,
      });
      bot.sendMessage(
        chatId,
        `✅ <b>Headless Mode:</b> ${userConf.headless ? 'Active' : 'Inactive'}`,
        { parse_mode: 'HTML' }
      );
    } else if (data === 'show_product_menu') {
      const options = {
        reply_markup: {
          inline_keyboard: [
            [{ text: 'Microsoft Copilot', callback_data: 'select_copilot' }],
            [{ text: 'Microsoft Teams Rooms Basic', callback_data: 'select_teams' }],
            [{ text: 'Business Apps (free)', callback_data: 'select_business_apps' }],
            [{ text: 'Microsoft 365 Lighthouse', callback_data: 'select_lighthouse' }],
          ],
        },
      };
      bot.sendMessage(chatId, '🛒 <b>Select a product to automate:</b>', {
        parse_mode: 'HTML',
        ...options,
      });
    } else if (data === 'select_copilot') {
      const userConf = await getUserConfig(chatId);
      userConf.microsoftUrl =
        'https://admin.cloud.microsoft/?#/catalog/m/offer-details/microsoft-365-copilot/CFQ7TTC0MM8R';
      userConf.updatedAt = new Date();
      await userConf.save();
      bot.answerCallbackQuery(callbackQuery.id, {
        text: '📞 Microsoft Copilot Selected',
      });
      bot.sendMessage(chatId, '✅ <b>Product Set to:</b> Microsoft 365 Copilot', {
        parse_mode: 'HTML',
      });
    } else if (data === 'select_teams') {
      const userConf = await getUserConfig(chatId);
      userConf.microsoftUrl =
        'https://admin.cloud.microsoft/?#/catalog/m/offer-details/microsoft-teams-rooms-basic/CFQ7TTC0QW5P';
      userConf.updatedAt = new Date();
      await userConf.save();
      bot.answerCallbackQuery(callbackQuery.id, {
        text: '📺 Microsoft Teams Room Selected',
      });
      bot.sendMessage(chatId, '✅ <b>Product Set to:</b> Microsoft Teams Rooms Basic', {
        parse_mode: 'HTML',
      });
    } else if (data === 'select_business_apps') {
      const userConf = await getUserConfig(chatId);
      userConf.microsoftUrl =
        'https://admin.cloud.microsoft/?#/catalog/m/offer-details/business-apps-free-/CFQ7TTC0LHZ0';
      userConf.updatedAt = new Date();
      await userConf.save();
      bot.answerCallbackQuery(callbackQuery.id, {
        text: '💼 Business Apps Free Selected',
      });
      bot.sendMessage(chatId, '✅ <b>Product Set to:</b> Business Apps (free)', {
        parse_mode: 'HTML',
      });
    } else if (data === 'select_lighthouse') {
      const userConf = await getUserConfig(chatId);
      userConf.microsoftUrl =
        'https://admin.cloud.microsoft/?ocid=cmm45ut5ap0#/catalog/m/offer-details/microsoft-365-lighthouse/CFQ7TTC0JW0V';
      userConf.updatedAt = new Date();
      await userConf.save();
      bot.answerCallbackQuery(callbackQuery.id, {
        text: '🏠 Microsoft 365 Lighthouse Selected',
      });
      bot.sendMessage(chatId, '✅ <b>Product Set to:</b> Microsoft 365 Lighthouse', {
        parse_mode: 'HTML',
      });
    }
  });

  bot.on('message', async (msg) => {
    const chatId = msg.chat.id;
    const text = msg.text;

    if (!text || text.startsWith('/')) return;
    const session = getSession(chatId);
    if (!session || session.step === 'IDLE') return;

    if (session.step === 'WAIT_ACCOUNT') {
      const lines = text.split('\n');
      let added = 0;
      for (const line of lines) {
        const parts = line.split('|').map((s) => s.trim());
        if (parts.length >= 2) {
          const email = parts[0];
          const password = parts[parts.length - 1];

          // Check for duplicate email in current session
          const isDuplicate = session.accounts.some(
            (acc) => acc.email.toLowerCase() === email.toLowerCase()
          );

          if (!isDuplicate) {
            session.accounts.push({
              email: email,
              password: password,
            });
            added++;
          }
        }
      }
      if (added > 0) {
        bot.sendMessage(chatId, `Successfully added ${added} accounts.`, mainMenu);
        session.step = 'IDLE';
      } else if (lines.length > 0) {
        bot.sendMessage(
          chatId,
          `No new accounts added (they might be duplicates or invalid format).`,
          mainMenu
        );
        session.step = 'IDLE';
      }
    } else if (session.step === 'SET_CONCURRENCY') {
      const num = parseInt(text);
      if (!isNaN(num) && num > 0) {
        const userConf = await getUserConfig(chatId);
        userConf.concurrencyLimit = num;
        await userConf.save();
        bot.sendMessage(
          chatId,
          `Concurrency updated to ${num}. Bot will now open up to ${num} windows together with 5s delay.`,
          mainMenu
        );
        session.step = 'IDLE';
      }
    }
  });
}

startBot();
