const dotenv = require("dotenv");
dotenv.config();

module.exports = {
  telegram: {
    token: process.env.TELEGRAM_BOT_TOKEN,
    logChatId: process.env.TELEGRAM_LOG_CHAT_ID, // Logs
  },
  headless: process.env.HEADLESS === "true",
};
