const dotenv = require('dotenv');
dotenv.config();

module.exports = {
  telegram: {
    token: process.env.TELEGRAM_BOT_TOKEN,
    logChatId: process.env.TELEGRAM_LOG_CHAT_ID,
  },
  database: {
    uri: process.env.MONGODB_URI,
  },
  headless: process.env.HEADLESS === 'true',
  chromiumPath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH,
};
