const dotenv = require("dotenv");
dotenv.config();

module.exports = {
  telegram: {
    token: process.env.TELEGRAM_BOT_TOKEN,
  },
  headless: process.env.HEADLESS === "true",
};
