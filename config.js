const dotenv = require("dotenv");
dotenv.config();

module.exports = {
  adsPower: {
    baseUrl: process.env.ADSPOWER_BASE_URL || "http://127.0.0.1:50325",
    apiKey: process.env.ADSPOWER_API_KEY,
    groupId: process.env.ADSPOWER_GROUP_ID,
  },
  proxy: {
    host: process.env.PROXY_HOST,
    port: process.env.PROXY_PORT,
    username: process.env.PROXY_USERNAME,
    password: process.env.PROXY_PASSWORD,
    type: "socks5",
  },
  headless: process.env.HEADLESS === "true",
};
