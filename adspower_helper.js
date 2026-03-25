const axios = require("axios");
const config = require("./config");

class AdsPowerHelper {
  constructor() {
    this.baseUrl = (config.adsPower.baseUrl || "").replace(/\/+$/, "");
  }

  randomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  randomUserAgent() {
    const chromeVersions = ["124", "125", "126", "127", "128", "145"];
    const version = chromeVersions[Math.floor(Math.random() * chromeVersions.length)];
    const minor = this.randomInt(0, 9999);
    const patch = this.randomInt(0, 999);
    return `Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/${version}.0.${minor}.${patch} Safari/537.36`;
  }

  async createProfile(profileName, proxyOverride = null) {
    try {
      const proxy = {
        proxy_type: "socks5",
        proxy_host: config.proxy.host,
        proxy_port: config.proxy.port,
        proxy_user: proxyOverride?.username || config.proxy.username,
        proxy_password: proxyOverride?.password || config.proxy.password,
        proxy_soft: "other",
      };

      const response = await axios.post(`${this.baseUrl}/api/v1/user/create`, {
        name: profileName,
        group_id: config.adsPower.groupId,
        domain_name: "",
        os_type: "windows",
        repeat_config: ["0"],
        user_proxy_config: proxy,
        fingerprint_config: {
          automatic_timezone: "1",
          language: ["en-US", "en"],
          flash: "block",
          fonts: ["all"],
          webrtc: "disabled",
          canvas: "0",
          webgl: "0",
          webgl_image: "0",
          hardware_concurrency: "default",
          ua: this.randomUserAgent(),
        },
      }, {
        headers: {
          Authorization: `Bearer ${config.adsPower.apiKey}`,
          "api-key": config.adsPower.apiKey,
        },
      });

      if (response.data.code !== 0) throw new Error(`AdsPower Profile Creation Failed: ${response.data.msg}`);
      return response.data.data.id;
    } catch (error) {
      console.error("Error creating AdsPower profile:", error.message);
      throw error;
    }
  }

  async startBrowser(profileId, headlessOverride = null) {
    try {
      const isHeadless = headlessOverride !== null ? headlessOverride : config.headless;
      const headlessParam = isHeadless ? "&open_tabs=1&headless=1" : "";
      const response = await axios.get(`${this.baseUrl}/api/v1/browser/start?user_id=${profileId}${headlessParam}`, {
        headers: {
          Authorization: `Bearer ${config.adsPower.apiKey}`,
          "api-key": config.adsPower.apiKey,
        },
      });

      if (response.data.code !== 0) throw new Error(`Failed to start AdsPower browser: ${response.data.msg}`);

      const wsUrl = response.data?.data?.ws?.puppeteer || response.data?.data?.ws?.selenium;
      if (!wsUrl) throw new Error("Could not find WebSocket URL");
      return { wsUrl };
    } catch (error) {
      console.error("Error starting browser:", error.message);
      throw error;
    }
  }
}

module.exports = new AdsPowerHelper();
