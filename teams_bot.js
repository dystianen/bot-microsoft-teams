const { chromium } = require("playwright-core");
const fs = require("fs");
const config = require("./config");

const SPINNER_SELECTOR = '[data-testid="spinner"], .ms-Spinner, [class*="spinner" i]';
const HARD_TIMEOUT = 1.5 * 60 * 1000;

class TeamsBot {
  constructor(wsUrl, accountConfig) {
    this.wsUrl = wsUrl;
    this.browser = null;
    this.context = null;
    this.page = null;
    this.accountConfig = accountConfig;
  }

  async humanDelay(min = 500, max = 1500) {
    const delay = Math.floor(Math.random() * (max - min + 1) + min);
    await this.page.waitForTimeout(delay);
  }

  async randomMouseMove() {
    const { width, height } = this.page.viewportSize() || { width: 1280, height: 720 };
    const x = Math.floor(Math.random() * width);
    const y = Math.floor(Math.random() * height);
    await this.page.mouse.move(x, y, { steps: 10 });
  }

  async runWithMonitor(promise, timeout = HARD_TIMEOUT) {
    let isDone = false;
    let errorMsg = null;

    const checkLoop = async () => {
      while (!isDone) {
        await this.page.waitForTimeout(2000).catch(() => { isDone = true; });
        if (isDone) break;
        // Simplified error check for now, can be expanded like in microsoft_bot.js
      }
    };

    const result = await Promise.race([promise, checkLoop()]).finally(() => {
      isDone = true;
    });

    if (errorMsg) {
      throw new Error(`MICROSOFT_ERROR: ${errorMsg}`);
    }

    return result;
  }

  async waitForSpinnerGone(extraDelay = 0) {
    const spinner = this.page.locator(SPINNER_SELECTOR).first();
    const spinnerVisible = await spinner.isVisible().catch(() => false);

    if (spinnerVisible) {
      console.log("[WAIT] Spinner detected, waiting until hidden...");
      try {
        await this.runWithMonitor(
          spinner.waitFor({ state: "hidden", timeout: HARD_TIMEOUT }),
        );
      } catch (e) {
        console.log("[WAIT] Spinner still visible or check failed, continuing...");
      }
      console.log("[WAIT] Spinner gone.");
    }

    if (extraDelay > 0) {
      await this.humanDelay(extraDelay, extraDelay + 300);
    }
  }

  async waitForVisible(locator) {
    await this.waitForSpinnerGone();
    await this.runWithMonitor(
      locator.waitFor({ state: "visible", timeout: HARD_TIMEOUT }),
    );
  }

  async clickButtonWithPossibleNames(names) {
    await this.waitForSpinnerGone();

    const keywords = names.flatMap((n) => n.trim().toLowerCase().split(/\s+/));
    const uniqueKeywords = [...new Set(keywords)];

    const found = await this.page.evaluate((keywords) => {
      const candidates = [
        ...document.querySelectorAll('button, [role="button"], a[role="button"], input[type="button"], input[type="submit"]'),
      ];

      const el = candidates.find((b) => {
        const text = (b.textContent || b.value || b.getAttribute("aria-label") || "").trim().toLowerCase();
        return text.length > 0 && text.length < 60 && keywords.some((kw) => text.includes(kw));
      });

      if (!el) return null;
      el.click();
      return el.textContent?.trim() || el.value || "unknown";
    }, uniqueKeywords);

    if (found) {
      console.log(`[INFO] Clicked: "${found}"`);
      return true;
    }

    console.log("[WARN] JS click not found, fallback to Playwright...");
    const pattern = new RegExp(names.map((n) => n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s*")).join("|"), "i");
    const button = this.page.getByRole("button", { name: pattern }).first();

    try {
      await button.waitFor({ state: "visible", timeout: HARD_TIMEOUT });
      await this.randomMouseMove();
      await this.humanDelay(500, 1000);
      await button.click({ timeout: 8000, force: true });
      const clickedText = await button.textContent().catch(() => "unknown");
      console.log(`[INFO] Clicked: "${clickedText?.trim()}"`);
      return true;
    } catch (err) {
      console.error(`[ERROR] Button not found for keywords:`, uniqueKeywords);
      throw err;
    }
  }

  getGenericLocator(keyword, elementType = "input") {
    return this.page
      .locator(`${elementType}[id*="${keyword}" i], ${elementType}[data-testid*="${keyword}" i], ${elementType}[name*="${keyword}" i], ${elementType}[aria-label*="${keyword}" i]`)
      .first();
  }

  async connect() {
    console.log("[STEP 1] Connecting to browser...");
    this.browser = await chromium.connectOverCDP(this.wsUrl);
    const contexts = this.browser.contexts();
    this.context = contexts.length > 0 ? contexts[0] : await this.browser.newContext();
    const pages = this.context.pages();
    this.page = pages.length > 0 ? pages[0] : await this.context.newPage();
  }

  async run() {
    try {
      await this.connect();
      
      // 2. buka https://admin.microsoft.com/
      console.log("[STEP 2] Opening https://admin.microsoft.com/...");
      await this.page.goto("https://admin.microsoft.com/", { waitUntil: "domcontentloaded", timeout: HARD_TIMEOUT });
      await this.waitForSpinnerGone();

      // 3. masukin email click next
      const email = this.accountConfig.microsoftAccount.email;
      console.log("[STEP 3] Entering email:", email);
      const emailInput = this.getGenericLocator("email");
      await this.waitForVisible(emailInput);
      await emailInput.fill(email);
      await this.humanDelay(500, 1000);
      await this.clickButtonWithPossibleNames(["Next", "Selanjutnya", "Berikutnya"]);

      // 3.5 Handle "Choose a way to sign in" if it appears
      console.log("[STEP 3.5] Checking for 'Choose a way to sign in' prompt...");
      const usePasswordPrompt = this.page.locator('div[role="button"][aria-label*="Use my password" i], div[role="button"]:has-text("Use my password")').first();
      try {
        await usePasswordPrompt.waitFor({ state: "visible", timeout: 5000 });
        console.log("[INFO] 'Choose a way to sign in' detected, clicking 'Use my password'...");
        await usePasswordPrompt.click();
        await this.humanDelay(1000, 2000);
      } catch (e) {
        console.log("[INFO] No 'Choose a way to sign in' prompt found, continuing...");
      }

      // 4. masukin password
      const password = this.accountConfig.microsoftAccount.password;
      console.log("[STEP 4] Entering password...");
      const passwordInput = this.page.locator('input[type="password"]').first();
      await this.waitForVisible(passwordInput);
      await passwordInput.fill(password);
      await this.humanDelay(500, 1000);
      await this.clickButtonWithPossibleNames(["Sign in", "Masuk"]);

      // 5. di halaman stay signed in, click yes
      console.log("[STEP 5] Handling 'Stay signed in'...");
      await this.clickButtonWithPossibleNames(["Yes", "Ya"]).catch(() => {
        console.log("[INFO] 'Stay signed in' prompt not found or failed, continuing...");
      });

      // 6. select menu users (collapse dia)
      console.log("[STEP 6] Selecting 'Users' menu...");
      
      // Wait for the side navigation to be present and stable
      const navLocator = this.page.locator('[data-hint="ReactLeftNav"]').first();
      await navLocator.waitFor({ state: "visible", timeout: 30000 });
      await this.humanDelay(2000, 3000); // Give it extra time to settle

      const usersMenu = this.page.locator('button[data-automation-id="LeftNavusersnodeNavToggler"], button[name="Users"], button:has-text("Users")').first();
      await this.waitForVisible(usersMenu);
      
      // Check if it's already expanded. If aria-expanded is 'true', don't click it again to collapse it.
      const isExpanded = await usersMenu.getAttribute("aria-expanded");
      if (isExpanded !== "true") {
        console.log("[INFO] Users menu is collapsed, clicking to expand...");
        await usersMenu.click();
        await this.humanDelay(1000, 2000);
      } else {
        console.log("[INFO] Users menu is already expanded.");
      }

      // 7. terus pilih activer users
      console.log("[STEP 7] Selecting 'Active users'...");
      const activeUsersLink = this.page.locator('a[data-automation-id="LeftNavactiveusersnodeNavToggler"], a:has-text("Active users"), a:has-text("Pengguna aktif")').first();
      await this.waitForVisible(activeUsersLink);
      await activeUsersLink.click();
      
      // Wait for the active users page to load
      console.log("[INFO] Waiting for Active users page to load...");
      await this.page.waitForLoadState("networkidle").catch(() => {});
      await this.waitForSpinnerGone(3000);

      // 8. terus click display name yg ada di list (Baris Pertama)
      console.log("[STEP 8] Clicking the first display name from the list...");
      
      const firstRowDisplayName = this.page.locator('div[data-automation-key="DisplayName"] span[role="button"], [role="gridcell"] button, [role="row"] button').first();
      await this.waitForVisible(firstRowDisplayName);
      
      const name = await firstRowDisplayName.textContent();
      console.log(`[INFO] Clicking display name: "${name?.trim()}"`);
      await firstRowDisplayName.click();
      
      await this.humanDelay(2000, 3000);
      console.log("[SUCCESS] Automation finished up to step 8.");
      return { success: true };

    } catch (error) {
      console.error("[ERROR] Automation failed:", error.message);
      return { success: false, error: error.message };
    }
  }

  async cleanup() {
    if (this.browser) {
      await this.browser.close();
    }
  }
}

module.exports = TeamsBot;
