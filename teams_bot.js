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

        const detectedError = await this.checkForError();
        if (detectedError) {
          errorMsg = detectedError;
          isDone = true;
          break;
        }
      }
    };

    const result = await Promise.race([
      promise, 
      checkLoop()
    ]).finally(() => {
      isDone = true;
    });

    if (errorMsg) {
      throw new Error(`MICROSOFT_ERROR: ${errorMsg}`);
    }

    return result;
  }

  async checkForError() {
    try {
      // 1. Cek pesan error validasi di field
      const fieldError = this.page.locator('[data-automation-id="error-message"], [id*="error" i]').first();
      if (await fieldError.isVisible().catch(() => false)) {
        const msg = (await fieldError.textContent().catch(() => "")).trim();
        if (msg) return `Field Error: ${msg}`;
      }

      // 2. Cek teks di SEMUA frame (termasuk iframe tersembunyi)
      const markers = [
        "something went wrong",
        "something happened",
        "terjadi sesuatu",
        "Terjadi kesalahan",
        "Melindungi akun Anda",
        "try a different way",
        "Protecting your account",
        "Please solve the puzzle",
        "error code",
        "715-123280",
      ];

      for (const frame of this.page.frames()) {
        try {
          const frameText = await frame.innerText('body').catch(() => "");
          const lowerFrameText = frameText.toLowerCase();
          const found = markers.find(m => lowerFrameText.includes(m.toLowerCase()));
          if (found) {
            return `Marker "${found}" detected in frame.`;
          }
        } catch (e) { /* skip inaccessible frames */ }
      }
    } catch (err) { /* ignore */ }
    return null;
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

    // 1. Coba klik di Main Page & Semua Frames menggunakan JS
    for (const frame of this.page.frames()) {
      try {
        const found = await frame.evaluate((keywords) => {
          const candidates = [...document.querySelectorAll('button, [role="button"], a[role="button"], input[type="button"], input[type="submit"]')];
          const el = candidates.find((b) => {
            const text = (b.textContent || b.value || b.getAttribute("aria-label") || "").trim().toLowerCase();
            return text.length > 0 && text.length < 60 && keywords.some((kw) => text.includes(kw));
          });
          if (!el) return null;
          el.click();
          return el.textContent?.trim() || el.value || "unknown";
        }, uniqueKeywords);

        if (found) {
          console.log(`[INFO] Clicked: "${found}" (in frame: ${frame.url() === this.page.url() ? "main" : "subframe"})`);
          return true;
        }
      } catch (e) {}
    }

    // 2. Fallback: Playwright native click di Main Page & Semua Frames
    console.log(`[INFO] Fallback to Playwright click for names: ${names.join(", ")}`);
    const pattern = new RegExp(names.map((n) => n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s*")).join("|"), "i");

    for (const frame of this.page.frames()) {
      try {
        const button = frame.getByRole("button", { name: pattern }).first();
        if (await button.isVisible({ timeout: 2000 }).catch(() => false)) {
          await this.randomMouseMove();
          const clickedText = await button.evaluate(el => (el.textContent || el.value || el.getAttribute("aria-label") || "").trim()).catch(() => "unknown");
          await button.click({ timeout: 5000, force: true });
          console.log(`[INFO] Clicked: "${clickedText || "unknown"}" (native, in frame)`);
          return true;
        }
      } catch (e) {}
    }

    console.error(`[ERROR] Button not found in any frame for names:`, names);
    throw new Error(`Button not found: ${names.join(", ")}`);
  }

  getGenericLocator(keyword, elementType = "input") {
    return this.page
      .locator(`${elementType}[id*="${keyword}" i], ${elementType}[data-testid*="${keyword}" i], ${elementType}[name*="${keyword}" i], ${elementType}[aria-label*="${keyword}" i]`)
      .first();
  }

  async connect() {
    if (this.wsUrl) {
      console.log("[STEP 1] Connecting to browser via Ads Power...");
      this.browser = await chromium.connectOverCDP(this.wsUrl);
      const contexts = this.browser.contexts();
      this.context = contexts.length > 0 ? contexts[0] : await this.browser.newContext();
    } else {
      console.log("[STEP 1] Launching local browser in incognito mode...");
      this.browser = await chromium.launch({ 
        headless: config.headless,
        args: ["--incognito", "--disable-blink-features=AutomationControlled"]
      });
      this.context = await this.browser.newContext();
    }
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

      // 5. Handling after Password (Robust State Monitoring)
      console.log("[STEP 5] Waiting for Dashboard or login prompts (KMSI/MFA)...");
      
      const dashboardMarker = this.page.locator('[data-hint="ReactLeftNav"], #admin-home-container').first();
      const loginLoopStart = Date.now();
      
      while (Date.now() - loginLoopStart < 120000) { // Max 2 menit menunggu dashboard
        // 5.1 Cek apakah sudah sampai Dashboard?
        if (await dashboardMarker.isVisible().catch(() => false)) {
          console.log("[SUCCESS] Dashboard detected!");
          break;
        }

        // 5.2 Cek rintangan: Stay signed in
        const yesBtn = this.page.locator('button:has-text("Yes"), input[value="Yes"], #idSIButton9').first();
        if (await yesBtn.isVisible().catch(() => false)) {
          console.log("[INFO] Handling 'Stay signed in'...");
          await yesBtn.click();
          await this.humanDelay(2000, 3000);
          continue;
        }

        // 5.3 Cek rintangan: MFA Skip
        const skipBtn = this.page.locator('a:has-text("Skip for now"), a:has-text("Lompati untuk sekarang"), button:has-text("Skip for now"), #idSecondaryButton').first();
        if (await skipBtn.isVisible().catch(() => false)) {
          console.log("[INFO] Handling MFA 'Skip for now'...");
          await skipBtn.click();
          await this.humanDelay(2000, 3000);
          continue;
        }

        // 5.4 Cek rintangan: Choose Account / Use Password
        const usePass = this.page.locator('text=Use my password, text=Gunakan kata sandi saya, #allowInterrupt').first();
        if (await usePass.isVisible().catch(() => false)) {
          console.log("[INFO] Handling 'Use my password' prompt...");
          await usePass.click();
          await this.humanDelay(2000, 3000);
          continue;
        }

        // 5.5 Cek Error Page
        const err = await this.checkForError();
        if (err) throw new Error(err);

        await this.page.waitForTimeout(2500); // Tunggu antar scan
      }

      // Verifikasi akhir sebelum lanjut ke Step 6
      if (!(await dashboardMarker.isVisible().catch(() => false))) {
        throw new Error("Login failed: Dashboard not reached within 2 minutes.");
      }

      // 6. select menu users (collapse dia)
      console.log("[STEP 6] Selecting 'Users' menu...");
      
      await this.waitForSpinnerGone(1000);
      const navLocator = this.page.locator('[data-hint="ReactLeftNav"]').first();
      await this.waitForVisible(navLocator);
      await this.humanDelay(1000, 2000);

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
      
      console.log("[INFO] Waiting for Active users list to appear...");
      await this.waitForSpinnerGone(1000);

      // 8. terus click display name yg ada di list (Baris Pertama)
      console.log("[STEP 8] Clicking the first display name from the list...");
      
      const firstRowDisplayName = this.page.locator('div[data-automation-key="DisplayName"] span[role="button"], [role="gridcell"] button, [role="row"] button').first();
      await this.waitForVisible(firstRowDisplayName);
      
      const name = await firstRowDisplayName.textContent();
      console.log(`[INFO] Clicking display name: "${name?.trim()}"`);
      await firstRowDisplayName.click();
      
      await this.humanDelay(2000, 3000);

      // 9. terus pilih yang licences dan apps
      console.log("[STEP 9] Selecting 'Licenses and apps' tab...");
      const licensesTab = this.page.locator('button[role="tab"]:has-text("Licenses and apps"), button:has-text("Licenses and apps")').first();
      await this.waitForVisible(licensesTab);
      await licensesTab.click();
      await this.waitForSpinnerGone(2000);

      // 10. uncheck office 365 itu
      console.log("[STEP 10] Unchecking Office 365 license...");
      // Typical selector for the license checkbox in the list
      const licenseCheckbox = this.page.locator('input[type="checkbox"][aria-label*="Office 365" i], input[type="checkbox"]:near(:text("Office 365"))').first();
      
      // If the above generic search fails, try searching specifically for the label "Office 365"
      const licenseLabel = this.page.locator('label:has-text("Office 365")').first();
      
      try {
        await this.page.waitForTimeout(2000); // Wait for load
        const isChecked = await licenseCheckbox.isChecked().catch(() => false);
        if (isChecked) {
          console.log("[INFO] License is checked, unchecking...");
          await licenseCheckbox.uncheck({ force: true });
        } else {
          // Alternative check via label click if checkbox locator is tricky
          console.log("[INFO] Checkbox state unsure, trying label click to ensure unchecked...");
          await licenseLabel.click().catch(() => {});
        }
      } catch (err) {
        console.warn("[WARN] Checkbox uncheck failed via standard ways, trying JS click on any match...");
        await this.page.evaluate(() => {
          const els = [...document.querySelectorAll('input[type="checkbox"]')];
          const office365 = els.find(el => el.parentElement?.textContent?.includes("Office 365") || el.getAttribute("aria-label")?.includes("Office 365"));
          if (office365 && office365.checked) office365.click();
        });
      }

      await this.humanDelay(1000, 2000);

      // 11. terus saves changes
      console.log("[STEP 11] Clicking 'Save changes'...");
      const saveBtn = this.page.locator('button:has-text("Save changes"), button[id*="save" i]').first();
      await this.waitForVisible(saveBtn);
      await saveBtn.click();
      
      // Wait for completion message
      console.log("[INFO] Waiting for save completion...");
      await this.waitForSpinnerGone(3000);
      // 12. New step: buka https://admin.cloud.microsoft/?#/catalog
      console.log("[STEP 12] Navigating to Marketplace catalog...");
      await this.page.goto("https://admin.cloud.microsoft/?#/catalog", { waitUntil: "domcontentloaded", timeout: HARD_TIMEOUT });
      await this.waitForSpinnerGone(3000);

      // Check for error: "You need a billing account owner..."
      const billingError = this.page.locator('div:has-text("You need a billing account owner or billing account contributor role to buy products")').first();
      const hasError = await billingError.isVisible({ timeout: 5000 }).catch(() => false);

      if (hasError) {
        console.warn("[WARN] Billing account role error detected. Stopping here.");
        return { success: false, error: "Billing account role error" };
      }

      // Check for 'Selecting a billing account' popup (blue popup in images)
      const billingAccountPopup = this.page.locator('div:has-text("Selecting a billing account")').first();
      const popupCloseBtn = this.page.locator('button[aria-label*="Close" i]').first();
      try {
        if (await billingAccountPopup.isVisible({ timeout: 5000 })) {
          console.log("[INFO] 'Selecting a billing account' popup detected, closing...");
          await popupCloseBtn.click();
          await this.humanDelay(1000, 2000);
        }
      } catch (e) {}

      // 13. Pilih tab all product
      console.log("[STEP 13] Selecting 'All products' tab...");
      const allProductsTab = this.page.locator('button[role="tab"]:has-text("All products"), button:has-text("All products")').first();
      await this.waitForVisible(allProductsTab);
      await allProductsTab.click();
      await this.waitForSpinnerGone(2000);

      // 14. Scroll ke bawah pilih yg copilot click details
      console.log("[STEP 14] Finding 'Microsoft 365 Copilot' and clicking 'Details'...");
      
      // Menggunakan data-automation-id persis dari HTML yang dberikan
      const copilotDetailsBtn = this.page.locator('[data-automation-id="NEW_PRODUCTS-Microsoft 365 Copilot-Tile"] button:has-text("Details"), button[aria-label*="View details for"][aria-label*="Microsoft 365 Copilot"]').first();
      
      try {
        await copilotDetailsBtn.scrollIntoViewIfNeeded();
        await this.waitForVisible(copilotDetailsBtn);
        await copilotDetailsBtn.click();
      } catch (err) {
        console.warn("[WARN] Primary Copilot details button locator failed, trying fallback...");
        const fallbackBtn = this.page.getByRole("heading", { name: "Microsoft 365 Copilot", exact: true })
          .locator("xpath=ancestor::div[contains(@class, 'offerTile')]//button[contains(., 'Details')]").first();
        await fallbackBtn.click();
      }

      // 15. Waiting spinner
      console.log("[STEP 15] Waiting for spinner after clicking Details...");
      await this.waitForSpinnerGone(5000);

      // 15.5 Select a plan
      console.log("[STEP 15.5] Selecting 'Microsoft 365 Copilot' plan...");
      const planDropdown = this.page.locator('div:has-text("Select a plan") select, [aria-label*="Select a plan" i], div:has-text("Select a plan") [role="combobox"]').first();
      await this.waitForVisible(planDropdown);
      
      try {
        const tagName = await planDropdown.evaluate(el => el.tagName.toLowerCase());
        if (tagName === "select") {
          await planDropdown.selectOption({ label: "Microsoft 365 Copilot" });
        } else {
          await planDropdown.click();
          await this.humanDelay(1000, 1500);
          const option = this.page.getByRole('option', {
            name: /^Microsoft 365 Copilot$/i
          });
          await option.click();
        }
      } catch (err) {
        console.warn("[WARN] Failed to select plan explicitly, it might be already selected. Continuing...");
      }
      await this.waitForSpinnerGone(2000);

      // 15.7 Select '1 year' commitment if present
      console.log("[STEP 15.7] Selecting '1 year' commitment...");
      try {
        // Menggunakan exact text agar bisa ngeklik langsung element aslinya tanpa peduli dia itu div, span, label, atau radio asli
        const oneYearText = this.page.getByText('1 year', { exact: true }).first();
        await oneYearText.waitFor({ state: "visible", timeout: 5000 });
        console.log("[INFO] '1 year' option found, clicking...");
        await oneYearText.click();
        await this.waitForSpinnerGone(2000);
      } catch (e) {
        console.log("[INFO] '1 year' option not found or already selected, continuing...");
      }

      // 16. Select 'Pay monthly'
      console.log("[STEP 16] Selecting 'Pay monthly' billing frequency...");
      try {
        const payMonthlyText = this.page.getByText('Pay monthly', { exact: true }).first();
        await payMonthlyText.waitFor({ state: "visible", timeout: 5000 });
        console.log("[INFO] 'Pay monthly' option found, clicking...");
        await payMonthlyText.click();
        await this.waitForSpinnerGone(2000);
      } catch (e) {
        console.log("[INFO] 'Pay monthly' option not found or already selected, continuing...");
      }

      // 17. Menunggu button buy muncul lalu click
      console.log("[STEP 17] Waiting for 'Buy' button and clicking...");
      const buyBtn = this.page.locator('button:has-text("Buy")').first();
      await this.waitForVisible(buyBtn);
      await buyBtn.click();
      await this.waitForSpinnerGone(3000);

      // 18. Muncul check box lalu check (OPTIONAL)
      console.log("[STEP 18] Checking authorization checkbox if present...");
      try {
        const authCheckbox = this.page.locator('input[type="checkbox"]').first();
        const checkboxFound = await authCheckbox.waitFor({ state: "visible", timeout: 3000 }).then(() => true).catch(() => false);
        
        if (checkboxFound) {
          const isAlreadyChecked = await authCheckbox.isChecked().catch(() => false);
          if (!isAlreadyChecked) {
            await authCheckbox.check({ force: true });
            console.log("[INFO] Authorization checkbox checked.");
          }
        } else {
          console.log("[INFO] No authorization checkbox found, proceeding to place order...");
        }
      } catch (err) {
        console.log("[INFO] Error handling checkbox, skipping...", err.message);
      }
      await this.humanDelay(1000, 2000);

      // 19. Place order
      console.log("[STEP 19] Clicking 'Place order'...");
      const placeOrderBtn = this.page.locator('button:has-text("Place order")').first();
      await this.waitForVisible(placeOrderBtn);
      await placeOrderBtn.click();
      
      // 20. Waiting spinner lagi
      console.log("[STEP 20] Waiting for spinner after placing order...");
      await this.waitForSpinnerGone(8000);

      // 21. Buka https://teams.microsoft.com/v2/ di tab baru
      console.log("[STEP 21] Opening Teams in a new tab...");
      const teamsPage = await this.context.newPage();
      await teamsPage.goto("https://teams.microsoft.com/v2/", { waitUntil: "domcontentloaded", timeout: HARD_TIMEOUT });
      await this.waitForSpinnerGone(); // Note: this waits for spinner in this.page, let's use a helper for teamsPage if needed

      // 22. Menunggu sampe button sign in muncul click
      console.log("[STEP 22] Waiting for 'Sign in' button in Teams...");
      const teamsSignInBtn = teamsPage.locator('button:has-text("Sign in"), a:has-text("Sign in"), button:has-text("Masuk"), a:has-text("Masuk")').first();
      try {
        await teamsSignInBtn.waitFor({ state: "visible", timeout: 30000 });
        await teamsSignInBtn.click();
        await teamsPage.waitForTimeout(5000);
      } catch (err) {
        console.log("[INFO] 'Sign in' button not found or already signed in Teams. Continuing...");
      }

      // 23. Menunggu start trial muncul lalu click
      console.log("[STEP 23] Waiting for 'Start trial' button in Teams...");
      const startTrialBtn = teamsPage.locator('button:has-text("Start trial"), button:has-text("Mulai uji coba"), [role="button"]:has-text("Start trial")').first();
      try {
        await startTrialBtn.waitFor({ state: "visible", timeout: 60000 });
        await startTrialBtn.click();
        await teamsPage.waitForTimeout(5000);
      } catch (err) {
        console.warn("[WARN] 'Start trial' button not found in Teams. Continuing...");
      }

      // Close the teams tab after trial
      await teamsPage.close().catch(() => {});

      // 24. Balik lagi ke admin user (original tab)
      console.log("[STEP 24] Returning to Admin Center to restore license...");
      await this.page.bringToFront();
      
      // Navigate back to Active Users
      await this.page.goto("https://admin.microsoft.com/#/users", { waitUntil: "domcontentloaded", timeout: HARD_TIMEOUT });
      await this.waitForSpinnerGone(3000);

      // 25. Pilih user yang sama lagi (Baris Pertama)
      console.log("[STEP 25] Re-selecting the first user to restore license...");
      const finalUserRow = this.page.locator('div[data-automation-key="DisplayName"] span[role="button"], [role="gridcell"] button, [role="row"] button').first();
      await this.waitForVisible(finalUserRow);
      await finalUserRow.click();
      await this.humanDelay(2000, 3000);

      // 26. Licenses and apps
      console.log("[STEP 26] Selecting 'Licenses and apps' tab...");
      const finalLicensesTab = this.page.locator('button[role="tab"]:has-text("Licenses and apps"), button:has-text("Licenses and apps")').first();
      await this.waitForVisible(finalLicensesTab);
      await finalLicensesTab.click();
      await this.waitForSpinnerGone(2000);

      // 27. Cantolin lagi lisensi sebelumnya (Check lagi)
      console.log("[STEP 27] Re-checking the license (Office 365)...");
      const finalLicenseCheckbox = this.page.locator('input[type="checkbox"][aria-label*="Office 365" i], input[type="checkbox"]:near(:text("Office 365"))').first();
      const finalLicenseLabel = this.page.locator('label:has-text("Office 365")').first();
      
      try {
        await this.page.waitForTimeout(2000);
        const isChecked = await finalLicenseCheckbox.isChecked().catch(() => false);
        if (!isChecked) {
          console.log("[INFO] License is unchecked, checking it again...");
          await finalLicenseCheckbox.check({ force: true });
        } else {
          console.log("[INFO] License is already checked.");
        }
      } catch (err) {
        console.warn("[WARN] Failed to re-check license via standard ways, trying label click...");
        await finalLicenseLabel.click().catch(() => {});
      }

      await this.humanDelay(1000, 2000);

      // 28. Save changes
      console.log("[STEP 28] Clicking 'Save changes'...");
      const finalSaveBtn = this.page.locator('button:has-text("Save changes"), button[id*="save" i]').first();
      await this.waitForVisible(finalSaveBtn);
      await finalSaveBtn.click();
      await this.waitForSpinnerGone(5000);

      console.log("[SUCCESS] Automation finished successfully.");
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
