const { chromium } = require("playwright-core");
const config = require("./config");
const remoteLogger = require("./remote_logger");

const SPINNER_SELECTOR =
  '[data-testid="spinner"], .ms-Spinner, [class*="spinner" i]';
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
    const { width, height } = this.page.viewportSize() || {
      width: 1280,
      height: 720,
    };
    const x = Math.floor(Math.random() * width);
    const y = Math.floor(Math.random() * height);
    await this.page.mouse.move(x, y, { steps: 10 });
  }

  async runWithMonitor(promise, timeout = HARD_TIMEOUT) {
    let isDone = false;
    let errorMsg = null;

    const checkLoop = async () => {
      while (!isDone) {
        await this.page.waitForTimeout(4000).catch(() => {
          isDone = true;
        });
        if (isDone) break;

        const detectedError = await this.checkForError();
        if (detectedError) {
          errorMsg = detectedError;
          isDone = true;
          break;
        }
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

  async checkForError() {
    try {
      // 1. Cek pesan error validasi di field
      const fieldError = this.page
        .locator('[data-automation-id="error-message"], [id*="error" i]')
        .first();
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
          const frameText = await frame.textContent("body").catch(() => "");
          const lowerFrameText = frameText.toLowerCase();
          const found = markers.find((m) =>
            lowerFrameText.includes(m.toLowerCase()),
          );
          if (found) {
            return `Marker "${found}" detected in frame.`;
          }
        } catch (e) {
          /* skip inaccessible frames */
        }
      }
    } catch (err) {
      /* ignore */
    }
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
        console.log(
          "[WAIT] Spinner still visible or check failed, continuing...",
        );
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
    const keywords = names.map((n) => n.trim().toLowerCase());

    // 1. Coba klik di Main Page & Semua Frames menggunakan JS
    for (const frame of this.page.frames()) {
      try {
        const found = await frame.evaluate((kws) => {
          const candidates = [
            ...document.querySelectorAll(
              'button, [role="button"], a[role="button"], input[type="button"], input[type="submit"]',
            ),
          ];
          const el = candidates.find((b) => {
            const text = (
              b.textContent ||
              b.value ||
              b.getAttribute("aria-label") ||
              ""
            )
              .trim()
              .toLowerCase();

            if (!text || text.length >= 60) return false;

            return kws.some((kw) => {
              const escaped = kw
                .replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
                .replace(/\s+/g, "\\s*");
              // Use word boundary to avoid partial matches like "no" in "notifications"
              return new RegExp(`\\b${escaped}\\b`, "i").test(text);
            });
          });
          if (!el) return null;
          el.click();
          return el.textContent?.trim() || el.value || "unknown";
        }, keywords);
        if (found) {
          console.log(
            `[INFO] Clicked: "${found}" (in frame: ${frame.url() === this.page.url() ? "main" : "subframe"})`,
          );
          return true;
        }
      } catch (e) {}
    }

    // 2. Fallback: Playwright native click di Main Page & Semua Frames
    console.log(
      `[INFO] Fallback to Playwright click for names: ${names.join(", ")}`,
    );
    const pattern = new RegExp(
      names
        .map((n) =>
          n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s*"),
        )
        .join("|"),
      "i",
    );

    for (const frame of this.page.frames()) {
      try {
        const button = frame.getByRole("button", { name: pattern }).first();
        if (await button.isVisible({ timeout: 2000 }).catch(() => false)) {
          await this.randomMouseMove();
          const clickedText = await button
            .evaluate((el) =>
              (
                el.textContent ||
                el.value ||
                el.getAttribute("aria-label") ||
                ""
              ).trim(),
            )
            .catch(() => "unknown");
          await button.click({ timeout: 5000, force: true });
          console.log(
            `[INFO] Clicked: "${clickedText || "unknown"}" (native, in frame)`,
          );
          return true;
        }
      } catch (e) {}
    }

    console.error(`[ERROR] Button not found in any frame for names:`, names);
    throw new Error(`Button not found: ${names.join(", ")}`);
  }

  async handlePopups() {
    console.log("[INFO] Checking for any popups to dismiss...");
    const names = [
      "Close",
      "Dismiss",
      "Maybe later",
      "Got it",
      "No thanks",
      "Tutup",
      "Lain kali",
      "Selesai",
      "X",
    ];
    // Don't split phrases into words to avoid false positives (e.g. "no" in "notifications")
    const keywords = names.map((n) => n.trim().toLowerCase());

    let foundSomethingVisible = true;
    let attempts = 0;
    while (foundSomethingVisible && attempts < 3) {
      foundSomethingVisible = false;
      attempts++;

      for (const frame of this.page.frames()) {
        try {
          const foundName = await frame.evaluate((kws) => {
            const candidates = [
              ...document.querySelectorAll(
                'button, [role="button"], a[role="button"], input[type="button"]',
              ),
            ];
            const el = candidates.find((b) => {
              const textContent = (b.textContent || "").trim().toLowerCase();
              const ariaLabel = (b.getAttribute("aria-label") || "")
                .trim()
                .toLowerCase();
              const val = (b.value || "").trim().toLowerCase();
              const titleMsg = (b.getAttribute("title") || "")
                .trim()
                .toLowerCase();

              const isVisible = !!(
                b.offsetWidth ||
                b.offsetHeight ||
                b.getClientRects().length
              );
              if (!isVisible) return false;

              const isMatch = kws.some((kw) => {
                const escaped = kw
                  .replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
                  .replace(/\s+/g, "\\s*");
                const regex = new RegExp(`\\b${escaped}\\b`, "i");
                return (
                  regex.test(textContent) ||
                  regex.test(ariaLabel) ||
                  regex.test(val) ||
                  regex.test(titleMsg)
                );
              });

              // Limit length to avoid clicking huge buttons accidentally
              const btnLength = Math.max(
                textContent.length,
                ariaLabel.length,
                val.length,
                titleMsg.length,
              );

              return isMatch && btnLength > 0 && btnLength < 35;
            });
            if (!el) return null;
            el.click();
            return (
              el.getAttribute("aria-label") ||
              el.textContent ||
              el.value ||
              "button"
            ).trim();
          }, keywords);

          if (foundName) {
            console.log(
              `[INFO] Dismissed popup button: "${foundName}" (Attempt ${attempts})`,
            );
            await this.humanDelay(1000, 2000);
            foundSomethingVisible = true;
            break; // Break the frame loop to start over from attempt
          }
        } catch (e) {}
      }
    }
  }

  getGenericLocator(keyword, elementType = "input") {
    return this.page
      .locator(
        `${elementType}[id*="${keyword}" i], ${elementType}[data-testid*="${keyword}" i], ${elementType}[name*="${keyword}" i], ${elementType}[aria-label*="${keyword}" i]`,
      )
      .first();
  }

  async connect() {
    if (this.wsUrl) {
      console.log("[STEP 1] Connecting to browser via Ads Power...");
      this.browser = await chromium.connectOverCDP(this.wsUrl);
      const contexts = this.browser.contexts();
      this.context =
        contexts.length > 0 ? contexts[0] : await this.browser.newContext();
    } else {
      console.log("[STEP 1] Launching local browser in incognito mode...");
      this.browser = await chromium.launch({
        headless:
          this.accountConfig?.headless !== undefined
            ? this.accountConfig.headless
            : config.headless,
        args: [
          "--incognito",
          "--disable-blink-features=AutomationControlled",
          "--disable-gpu",
          "--disable-dev-shm-usage",
          "--disable-software-rasterizer",
          "--mute-audio",
        ],
      });
      this.context = await this.browser.newContext();
    }
    const pages = this.context.pages();
    this.page = pages.length > 0 ? pages[0] : await this.context.newPage();
  }

  async run() {
    const email = this.accountConfig.microsoftAccount.email;
    try {
      await this.connect();

      await remoteLogger.logStep(
        email,
        2,
        "🌐 Membuka halaman Microsoft Admin Center...",
      );
      await this.page.goto("https://admin.microsoft.com/", {
        waitUntil: "domcontentloaded",
        timeout: HARD_TIMEOUT,
      });
      await this.waitForSpinnerGone();

      // 3. masukin email click next
      await remoteLogger.logStep(email, 3, `📧 Memasukkan email: ${email}`);
      const emailInput = this.getGenericLocator("email");
      await this.waitForVisible(emailInput);
      await emailInput.fill(email);
      await this.humanDelay(500, 1000);
      await this.clickButtonWithPossibleNames([
        "Next",
        "Selanjutnya",
        "Berikutnya",
      ]);

      // 3.5 Handle "Choose a way to sign in" if it appears
      console.log(
        "[STEP 3.5] Checking for 'Choose a way to sign in' prompt...",
      );
      const usePasswordPrompt = this.page
        .locator(
          'div[role="button"][aria-label*="Use my password" i], div[role="button"]:has-text("Use my password"), div[role="button"][aria-label*="Gunakan kata sandi saya" i], div[role="button"]:has-text("Gunakan kata sandi saya")',
        )
        .first();
      try {
        await usePasswordPrompt.waitFor({ state: "visible", timeout: 5000 });
        console.log(
          "[INFO] 'Choose a way to sign in' detected, clicking 'Use my password'...",
        );
        await usePasswordPrompt.click();
        await this.humanDelay(1000, 2000);
      } catch (e) {
        console.log(
          "[INFO] No 'Choose a way to sign in' prompt found, continuing...",
        );
      }

      // 4. masukin password
      const password = this.accountConfig.microsoftAccount.password;
      await remoteLogger.logStep(email, 4, "🔑 Memasukkan password akun...");
      const passwordInput = this.page.locator('input[type="password"]').first();
      await this.waitForVisible(passwordInput);
      await passwordInput.fill(password);
      await this.humanDelay(500, 1000);
      await this.clickButtonWithPossibleNames(["Sign in", "Masuk"]);

      // 5. Handling after Password (Robust State Monitoring)
      await remoteLogger.logStep(
        email,
        5,
        "⏳ Menunggu dashboard atau konfirmasi login (KMSI/MFA)...",
      );

      const dashboardMarker = this.page
        .locator('[data-hint="ReactLeftNav"], #admin-home-container')
        .first();
      const loginLoopStart = Date.now();

      while (Date.now() - loginLoopStart < 120000) {
        // Max 2 menit menunggu dashboard
        // 5.1 Cek apakah sudah sampai Dashboard?
        if (await dashboardMarker.isVisible().catch(() => false)) {
          console.log("[SUCCESS] Dashboard detected!");
          await this.humanDelay(2000, 4000); // Wait for potential popups to load
          await this.handlePopups();
          break;
        }

        // 5.2 Cek rintangan: Stay signed in
        const yesBtn = this.page
          .locator(
            'button:has-text("Yes"), input[value="Yes"], button:has-text("Ya"), input[value="Ya"], #idSIButton9',
          )
          .first();
        if (await yesBtn.isVisible().catch(() => false)) {
          console.log("[INFO] Handling 'Stay signed in'...");
          await yesBtn.click();
          await this.humanDelay(2000, 3000);
          continue;
        }

        // 5.3 Check for popups
        await this.handlePopups();

        // 5.4 Cek rintangan: MFA Skip
        const skipBtn = this.page
          .locator(
            'a:has-text("Skip for now"), a:has-text("Lompati untuk sekarang"), a:has-text("Lewati untuk sekarang"), button:has-text("Skip for now"), #idSecondaryButton',
          )
          .first();
        if (await skipBtn.isVisible().catch(() => false)) {
          console.log("[INFO] Handling MFA 'Skip for now'...");
          await skipBtn.click();
          await this.humanDelay(2000, 3000);
          continue;
        }

        // 5.4 Cek rintangan: Choose Account / Use Password
        const usePass = this.page
          .locator(
            "text=Use my password, text=Gunakan kata sandi saya, #allowInterrupt",
          )
          .first();
        if (await usePass.isVisible().catch(() => false)) {
          console.log("[INFO] Handling 'Use my password' prompt...");
          await usePass.click();
          await this.humanDelay(2000, 3000);
          continue;
        }

        // 5.6 Cek Error Page
        const err = await this.checkForError();
        if (err) throw new Error(err);

        await this.page.waitForTimeout(2500); // Tunggu antar scan
      }

      // Verifikasi akhir sebelum lanjut ke Step 6
      if (!(await dashboardMarker.isVisible().catch(() => false))) {
        throw new Error(
          "Login failed: Dashboard not reached within 2 minutes.",
        );
      }

      // 6. select menu users (collapse dia)
      await remoteLogger.logStep(
        email,
        6,
        "📂 Membuka menu 'Pengguna' di panel navigasi kiri...",
      );

      await this.waitForSpinnerGone(1000);
      await this.handlePopups(); // One more check before interacting
      const navLocator = this.page
        .locator('[data-hint="ReactLeftNav"]')
        .first();
      await this.waitForVisible(navLocator);
      await this.humanDelay(1000, 2000);

      const usersMenu = this.page
        .locator(
          'button[data-automation-id="LeftNavusersnodeNavToggler"], button[name="Users"], button:has-text("Users"), button[name="Pengguna"], button:has-text("Pengguna")',
        )
        .first();
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
      await remoteLogger.logStep(
        email,
        7,
        "👥 Memilih sub-menu 'Pengguna Aktif'...",
      );
      const activeUsersLink = this.page
        .locator(
          'a[data-automation-id="LeftNavactiveusersnodeNavToggler"], a:has-text("Active users"), a:has-text("Pengguna aktif")',
        )
        .first();
      await this.waitForVisible(activeUsersLink);
      await activeUsersLink.click();

      console.log("[INFO] Waiting for Active users list to appear...");
      await this.waitForSpinnerGone(1000);

      // 8. Search account by email in the user list and select
      const fullEmail = this.accountConfig.microsoftAccount.email;
      await remoteLogger.logStep(
        email,
        8,
        `🔍 Mencari akun pengguna: ${fullEmail}...`,
      );

      const searchInput = this.page
        .locator('[data-automation-id="UserListV2,CommandBarSearchInputBox"]')
        .first();
      await this.waitForVisible(searchInput);
      await searchInput.fill(fullEmail);
      await this.page.keyboard.press("Enter");

      await this.waitForSpinnerGone(2000);

      const userRow = this.page
        .locator(
          'div[data-automation-key="DisplayName"] span[role="button"], [role="gridcell"] button, [role="row"] button',
        )
        .first();
      await this.waitForVisible(userRow);

      const nameFound = await userRow.textContent();
      console.log(
        `[INFO] Clicking display name: "${nameFound?.trim()}" (Found after searching for ${fullEmail})`,
      );
      await userRow.click();

      await this.humanDelay(2000, 3000);

      // 9. terus pilih yang licences dan apps
      await remoteLogger.logStep(
        email,
        9,
        "📋 Membuka tab 'Lisensi dan Aplikasi' milik pengguna...",
      );
      const licensesTab = this.page
        .locator(
          'button[role="tab"]:has-text("Licenses and apps"), button:has-text("Licenses and apps"), button[role="tab"]:has-text("Lisensi dan aplikasi"), button:has-text("Lisensi dan aplikasi")',
        )
        .first();
      await this.waitForVisible(licensesTab);
      await licensesTab.click();
      await this.waitForSpinnerGone(2000);

      // 10. uncheck all checked checkboxes
      await remoteLogger.logStep(
        email,
        10,
        "🔲 Menonaktifkan semua lisensi yang sedang aktif...",
      );
      try {
        await this.waitForSpinnerGone(1000); // Ensure no spinner blocks the initial state
        const checkboxSelector = 'input[type="checkbox"]';
        await this.page
          .locator(checkboxSelector)
          .first()
          .waitFor({ state: "visible", timeout: 15000 })
          .catch(() => {});
        await this.page.waitForTimeout(3000);

        for (let attempt = 1; attempt <= 3; attempt++) {
          await this.waitForSpinnerGone(1000);
          const checkboxes = await this.page.locator(checkboxSelector).all();
          let changed = 0;
          for (const cb of checkboxes) {
            if (await cb.isChecked()) {
              await cb.click({ force: true });
              changed++;
              await this.page.waitForTimeout(300);
            }
          }

          // Verification
          await this.page.waitForTimeout(2000);
          const remaining = await this.page
            .locator('input[type="checkbox"]:checked')
            .count();

          if (remaining === 0) {
            await remoteLogger.logStep(
              email,
              10,
              `✅ Semua lisensi berhasil dinonaktifkan (Percobaan ke-${attempt}).`,
            );
            break;
          } else {
            await remoteLogger.logStep(
              email,
              10,
              `⚠️ Percobaan ke-${attempt}: Masih ada ${remaining} lisensi yang aktif. Mencoba ulang...`,
            );
            if (attempt === 3)
              throw new Error(
                `UNCHECK_ALL_FAILED: Still have ${remaining} checkboxes checked after 3 attempts.`,
              );
            await this.waitForSpinnerGone(2000);
          }
        }
      } catch (err) {
        await remoteLogger.logError(
          email,
          "❌ Langkah 10 Gagal: Tidak dapat menonaktifkan semua lisensi",
          err.message,
        );
        throw err;
      }

      await this.humanDelay(1000, 2000);

      // 11. terus saves changes
      await remoteLogger.logStep(
        email,
        11,
        "💾 Menyimpan perubahan lisensi (nonaktifkan semua)...",
      );
      const saveBtn = this.page
        .locator(
          'button:has-text("Save changes"), button[id*="save" i], button:has-text("Simpan perubahan")',
        )
        .first();
      await this.waitForVisible(saveBtn);
      await saveBtn.click();

      // Wait for completion message
      console.log("[INFO] Waiting for save completion...");
      await this.waitForSpinnerGone(3000);
      // 12. Navigating to product URL from config
      const catalogUrl =
        this.accountConfig.productUrl ||
        "https://admin.cloud.microsoft/?#/catalog/m/offer-details/microsoft-teams-rooms-basic/CFQ7TTC0QW5P";
      const isTeamsRooms = catalogUrl.includes("microsoft-teams-rooms-basic");
      const isPhoneSystem = catalogUrl.includes("phone-system");
      const isCopilot = catalogUrl.includes("copilot");
      const isBusinessAppsFree = catalogUrl.includes("business-apps-free-");

      let planName = "Microsoft 365 Copilot"; // Fallback
      if (isTeamsRooms) planName = "Microsoft Teams Rooms Basic";
      else if (isPhoneSystem) planName = "Microsoft 365 Phone System";
      else if (isBusinessAppsFree) planName = "Business Apps (free)";

      await remoteLogger.logStep(
        email,
        12,
        `🛒 Membuka halaman Marketplace untuk produk: ${planName}...`,
      );
      await this.page.goto(catalogUrl, {
        waitUntil: "commit",
        timeout: HARD_TIMEOUT,
      });

      console.log("[INFO] Waiting for spinner to appear and finish...");
      const spinnerLocator = this.page.locator(SPINNER_SELECTOR).first();
      await spinnerLocator
        .waitFor({ state: "visible", timeout: 10000 })
        .catch(() => {});
      await this.waitForSpinnerGone(3000);

      // Remaining steps will adapt based on the product defined above.

      // 15.5 Select a plan
      if (!isBusinessAppsFree) {
        console.log(`[STEP 15.5] Selecting '${planName}' plan...`);
        const planDropdown = this.page
          .locator(
            'div:has-text("Select a plan") select, [aria-label*="Select a plan" i], div:has-text("Select a plan") [role="combobox"], div:has-text("Pilih paket") select, [aria-label*="Pilih paket" i], div:has-text("Pilih paket") [role="combobox"]',
          )
          .first();
        await this.waitForVisible(planDropdown);

        try {
          const tagName = await planDropdown.evaluate((el) =>
            el.tagName.toLowerCase(),
          );
          if (tagName === "select") {
            await planDropdown.selectOption({ label: planName });
          } else {
            await planDropdown.click();
            await this.humanDelay(1000, 1500);
            const option = this.page.getByRole("option", {
              name: new RegExp(
                `^${planName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`,
                "i",
              ),
            });
            await option.click();
          }
        } catch (err) {
          console.warn(
            `[WARN] Failed to select plan '${planName}' explicitly, it might be already selected. Continuing...`,
          );
        }
      }

      // 15.7 Select '1 year' commitment if present (Only for Copilot)
      if (isCopilot) {
        console.log("[STEP 15.7] Selecting '1 year' commitment...");
        try {
          const oneYearText = this.page
            .locator(
              ':text-is("1 year"), :text-is("1 tahun"), :text-is("1 Tahun")',
            )
            .first();
          await oneYearText.waitFor({ state: "visible", timeout: 5000 });
          console.log("[INFO] '1 year' option found, clicking...");
          await oneYearText.click();
        } catch (e) {
          console.log(
            "[INFO] '1 year' option not found or already selected, continuing...",
          );
        }
      } else if (isBusinessAppsFree) {
        console.log("[STEP 15.7] Selecting '1 month' commitment...");
        try {
          const oneMonthText = this.page
            .locator(
              ':text-is("1 month"), :text-is("1 bulan"), :text-is("1 Bulan")',
            )
            .first();
          await oneMonthText.waitFor({ state: "visible", timeout: 5000 });
          console.log("[INFO] '1 month' option found, clicking...");
          await oneMonthText.click();
        } catch (e) {
          console.log(
            "[INFO] '1 month' option not found or already selected, continuing...",
          );
        }
      } else {
        console.log(
          "[STEP 15.7] Skipping commitment selection for non-Copilot product.",
        );
      }

      // 16. Select 'Pay monthly'
      console.log("[STEP 16] Selecting 'Pay monthly' billing frequency...");
      try {
        const payMonthlyText = this.page
          .locator(':text-is("Pay monthly"), :text-is("Bayar bulanan")')
          .first();
        await payMonthlyText.waitFor({ state: "visible", timeout: 5000 });
        console.log("[INFO] 'Pay monthly' option found, clicking...");
        await payMonthlyText.click();
        await this.waitForSpinnerGone(2000);
      } catch (e) {
        console.log(
          "[INFO] 'Pay monthly' option not found or already selected, continuing...",
        );
      }

      // 17. Menunggu button buy muncul lalu click
      await remoteLogger.logStep(
        email,
        17,
        "🛍️ Menunggu tombol 'Beli' muncul dan mengkliknya...",
      );
      const buyBtn = this.page
        .locator('button:has-text("Buy"), button:has-text("Beli")')
        .first();
      await this.waitForVisible(buyBtn);
      await buyBtn.click();
      await this.waitForSpinnerGone(15000);

      console.log("[STEP 18] Checking for authorization checkboxes...");
      try {
        // Cari container .ms-Checkbox yang mengandung teks recurring payments
        const checkboxContainer = this.page
          .locator(
            '.ms-Checkbox:has-text("authorize recurring payments"), .ms-Checkbox:has-text("pembayaran berulang")',
          )
          .first();

        const isVisible = await checkboxContainer
          .isVisible({ timeout: 10000 })
          .catch(() => false);

        if (isVisible) {
          console.log(
            "[INFO] Authorization checkbox detected. Attempting to check...",
          );

          // Ambil input langsung dari dalam container (sibling bisa diakses lewat parent)
          const checkboxInput = checkboxContainer
            .locator('input[type="checkbox"]')
            .first();

          for (let attempt = 1; attempt <= 3; attempt++) {
            const isChecked = await checkboxInput
              .isChecked()
              .catch(() => false);
            if (isChecked) {
              console.log("[INFO] Checkbox is already checked.");
              break;
            }

            console.log(`[INFO] Checking attempt ${attempt}...`);

            // Klik langsung di visual checkbox (kotak kecilnya), bukan label teks
            const visualCheckbox = checkboxContainer
              .locator(".ms-Checkbox-checkbox")
              .first();

            await visualCheckbox.click({ force: true });
            await this.page.waitForTimeout(1500);

            const nowChecked = await checkboxInput
              .isChecked()
              .catch(() => false);
            if (nowChecked) {
              console.log("[SUCCESS] Checkbox is now checked.");
              break;
            }

            // Fallback: klik via JavaScript langsung ke input
            if (attempt === 2) {
              console.log("[INFO] Fallback: clicking input directly via JS...");
              await checkboxInput.evaluate((el) => el.click());
              await this.page.waitForTimeout(1500);
            }

            if (attempt === 3) {
              throw new Error(
                "Gagal mencentang checkbox otorisasi setelah 3x percobaan.",
              );
            }
          }
        } else {
          console.log("[INFO] No authorization checkbox found, skipping...");
        }
      } catch (err) {
        throw new Error(`CRITICAL_ERROR di Step 18: ${err.message}`);
      }

      // 19. Place order (STRICT MODE)
      await remoteLogger.logStep(
        email,
        19,
        "📦 Mengklik tombol 'Buat Pesanan' untuk konfirmasi pembelian...",
      );
      const placeOrderBtn = this.page
        .locator(
          'button:has-text("Place order"), button:has-text("Buat pesanan"), button:has-text("Tempatkan pesanan")',
        )
        .first();

      console.log(
        "[INFO] Waiting for 'Place order' button to become enabled...",
      );
      try {
        await this.page.waitForFunction(
          (btn) => {
            return (
              btn &&
              !btn.disabled &&
              btn.getAttribute("aria-disabled") !== "true" &&
              !btn.classList.contains("is-disabled")
            );
          },
          await placeOrderBtn.elementHandle(),
          { timeout: 30000 },
        );
        console.log("[INFO] Button 'Place order' is now enabled.");
      } catch (e) {
        throw new Error(
          "PLACE_ORDER_DISABLED: Tombol tidak aktif dalam 30 detik. Kemungkinan otorisasi gagal atau field ada yang kurang.",
        );
      }

      await placeOrderBtn.click({ timeout: 10000 }).catch(async (e) => {
        console.log("[INFO] Click failed, trying force click...");
        await placeOrderBtn.click({ force: true });
      });
      console.log("[INFO] 'Place order' clicked. Waiting for confirmation...");

      // 20. Verifikasi Transaksi Berhasil (STRICT)
      await remoteLogger.logStep(
        email,
        20,
        "🔎 Memverifikasi keberhasilan pesanan sebelum melanjutkan...",
      );
      let isSuccess = false;
      const verifyStart = Date.now();

      while (Date.now() - verifyStart < 60000) {
        // Max 1 menit menunggu konfirmasi
        // 20.1 Cek apakah tombol sudah hilang?
        const isBtnHidden = await placeOrderBtn.isHidden().catch(() => true);

        // 20.2 Cek apakah URL menunjukkan konfirmasi atau ada teks sukses?
        const currentUrl = this.page.url().toLowerCase();
        const bodyContent = await this.page.innerText("body").catch(() => "");
        const successKeywords = [
          "all set",
          "confirmation",
          "thanks",
          "terima kasih",
          "detail pesanan",
          "order details",
        ];
        const foundKeyword = successKeywords.find((kw) =>
          bodyContent.toLowerCase().includes(kw),
        );

        if (
          isBtnHidden &&
          (currentUrl.includes("confirmation") || foundKeyword)
        ) {
          console.log(
            `[SUCCESS] Order placement verified! (Keyword found: "${foundKeyword || "URL Confirmation"}")`,
          );
          isSuccess = true;
          break;
        }

        // 20.3 Cek apakah ada error muncul?
        const detectedError = await this.checkForError();
        if (detectedError) {
          throw new Error(`PLACE_ORDER_FAILED: ${detectedError}`);
        }

        await this.page.waitForTimeout(2500); // Scan tiap 2.5s
      }

      if (!isSuccess) {
        console.warn(
          "[WARN] Order confirmation not clearly detected, but button is gone. Proceeding with caution...",
        );
      }

      await this.waitForSpinnerGone(5000);

      // 21. Buka https://teams.microsoft.com/v2/ di tab baru
      await remoteLogger.logStep(
        email,
        21,
        "🚀 Membuka Microsoft Teams di tab baru untuk aktivasi trial...",
      );
      const teamsPage = await this.context.newPage();
      await teamsPage.goto("https://teams.microsoft.com/v2/", {
        waitUntil: "domcontentloaded",
        timeout: HARD_TIMEOUT,
      });
      await this.waitForSpinnerGone(); // Note: this waits for spinner in this.page, let's use a helper for teamsPage if needed

      // 22. Menunggu sampe button sign in muncul click
      await remoteLogger.logStep(
        email,
        22,
        "⏳ Menunggu tombol 'Masuk' muncul di Teams (atau deteksi error izin)...",
      );
      const teamsSignInBtn = teamsPage
        .locator(
          'button:has-text("Sign in"), a:has-text("Sign in"), button:has-text("Masuk"), a:has-text("Masuk")',
        )
        .first();

      const permissionErrorLocator = teamsPage
        .getByText("You don't have the required permissions to access this org")
        .first();

      try {
        await teamsSignInBtn
          .or(permissionErrorLocator)
          .waitFor({ state: "visible", timeout: 30000 });

        if (await permissionErrorLocator.isVisible().catch(() => false)) {
          console.error(
            "[ERROR] Permission error page detected before Sign in.",
          );
          throw new Error(
            "Don't have the required permissions to access this org",
          );
        }

        await teamsSignInBtn.click();
        await teamsPage.waitForTimeout(5000);
      } catch (err) {
        if (
          err.message ===
          "Don't have the required permissions to access this org"
        ) {
          throw err;
        }
        console.log(
          "[INFO] 'Sign in' button not found or already signed in Teams. Continuing...",
        );
      }

      // 23. Menunggu start trial muncul lalu click
      console.log("[STEP 23] Waiting for Teams loading screen to finish...");

      const startTrialBtn = teamsPage
        .locator(
          'button:has-text("Start trial"), button:has-text("Mulai uji coba"), [role="button"]:has-text("Start trial")',
        )
        .first();

      try {
        // Race: tunggu loading hilang ATAU button muncul, mana duluan
        await Promise.race([
          teamsPage
            .locator("#loading-screen")
            .waitFor({ state: "hidden", timeout: 120000 })
            .catch(() => {}),
          startTrialBtn.waitFor({ state: "visible", timeout: 120000 }),
        ]);

        // Pastikan button visible sebelum klik
        const isBtnVisible = await startTrialBtn.isVisible().catch(() => false);
        if (!isBtnVisible) {
          // Kalau belum visible setelah race, tunggu sebentar lagi
          await startTrialBtn.waitFor({ state: "visible", timeout: 30000 });
        }

        await remoteLogger.logStep(
          email,
          23.5,
          "▶️ Mengklik tombol 'Mulai Uji Coba' di Microsoft Teams...",
        );

        // Scroll ke button lalu klik
        await startTrialBtn.scrollIntoViewIfNeeded().catch(() => {});
        await startTrialBtn.click();

        // Menunggu loading setelah klik start trial selesai sebelum close
        await remoteLogger.logStep(
          email,
          23.6,
          "⏳ Menunggu proses aktivasi uji coba selesai (loading)...",
        );

        // Tunggu spinner muncul dulu (max 5s), baru tunggu hilang
        const teamsSpinner = teamsPage.locator(SPINNER_SELECTOR).first();
        await teamsSpinner
          .waitFor({ state: "visible", timeout: 5000 })
          .catch(() => {});
        const isSpinning = await teamsSpinner.isVisible().catch(() => false);
        if (isSpinning) {
          await teamsSpinner
            .waitFor({ state: "hidden", timeout: 60000 })
            .catch(() => {
              console.log(
                "[WARN] Teams trial setup spinner still visible, continuing anyway.",
              );
            });
        }

        await remoteLogger.logStep(
          email,
          23.7,
          "✅ Aktivasi uji coba selesai. Menutup tab Teams...",
        );
      } catch (err) {
        await teamsPage.close().catch(() => {});
        throw new Error(
          "START_TRIAL_NOT_FOUND: Tombol 'Start trial' gagal ditemukan setelah Teams terbuka.",
        );
      }
      // Close the teams tab after trial
      await teamsPage.close().catch(() => {});

      // 24. Balik lagi ke admin user (original tab)
      await remoteLogger.logStep(
        email,
        24,
        "↩️ Kembali ke Admin Center untuk memulihkan lisensi pengguna...",
      );
      await this.page.bringToFront();

      // Navigate back to Active Users
      await this.page.goto("https://admin.microsoft.com/#/users", {
        waitUntil: "domcontentloaded",
        timeout: HARD_TIMEOUT,
      });
      await this.waitForSpinnerGone(3000);

      // 25. Search the same user again and select
      await remoteLogger.logStep(
        email,
        25,
        `🔍 Mencari ulang pengguna: ${fullEmail} untuk pemulihan lisensi...`,
      );

      const finalSearchInput = this.page
        .locator('[data-automation-id="UserListV2,CommandBarSearchInputBox"]')
        .first();
      await this.waitForVisible(finalSearchInput);
      await finalSearchInput.fill(fullEmail);
      await this.page.keyboard.press("Enter");

      await this.waitForSpinnerGone(2000);

      const finalUserRow = this.page
        .locator(
          'div[data-automation-key="DisplayName"] span[role="button"], [role="gridcell"] button, [role="row"] button',
        )
        .first();
      await this.waitForVisible(finalUserRow);
      await finalUserRow.click();
      await this.humanDelay(2000, 3000);

      // 26. Licenses and apps
      await remoteLogger.logStep(
        email,
        26,
        "📋 Membuka kembali tab 'Lisensi dan Aplikasi' untuk pemulihan...",
      );
      const finalLicensesTab = this.page
        .locator(
          'button[role="tab"]:has-text("Licenses and apps"), button:has-text("Licenses and apps"), button[role="tab"]:has-text("Lisensi dan aplikasi"), button:has-text("Lisensi dan aplikasi")',
        )
        .first();
      await this.waitForVisible(finalLicensesTab);
      await finalLicensesTab.click();
      await this.waitForSpinnerGone(2000);

      // 27. Restore license - search by name from a prioritized list
      const licenseNames = [
        "Microsoft 365 Business Standard",
        "Microsoft 365 Business Basic",
        "Microsoft 365 Business Premium",
        "Microsoft 365 E3",
        "Microsoft 365 E5",
        "Office 365 E1",
        "Office 365 E3",
        "Office 365 E5",
      ];

      await remoteLogger.logStep(
        email,
        27,
        "🔍 Mencari lisensi yang dikenal di daftar untuk dipulihkan...",
      );

      try {
        await this.waitForSpinnerGone(1000);

        // Tunggu sampai minimal satu checkbox muncul
        await this.page
          .locator('input[type="checkbox"]')
          .first()
          .waitFor({ state: "visible", timeout: 15000 })
          .catch(() => {});

        // Debug: log semua lisensi yang tersedia di halaman
        const allLicenseTexts = await this.page
          .locator('[data-automation-id^="LicenseText_"]')
          .all();

        for (const el of allLicenseTexts) {
          const text = await el.innerText().catch(() => "N/A");
          const automationId = await el
            .getAttribute("data-automation-id")
            .catch(() => "N/A");
          await remoteLogger.logStep(
            email,
            27,
            `🔎 Lisensi tersedia di halaman: "${text}" (${automationId})`,
          );
        }

        // Find the first matching license checkbox by data-automation-id
        let targetCheckbox = null;
        let foundLicenseName = null;

        for (const licenseName of licenseNames) {
          // Cari elemen teks lisensi via data-automation-id yang spesifik
          const licenseTextEl = this.page
            .locator(`[data-automation-id="LicenseText_${licenseName}"]`)
            .first();

          const isVisible = await licenseTextEl
            .isVisible({ timeout: 3000 })
            .catch(() => false);

          if (!isVisible) continue;

          // Naik ke ancestor .ms-Checkbox, lalu ambil input checkbox di dalamnya
          const checkbox = licenseTextEl
            .locator('xpath=ancestor::div[contains(@class,"ms-Checkbox")][1]')
            .locator('input[type="checkbox"]');

          const isCheckboxVisible = await checkbox
            .isVisible()
            .catch(() => false);

          if (isCheckboxVisible) {
            targetCheckbox = checkbox;
            foundLicenseName = licenseName;
            await remoteLogger.logStep(
              email,
              27,
              `✅ Lisensi ditemukan: '${licenseName}' — akan diaktifkan kembali.`,
            );
            break;
          }
        }

        if (!targetCheckbox) {
          throw new Error(
            `LICENSE_NOT_FOUND: None of the known licenses found in the checklist. Checked: ${licenseNames.join(", ")}`,
          );
        }

        // Coba centang checkbox hingga 3 kali
        let verifyChecked = false;
        for (let attempt = 1; attempt <= 3; attempt++) {
          await this.waitForSpinnerGone(500);

          const isChecked = await targetCheckbox.isChecked().catch(() => false);

          if (!isChecked) {
            await remoteLogger.logStep(
              email,
              27,
              `🖱️ Percobaan ke-${attempt}: Mengaktifkan centang lisensi '${foundLicenseName}'...`,
            );
            await targetCheckbox.click({ force: true });
            await this.page.waitForTimeout(1500);
          } else {
            await remoteLogger.logStep(
              email,
              27,
              `ℹ️ Percobaan ke-${attempt}: Lisensi '${foundLicenseName}' sudah tercentang.`,
            );
          }

          // Verifikasi status centang
          verifyChecked = await targetCheckbox.isChecked().catch(() => false);

          if (verifyChecked) {
            await remoteLogger.logStep(
              email,
              27,
              `✅ Lisensi '${foundLicenseName}' berhasil diaktifkan kembali dan terverifikasi.`,
            );
            break;
          } else {
            await remoteLogger.logStep(
              email,
              27,
              `⚠️ Percobaan ke-${attempt}: Lisensi masih belum tercentang. Mencoba ulang...`,
            );
            await this.waitForSpinnerGone(1000);
          }

          if (attempt === 3 && !verifyChecked) {
            throw new Error(
              `STRICT_CHECKBOX_FAILED: Failed to check '${foundLicenseName}' license after 3 attempts.`,
            );
          }
        }
      } catch (err) {
        await remoteLogger.logError(
          email,
          "❌ Langkah 27 Gagal: Tidak dapat memulihkan lisensi",
          err.message,
        );
        throw err;
      }

      await this.humanDelay(1000, 2000);

      // 28. Save changes
      await remoteLogger.logStep(
        email,
        28,
        "💾 Menyimpan perubahan lisensi yang telah dipulihkan...",
      );
      const finalSaveBtn = this.page
        .locator(
          'button:has-text("Save changes"), button[id*="save" i], button:has-text("Simpan perubahan")',
        )
        .first();
      await this.waitForVisible(finalSaveBtn);
      await finalSaveBtn.click();
      await this.waitForSpinnerGone(5000);

      await remoteLogger.logSuccess(
        email,
        "🎉 Proses otomasi selesai dengan sukses! Semua langkah berhasil dijalankan.",
      );
      return { success: true };
    } catch (error) {
      await remoteLogger.logError(
        this.accountConfig?.microsoftAccount?.email,
        "❌ Otomasi gagal — proses dihentikan",
        error.message,
      );
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
