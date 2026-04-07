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
        await this.page.waitForTimeout(1500).catch(() => {
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
      // "X" sengaja dihapus — terlalu generic, bisa salah klik tombol di halaman MFA
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

      // --- STEP 3 VERIFICATION (CHECKPOINT) ---
      console.log(
        "[STEP 3 VERIFY] Waiting for Password input or Choose method prompt...",
      );
      const passwordOrPrompt = this.page.locator(
        'input[type="password"], div[role="button"]:has-text("Use my password"), div[role="button"]:has-text("Gunakan kata sandi saya")',
      );
      await passwordOrPrompt
        .first()
        .waitFor({ state: "visible", timeout: 15000 })
        .catch(() => {
          throw new Error(
            "EMAIL_TRANSITION_FAILED: Gagal lanjut ke pengisian password. Cek apakah email sudah benar atau ada error di halaman.",
          );
        });

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
        await this.humanDelay(400, 800);
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
          await this.humanDelay(1000, 1500); // Wait for potential popups to load
          await this.handlePopups();
          break;
        }

        // 5.2 Cek rintangan: Stay signed in (KMSI)
        const yesBtn = this.page
          .locator(
            'button:has-text("Yes"), input[value="Yes"], button:has-text("Ya"), input[value="Ya"], #idSIButton9',
          )
          .first();
        if (await yesBtn.isVisible().catch(() => false)) {
          console.log("[INFO] Handling 'Stay signed in'...");
          try {
            await yesBtn.click({ timeout: 5000 });
            await this.humanDelay(1000, 1500);
            continue;
          } catch (e) {
            console.log(
              "[WARN] 'Yes' button blocked or failed, checking popups...",
            );
          }
        }

        // 5.3 Cek rintangan: MFA Skip (PRIORITAS — harus dicek SEBELUM handlePopups)
        const skipBtn = this.page
          .locator(
            'a:has-text("Skip for now"), a:has-text("Lompati untuk sekarang"), a:has-text("Lewati untuk sekarang"), button:has-text("Skip for now"), #idSecondaryButton',
          )
          .first();
        if (await skipBtn.isVisible().catch(() => false)) {
          console.log("[INFO] Handling MFA 'Skip for now'...");
          try {
            await skipBtn.click({ timeout: 5000 });
            await this.humanDelay(1000, 1500);
            continue;
          } catch (e) {
            console.log("[WARN] 'Skip for now' blocked, checking popups...");
          }
        }

        // 5.4 Cek rintangan: Use Password prompt
        const usePass = this.page
          .locator(
            "text=Use my password, text=Gunakan kata sandi saya, #allowInterrupt",
          )
          .first();
        if (await usePass.isVisible().catch(() => false)) {
          console.log("[INFO] Handling 'Use my password' prompt...");
          try {
            await usePass.click({ timeout: 5000 });
            await this.humanDelay(1000, 1500);
            continue;
          } catch (e) {
            console.log(
              "[WARN] 'Use Password' prompt blocked, checking popups...",
            );
          }
        }

        // 5.5 Cek Error Page
        const err = await this.checkForError();
        if (err) throw new Error(err);

        // 5.6 Tangani popup umum SETELAH semua pengecekan MFA selesai
        await this.handlePopups();

        await this.page.waitForTimeout(800); // Tunggu antar scan
      }

      // Verifikasi akhir sebelum lanjut ke Step 6
      if (!(await dashboardMarker.isVisible().catch(() => false))) {
        throw new Error(
          "Login failed: Dashboard not reached within 2 minutes.",
        );
      }

      // 6. Direct navigation to Active Users
      await remoteLogger.logStep(
        email,
        6,
        "📂 Membuka halaman 'Pengguna Aktif' langsung via URL...",
      );

      await this.page.goto("https://admin.cloud.microsoft/?#/users", {
        waitUntil: "domcontentloaded",
        timeout: HARD_TIMEOUT,
      });
      await this.waitForSpinnerGone(200);
      await this.handlePopups(); // One check for page popups

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

      await this.humanDelay(800, 1500);

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
      await this.waitForSpinnerGone(500);

      // 10. uncheck all checked checkboxes
      await remoteLogger.logStep(
        email,
        10,
        "🔲 Menonaktifkan semua lisensi yang sedang aktif...",
      );
      try {
        await this.waitForSpinnerGone(200); // Ensure no spinner blocks the initial state
        const checkboxSelector = 'input[type="checkbox"]';
        await this.page
          .locator(checkboxSelector)
          .first()
          .waitFor({ state: "visible", timeout: 15000 })
          .catch(() => {});
        await this.page.waitForTimeout(800);

        for (let attempt = 1; attempt <= 3; attempt++) {
          await this.waitForSpinnerGone(200);
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
          await this.page.waitForTimeout(800);
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
            await this.waitForSpinnerGone(500);
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
      // --- PURCHASE RETRY LOOP (Step 12 - 20) ---
      let purchaseSuccess = false;
      for (let purchaseAttempt = 1; purchaseAttempt <= 3; purchaseAttempt++) {
        try {
          if (purchaseAttempt > 1) {
            await remoteLogger.logStep(
              email,
              12,
              `🔄 Mencoba ulang proses pembelian (Percobaan ke-${purchaseAttempt}/3)...`,
            );
            await this.page.reload({ waitUntil: "domcontentloaded" });
            await this.waitForSpinnerGone(2000);
          }

          // 12. Navigating to product URL from config
          const catalogUrl =
            this.accountConfig.productUrl ||
            "https://admin.cloud.microsoft/?#/catalog/m/offer-details/microsoft-teams-rooms-basic/CFQ7TTC0QW5P";
          const isTeamsRooms = catalogUrl.includes(
            "microsoft-teams-rooms-basic",
          );
          const isPhoneSystem = catalogUrl.includes("phone-system");
          const isCopilot = catalogUrl.includes("copilot");
          const isBusinessAppsFree = catalogUrl.includes("business-apps-free-");

          let planName = "Microsoft 365 Copilot"; // Fallback
          if (isTeamsRooms) planName = "Microsoft Teams Rooms Basic";
          else if (isPhoneSystem) planName = "Microsoft 365 Phone System";
          else if (isBusinessAppsFree) planName = "Business Apps (free)";

          if (purchaseAttempt === 1 || this.page.url() !== catalogUrl) {
            await remoteLogger.logStep(
              email,
              12,
              `🛒 Membuka halaman Marketplace untuk produk: ${planName}...`,
            );
            await this.page.goto(catalogUrl, {
              waitUntil: "commit",
              timeout: HARD_TIMEOUT,
            });
            await this.waitForSpinnerGone(800);
          }

          // Check for 'This product is unavailable' or 'You are not eligible'
          console.log("[INFO] Checking for product availability...");
          const unavailableMarker = this.page
            .locator(
              'text="This product is unavailable", text="Produk ini tidak tersedia", text="You are not eligible", text="Anda tidak memenuhi syarat"',
            )
            .first();

          const isUnavailable = await unavailableMarker
            .isVisible({ timeout: 5000 })
            .catch(() => false);
          if (isUnavailable) {
            const errorText = await unavailableMarker
              .innerText()
              .catch(() => "Produk tidak tersedia");
            throw new Error(
              `MARKETPLACE_ERROR: ${errorText}. Bot tidak dapat melanjutkan dengan akun ini.`,
            );
          }

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

          // 17. Menunggu button buy muncul lalu click
          await remoteLogger.logStep(
            email,
            17,
            "🛍️ Menunggu tombol 'Beli' muncul dan mengkliknya...",
          );

          const buyBtn = this.page
            .locator(
              `
        button:has-text("Buy"), button:has-text("Beli"), 
        [role="button"]:has-text("Buy"), [role="button"]:has-text("Beli"), 
        a:has-text("Buy"), a:has-text("Beli"),
        button:has-text("Get"), button:has-text("Dapatkan"), 
        [role="button"]:has-text("Get"), [role="button"]:has-text("Dapatkan"),
        a:has-text("Get"), a:has-text("Dapatkan"),
        button:has-text("Checkout"), [role="button"]:has-text("Checkout"), a:has-text("Checkout")
      `,
            )
            .first();

          await this.waitForVisible(buyBtn).catch(() => {
            throw new Error("BUY_BUTTON_NOT_FOUND");
          });

          // 15.7 & 16: Pemilihan Commitment & Billing Frequency (dengan Re-check silang)
          for (let retry = 1; retry <= 3; retry++) {
            console.log(
              `[STEP 15.7/16] Attempt ${retry}: Ensuring options are selected...`,
            );

            if (isCopilot) {
              const oneYear = this.page
                .locator(
                  'label:has-text("1 year"), label:has-text("1 tahun"), label:has-text("1 Tahun"), :text-is("1 year"), :text-is("1 tahun"), :text-is("1 Tahun")',
                )
                .first();
              await oneYear.click({ timeout: 5000 }).catch(async () => {
                // Fallback: click visual container
                await this.page
                  .locator(
                    '.ms-Checkbox:has-text("1 year"), .ms-Checkbox:has-text("1 tahun")',
                  )
                  .first()
                  .click({ force: true, timeout: 3000 })
                  .catch(() => {});
              });
            }

            if (isBusinessAppsFree) {
              const oneMonth = this.page
                .locator(
                  'label:has-text("1 month"), label:has-text("1 bulan"), label:has-text("1 Bulan"), :text-is("1 month"), :text-is("1 bulan"), :text-is("1 Bulan")',
                )
                .first();
              await oneMonth.click({ timeout: 5000 }).catch(async () => {
                // Fallback: click visual container
                await this.page
                  .locator(
                    '.ms-Checkbox:has-text("1 month"), .ms-Checkbox:has-text("1 bulan")',
                  )
                  .first()
                  .click({ force: true, timeout: 3000 })
                  .catch(() => {});
              });
            }

            const payMonthly = this.page
              .locator(
                'label:has-text("Pay monthly"), label:has-text("Bayar bulanan"), :text-is("Pay monthly"), :text-is("Bayar bulanan")',
              )
              .first();
            await payMonthly.click({ timeout: 5000 }).catch(async () => {
              // Fallback: click visual container
              await this.page
                .locator(
                  '.ms-Checkbox:has-text("Pay monthly"), .ms-Checkbox:has-text("Bayar bulanan")',
                )
                .first()
                .click({ force: true, timeout: 3000 })
                .catch(() => {});
            });

            await this.page.waitForTimeout(1000); // Wait for potential state changes

            const isStillDisabled = await buyBtn
              .evaluate((btn) => {
                return (
                  btn.disabled ||
                  btn.getAttribute("aria-disabled") === "true" ||
                  btn.classList.contains("is-disabled") ||
                  btn.getAttribute("disabled") !== null
                );
              })
              .catch(() => true);

            if (!isStillDisabled) {
              console.log(
                "[SUCCESS] Options registered, Buy button is now ENABLED.",
              );
              break;
            }

            console.warn(
              `[WARN] Buy button still disabled after attempt ${retry}. Retrying selection...`,
            );

            if (retry === 3) {
              throw new Error("BUY_BUTTON_LOCKED");
            }
          }

          let clickedSuccessfully = false;
          for (let attempt = 1; attempt <= 3; attempt++) {
            console.log(
              `[INFO] Attempting to click 'Buy' button (Attempt ${attempt})...`,
            );
            const oldUrl = this.page.url();

            // Gunakan force: true hanya sebagai pengaman tambahan jika transisi butuh trigger lebih keras
            await buyBtn.click({ timeout: 10000, force: true }).catch(() => {});
            await this.page.waitForTimeout(1500);

            // Verifikasi apakah klik berhasil (URL berubah, tombol hilang, atau elemen step berikutnya muncul)
            const newUrl = this.page.url();
            const isBtnStillVisible = await buyBtn
              .isVisible()
              .catch(() => false);

            const nextStepMarker = this.page
              .locator(
                `
          .ms-Checkbox:has-text("authorize recurring payments"), 
          .ms-Checkbox:has-text("pembayaran berulang"),
          button:has-text("Place order"), 
          button:has-text("Buat pesanan"),
          button:has-text("Tempatkan pesanan")
        `,
              )
              .first();
            const isNextStepVisible = await nextStepMarker
              .isVisible()
              .catch(() => false);

            if (newUrl !== oldUrl || !isBtnStillVisible || isNextStepVisible) {
              console.log(
                "[SUCCESS] 'Buy' click triggered transition (URL change or Next Step detected).",
              );
              clickedSuccessfully = true;
              break;
            }

            console.warn(
              "[WARN] 'Buy' click didn't seem to trigger anything. Retrying...",
            );
            await this.page.waitForTimeout(1000);
          }

          if (!clickedSuccessfully) {
            throw new Error("BUY_CLICK_NOP");
          }
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
                await this.page.waitForTimeout(800);

                const nowChecked = await checkboxInput
                  .isChecked()
                  .catch(() => false);
                if (nowChecked) {
                  console.log("[SUCCESS] Checkbox is now checked.");
                  break;
                }

                // Fallback: klik via JavaScript langsung ke input
                if (attempt === 2) {
                  console.log(
                    "[INFO] Fallback: clicking input directly via JS...",
                  );
                  await checkboxInput.evaluate((el) => el.click());
                  await this.page.waitForTimeout(800);
                }

                if (attempt === 3) {
                  throw new Error(
                    "Gagal mencentang checkbox otorisasi setelah 3x percobaan.",
                  );
                }
              }
            } else {
              console.log(
                "[INFO] No authorization checkbox found, skipping...",
              );
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
          console.log(
            "[INFO] 'Place order' clicked. Waiting for confirmation...",
          );

          // 20. Verifikasi Transaksi Berhasil (STRICT)
          await remoteLogger.logStep(
            email,
            20,
            "🔎 Memverifikasi keberhasilan pesanan sebelum melanjutkan...",
          );
          const verifyStart = Date.now();

          while (Date.now() - verifyStart < 60000) {
            // Max 1 menit menunggu konfirmasi
            // 20.1 Cek apakah tombol sudah hilang?
            const isBtnHidden = await placeOrderBtn
              .isHidden()
              .catch(() => true);

            // 20.2 Cek apakah URL menunjukkan konfirmasi atau ada teks sukses?
            const currentUrl = this.page.url().toLowerCase();
            const bodyContent = await this.page
              .innerText("body")
              .catch(() => "");
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
              purchaseSuccess = true;
              break;
            }

            // 20.3 Cek apakah ada error muncul?
            const detectedError = await this.checkForError();
            if (detectedError) {
              throw new Error(`PLACE_ORDER_FAILED: ${detectedError}`);
            }

            await this.page.waitForTimeout(800);
          }

          if (!purchaseSuccess) {
            console.warn(
              "[WARN] Order confirmation not clearly detected, but button is gone. Proceeding with caution...",
            );
          }

          await this.waitForSpinnerGone(1000);
          break; // EXIT RETRY LOOP IF SUCCESS!
        } catch (err) {
          if (
            (err.message.includes("something happened") ||
              err.message.includes("Terjadi kesalahan")) &&
            purchaseAttempt < 3
          ) {
            console.warn(
              `[RETRY] Purchase Failed with 'Something happened'. Attempt ${purchaseAttempt}/3. Reloading...`,
            );
            await this.page.reload({ waitUntil: "domcontentloaded" });
            await this.page.waitForTimeout(5000);
            continue;
          }
          throw err; // Persist other errors
        }
      }

      // 21. Buka https://teams.microsoft.com/v2/ di tab baru
      await remoteLogger.logStep(
        email,
        21,
        "🚀 Membuka Microsoft Teams di tab baru untuk aktivasi trial...",
      );
      const teamsPage = await this.context.newPage();
      try {
        await teamsPage.goto("https://teams.microsoft.com/v2/", {
          waitUntil: "domcontentloaded",
          timeout: HARD_TIMEOUT,
        });

        // Final robustness check: if page is totally blank, try one refresh
        const bodyText = await teamsPage.innerText("body").catch(() => "");
        if (!bodyText || bodyText.trim().length === 0) {
          console.log("[INFO] Teams page is blank, refreshing...");
          await teamsPage.reload({ waitUntil: "domcontentloaded" });
        }
      } catch (e) {
        console.warn("[WARN] Initial Teams navigation failed, retrying...");
        await teamsPage.goto("https://teams.microsoft.com/v2/", {
          waitUntil: "domcontentloaded",
        });
      }

      // 22 & 23 Combined: Faster detection for Teams state
      await remoteLogger.logStep(
        email,
        22,
        "⏳ Menunggu Teams siap (Sign in atau Start Trial)...",
      );

      // Wait for page to stabilize after potential blank check
      await teamsPage.waitForTimeout(2000);

      const teamsSignInBtn = teamsPage
        .locator(
          'button:has-text("Sign in"), a:has-text("Sign in"), button:has-text("Masuk"), a:has-text("Masuk")',
        )
        .first();

      const startTrialBtn = teamsPage
        .locator(
          'button:has-text("Start trial"), button:has-text("Mulai uji coba"), [role="button"]:has-text("Start trial"), button:has-text("Get started"), button:has-text("Mulai"), button:has-text("Try now"), a:has-text("Get started")',
        )
        .first();

      const pickAccountHeader = teamsPage
        .locator(
          'div:has-text("Pick an account"), div:has-text("Pilih akun"), h1:has-text("Pick an account"), h1:has-text("Pilih akun")',
        )
        .first();

      const permissionErrorLocator = teamsPage
        .getByText(
          /You don't have the required permissions to access this org|Anda tidak memiliki izin yang diperlukan untuk mengakses organisasi ini/i,
        )
        .first();

      try {
        console.log(
          "[INFO] Waiting for Sign In, Start Trial, Pick Account, or Permission (Race 60s)...",
        );
        // Race between elements with an extended 60s timeout
        await teamsSignInBtn
          .or(startTrialBtn)
          .or(pickAccountHeader)
          .or(permissionErrorLocator)
          .waitFor({ state: "visible", timeout: 60000 });

        if (await permissionErrorLocator.isVisible().catch(() => false)) {
          console.error("[ERROR] Permission error page detected.");
          throw new Error(
            "Don't have the required permissions to access this org",
          );
        }

        // Handle "Pick an account" screen if it appears
        if (await pickAccountHeader.isVisible().catch(() => false)) {
          console.log(
            "[INFO] 'Pick an account' detected, looking for current user email...",
          );
          const targetAccountItem = teamsPage
            .locator(
              `div[role="listitem"]:has-text("${email}"), [data-test-id="${email}"]`,
            )
            .first();
          if (await targetAccountItem.isVisible().catch(() => false)) {
            await targetAccountItem.click();
            await teamsPage.waitForTimeout(2000);
          } else {
            // Fallback: click first account tile
            await teamsPage
              .locator('div[role="listitem"], .tile-container')
              .first()
              .click()
              .catch(() => {});
            await teamsPage.waitForTimeout(2000);
          }
        }

        if (await teamsSignInBtn.isVisible().catch(() => false)) {
          console.log("[INFO] 'Sign in' button detected, clicking...");
          await teamsSignInBtn.click();
          await teamsPage.waitForTimeout(1500);
          console.log(
            "[INFO] Waiting for 'Start Trial' button to appear after Sign in...",
          );
          await startTrialBtn.waitFor({ state: "visible", timeout: 45000 });
        } else {
          console.log("[INFO] Proceeding to Start Trial steps.");
        }

        await remoteLogger.logStep(
          email,
          23,
          "▶️ Mengklik tombol 'Mulai Uji Coba' di Microsoft Teams...",
        );

        // Scroll ke button lalu klik
        await startTrialBtn.scrollIntoViewIfNeeded().catch(() => {});
        await startTrialBtn.click();

        // Menunggu loading setelah klik start trial selesai sebelum close
        await remoteLogger.logStep(
          email,
          23.5,
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
          23.6,
          "✅ Aktivasi uji coba selesai. Menutup tab Teams...",
        );
      } catch (err) {
        await teamsPage.close().catch(() => {});
        if (err.message.includes("permissions")) throw err;
        throw new Error(
          `START_TRIAL_FAILED: ${err.message || "Gagal aktivasi trial Teams."}`,
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
      await this.page.goto("https://admin.cloud.microsoft/?#/users", {
        waitUntil: "domcontentloaded",
        timeout: HARD_TIMEOUT,
      });
      await this.waitForSpinnerGone(800);

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
      await this.humanDelay(800, 1500);

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
      await this.waitForSpinnerGone(500);

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
        await this.waitForSpinnerGone(200);

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
            await this.page.waitForTimeout(800);
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
            await this.waitForSpinnerGone(200);
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

      await this.humanDelay(500, 1000);

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
      await this.waitForSpinnerGone(1000);

      await remoteLogger.logSuccess(
        email,
        "🎉 Proses otomasi selesai dengan sukses! Semua langkah berhasil dijalankan.",
      );
      return { success: true };
    } catch (error) {
      let userMsg = "❌ Otomasi gagal — proses dihentikan";
      const errMsg = error.message || "";

      // Error Mapping Dictionary (Human Friendly)
      if (errMsg.includes("BUY_BUTTON_NOT_FOUND")) {
        userMsg =
          "❌ Step 17 Gagal: Tombol 'Beli' tidak ditemukan. Halaman Marketplace mungkin tidak memuat produk ini.";
      } else if (errMsg.includes("BUY_BUTTON_LOCKED")) {
        userMsg =
          "❌ Step 17 Gagal: Tombol 'Beli' terkunci (abu-abu). Cek kelengkapan data penagihan atau apakah produk masih tersedia.";
      } else if (errMsg.includes("PLACE_ORDER_FAILED")) {
        userMsg =
          "❌ Step 19 Gagal: Gagal saat menekan 'Buat Pesanan'. Microsoft mungkin menolak transaksi ini.";
      } else if (errMsg.includes("START_TRIAL_FAILED")) {
        userMsg =
          "❌ Step 23 Gagal: Gagal aktivasi trial Teams. Mungkin trial sudah pernah atau sedang aktif.";
      } else if (errMsg.includes("LOGIN_FAILED")) {
        userMsg =
          "❌ Step 5 Gagal: Gagal login ke Dashboard. Admin Center tidak dapat diakses.";
      } else if (errMsg.includes("EMAIL_TRANSITION_FAILED")) {
        userMsg =
          "❌ Step 3 Gagal: Gagal lanjut ke pengisian password. Sistem mentok di pengisian email.";
      } else if (errMsg.includes("MARKETPLACE_ERROR")) {
        userMsg = "❌ Step 12 Gagal: Produk tidak tersedia untuk akun ini.";
      } else if (errMsg.includes("BUY_CLICK_NOP")) {
        userMsg =
          "❌ Step 17 Gagal: Klik tombol 'Beli' tidak direspon oleh halaman Marketplace. Cek kondisi server Microsoft.";
      } else if (errMsg.includes("timeout") || errMsg.includes("waiting")) {
        userMsg =
          "❌ Koneksi Lambat: Proses berhenti karena waktu tunggu habis (Timeout).";
      }

      await remoteLogger.logError(
        this.accountConfig?.microsoftAccount?.email,
        userMsg,
        errMsg,
      );
      return { success: false, error: errMsg };
    }
  }

  async cleanup() {
    if (this.browser) {
      await this.browser.close();
    }
  }
}

module.exports = TeamsBot;
