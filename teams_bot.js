const { chromium } = require('playwright-core');
const config = require('./config');
const remoteLogger = require('./remote_logger');

const SPINNER_SELECTOR = '[data-testid="spinner"], .ms-Spinner, [class*="spinner" i]';
const HARD_TIMEOUT = 1.5 * 60 * 1000;

const SELECTORS = {
  searchInput: '[data-automation-id="UserListV2,CommandBarSearchInputBox"]',
  userRow:
    'div[data-automation-key="DisplayName"] span[role="button"], [role="gridcell"] button, [role="row"] button',
  licensesTab:
    'button[role="tab"]:has-text("Licenses and apps"), button:has-text("Licenses and apps"), button[role="tab"]:has-text("Lisensi dan aplikasi"), button:has-text("Lisensi dan aplikasi")',
  saveBtn:
    'button:has-text("Save changes"), button[id*="save" i], button:has-text("Simpan perubahan")',
};

class TeamsBot {
  constructor(wsUrl, accountConfig) {
    this.wsUrl = wsUrl;
    this.browser = null;
    this.context = null;
    this.page = null;
    this.accountConfig = accountConfig;
  }

  _escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s*');
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
      const errorSelectors = [
        '[data-automation-id="error-message"]',
        '[id*="error" i]',
        '.error',
        '#passwordError',
        '#usernameError',
      ];

      const markers = [
        'something went wrong',
        'something happened',
        'terjadi sesuatu',
        'Terjadi kesalahan',
        'Melindungi akun Anda',
        'try a different way',
        'Protecting your account',
        'Please solve the puzzle',
        'error code',
        '715-123280',
        'incorrect password',
        'password incorrect',
        'sandi salah',
        'salah sandi',
        'kata sandi salah',
        'password salah',
        'account or password is incorrect',
        'akun atau kata sandi anda salah',
        'is not recognized',
        'tidak dikenali',
        'tidak dapat menemukan akun',
        "couldn't find an account",
      ];

      for (const frame of this.page.frames()) {
        try {
          // 1. Cek selector error di tiap frame
          for (const selector of errorSelectors) {
            const el = frame.locator(selector).first();
            if (await el.isVisible({ timeout: 500 }).catch(() => false)) {
              const msg = (await el.innerText().catch(() => '')).trim();
              if (msg) return `Field Error: ${msg}`;
            }
          }

          // 2. Cek marker teks di tiap frame
          const frameText = await frame
            .locator('body')
            .innerText()
            .catch(() => '');
          const lowerText = frameText.toLowerCase();
          const found = markers.find((m) => lowerText.includes(m.toLowerCase()));
          if (found) {
            return `Marker "${found}" detected.`;
          }
        } catch (e) {}
      }
    } catch (err) {}
    return null;
  }

  async waitForSpinnerGone(minPostDelay = 0) {
    const spinner = this.page.locator(SPINNER_SELECTOR).first();
    const spinnerVisible = await spinner.isVisible().catch(() => false);

    if (spinnerVisible) {
      console.log('[WAIT] Spinner detected, waiting until hidden...');
      try {
        await spinner.waitFor({ state: 'hidden', timeout: HARD_TIMEOUT });
      } catch (e) {
        console.log('[WAIT] Spinner still visible or check failed, continuing...');
      }
      console.log('[WAIT] Spinner gone.');
    }

    if (minPostDelay > 0) {
      await this.humanDelay(minPostDelay, minPostDelay + 300);
    }
  }

  async waitForVisible(locator) {
    await this.waitForSpinnerGone();
    await this.runWithMonitor(locator.waitFor({ state: 'visible', timeout: HARD_TIMEOUT }));
  }

  async clickButtonWithPossibleNames(names) {
    await this.waitForSpinnerGone();

    // 1. Coba Playwright native dulu (lebih stabil)
    for (const name of names) {
      try {
        const btn = this.page.getByRole('button', {
          name: new RegExp(`^${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i'),
        });
        if (await btn.isVisible({ timeout: 3000 }).catch(() => false)) {
          await this.randomMouseMove();
          await btn.click({ timeout: 5000 });
          console.log(`[SUCCESS] Clicked: "${name}"`);
          return true;
        }
      } catch (e) {
        // lanjut ke nama berikutnya
      }
    }

    // 2. Fallback ke JavaScript hanya jika native gagal semua
    for (const frame of this.page.frames()) {
      try {
        const found = await frame.evaluate((kws) => {
          const escapeRegex = (s) =>
            s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s*');

          const candidates = Array.from(
            document.querySelectorAll('button, [role="button"]')
          ).filter((b) => b.offsetParent !== null);

          const el = candidates.find((b) => {
            const text = (b.textContent || b.value || b.getAttribute('aria-label') || '').trim();
            return (
              text.length > 0 &&
              text.length < 60 &&
              kws.some((kw) => new RegExp(`\\b${escapeRegex(kw)}\\b`, 'i').test(text))
            );
          });

          if (el) {
            el.click();
            return el.textContent?.trim() || 'button';
          }
          return null;
        }, names);

        if (found) {
          console.log(`[SUCCESS] Clicked via JS: "${found}"`);
          return true;
        }
      } catch (e) {}
    }

    console.error(`[ERROR] Button not found: ${names.join(', ')}`);
    throw new Error(`Button not found: ${names.join(', ')}`);
  }

  async handlePopups() {
    console.log('[INFO] Checking for any popups to dismiss...');
    const names = [
      'Close',
      'Dismiss',
      'Maybe later',
      'Got it',
      'No thanks',
      'Tutup',
      'Lain kali',
      'Selesai',
    ];
    const keywords = names.map((n) => n.trim().toLowerCase());

    const dismissed = new Set();
    let foundSomethingVisible = true;
    let attempts = 0;

    while (foundSomethingVisible && attempts < 3) {
      foundSomethingVisible = false;
      attempts++;

      for (const frame of this.page.frames()) {
        try {
          const foundName = await frame.evaluate((kws) => {
            const escapeRegex = (s) =>
              s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s*');

            const candidates = [
              ...document.querySelectorAll(
                'button, [role="button"], a[role="button"], input[type="button"]'
              ),
            ];
            const el = candidates.find((b) => {
              const textContent = (b.textContent || '').trim().toLowerCase();
              const ariaLabel = (b.getAttribute('aria-label') || '').trim().toLowerCase();
              const val = (b.value || '').trim().toLowerCase();
              const titleMsg = (b.getAttribute('title') || '').trim().toLowerCase();

              const isVisible = !!(b.offsetWidth || b.offsetHeight || b.getClientRects().length);
              if (!isVisible) return false;

              const isMatch = kws.some((kw) => {
                const escaped = escapeRegex(kw);
                const regex = new RegExp(`\\b${escaped}\\b`, 'i');
                return (
                  regex.test(textContent) ||
                  regex.test(ariaLabel) ||
                  regex.test(val) ||
                  regex.test(titleMsg)
                );
              });

              const btnLength = Math.max(
                textContent.length,
                ariaLabel.length,
                val.length,
                titleMsg.length
              );
              return isMatch && btnLength > 0 && btnLength < 35;
            });

            if (!el) return null;
            el.click();
            return (el.getAttribute('aria-label') || el.textContent || el.value || 'button').trim();
          }, keywords);

          if (foundName && !dismissed.has(foundName)) {
            dismissed.add(foundName);
            console.log(`[INFO] Dismissed popup button: "${foundName}" (Attempt ${attempts})`);
            await this.humanDelay(1000, 2000);
            foundSomethingVisible = true;
            break;
          }
        } catch (e) {}
      }
    }
  }

  getGenericLocator(keyword, elementType = 'input') {
    return this.page
      .locator(
        `${elementType}[id*="${keyword}" i], ${elementType}[data-testid*="${keyword}" i], ${elementType}[name*="${keyword}" i], ${elementType}[aria-label*="${keyword}" i]`
      )
      .first();
  }

  async connect() {
    if (this.wsUrl) {
      console.log('[STEP 1] Connecting to browser via Ads Power...');
      this.browser = await chromium.connectOverCDP(this.wsUrl);
      const contexts = this.browser.contexts();
      this.context = contexts.length > 0 ? contexts[0] : await this.browser.newContext();
    } else {
      console.log('[STEP 1] Launching local browser in incognito mode...');
      this.browser = await chromium.launch({
        headless:
          this.accountConfig?.headless !== undefined
            ? this.accountConfig.headless
            : config.headless,
        args: [
          '--incognito',
          '--disable-blink-features=AutomationControlled',
          '--disable-gpu',
          '--disable-dev-shm-usage',
          '--disable-software-rasterizer',
          '--mute-audio',
        ],
      });
      this.context = await this.browser.newContext();
    }
    const pages = this.context.pages();
    this.page = pages.length > 0 ? pages[0] : await this.context.newPage();
  }

  async _loginToAdminCenter(email, password) {
    await remoteLogger.logStep(email, 2, '🌐 Membuka halaman Microsoft Admin Center...');
    await this.page.goto('https://admin.microsoft.com/', {
      waitUntil: 'domcontentloaded',
      timeout: HARD_TIMEOUT,
    });
    await this.waitForSpinnerGone();

    await remoteLogger.logStep(email, 3, `📧 Memasukkan email: ${email}`);
    const emailInput = this.getGenericLocator('email');
    await this.waitForVisible(emailInput);
    await emailInput.fill(email);
    await this.humanDelay(500, 1000);
    await this.clickButtonWithPossibleNames(['Next', 'Selanjutnya', 'Berikutnya']);

    console.log('[STEP 3 VERIFY] Waiting for Password input or Choose method prompt...');
    const passwordOrPrompt = this.page.locator(
      'input[type="password"], div[role="button"]:has-text("Use my password"), div[role="button"]:has-text("Gunakan kata sandi saya")'
    );
    await passwordOrPrompt
      .first()
      .waitFor({ state: 'visible', timeout: 15000 })
      .catch(() => {
        throw new Error(
          'EMAIL_TRANSITION_FAILED: Gagal lanjut ke pengisian password. Cek apakah email sudah benar atau ada error di halaman.'
        );
      });

    console.log("[STEP 3.5] Checking for 'Choose a way to sign in' prompt...");
    const usePasswordPrompt = this.page
      .locator(
        'div[role="button"][aria-label*="Use my password" i], div[role="button"]:has-text("Use my password"), div[role="button"][aria-label*="Gunakan kata sandi saya" i], div[role="button"]:has-text("Gunakan kata sandi saya")'
      )
      .first();
    try {
      await usePasswordPrompt.waitFor({ state: 'visible', timeout: 5000 });
      console.log("[INFO] 'Choose a way to sign in' detected, clicking 'Use my password'...");
      await usePasswordPrompt.click();
      await this.humanDelay(400, 800);
    } catch (e) {
      console.log("[INFO] No 'Choose a way to sign in' prompt found, continuing...");
    }

    await remoteLogger.logStep(email, 4, '🔑 Memasukkan password akun...');
    const passwordInput = this.page.locator('input[type="password"]').first();
    await this.waitForVisible(passwordInput);
    await passwordInput.fill(password);
    await this.humanDelay(500, 1000);
    await this.clickButtonWithPossibleNames(['Sign in', 'Masuk']);

    await remoteLogger.logStep(
      email,
      5,
      '⏳ Menunggu dashboard atau konfirmasi login (KMSI/MFA)...'
    );

    const dashboardMarker = this.page
      .locator('[data-hint="ReactLeftNav"], #admin-home-container')
      .first();
    const loginLoopStart = Date.now();

    while (Date.now() - loginLoopStart < 120000) {
      if (await dashboardMarker.isVisible().catch(() => false)) {
        console.log('[SUCCESS] Dashboard detected!');
        await this.humanDelay(1000, 1500);
        await this.handlePopups();
        return;
      }

      const yesBtn = this.page
        .locator(
          'button:has-text("Yes"), input[value="Yes"], button:has-text("Ya"), input[value="Ya"], #idSIButton9'
        )
        .first();
      if (await yesBtn.isVisible().catch(() => false)) {
        console.log("[INFO] Handling 'Stay signed in'...");
        try {
          await yesBtn.click({ timeout: 5000 });
          await this.humanDelay(1000, 1500);
          continue;
        } catch (e) {
          console.log("[WARN] 'Yes' button blocked or failed.");
        }
      }

      const skipBtn = this.page
        .locator(
          'a:has-text("Skip for now"), a:has-text("Lompati untuk sekarang"), a:has-text("Lewati untuk sekarang"), button:has-text("Skip for now"), #idSecondaryButton'
        )
        .first();
      if (await skipBtn.isVisible().catch(() => false)) {
        console.log("[INFO] Handling MFA 'Skip for now'...");
        try {
          await skipBtn.click({ timeout: 5000 });
          await this.humanDelay(1000, 1500);
          continue;
        } catch (e) {
          console.log("[WARN] 'Skip for now' blocked.");
        }
      }

      const usePass = this.page
        .locator('text=Use my password, text=Gunakan kata sandi saya, #allowInterrupt')
        .first();
      if (await usePass.isVisible().catch(() => false)) {
        console.log("[INFO] Handling 'Use my password' prompt...");
        try {
          await usePass.click({ timeout: 5000 });
          await this.humanDelay(1000, 1500);
          continue;
        } catch (e) {
          console.log("[WARN] 'Use Password' prompt blocked.");
        }
      }

      const err = await this.checkForError();
      if (err) throw new Error(`LOGIN_FAILED: ${err}`);

      await this.handlePopups();
      await this.page.waitForTimeout(800);
    }

    if (!(await dashboardMarker.isVisible().catch(() => false))) {
      throw new Error('LOGIN_FAILED: Dashboard not reached within 2 minutes.');
    }
  }

  async _deactivateAllLicenses(email) {
    await remoteLogger.logStep(email, 6, "📂 Membuka halaman 'Pengguna Aktif' langsung via URL...");
    await this.page.goto('https://admin.cloud.microsoft/?#/users', {
      waitUntil: 'domcontentloaded',
      timeout: HARD_TIMEOUT,
    });
    await this.waitForSpinnerGone(200);
    await this.handlePopups();

    await remoteLogger.logStep(email, 8, `🔍 Mencari akun pengguna: ${email}...`);
    const searchInput = this.page.locator(SELECTORS.searchInput).first();
    await this.waitForVisible(searchInput);
    await searchInput.fill(email);
    await this.page.keyboard.press('Enter');
    await this.waitForSpinnerGone(2000);

    const userRow = this.page.locator(SELECTORS.userRow).first();
    await this.waitForVisible(userRow);
    const nameFound = await userRow.textContent();
    console.log(`[INFO] Clicking display name: "${nameFound?.trim()}" (Found for ${email})`);
    await userRow.click();
    await this.humanDelay(800, 1500);

    await remoteLogger.logStep(email, 9, "📋 Membuka tab 'Lisensi dan Aplikasi'...");
    const licensesTab = this.page.locator(SELECTORS.licensesTab).first();
    await this.waitForVisible(licensesTab);
    await licensesTab.click();
    await this.waitForSpinnerGone(500);

    await remoteLogger.logStep(email, 10, '🔲 Menonaktifkan semua lisensi yang sedang aktif...');
    try {
      await this.waitForSpinnerGone(200);
      await this.page
        .locator('input[type="checkbox"]')
        .first()
        .waitFor({ state: 'visible', timeout: 15000 })
        .catch(() => {});
      await this.page.waitForTimeout(800);

      for (let attempt = 1; attempt <= 3; attempt++) {
        await this.waitForSpinnerGone(200);

        let uncheckedSomething = true;
        while (uncheckedSomething) {
          uncheckedSomething = false;
          const checked = await this.page.locator('input[type="checkbox"]:checked').all();
          for (const cb of checked) {
            await cb.click({ force: true }).catch(() => {});
            uncheckedSomething = true;
            await this.page.waitForTimeout(300);
          }
        }

        await this.page.waitForTimeout(800);
        const remaining = await this.page.locator('input[type="checkbox"]:checked').count();

        if (remaining === 0) {
          await remoteLogger.logStep(
            email,
            10,
            `✅ Semua lisensi dinonaktifkan (Percobaan ${attempt}).`
          );
          break;
        } else {
          await remoteLogger.logStep(
            email,
            10,
            `⚠️ Percobaan ${attempt}: ${remaining} masih aktif. Retry...`
          );
          if (attempt === 3)
            throw new Error(
              `UNCHECK_ALL_FAILED: Still have ${remaining} checked after 3 attempts.`
            );
          await this.waitForSpinnerGone(500);
        }
      }
    } catch (err) {
      await remoteLogger.logError(email, '❌ Tidak dapat menonaktifkan semua lisensi', err.message);
      throw err;
    }

    await this.humanDelay(1000, 2000);

    await remoteLogger.logStep(email, 11, '💾 Menyimpan perubahan lisensi (nonaktifkan semua)...');
    const saveBtn = this.page.locator(SELECTORS.saveBtn).first();
    await this.waitForVisible(saveBtn);
    await saveBtn.click();
    await this.waitForSpinnerGone(500);
    console.log('[INFO] Waiting for license save to settle...');
    await this.humanDelay(8000, 12000);
  }

  async _purchaseProductTrial(email, productUrl) {
    let purchaseSuccess = false;

    for (let purchaseAttempt = 1; purchaseAttempt <= 2; purchaseAttempt++) {
      try {
        if (purchaseAttempt > 1) {
          await remoteLogger.logStep(email, 12, `🔄 Retry pembelian (${purchaseAttempt}/2)...`);
          await this.page.reload({ waitUntil: 'domcontentloaded' });
          await this.waitForSpinnerGone(2000);
          await this.humanDelay(3000, 5000);
        }

        const isTeamsRooms = productUrl.includes('microsoft-teams-rooms-basic');
        const isPhoneSystem = productUrl.includes('phone-system');
        const isCopilot = productUrl.includes('copilot');
        const isBusinessAppsFree = productUrl.includes('business-apps-free-');

        let planName = 'Microsoft 365 Copilot';
        if (isTeamsRooms) planName = 'Microsoft Teams Rooms Basic';
        else if (isPhoneSystem) planName = 'Microsoft 365 Phone System';
        else if (isBusinessAppsFree) planName = 'Business Apps (free)';

        if (purchaseAttempt === 1 || this.page.url() !== productUrl) {
          await remoteLogger.logStep(email, 12, `🛒 Membuka Marketplace: ${planName}...`);
          await this.page.goto(productUrl, {
            waitUntil: 'domcontentloaded',
            timeout: HARD_TIMEOUT,
          });
          await Promise.race([
            this.page.waitForLoadState('networkidle', { timeout: 20000 }),
            this.page.waitForTimeout(12000),
          ]).catch(() => {});
          await this.waitForSpinnerGone(500);
        }

        const unavailableMarker = this.page
          .locator(
            'text="This product is unavailable", text="Produk ini tidak tersedia", text="You are not eligible", text="Anda tidak memenuhi syarat"'
          )
          .first();
        if (await unavailableMarker.isVisible({ timeout: 5000 }).catch(() => false)) {
          const errorText = await unavailableMarker
            .innerText()
            .catch(() => 'Produk tidak tersedia');
          throw new Error(`MARKETPLACE_ERROR: ${errorText}`);
        }

        if (!isBusinessAppsFree) {
          console.log(`[STEP 15.5] Selecting plan: '${planName}'...`);
          const planDropdown = this.page
            .locator(
              'div:has-text("Select a plan") select, [aria-label*="Select a plan" i], div:has-text("Select a plan") [role="combobox"], div:has-text("Pilih paket") select, [aria-label*="Pilih paket" i], div:has-text("Pilih paket") [role="combobox"]'
            )
            .first();
          await this.waitForVisible(planDropdown);

          const uniquePlanWord = planName.split(' ').pop().toLowerCase();
          let planSelected = false;
          for (let planAttempt = 1; planAttempt <= 3; planAttempt++) {
            try {
              const tagName = await planDropdown.evaluate((el) => el.tagName.toLowerCase());
              if (tagName === 'select') {
                await planDropdown.selectOption({ label: planName });
              } else {
                await planDropdown.click();
                await this.humanDelay(800, 1200);
                const option = this.page.getByRole('option', {
                  name: new RegExp(`^${planName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i'),
                });
                await option.waitFor({ state: 'visible', timeout: 5000 });
                await option.click();
              }
              await this.page.waitForTimeout(800);
              const selectedText = await planDropdown
                .evaluate((el) => {
                  if (el.tagName === 'SELECT') return el.options[el.selectedIndex]?.text || '';
                  return el.textContent || el.getAttribute('aria-label') || '';
                })
                .catch(() => '');

              if (selectedText.toLowerCase().includes(uniquePlanWord)) {
                console.log(`[INFO] Plan verified: "${selectedText}"`);
                planSelected = true;
                break;
              }
              console.warn(
                `[WARN] Plan attempt ${planAttempt}: "${selectedText}" != "${planName}"`
              );
            } catch (err) {
              console.warn(`[WARN] Plan attempt ${planAttempt} failed: ${err.message}`);
            }
            await this.page.waitForTimeout(1000);
          }
          if (!planSelected) console.warn('[WARN] Plan unverified, proceeding...');
        }

        await remoteLogger.logStep(email, 16, '⚙️ Mengatur opsi commitment & billing...');
        let buyBtn = null;

        for (let retry = 1; retry <= 4; retry++) {
          console.log(`[STEP 16] Options attempt ${retry}/4...`);
          await this.waitForSpinnerGone(200);

          if (isCopilot) {
            const oneYear = this.page
              .locator(
                'label:has-text("1 year"), label:has-text("1 tahun"), :text-is("1 year"), :text-is("1 tahun")'
              )
              .first();
            if (await oneYear.isVisible({ timeout: 5000 }).catch(() => false)) {
              await oneYear.click({ force: true }).catch(() => {});
              await this.page.waitForTimeout(500);
            }
          }

          if (isBusinessAppsFree) {
            const oneMonth = this.page
              .locator(
                'label:has-text("1 month"), label:has-text("1 bulan"), :text-is("1 month"), :text-is("1 bulan")'
              )
              .first();
            if (await oneMonth.isVisible({ timeout: 5000 }).catch(() => false)) {
              await oneMonth.click({ force: true }).catch(() => {});
              await this.page.waitForTimeout(500);
            }
            const payMonthlyVisible = await this.page
              .locator('label:has-text("Pay monthly"), label:has-text("Bayar bulanan")')
              .first()
              .isVisible({ timeout: 3000 })
              .catch(() => false);
            console.log(
              `[INFO] Business Apps Free: pay monthly visible=${payMonthlyVisible} (auto-selected)`
            );
          } else {
            const payMonthly = this.page
              .locator(
                'label:has-text("Pay monthly"), label:has-text("Bayar bulanan"), :text-is("Pay monthly"), :text-is("Bayar bulanan")'
              )
              .first();
            if (await payMonthly.isVisible({ timeout: 5000 }).catch(() => false)) {
              await payMonthly.click({ force: true }).catch(() => {});
              await this.page.waitForTimeout(500);
            } else if (retry === 1) {
              console.log('[WARN] Pay monthly not visible yet, waiting...');
              await this.page.waitForTimeout(2000);
              continue;
            }
          }

          await this.page.waitForTimeout(1000);

          const buyBtnLocator = this.page
            .locator(
              `
              button:has-text("Buy"), button:has-text("Beli"),
              [role="button"]:has-text("Buy"), [role="button"]:has-text("Beli"),
              a:has-text("Buy"), a:has-text("Beli"),
              button:has-text("Get"), button:has-text("Dapatkan"),
              [role="button"]:has-text("Get"), [role="button"]:has-text("Dapatkan"),
              button:has-text("Checkout"), [role="button"]:has-text("Checkout"),
              button:has-text("Subscribe"), button:has-text("Berlangganan"),
              button:has-text("Try free"), button:has-text("Coba gratis")
            `
            )
            .first();

          const isBtnVisible = await buyBtnLocator.isVisible({ timeout: 3000 }).catch(() => false);
          if (isBtnVisible) {
            const isDisabled = await buyBtnLocator
              .evaluate(
                (btn) =>
                  btn.disabled ||
                  btn.getAttribute('aria-disabled') === 'true' ||
                  btn.classList.contains('is-disabled') ||
                  btn.getAttribute('disabled') !== null
              )
              .catch(() => true);

            if (!isDisabled) {
              console.log(`[SUCCESS] Buy button enabled at attempt ${retry}.`);
              buyBtn = buyBtnLocator;
              break;
            }
            console.warn(`[WARN] Buy button visible but disabled (attempt ${retry}).`);
          } else {
            console.warn(`[WARN] Buy button not visible (attempt ${retry}).`);
          }

          if (retry === 4)
            throw new Error('BUY_BUTTON_NOT_FOUND: Button tidak ditemukan setelah 4 attempts.');
          await this.page.waitForTimeout(1500);
        }

        await remoteLogger.logStep(email, 17, "🛍️ Mengklik tombol 'Beli'...");
        await this.waitForVisible(buyBtn);

        let clickedSuccessfully = false;
        for (let attempt = 1; attempt <= 3; attempt++) {
          const oldUrl = this.page.url();
          await buyBtn.click({ timeout: 10000, force: true }).catch(() => {});
          await this.page.waitForTimeout(1500);

          const newUrl = this.page.url();
          const isBtnStillVisible = await buyBtn.isVisible().catch(() => false);
          const isNextStepVisible = await this.page
            .locator(
              '.ms-Checkbox:has-text("authorize recurring payments"), .ms-Checkbox:has-text("pembayaran berulang"), button:has-text("Place order"), button:has-text("Buat pesanan"), button:has-text("Tempatkan pesanan")'
            )
            .first()
            .isVisible()
            .catch(() => false);

          if (newUrl !== oldUrl || !isBtnStillVisible || isNextStepVisible) {
            console.log('[SUCCESS] Buy click triggered transition.');
            clickedSuccessfully = true;
            break;
          }
          console.warn(`[WARN] Buy click NOP, retrying (${attempt}/3)...`);
          await this.page.waitForTimeout(1000);
        }
        if (!clickedSuccessfully) throw new Error('BUY_CLICK_NOP');
        await this.waitForSpinnerGone(15000);

        console.log('[STEP 18] Checking for authorization checkboxes...');
        try {
          const checkboxContainer = this.page
            .locator(
              '.ms-Checkbox:has-text("authorize recurring payments"), .ms-Checkbox:has-text("pembayaran berulang")'
            )
            .first();
          const isVisible = await checkboxContainer
            .isVisible({ timeout: 10000 })
            .catch(() => false);

          if (isVisible) {
            console.log('[INFO] Authorization checkbox detected. Attempting to check...');
            const checkboxInput = checkboxContainer.locator('input[type="checkbox"]').first();

            for (let attempt = 1; attempt <= 3; attempt++) {
              const isChecked = await checkboxInput.isChecked().catch(() => false);
              if (isChecked) {
                console.log('[INFO] Checkbox is already checked.');
                break;
              }
              console.log(`[INFO] Checking attempt ${attempt}...`);
              const visualCheckbox = checkboxContainer.locator('.ms-Checkbox-checkbox').first();
              await visualCheckbox.click({ force: true });
              await this.page.waitForTimeout(800);

              const nowChecked = await checkboxInput.isChecked().catch(() => false);
              if (nowChecked) {
                console.log('[SUCCESS] Checkbox is now checked.');
                break;
              }
              if (attempt === 2) {
                console.log('[INFO] Fallback: clicking input directly via JS...');
                await checkboxInput.evaluate((el) => el.click());
                await this.page.waitForTimeout(800);
              }
              if (attempt === 3) throw new Error('Checkbox authorization failed after 3 attempts.');
            }
          } else {
            console.log('[INFO] No authorization checkbox found, skipping...');
          }
        } catch (err) {
          throw new Error(`CRITICAL_ERROR di Step 18: ${err.message}`);
        }

        await remoteLogger.logStep(email, 19, "📦 Mengklik tombol 'Buat Pesanan'...");
        const placeOrderBtn = this.page
          .locator(
            'button:has-text("Place order"), button:has-text("Buat pesanan"), button:has-text("Tempatkan pesanan")'
          )
          .first();

        console.log("[INFO] Waiting for 'Place order' button to become enabled...");
        try {
          await this.page.waitForFunction(
            (btn) =>
              btn &&
              !btn.disabled &&
              btn.getAttribute('aria-disabled') !== 'true' &&
              !btn.classList.contains('is-disabled'),
            await placeOrderBtn.elementHandle(),
            { timeout: 30000 }
          );
          console.log("[INFO] 'Place order' is now enabled.");
        } catch (e) {
          throw new Error(
            'PLACE_ORDER_DISABLED: Tombol tidak aktif dalam 30 detik. Kemungkinan otorisasi gagal.'
          );
        }

        await placeOrderBtn
          .click({ timeout: 10000 })
          .catch(() => placeOrderBtn.click({ force: true }));
        console.log("[INFO] 'Place order' clicked. Waiting for confirmation...");

        await remoteLogger.logStep(email, 20, '🔎 Memverifikasi keberhasilan pesanan...');
        const verifyStart = Date.now();
        while (Date.now() - verifyStart < 60000) {
          const isBtnHidden = await placeOrderBtn.isHidden().catch(() => true);

          if (isBtnHidden) {
            const currentUrl = this.page.url().toLowerCase();
            const bodyContent = await this.page.innerText('body').catch(() => '');
            const successKeywords = [
              'all set',
              'confirmation',
              'thanks',
              'terima kasih',
              'detail pesanan',
              'order details',
            ];
            const foundKeyword = successKeywords.find((kw) =>
              bodyContent.toLowerCase().includes(kw)
            );

            if (currentUrl.includes('confirmation') || foundKeyword) {
              console.log(
                `[SUCCESS] Order verified! (Keyword: "${foundKeyword || 'URL confirmation'}")`
              );
              purchaseSuccess = true;
              break;
            }
          }

          const detectedError = await this.checkForError();
          if (detectedError) throw new Error(`PLACE_ORDER_FAILED: ${detectedError}`);
          await this.page.waitForTimeout(800);
        }

        if (!purchaseSuccess) {
          console.warn(
            '[WARN] Order confirmation not clearly detected, but button is gone. Proceeding...'
          );
        }

        await this.waitForSpinnerGone(1000);
        break; // EXIT RETRY LOOP IF SUCCESS
      } catch (err) {
        if (
          purchaseAttempt < 2 &&
          (err.message.includes('something happened') ||
            err.message.includes('Terjadi kesalahan') ||
            err.message.includes('BUY_BUTTON_NOT_FOUND'))
        ) {
          console.warn(`[RETRY] Purchase failed. Attempt ${purchaseAttempt}/2. Reloading...`);
          await this.page.reload({ waitUntil: 'domcontentloaded' });
          await this.page.waitForTimeout(5000);
          continue;
        }
        throw err;
      }
    }
  }

  async _activateTeamsTrial(email) {
    await remoteLogger.logStep(
      email,
      21,
      '🚀 Membuka Microsoft Teams di tab baru untuk aktivasi trial...'
    );
    const teamsPage = await this.context.newPage();
    try {
      try {
        await teamsPage.goto('https://teams.microsoft.com/v2/', {
          waitUntil: 'domcontentloaded',
          timeout: HARD_TIMEOUT,
        });
        const bodyText = await teamsPage.innerText('body').catch(() => '');
        if (!bodyText || bodyText.trim().length === 0) {
          console.log('[INFO] Teams page is blank, refreshing...');
          await teamsPage.reload({ waitUntil: 'domcontentloaded' });
        }
      } catch (e) {
        console.warn('[WARN] Initial Teams navigation failed, retrying...');
        await teamsPage.goto('https://teams.microsoft.com/v2/', {
          waitUntil: 'domcontentloaded',
        });
      }

      await remoteLogger.logStep(email, 22, '⏳ Menunggu Teams siap (Sign in atau Start Trial)...');
      await teamsPage.waitForTimeout(2000);

      const teamsSignInBtn = teamsPage
        .locator(
          'button:has-text("Sign in"), a:has-text("Sign in"), button:has-text("Masuk"), a:has-text("Masuk")'
        )
        .first();
      const startTrialBtn = teamsPage
        .locator(
          'button:has-text("Start trial"), button:has-text("Mulai uji coba"), [role="button"]:has-text("Start trial"), button:has-text("Get started"), button:has-text("Mulai"), button:has-text("Try now"), a:has-text("Get started")'
        )
        .first();
      const pickAccountHeader = teamsPage
        .locator(
          'div:has-text("Pick an account"), div:has-text("Pilih akun"), h1:has-text("Pick an account"), h1:has-text("Pilih akun")'
        )
        .first();
      const permissionErrorLocator = teamsPage
        .getByText(
          /You don't have the required permissions to access this org|Anda tidak memiliki izin yang diperlukan untuk mengakses organisasi ini/i
        )
        .first();

      console.log('[INFO] Waiting for Sign In, Start Trial, Pick Account, or Permission (60s)...');
      await teamsSignInBtn
        .or(startTrialBtn)
        .or(pickAccountHeader)
        .or(permissionErrorLocator)
        .waitFor({ state: 'visible', timeout: 60000 });

      if (await permissionErrorLocator.isVisible().catch(() => false)) {
        console.error('[ERROR] Permission error page detected.');
        throw new Error(
          "PERMISSION_ERROR: Don't have the required permissions to access this org."
        );
      }

      if (await pickAccountHeader.isVisible().catch(() => false)) {
        console.log("[INFO] 'Pick an account' detected, selecting current user...");
        const targetAccountItem = teamsPage
          .locator(`div[role="listitem"]:has-text("${email}"), [data-test-id="${email}"]`)
          .first();
        if (await targetAccountItem.isVisible().catch(() => false)) {
          await targetAccountItem.click();
        } else {
          await teamsPage
            .locator('div[role="listitem"], .tile-container')
            .first()
            .click()
            .catch(() => {});
        }
        await teamsPage.waitForTimeout(2000);
      }

      if (await teamsSignInBtn.isVisible().catch(() => false)) {
        console.log("[INFO] 'Sign in' button detected, clicking...");
        await teamsSignInBtn.click();
        await teamsPage.waitForTimeout(1500);
        console.log("[INFO] Waiting for 'Start Trial' button after Sign in...");
        await startTrialBtn.waitFor({ state: 'visible', timeout: 45000 });
      }

      await remoteLogger.logStep(
        email,
        23,
        "▶️ Mengklik tombol 'Mulai Uji Coba' di Microsoft Teams..."
      );
      await startTrialBtn.scrollIntoViewIfNeeded().catch(() => {});
      await startTrialBtn.click();

      await remoteLogger.logStep(email, 23.5, '⏳ Menunggu proses aktivasi uji coba selesai...');
      const teamsSpinner = teamsPage.locator(SPINNER_SELECTOR).first();
      await teamsSpinner.waitFor({ state: 'visible', timeout: 5000 }).catch(() => {});
      if (await teamsSpinner.isVisible().catch(() => false)) {
        await teamsSpinner.waitFor({ state: 'hidden', timeout: 60000 }).catch(() => {
          console.log('[WARN] Teams trial spinner still visible, continuing anyway.');
        });
      }

      await remoteLogger.logStep(email, 23.6, '✅ Aktivasi uji coba selesai. Menutup tab Teams...');
    } catch (err) {
      // Jika error permission, lempar tanpa wrap agar bisa dibedakan
      if (err.message.startsWith('PERMISSION_ERROR')) throw err;
      // Semua error lain di-wrap START_TRIAL_FAILED agar error mapping di run() bisa menangkap
      throw new Error(`START_TRIAL_FAILED: ${err.message || 'Gagal aktivasi trial Teams.'}`);
    } finally {
      await teamsPage.close().catch(() => {});
    }
  }

  async _restorePrimaryLicense(email) {
    await remoteLogger.logStep(
      email,
      24,
      '↩️ Kembali ke Admin Center untuk memulihkan lisensi pengguna...'
    );
    await this.page.bringToFront();
    await this.page.goto('https://admin.cloud.microsoft/?#/users', {
      waitUntil: 'domcontentloaded',
      timeout: HARD_TIMEOUT,
    });
    await this.waitForSpinnerGone(800);

    await remoteLogger.logStep(email, 25, `🔍 Mencari ulang pengguna: ${email}...`);
    const searchInput = this.page.locator(SELECTORS.searchInput).first();
    await this.waitForVisible(searchInput);
    await searchInput.fill(email);
    await this.page.keyboard.press('Enter');
    await this.waitForSpinnerGone(2000);

    const userRow = this.page.locator(SELECTORS.userRow).first();
    await this.waitForVisible(userRow);
    await userRow.click();
    await this.humanDelay(800, 1500);

    await remoteLogger.logStep(email, 26, "📋 Membuka tab 'Lisensi dan Aplikasi'...");
    const tab = this.page.locator(SELECTORS.licensesTab).first();
    await this.waitForVisible(tab);
    await tab.click();
    await this.waitForSpinnerGone(500);

    const licenseNames = [
      'Microsoft 365 Business Standard',
      'Microsoft 365 Business Basic',
      'Microsoft 365 Business Premium',
      'Microsoft 365 E3',
      'Microsoft 365 E5',
      'Office 365 E1',
      'Office 365 E3',
      'Office 365 E5',
    ];

    await remoteLogger.logStep(email, 27, '🔍 Mencari lisensi yang dikenal untuk dipulihkan...');
    try {
      await this.waitForSpinnerGone(200);
      await this.page
        .locator('input[type="checkbox"]')
        .first()
        .waitFor({ state: 'visible', timeout: 15000 })
        .catch(() => {});

      const allLicenseTexts = await this.page.locator('[data-automation-id^="LicenseText_"]').all();
      for (const el of allLicenseTexts) {
        const text = await el.innerText().catch(() => 'N/A');
        const automationId = await el.getAttribute('data-automation-id').catch(() => 'N/A');
        await remoteLogger.logStep(
          email,
          27,
          `🔎 Lisensi tersedia di halaman: "${text}" (${automationId})`
        );
      }

      let targetCheckbox = null;
      let foundLicenseName = null;
      for (const name of licenseNames) {
        const textEl = this.page.locator(`[data-automation-id="LicenseText_${name}"]`).first();
        const isVisible = await textEl.isVisible({ timeout: 3000 }).catch(() => false);
        if (!isVisible) continue;

        const checkbox = textEl
          .locator('xpath=ancestor::div[contains(@class,"ms-Checkbox")][1]')
          .locator('input[type="checkbox"]');
        const isCheckboxVisible = await checkbox.isVisible().catch(() => false);
        if (isCheckboxVisible) {
          targetCheckbox = checkbox;
          foundLicenseName = name;
          await remoteLogger.logStep(
            email,
            27,
            `✅ Lisensi ditemukan: '${name}' — akan diaktifkan kembali.`
          );
          break;
        }
      }

      if (!targetCheckbox)
        throw new Error(
          `LICENSE_NOT_FOUND: None of the known licenses found. Checked: ${licenseNames.join(', ')}`
        );

      for (let attempt = 1; attempt <= 3; attempt++) {
        await this.waitForSpinnerGone(500);
        if (!(await targetCheckbox.isChecked().catch(() => false))) {
          await remoteLogger.logStep(
            email,
            27,
            `🖱️ Percobaan ${attempt}: Mengaktifkan '${foundLicenseName}'...`
          );
          await targetCheckbox.click({ force: true });
          await this.page.waitForTimeout(800);
        }

        if (await targetCheckbox.isChecked().catch(() => false)) {
          await remoteLogger.logStep(email, 27, `✅ '${foundLicenseName}' berhasil diaktifkan.`);
          break;
        }

        await remoteLogger.logStep(
          email,
          27,
          `⚠️ Percobaan ${attempt}: Masih belum tercentang. Mencoba ulang...`
        );
        if (attempt === 3)
          throw new Error(`STRICT_CHECKBOX_FAILED: Failed to check '${foundLicenseName}'.`);
      }
    } catch (err) {
      await remoteLogger.logError(email, '❌ Tidak dapat memulihkan lisensi', err.message);
      throw err;
    }

    await this.humanDelay(500, 1000);
    await remoteLogger.logStep(
      email,
      28,
      '💾 Menyimpan perubahan lisensi yang telah dipulihkan...'
    );
    const saveBtn = this.page.locator(SELECTORS.saveBtn).first();
    await this.waitForVisible(saveBtn);
    await saveBtn.click();
    await this.waitForSpinnerGone(1000);
  }

  async run() {
    const email = this.accountConfig.microsoftAccount.email;
    const password = this.accountConfig.microsoftAccount.password;
    const productUrl = this.accountConfig.productUrl;

    try {
      await this.connect();

      await this._loginToAdminCenter(email, password);
      await this._deactivateAllLicenses(email);
      await this._purchaseProductTrial(email, productUrl);
      await this._activateTeamsTrial(email);
      await this._restorePrimaryLicense(email);

      await remoteLogger.logSuccess(
        email,
        '🎉 Proses otomasi selesai dengan sukses! Semua langkah berhasil dijalankan.'
      );
      return { success: true };
    } catch (error) {
      let userMsg = '❌ Otomasi gagal — proses dihentikan';
      const errMsg = error.message || '';

      if (errMsg.includes('BUY_BUTTON_NOT_FOUND')) {
        userMsg =
          "❌ Step 17 Gagal: Tombol 'Beli' tidak ditemukan. Halaman Marketplace mungkin tidak memuat produk ini.";
      } else if (errMsg.includes('BUY_BUTTON_LOCKED')) {
        userMsg =
          "❌ Step 17 Gagal: Tombol 'Beli' terkunci (abu-abu). Cek kelengkapan data penagihan atau apakah produk masih tersedia.";
      } else if (errMsg.includes('PLACE_ORDER_FAILED') || errMsg.includes('PLACE_ORDER_DISABLED')) {
        userMsg =
          "❌ Step 19 Gagal: Gagal saat menekan 'Buat Pesanan'. Microsoft mungkin menolak transaksi ini.";
      } else if (errMsg.includes('START_TRIAL_FAILED')) {
        userMsg =
          '❌ Step 23 Gagal: Gagal aktivasi trial Teams. Mungkin trial sudah pernah atau sedang aktif.';
      } else if (errMsg.includes('PERMISSION_ERROR')) {
        userMsg =
          '❌ Step 22 Gagal: Akun tidak memiliki izin untuk mengakses organisasi Teams ini.';
      } else if (errMsg.includes('LOGIN_FAILED')) {
        if (errMsg.toLowerCase().includes('password') || errMsg.toLowerCase().includes('sandi')) {
          userMsg = '❌ Login Gagal: Kata sandi salah. Silakan cek kembali password akun.';
        } else if (
          errMsg.toLowerCase().includes('recognized') ||
          errMsg.toLowerCase().includes('dikenali')
        ) {
          userMsg = '❌ Login Gagal: Akun tidak dikenali atau email salah.';
        } else {
          userMsg = '❌ Step 5 Gagal: Gagal login ke Dashboard. Admin Center tidak dapat diakses.';
        }
      } else if (errMsg.includes('EMAIL_TRANSITION_FAILED')) {
        userMsg =
          '❌ Step 3 Gagal: Gagal lanjut ke pengisian password. Sistem mentok di pengisian email.';
      } else if (errMsg.includes('MARKETPLACE_ERROR')) {
        userMsg = '❌ Step 12 Gagal: Produk tidak tersedia untuk akun ini.';
      } else if (errMsg.includes('BUY_CLICK_NOP')) {
        userMsg = "❌ Step 17 Gagal: Klik tombol 'Beli' tidak direspon oleh halaman Marketplace.";
      } else if (errMsg.includes('UNCHECK_ALL_FAILED')) {
        userMsg = '❌ Step 10 Gagal: Tidak dapat menonaktifkan semua lisensi setelah 3 percobaan.';
      } else if (errMsg.includes('LICENSE_NOT_FOUND')) {
        userMsg = '❌ Step 27 Gagal: Tidak ada lisensi yang dikenal ditemukan untuk dipulihkan.';
      } else if (errMsg.includes('STRICT_CHECKBOX_FAILED')) {
        userMsg = '❌ Step 27 Gagal: Gagal mencentang lisensi untuk dipulihkan.';
      } else if (errMsg.includes('timeout') || errMsg.includes('waiting')) {
        userMsg = '❌ Koneksi Lambat: Proses berhenti karena waktu tunggu habis (Timeout).';
      }

      await remoteLogger.logError(email, userMsg, errMsg);
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
